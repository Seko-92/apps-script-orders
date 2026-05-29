// =======================================================================================
// N8N_INTEGRATION.gs - v3.0 with Timestamp Support
// =======================================================================================

// Note: N8N_WEBHOOK_URL is now defined in Secrets.js
// Make sure Secrets.js is uploaded to your Apps Script project 

/**
 * Triggers the n8n Awaiting Shipments workflow via webhook
 * Called from the Sidebar when user clicks "Sync Orders"
 * NOW UPDATES TIMESTAMP IN CELL F2!
 */
function triggerN8NWebhook() {
  if (!N8N_WEBHOOK_URL) {
    return "⚠️ Webhook URL not configured.";
  }
  
  try {
    var options = {
      'method': 'get',
      'muteHttpExceptions': true,
      'followRedirects': true,
      'timeout': 30000,
      'headers': {
        'ngrok-skip-browser-warning': 'true',
        'User-Agent': 'GoogleAppsScript'
      }
    };
    
    var response = UrlFetchApp.fetch(N8N_WEBHOOK_URL, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    if (responseCode === 200) {
      // 1. Update stats and timestamp
      try { updateOrderStatsInSheet(); } catch(e) {}  // no-op since 2026-05-17
      try { updateLastSyncTimestamp(); } catch(e) {}  // no-op since 2026-05-17

      // 2. REMOVED 2026-05-19 — the legacy updateLastOrderTimestamp("F2") call
      //    used to stamp "📦 Last Order: M/d/yyyy h:mm AM/PM" into F2. That
      //    worked silently in the old layout (F2 was inside the A2:F2 logo
      //    merge — hidden, no validation). After today's manual layout
      //    compaction, F2 is the Pick ID for Shipping dropdown anchor WITH
      //    data validation — writing a timestamp string violates the rule
      //    and Sheets throws. The legacy feature is also redundant: the
      //    Service Bay v6 banner already has live System Pulse in E1 reading
      //    Last sync + ALIVE/IDLE/STALE state from the Activity Log.

      // 3. Parse response
      try {
        var data = JSON.parse(responseText);
        if (data.message) return "✅ " + data.message;
        if (data.added) return "✅ Synced! " + data.added + " orders added.";
      } catch(e) {}
      
      return "✅ Sync complete!";
      
    } else if (responseCode === 404) {
      return "❌ 404: Webhook not found. Is n8n workflow ACTIVE?";
    } else if (responseCode === 502) {
      return "❌ 502: ngrok cannot reach n8n. Is n8n running?";
    } else {
      return "⚠️ Error: n8n responded with code " + responseCode;
    }
    
  } catch (error) {
    Logger.log("n8n webhook error: " + error.toString());
    return "❌ Connection error: " + error.message;
  }
}

/**
 * Tests the n8n webhook connection
 */
function testN8NConnection() {
  try {
    var response = UrlFetchApp.fetch(N8N_WEBHOOK_URL, {
      'method': 'get',
      'muteHttpExceptions': true,
      'headers': {
        'ngrok-skip-browser-warning': 'true'
      }
    });
    
    var code = response.getResponseCode();
    if (code === 200) return "✅ Connection successful!";
    return "⚠️ Connection failed. Code: " + code;
  } catch (e) {
    return "❌ Connection failed: " + e.message;
  }
}

function getN8NStatus() {
  return {
    configured: !!N8N_WEBHOOK_URL,
    url: N8N_WEBHOOK_URL ? "(configured)" : "(not set)"
  };
}


// =======================================================================================
// MANUAL-TRIGGER WEBHOOKS (added 2026-05-06)
// =======================================================================================
// These two functions fire the new n8n manual-trigger webhooks added on 2026-05-06.
// Both are guarded with Header Auth (X-API-Token = APP_SECRET_TOKEN) so the
// endpoints reject unauthenticated requests if exposed via ngrok.
//
// Pattern: identical to triggerN8NWebhook above, with two differences —
//   1. POST instead of GET (matches the new webhooks' configured method)
//   2. X-API-Token header on every request (Header Auth on n8n side)
//
// Return shape: { ok: boolean, message: string, code: number }
// The sidebar uses this shape to drive status-bar feedback.
// =======================================================================================

/**
 * Triggers the API Usage Monitor workflow on demand.
 * Called from the sidebar's "Refresh Now" button on the API Status card.
 *
 * After this fires, the sidebar should wait ~2-3 seconds then re-poll
 * getLatestApiMetrics() — that gives n8n time to fetch fresh quota data
 * from eBay and write the API Usage sheet.
 */
function triggerApiUsageRefresh() {
  return _fireAuthedWebhook(N8N_API_USAGE_WEBHOOK_URL, "API quotas refreshed");
}

/**
 * Triggers the Shipped Status Check sub-workflow on demand.
 * Called from the sidebar's "Check Shipped Status" button.
 *
 * After this fires, the workflow walks all PREPARING/PENDING orders and
 * checks eBay for status changes (mostly: PREPARING → SHIPPED when tracking
 * lands). Any changes flow back into the All Orders sheet via doPost.
 */
function triggerStatusCheck() {
  return _fireAuthedWebhook(N8N_STATUS_CHECK_WEBHOOK_URL, "Status check kicked off");
}

/**
 * Triggers the Inventory Lite Sync workflow on demand.
 * Called from the sidebar's "Sync Inventory Now" button.
 *
 * After this fires, the workflow downloads the LMS active inventory report
 * (REST sell.feed — free quota), diffs against MI in memory, and upserts
 * only changed rows. Then n8n calls back here with action=recomputeHand
 * so HAND values in All Orders reflect the new inventory within ~1s.
 *
 * Normal operation: schedule fires every 10 min automatically. This button
 * is for "I just made a change on eBay, sync now" — instant verification.
 */
function triggerInventoryLiteSync() {
  return _fireAuthedWebhook(N8N_INVENTORY_LITE_WEBHOOK_URL, "Inventory sync kicked off");
}

/**
 * Triggers the SHIPPED Verification workflow on demand.
 * Called from the sidebar's "Verify SHIPPED" button.
 *
 * Reverse safety net: walks every SHIPPED row in the sheet, asks eBay's
 * Fulfillment API whether it's *actually* fulfilled, and reverts to PENDING
 * any row that the sheet says is SHIPPED but eBay says is not. Posts a
 * Telegram alert if anything was reverted. Source on the revert event in
 * Activity Log is `n8n-verify` — non-zero counts are the canary.
 *
 * Heavier than the forward status check (queries every SHIPPED order, not
 * just the open ones), so the sidebar should use a longer cooldown
 * (~60s recommended) to avoid burning sell.fulfillment quota during testing.
 */
function triggerVerifyShipped() {
  return _fireAuthedWebhook(N8N_VERIFY_SHIPPED_WEBHOOK_URL, "Verification sweep kicked off");
}

/**
 * Triggers the Zoho SO Backfill workflow on demand and waits for the result.
 * Called from the sidebar's "Fetch from Zoho" button when a Preview returns
 * "not in Pending" — picker types SO# or INV#, this fetches the full SO from
 * Zoho via n8n, n8n forwards it to doPost which upserts the Pending row, and
 * we return the result to the sidebar.
 *
 * DIFFERENT from _fireAuthedWebhook: that helper sends an empty body and
 * doesn't consume the response body. This one sends `{query: "SO-X" | "INV-X"}`
 * and returns the parsed JSON response (so the sidebar can show "found SO-22815
 * via invoice INV-022496 · 3 line items").
 *
 * @param {string} query — SO# (e.g., "SO-22815") or invoice # (e.g., "INV-022496")
 * @returns {{ok: boolean, message: string, data?: object, code?: number}}
 */
function triggerZohoBackfill(query) {
  var url = N8N_ZOHO_FETCH_WEBHOOK_URL;
  if (!url) {
    return { ok: false, message: "Backfill webhook URL not configured (check Secrets.js)" };
  }
  if (!query || !String(query).trim()) {
    return { ok: false, message: "Query is empty — type an SO# or INV# first" };
  }

  try {
    var options = {
      method:              'post',
      muteHttpExceptions:  true,
      followRedirects:     true,
      // Backfill chain has latency: get-token + lookup + fetch-full-SO +
      // forward-to-AppsScript. Bump timeout to allow for that.
      // (Apps Script /exec is the slowest leg at ~1-2s.)
      // Use the longest realistic allowance to avoid premature failure on
      // first-call cold paths.
      // Safe to wait since this is a foreground sidebar action with spinner.
      // NOTE: UrlFetchApp's max timeout is bounded at the platform level; we
      // request what we want and the platform clamps if needed.
      headers: {
        'X-API-Token':                APP_SECRET_TOKEN,
        'ngrok-skip-browser-warning': 'true',
        'User-Agent':                 'GoogleAppsScript'
      },
      contentType: 'application/json',
      payload:     JSON.stringify({ query: String(query).trim() })
    };

    var response = UrlFetchApp.fetch(url, options);
    var code     = response.getResponseCode();
    var body     = response.getContentText();

    if (code !== 200 && code !== 204) {
      if (code === 401 || code === 403) {
        return { ok: false, code: code, message: "Auth rejected by n8n (X-API-Token mismatch)" };
      }
      if (code === 404) {
        return { ok: false, code: code, message: "Workflow not active in n8n — toggle it on" };
      }
      if (code === 500) {
        // n8n returns 500 on workflow execution errors (e.g., 'Not found in Zoho')
        // Try to parse the error message from the response body
        var errMsg = body;
        try {
          var parsed = JSON.parse(body);
          errMsg = parsed.message || parsed.error || body;
        } catch (_) {}
        return { ok: false, code: code, message: errMsg };
      }
      return { ok: false, code: code, message: "n8n returned HTTP " + code + ": " + body };
    }

    // Parse the JSON response from the workflow's "Respond to Caller" node
    var parsed;
    try {
      parsed = JSON.parse(body);
    } catch (parseErr) {
      return { ok: false, code: code, message: "n8n returned malformed JSON: " + body.substring(0, 200) };
    }

    return {
      ok:      true,
      code:    code,
      message: "Fetched from Zoho",
      data:    parsed
    };
  } catch (err) {
    return { ok: false, message: "Backfill request failed: " + (err.message || err) };
  }
}

/**
 * Bulk-fetch every Zoho item with selling_price. Calls the Zoho Items Bulk
 * Fetch Proxy n8n workflow, which paginates `GET /items` and aggregates the
 * full active-items catalog. Used by the Price Audit feature (PriceAudit.js)
 * to compare Zoho's stored selling_price against MI's currentPrice.
 *
 * Returns: {
 *   ok: boolean,
 *   message: string,
 *   data: {                                — parsed n8n response body
 *     items: [{ sku, item_id, item_name, selling_price, status, reference_id }],
 *     totalFetched: number,
 *     totalPages:   number
 *   },
 *   code: number
 * }
 *
 * Latency: ~10-15s for ~3,500 items (18 pages × 500ms + OAuth). Foreground
 * call with sidebar spinner — picker waits, no async.
 */
function triggerZohoBulkItemsFetch() {
  var url = N8N_ZOHO_BULK_ITEMS_WEBHOOK_URL;
  if (!url) {
    return { ok: false, message: "Bulk items webhook URL not configured (check Secrets.js)" };
  }

  try {
    var options = {
      method:              'post',
      muteHttpExceptions:  true,
      followRedirects:     true,
      headers: {
        'X-API-Token':                APP_SECRET_TOKEN,
        'ngrok-skip-browser-warning': 'true',
        'User-Agent':                 'GoogleAppsScript'
      },
      contentType: 'application/json',
      payload:     JSON.stringify({})
    };

    var response = UrlFetchApp.fetch(url, options);
    var code     = response.getResponseCode();
    var body     = response.getContentText();

    if (code !== 200 && code !== 204) {
      if (code === 401 || code === 403) {
        return { ok: false, code: code, message: "Auth rejected by n8n (X-API-Token mismatch)" };
      }
      if (code === 404) {
        return { ok: false, code: code, message: "Workflow not active in n8n — toggle it on" };
      }
      return { ok: false, code: code, message: "n8n returned HTTP " + code + ": " + body.substring(0, 200) };
    }

    var parsed;
    try { parsed = JSON.parse(body); }
    catch (parseErr) {
      return { ok: false, code: code, message: "n8n returned malformed JSON: " + body.substring(0, 200) };
    }

    return {
      ok:      true,
      code:    code,
      message: "Fetched " + ((parsed.items && parsed.items.length) || 0) + " items from Zoho",
      data:    parsed
    };
  } catch (err) {
    return { ok: false, message: "Bulk fetch failed: " + (err.message || err) };
  }
}


/**
 * Write ONE item's selling price (rate) back to Zoho via the Zoho Item Price
 * Write Proxy. This is the only path in the codebase that WRITES to Zoho —
 * every other Zoho integration is read-only. See PriceWriteback.js for the
 * SKU-resolving caller and Zoho Item Price Write Proxy_v1.json for the n8n side.
 *
 * The proxy validates + sanity-gates the price, GETs the item first (captures
 * the before-rate), PUTs only the `rate` field, and returns before/after — so
 * we get a real confirmation of what Zoho actually stored, not just "200 OK".
 *
 * @param {string} itemId  Zoho internal item_id (from the Zoho Stock sheet)
 * @param {string} sku     SKU — passed through for the response message only
 * @param {number} price   new selling price (eBay's authoritative price)
 * @returns {{ok: boolean, message: string, data?: object, code?: number}}
 *          data = { item_id, sku, name, before_rate, after_rate, requested_rate, message }
 */
function triggerZohoPriceWrite(itemId, sku, price) {
  var url = N8N_ZOHO_PRICE_WRITE_WEBHOOK_URL;
  if (!url) {
    return { ok: false, message: "Price-write webhook URL not configured (check Secrets.js)" };
  }
  itemId = String(itemId || "").trim();
  if (!itemId) {
    return { ok: false, message: "item_id is empty — can't write" };
  }
  var p = parseFloat(price);
  if (!isFinite(p) || p <= 0) {
    return { ok: false, message: "price must be a positive number (got " + price + ")" };
  }

  try {
    var options = {
      method:             'post',
      muteHttpExceptions: true,
      followRedirects:    true,
      headers: {
        'X-API-Token':                APP_SECRET_TOKEN,
        'ngrok-skip-browser-warning': 'true',
        'User-Agent':                 'GoogleAppsScript'
      },
      contentType: 'application/json',
      payload:     JSON.stringify({ item_id: itemId, sku: String(sku || ""), price: p })
    };

    var response = UrlFetchApp.fetch(url, options);
    var code     = response.getResponseCode();
    var body     = response.getContentText();

    if (code !== 200 && code !== 204) {
      if (code === 401 || code === 403) {
        return { ok: false, code: code, message: "Auth rejected by n8n (X-API-Token mismatch)" };
      }
      if (code === 404) {
        return { ok: false, code: code, message: "Workflow not active in n8n — toggle it on" };
      }
      if (code === 500) {
        // n8n returns 500 on a thrown node error (sanity-gate reject, bad
        // item_id, Zoho auth/scope failure). Surface the underlying message.
        var errMsg = body;
        try {
          var parsed = JSON.parse(body);
          errMsg = parsed.message || parsed.error || body;
        } catch (_) {}
        return { ok: false, code: code, message: errMsg };
      }
      return { ok: false, code: code, message: "n8n returned HTTP " + code + ": " + body.substring(0, 200) };
    }

    var parsedBody;
    try { parsedBody = JSON.parse(body); }
    catch (parseErr) {
      return { ok: false, code: code, message: "n8n returned malformed JSON: " + body.substring(0, 200) };
    }

    return {
      ok:      true,
      code:    code,
      message: "Price written to Zoho",
      data:    parsedBody
    };
  } catch (err) {
    return { ok: false, message: "Price-write request failed: " + (err.message || err) };
  }
}


/**
 * Write a BATCH of item prices back to Zoho via the bulk write proxy, and
 * wait for the per-SKU result (synchronous — the modal renders it instantly).
 * The proxy writes rate + purchase_rate (= eBay price) for each item, throttled.
 *
 * Keep batches small (~25-30) — Apps Script's request window bounds how long
 * we can wait. The caller (applyPricePush) enforces the cap before calling.
 *
 * @param {Array<{item_id:string, sku:string, before:(number|null), price:number}>} items
 * @returns {{ok:boolean, message:string, code?:number, data?:{
 *            ok:boolean, total:number, pushed:number, failed:number,
 *            results:Array<{sku, item_id, ok, before, after, error}>}}}
 */
function triggerZohoPriceBulkWrite(items) {
  var url = N8N_ZOHO_PRICE_BULK_WRITE_WEBHOOK_URL;
  if (!url) {
    return { ok: false, message: "Bulk price-write webhook URL not configured (check Secrets.js)" };
  }
  if (!Array.isArray(items) || items.length === 0) {
    return { ok: false, message: "No items to write" };
  }

  try {
    var options = {
      method:             'post',
      muteHttpExceptions: true,
      followRedirects:    true,
      headers: {
        'X-API-Token':                APP_SECRET_TOKEN,
        'ngrok-skip-browser-warning': 'true',
        'User-Agent':                 'GoogleAppsScript'
      },
      contentType: 'application/json',
      payload:     JSON.stringify({ items: items })
    };

    var response = UrlFetchApp.fetch(url, options);
    var code     = response.getResponseCode();
    var body     = response.getContentText();

    if (code !== 200 && code !== 204) {
      if (code === 401 || code === 403) {
        return { ok: false, code: code, message: "Auth rejected by n8n (X-API-Token mismatch)" };
      }
      if (code === 404) {
        return { ok: false, code: code, message: "Workflow not active in n8n — toggle it on" };
      }
      if (code === 500) {
        var errMsg = body;
        try { var parsed = JSON.parse(body); errMsg = parsed.message || parsed.error || body; } catch (_) {}
        return { ok: false, code: code, message: errMsg };
      }
      return { ok: false, code: code, message: "n8n returned HTTP " + code + ": " + body.substring(0, 200) };
    }

    var parsedBody;
    try { parsedBody = JSON.parse(body); }
    catch (parseErr) {
      return { ok: false, code: code, message: "n8n returned malformed JSON: " + body.substring(0, 200) };
    }

    return { ok: true, code: code, message: "Bulk price write complete", data: parsedBody };
  } catch (err) {
    return { ok: false, message: "Bulk price-write request failed: " + (err.message || err) };
  }
}


/**
 * Internal: fire a Header-Auth-protected webhook and translate the response
 * into the {ok, message, code} shape the sidebar expects.
 *
 * Catches network errors, n8n errors, and 401/404/502 cases distinctly so the
 * status bar can show a useful message instead of a generic "failed."
 */
function _fireAuthedWebhook(url, successMessage) {
  if (!url) {
    return { ok: false, message: "Webhook URL not configured (check Secrets.js)", code: 0 };
  }

  try {
    var options = {
      'method':              'post',
      'muteHttpExceptions':  true,
      'followRedirects':     true,
      'timeout':             30000,
      'headers': {
        'X-API-Token':                  APP_SECRET_TOKEN,
        'ngrok-skip-browser-warning':   'true',
        'User-Agent':                   'GoogleAppsScript'
      },
      // Empty JSON body — n8n's POST webhook node accepts no-body, but some
      // intermediate proxies prefer a non-zero content-length. Belt + braces.
      'contentType': 'application/json',
      'payload':     '{}'
    };

    var response = UrlFetchApp.fetch(url, options);
    var code     = response.getResponseCode();

    if (code === 200 || code === 204) {
      return { ok: true, message: successMessage, code: code };
    }
    if (code === 401 || code === 403) {
      return { ok: false, message: "Auth rejected (token mismatch — check Header Auth credential in n8n)", code: code };
    }
    if (code === 404) {
      return { ok: false, message: "Webhook not found — is the workflow Active in n8n?", code: code };
    }
    if (code === 502) {
      return { ok: false, message: "ngrok can't reach n8n — is n8n running on port 5678?", code: code };
    }
    return { ok: false, message: "n8n returned HTTP " + code, code: code };

  } catch (err) {
    Logger.log("Webhook fire error (" + url + "): " + err.toString());
    return { ok: false, message: "Connection error: " + (err && err.message ? err.message : err), code: 0 };
  }
}
