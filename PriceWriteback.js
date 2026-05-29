// =======================================================================================
// PriceWriteback.js — push eBay's authoritative price into Zoho's selling_price
// Phase 1 (2026-05-29): single-item SAFE TEST only. No bulk, no sheet-wide writes.
// =======================================================================================
//
// WHY
// ---
// Price Audit (PriceAudit.js) surfaces ~500 items where Zoho's selling_price has
// drifted from eBay's live price. Today they're fixed by hand in Zoho, one at a
// time, because the team has to re-check eBay before quoting every direct order.
// eBay is the price authority (~99% of the time), so the fix is: copy eBay's
// price → Zoho. This file is the WRITE half the audit always lacked.
//
// SAFETY MODEL (this phase)
// -------------------------
//   - ONE item per call, by explicit SKU. Nothing is bulk yet.
//   - eBay is the source: we only ever write eBay's active-listing price. If the
//     SKU has no active eBay price, we REFUSE (nothing authoritative to copy).
//   - Sanity gate on BOTH sides (here + the n8n proxy): refuse 0 / negative /
//     NaN / absurd values. This is the guard against the 700-priceless-items
//     class of bad-data write.
//   - The proxy GETs the item first and returns before/after, so the caller sees
//     exactly what Zoho stored — the real confirmation, not just "200 OK".
//   - We write eBay's price into BOTH `rate` (selling) and `purchase_rate` (cost).
//     Zoho is warehouse-management-only here, not accounting, so the two are kept
//     equal (per user, 2026-05-29). We touch nothing else — not stock, not name,
//     not status.
//
// HARD PREREQUISITE: a Zoho refresh token minted with `ZohoInventory.items.UPDATE`
// scope, pasted into the Zoho Item Price Write Proxy workflow (n8n side). Until
// that exists, triggerZohoPriceWrite returns a Zoho auth error and nothing changes.
//
// NEXT PHASE (not built): row-selection on the Price Audit sheet → push the
// selected/approved rows in a throttled batch, with an Activity-Log entry per write.
// =======================================================================================


/**
 * Resolve a SKU to its Zoho item_id + eBay price, then write eBay's price into
 * Zoho's selling_price for that one item. Returns a structured result with the
 * before/after rate Zoho reported.
 *
 * Resolution sources (both already used by Price Audit, so the numbers match
 * what you saw in the audit sheet):
 *   - item_id + current Zoho price : buildZohoStockMap()  (ZohoStock.js)
 *   - eBay authoritative price      : _buildActiveEbayMaps() (PriceAudit.js),
 *                                     active listings only
 *
 * @param {string} sku
 * @returns {{ok:boolean, message:string, sku?:string, itemId?:string,
 *            zohoBefore?:number, ebayTarget?:number, zohoAfter?:number, detail?:object}}
 */
function pushSinglePriceToZohoBySku(sku) {
  sku = String(sku || "").trim();
  if (!sku) return { ok: false, message: "No SKU given." };
  var skuLower = sku.toLowerCase();

  // --- Resolve Zoho item_id + current Zoho price from the Zoho Stock mirror ---
  var zMap = buildZohoStockMap();
  if (!zMap || zMap.size === 0) {
    return { ok: false, message: "Zoho Stock sheet is empty — run Sync Zoho Stock first." };
  }
  var z = zMap.get(skuLower);
  if (!z) {
    return { ok: false, message: "SKU not found in Zoho Stock sheet: " + sku };
  }
  if (!z.itemId) {
    return { ok: false, message: "No Zoho item_id recorded for SKU " + sku + " — can't target the write." };
  }

  // --- Resolve eBay's authoritative price (active listings only) ---
  var miMaps = _buildActiveEbayMaps();
  var ebayPrice = miMaps.prices.get(skuLower);
  if (ebayPrice == null || !(ebayPrice > 0)) {
    return { ok: false,
             message: "No active eBay price for SKU " + sku + " — refusing to write " +
                      "(eBay is the authority; nothing to copy). If the listing ended, this " +
                      "is an INACTIVE candidate, not a price fix." };
  }

  // --- Sanity gate (mirrors the n8n proxy's gate; fail fast before any network) ---
  if (!isFinite(ebayPrice) || ebayPrice < 0.01 || ebayPrice > 100000) {
    return { ok: false, message: "eBay price " + ebayPrice + " is outside sanity bounds [0.01, 100000] — refusing to write." };
  }

  var zohoBefore = z.sellingPrice || 0;

  // NOTE: no no-op short-circuit here. We now sync BOTH selling and cost price,
  // but the Zoho Stock sheet only mirrors selling price — we can't see the cost
  // from here to know whether it's already in sync. Writing is idempotent
  // (re-sending the same values is harmless), so we always send and let the
  // proxy's before/after report what actually moved.

  try {
    console.log("PRICE WRITEBACK (test) → SKU=" + sku + " item_id=" + z.itemId +
                " | Zoho sell now $" + zohoBefore + " → setting BOTH sell+cost to eBay $" + ebayPrice);
  } catch (_) {}

  var res = triggerZohoPriceWrite(z.itemId, sku, ebayPrice);

  if (!res || !res.ok) {
    return { ok: false,
             message: "Write failed: " + (res && res.message ? res.message : "unknown"),
             sku: sku, itemId: z.itemId, zohoBefore: zohoBefore, ebayTarget: ebayPrice,
             detail: res };
  }

  var data = res.data || {};
  return {
    ok:         true,
    message:    "✅ " + (data.message || ("Wrote $" + ebayPrice.toFixed(2) + " (sell+cost) to " + sku)),
    sku:        sku,
    itemId:     z.itemId,
    name:       data.name || z.itemName || "",
    ebayTarget: ebayPrice,
    sellBefore: data.before_rate != null ? data.before_rate : zohoBefore,
    sellAfter:  data.after_rate != null ? data.after_rate : null,
    costBefore: data.before_purchase_rate != null ? data.before_purchase_rate : null,
    costAfter:  data.after_purchase_rate  != null ? data.after_purchase_rate  : null,
    detail:     data
  };
}


/**
 * EDITOR-RUN TEST WRAPPER.
 * Type the SKU you want to test into TEST_SKU below, then Run this function
 * from the Apps Script editor. Logs the full before/after result.
 *
 * SAFE BY DEFAULT: with TEST_SKU left as the placeholder, it refuses to write
 * (the SKU won't resolve) — you must deliberately set a real SKU to do anything.
 */
function runSinglePriceWriteTest() {
  var TEST_SKU = "REPLACE_WITH_SKU_TO_TEST";   // <-- type one SKU here, then Run

  var r = pushSinglePriceToZohoBySku(TEST_SKU);
  try { console.log(JSON.stringify(r, null, 2)); } catch (_) { console.log(r); }
  return r;
}


// =======================================================================================
// BULK — selection-driven Price Push modal
// =======================================================================================
//
// Flow: select drift rows on the Price Audit sheet → sidebar button →
// openPricePushModal() reads the selection, resolves item_id + before/after,
// caches the candidates, opens PricePushModal.html → picker unchecks anything →
// Apply (passphrase + confirm) → applyPricePush() fires the n8n bulk write proxy
// SYNCHRONOUSLY and returns a per-SKU report the modal renders + a Price Push
// Log row per write.
//
// SECURITY: everyone uses the public link (no per-email identity), so the gate
// is a server-checked passphrase stored in a Script Property (set once via
// setPricePushPassphrase from the editor — never in the sidebar HTML).
//
// CAP: each push is bounded (~25-30) so the synchronous n8n round-trip finishes
// inside Apps Script's request window. Clear the 500 backlog in chunks.
// =======================================================================================

var PRICE_PUSH_CAP = 30;                       // max rows per push (sync window bound)
var PRICE_PUSH_PASS_KEY = "PRICE_PUSH_PASSPHRASE";   // Script Property slot NAME (a label) — NOT the secret. The secret is the argument you pass to setPricePushPassphrase().
var PRICE_PUSH_BIG_SWING = 0.5;                // |after-before|/before > this → amber flag


/** EDITOR-RUN: set the price-push passphrase. Run once, privately. */
function setPricePushPassphrase(passphrase) {
  passphrase = String(passphrase == null ? "" : passphrase);
  if (!passphrase) return "❌ Empty passphrase — pass a non-empty string.";
  PropertiesService.getScriptProperties().setProperty(PRICE_PUSH_PASS_KEY, passphrase);
  return "✅ Price-push passphrase set (" + passphrase.length + " chars).";
}

/** EDITOR-RUN: select THIS function in the editor and click Run to set the
 *  passphrase (the Run button can't pass arguments to setPricePushPassphrase
 *  directly, so this wrapper supplies it). Change the secret here if you ever
 *  rotate it. You can blank the literal back to "" after running if you like —
 *  the value is already saved in the Script Property by then. */
function setMyPricePushPassphraseNow() {
  return setPricePushPassphrase("");
}

/** True if a passphrase has been configured. */
function hasPricePushPassphrase() {
  return !!PropertiesService.getScriptProperties().getProperty(PRICE_PUSH_PASS_KEY);
}

/** Server-side passphrase check. Refuses if none is configured. */
function _checkPricePushPassphrase(passphrase) {
  var stored = PropertiesService.getScriptProperties().getProperty(PRICE_PUSH_PASS_KEY);
  if (!stored) return false;
  return String(passphrase == null ? "" : passphrase) === stored;
}


// =======================================================================================
// PRICE PUSH LOG — append-only history of every price write
// =======================================================================================

var PRICE_PUSH_LOG = {
  sheetName: "Price Push Log",
  cols: { TIMESTAMP: 1, SKU: 2, ITEM_NAME: 3, BEFORE: 4, AFTER: 5, RESULT: 6, DETAIL: 7 },
  dataWidth: 7,
  headerRow: 1,
  dataStartRow: 2,
  headers: ["TIMESTAMP", "SKU", "ITEM NAME", "BEFORE", "AFTER", "RESULT", "DETAIL"]
};

function setupPricePushLogSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PRICE_PUSH_LOG.sheetName);
  if (!sheet) sheet = ss.insertSheet(PRICE_PUSH_LOG.sheetName);

  sheet.getRange(PRICE_PUSH_LOG.headerRow, 1, 1, PRICE_PUSH_LOG.dataWidth)
    .setValues([PRICE_PUSH_LOG.headers])
    .setBackground('#1d1d1b').setFontColor('#ffd966').setFontFamily('Oswald')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(PRICE_PUSH_LOG.headerRow, 1, 1, PRICE_PUSH_LOG.dataWidth)
    .setBorder(null, null, true, null, null, null, '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);

  sheet.setColumnWidth(PRICE_PUSH_LOG.cols.TIMESTAMP, 150);
  sheet.setColumnWidth(PRICE_PUSH_LOG.cols.SKU,       100);
  sheet.setColumnWidth(PRICE_PUSH_LOG.cols.ITEM_NAME, 320);
  sheet.setColumnWidth(PRICE_PUSH_LOG.cols.BEFORE,     90);
  sheet.setColumnWidth(PRICE_PUSH_LOG.cols.AFTER,      90);
  sheet.setColumnWidth(PRICE_PUSH_LOG.cols.RESULT,     80);
  sheet.setColumnWidth(PRICE_PUSH_LOG.cols.DETAIL,    320);

  var maxDataRow = 8000;
  var dataRows = maxDataRow - PRICE_PUSH_LOG.dataStartRow + 1;
  sheet.getRange(PRICE_PUSH_LOG.dataStartRow, PRICE_PUSH_LOG.cols.TIMESTAMP, dataRows, 1)
    .setNumberFormat('M/d/yy h:mm am/pm').setFontFamily('Roboto Mono').setFontSize(9)
    .setFontColor('#5f5f5f').setHorizontalAlignment('center');
  sheet.getRange(PRICE_PUSH_LOG.dataStartRow, PRICE_PUSH_LOG.cols.SKU, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(PRICE_PUSH_LOG.dataStartRow, PRICE_PUSH_LOG.cols.ITEM_NAME, dataRows, 1)
    .setFontFamily('Roboto').setFontSize(10).setHorizontalAlignment('left');
  sheet.getRange(PRICE_PUSH_LOG.dataStartRow, PRICE_PUSH_LOG.cols.BEFORE, dataRows, 2)
    .setNumberFormat('$#,##0.00').setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(PRICE_PUSH_LOG.dataStartRow, PRICE_PUSH_LOG.cols.RESULT, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(PRICE_PUSH_LOG.dataStartRow, PRICE_PUSH_LOG.cols.DETAIL, dataRows, 1)
    .setFontFamily('Roboto').setFontSize(9).setFontColor('#5f5f5f').setHorizontalAlignment('left');

  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  var band = sheet.getRange(1, 1, maxDataRow, PRICE_PUSH_LOG.dataWidth)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b').setFirstRowColor('#ffffff').setSecondRowColor('#fff8e7');

  // RESULT tint: OK green, FAIL red
  var existing = sheet.getConditionalFormatRules() || [];
  var keep = existing.filter(function(r) {
    var rs = r.getRanges();
    if (!rs || rs.length === 0) return true;
    return !rs.some(function(rg) {
      return rg.getSheet().getName() === PRICE_PUSH_LOG.sheetName &&
             rg.getColumn() === PRICE_PUSH_LOG.cols.RESULT;
    });
  });
  var resRange = sheet.getRange(PRICE_PUSH_LOG.dataStartRow, PRICE_PUSH_LOG.cols.RESULT, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('OK').setBackground('#c8e6c9').setFontColor('#1b5e20')
    .setRanges([resRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('FAIL').setBackground('#ffcdd2').setFontColor('#b71c1c')
    .setRanges([resRange]).build());
  sheet.setConditionalFormatRules(keep);

  sheet.setFrozenRows(1);
  return "✅ Price Push Log sheet ready.";
}

/** Sidebar: open the Price Push Log sheet. */
function openPricePushLog() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet";
  var sheet = ss.getSheetByName(PRICE_PUSH_LOG.sheetName);
  if (!sheet) { setupPricePushLogSheet(); sheet = ss.getSheetByName(PRICE_PUSH_LOG.sheetName); }
  ss.setActiveSheet(sheet);
  return "✅ Opened Price Push Log";
}

/** Append result rows (best-effort — a log failure never blocks the push). */
function _appendPricePushLog(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return;
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(PRICE_PUSH_LOG.sheetName);
    if (!sheet) { setupPricePushLogSheet(); sheet = ss.getSheetByName(PRICE_PUSH_LOG.sheetName); }
    var startRow = Math.max(sheet.getLastRow() + 1, PRICE_PUSH_LOG.dataStartRow);
    sheet.getRange(startRow, 1, rows.length, PRICE_PUSH_LOG.dataWidth).setValues(rows);
  } catch (e) {
    try { console.log("_appendPricePushLog error: " + e); } catch (_) {}
  }
}


// =======================================================================================
// MODAL: openPricePushModal() — reads selection, builds candidates, opens modal
// =======================================================================================

/**
 * Sidebar entry. Reads the highlighted rows on the Price Audit sheet, resolves
 * each to a push candidate (item_id + before/after price), caches them, and
 * opens PricePushModal.html. Returns a status-only result if nothing pushable
 * is selected (no modal opens).
 *
 * @returns {{ok:boolean, modalOpened:boolean, reason:string,
 *            pushable:number, skipped:number}}
 */
function openPricePushModal() {
  try {
    var ss = SpreadsheetApp.getActive();
    if (!ss) return { ok: false, modalOpened: false, reason: "No active spreadsheet.", pushable: 0, skipped: 0 };
    var sheet = ss.getActiveSheet();
    if (!sheet || sheet.getName() !== PRICE_AUDIT.sheetName) {
      return { ok: false, modalOpened: false,
               reason: "Open the Price Audit sheet and highlight the rows to push first.",
               pushable: 0, skipped: 0 };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < PRICE_AUDIT.dataStartRow) {
      return { ok: false, modalOpened: false, reason: "Price Audit sheet is empty — run the audit first.", pushable: 0, skipped: 0 };
    }

    // Gather highlighted data-row numbers
    var rangeList = sheet.getActiveRangeList();
    var ranges = rangeList ? rangeList.getRanges() : [sheet.getActiveRange()];
    var rowSet = {};
    ranges.forEach(function(rg) {
      if (!rg) return;
      var start = rg.getRow();
      var num = rg.getNumRows();
      for (var r = start; r < start + num; r++) {
        if (r >= PRICE_AUDIT.dataStartRow && r <= lastRow) rowSet[r] = true;
      }
    });
    var rows = Object.keys(rowSet).map(Number).sort(function(a, b) { return a - b; });
    if (rows.length === 0) {
      return { ok: false, modalOpened: false, reason: "No data rows highlighted — select drift rows on the Price Audit sheet.", pushable: 0, skipped: 0 };
    }

    // Read the data block once + the Zoho Stock map (for item_id join)
    var block = sheet.getRange(PRICE_AUDIT.dataStartRow, 1, lastRow - PRICE_AUDIT.dataStartRow + 1, PRICE_AUDIT.dataWidth).getValues();
    var zohoMap = buildZohoStockMap();

    var candidates = [];
    var pushableCount = 0;
    rows.forEach(function(rowNum) {
      var vals = block[rowNum - PRICE_AUDIT.dataStartRow];
      if (!vals) return;
      var sku = String(vals[PRICE_AUDIT.idx("SKU")] || "").trim();
      if (!sku) return;
      var name = String(vals[PRICE_AUDIT.idx("ITEM_NAME")] || "");
      var ebay = parseFloat(vals[PRICE_AUDIT.idx("EBAY_LIVE")]);
      var zoho = parseFloat(vals[PRICE_AUDIT.idx("ZOHO_LIVE")]);
      var direction = String(vals[PRICE_AUDIT.idx("DIRECTION")] || "").trim().toUpperCase();

      var zRec = zohoMap.get(sku.toLowerCase());
      var itemId = zRec ? zRec.itemId : "";
      var zohoBefore = isFinite(zoho) ? zoho : (zRec && zRec.sellingPrice ? zRec.sellingPrice : null);

      var pushable = true, skipReason = "";
      if (direction !== "ZOHO HIGH" && direction !== "ZOHO LOW") {
        pushable = false; skipReason = direction || "not a price drift";
      } else if (!isFinite(ebay) || ebay <= 0) {
        pushable = false; skipReason = "no eBay price to copy";
      } else if (!itemId) {
        pushable = false; skipReason = "no Zoho item_id (run Sync Zoho Stock)";
      } else if (ebay < 0.01 || ebay > 100000) {
        pushable = false; skipReason = "eBay price failed sanity bounds";
      }

      var ebayTarget = (isFinite(ebay) && ebay > 0) ? ebay : null;
      var delta = (ebayTarget != null && zohoBefore != null) ? (ebayTarget - zohoBefore) : null;
      var bigSwing = (zohoBefore && zohoBefore > 0 && ebayTarget != null)
                   ? (Math.abs(ebayTarget - zohoBefore) / zohoBefore > PRICE_PUSH_BIG_SWING)
                   : (ebayTarget != null);  // unknown before → treat as notable

      if (pushable) pushableCount++;
      candidates.push({
        sku: sku, name: name, direction: direction,
        zohoBefore: zohoBefore, ebayTarget: ebayTarget, delta: delta,
        bigSwing: bigSwing, itemId: itemId,
        pushable: pushable, skipReason: skipReason
      });
    });

    if (candidates.length === 0) {
      return { ok: false, modalOpened: false, reason: "Highlighted rows have no SKUs.", pushable: 0, skipped: 0 };
    }
    if (pushableCount === 0) {
      return { ok: false, modalOpened: false,
               reason: "Nothing pushable in selection — rows are INACTIVE/OOS, missing a Zoho item_id, or have no eBay price.",
               pushable: 0, skipped: candidates.length };
    }
    if (pushableCount > PRICE_PUSH_CAP) {
      return { ok: false, modalOpened: false,
               reason: "Too many pushable rows selected (" + pushableCount + "). Keep each push ≤ " + PRICE_PUSH_CAP + " — highlight fewer and run in chunks.",
               pushable: pushableCount, skipped: candidates.length - pushableCount };
    }

    // Cache candidates for the apply step
    var sessionId = Utilities.getUuid();
    CacheService.getUserCache().put("pricepush_" + sessionId, JSON.stringify(candidates), 1800);

    var syncedAt = null;
    try { var s = getZohoStockSyncedAt(); syncedAt = s ? s.getTime() : null; } catch (_) {}

    var payload = {
      sessionId: sessionId,
      candidates: candidates,
      meta: {
        pushable: pushableCount,
        skipped: candidates.length - pushableCount,
        cap: PRICE_PUSH_CAP,
        hasPassphrase: hasPricePushPassphrase(),
        zohoSyncedAt: syncedAt
      }
    };

    // Force-unescaped JSON injection (see ZohoPullModal pattern). Pre-substitute
    // </ so any value can't close the script tag early.
    var dataJson = JSON.stringify(payload).replace(/<\//g, "<\\/");
    var template = HtmlService.createTemplateFromFile("PricePushModal");
    template.dataJson = dataJson;
    var html = template.evaluate().setWidth(880).setHeight(620);

    SpreadsheetApp.getUi().showModalDialog(html,
      "Push Prices → Zoho · " + pushableCount + " ready"
      + (candidates.length - pushableCount > 0 ? " · " + (candidates.length - pushableCount) + " skipped" : ""));

    return { ok: true, modalOpened: true, reason: "", pushable: pushableCount, skipped: candidates.length - pushableCount };
  } catch (err) {
    try { console.log("openPricePushModal error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, modalOpened: false, reason: "Failed to open Price Push modal: " + (err.message || err), pushable: 0, skipped: 0 };
  }
}


// =======================================================================================
// APPLY: applyPricePush(sessionId, checkedSkus, passphrase)
// =======================================================================================

/**
 * Commit the checked rows. Validates the passphrase, reads the cached
 * candidates, filters to pushable+checked, fires the n8n bulk write proxy
 * synchronously, logs every result, and returns a per-SKU report.
 *
 * @param {string} sessionId
 * @param {string[]} checkedSkus
 * @param {string} passphrase
 * @returns {{ok:boolean, reason?:string, pushed?:number, failed?:number,
 *            total?:number, results?:Array}}
 */
function applyPricePush(sessionId, checkedSkus, passphrase) {
  try {
    if (!_checkPricePushPassphrase(passphrase)) {
      return { ok: false, reason: hasPricePushPassphrase()
        ? "Passphrase incorrect."
        : "No passphrase configured — admin must run setPricePushPassphrase first." };
    }

    var raw = CacheService.getUserCache().get("pricepush_" + String(sessionId || ""));
    if (!raw) {
      return { ok: false, reason: "Session expired — close and re-open the modal." };
    }
    var candidates;
    try { candidates = JSON.parse(raw); } catch (e) { return { ok: false, reason: "Cached selection corrupt — re-open the modal." }; }

    var checkedSet = {};
    (Array.isArray(checkedSkus) ? checkedSkus : []).forEach(function(s) { checkedSet[String(s)] = true; });

    var nameBySku = {};
    var toPush = [];
    candidates.forEach(function(c) {
      nameBySku[c.sku] = c.name || "";
      if (c.pushable && checkedSet[c.sku]) {
        toPush.push({ item_id: c.itemId, sku: c.sku, before: c.zohoBefore, price: c.ebayTarget });
      }
    });

    if (toPush.length === 0) {
      return { ok: true, pushed: 0, failed: 0, total: 0, results: [], reason: "Nothing selected." };
    }
    if (toPush.length > PRICE_PUSH_CAP) {
      return { ok: false, reason: "Too many selected (" + toPush.length + ") — cap is " + PRICE_PUSH_CAP + " per push." };
    }

    var res = triggerZohoPriceBulkWrite(toPush);
    if (!res || !res.ok || !res.data) {
      return { ok: false, reason: "Bulk write failed: " + (res && res.message ? res.message : "unknown") };
    }

    var data = res.data;
    var results = Array.isArray(data.results) ? data.results : [];

    // Append every result to the Price Push Log
    var now = new Date();
    var logRows = results.map(function(r) {
      return [
        now,
        r.sku || "",
        nameBySku[r.sku] || "",
        (r.before != null ? r.before : ""),
        (r.after  != null ? r.after  : ""),
        r.ok ? "OK" : "FAIL",
        r.ok ? "" : (r.error || "")
      ];
    });
    _appendPricePushLog(logRows);

    return {
      ok: true,
      pushed: data.pushed || 0,
      failed: data.failed || 0,
      total: data.total || results.length,
      results: results
    };
  } catch (err) {
    try { console.log("applyPricePush error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: "Apply failed: " + (err.message || err) };
  }
}
