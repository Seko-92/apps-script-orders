// =======================================================================================
// ORDER_SERVICE.gs - COMPLETE with Hidden Sheet Message ID Storage/
// =======================================================================================

// Note: TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID are now defined in Secrets.js
// Make sure Secrets.js is uploaded to your Apps Script project
var HIDDEN_SHEET_NAME = "Telegram_Messages"; // Hidden sheet for message IDs

/**
 * The "Front Door" for n8n - Receives POST requests
 */
/**
 * The "Front Door" for n8n - Receives POST requests
 */
/* ⚠⚠ READ-ONLY ACTIONS MUST NOT TAKE THE WRITE LOCK (2026-08-17).
   doPost took LockService.getScriptLock() with a 30s wait at the top for EVERY
   request — including the pure reads the Floor Board fires constantly. So every
   read serialised against every write and against every other read, through one
   global lock, and a board tap had to wait behind whatever else was queued.

   THE SIGNATURE IS UNMISTAKABLE ONCE YOU SEE IT: the failures measured on the
   floor came back at 32.3 / 32.8 / 32.9 / 33.5 SECONDS — that is waitLock(30000)
   expiring, plus overhead, returning {"status":"error","message":"Server Busy"}.
   Not CPU, not Google being slow, not the network: lock starvation, self-
   inflicted. And "Server Busy" carries no `ok` key, so on the old board it also
   produced a SILENT un-pick.

   Reads touch nothing that needs serialising — a tick is a cached cell, a part
   dossier is a lookup. Only the mutating actions take the lock now.

   ⚠ THE ALLOWLIST IS THE SAFETY BOUNDARY AND IT IS DELIBERATELY SHORT. Anything
   not named here keeps the lock, so a new action is locked BY DEFAULT and a
   mistake fails safe. Never add an action here without checking it writes
   nothing — including transitively. */
var DOPOST_LOCK_FREE = {
  boardTick: 1, boardRadio: 1, boardPickers: 1, boardPrint: 1,
  boardPart: 1, boardPartLite: 1, boardOrder: 1
};

/* ⭐ DOES THIS REQUEST NEED THE SCRIPT LOCK? (2026-08-28)
   ─────────────────────────────────────────────────────────────────────────────
   PURE — no SpreadsheetApp, no LockService, no clock — so the decision that
   guards the floor's most contended resource is Node-testable in isolation.

   ⚠⚠ WHY THE SECOND CLAUSE EXISTS. eBay is connected to Zoho, so EVERY eBay
   order becomes a Zoho Sales Order and fires the SO webhook. upsertPendingSalesOrder
   already discards them — but the filter lives INSIDE it, i.e. AFTER waitLock(30000).
   So each one took a full place in the lock queue and only then discovered it had
   nothing to do.

   Measured 2026-08-28: bursts of 4–12+ webhooks in the SAME SECOND, all different
   SO numbers, all `sales_channel: ebay_us`, all returning
   `skipped · Non-direct channel: ebay_us`. Roughly 130–180 a day — about 6–9
   minutes of runtime and that many lock acquisitions, spent on orders we throw away.
   Zoho polls eBay on a schedule and batch-creates the orders, which is why they
   arrive all at once rather than spread out.

   ⚠ IT IS PROVABLY SAFE TO SKIP THE LOCK HERE. Read upsertPendingSalesOrder's
   opening: a type check, a salesorder_number read, then the channel filter returns.
   The first touch of anything shared — SpreadsheetApp.openById — is AFTER that
   filter. A non-direct SO never reaches a sheet, a cache or a property.

   ⚠ FOUR THINGS THIS MUST NOT BREAK, each pinned by a test:
     1. INVOICES still lock. They arrive on the same action with `invoice` rather
        than `salesorder`, and they DO write (the INVOICE column).
     2. A BLANK channel still locks — mirroring `if (channel && ...)` inside
        upsertPendingSalesOrder exactly. If Zoho ever stops sending the field,
        nothing silently disappears.
     3. Scoped to action `zohoSalesOrder`. The backfill path is left alone.
     4. An unparseable body means `payload` is undefined → lock. Fail safe, as before.

   ⚠ The n8n proxy should ALSO filter these upstream so they never arrive at all
   (that is the bigger win — this saves the lock slot, not the container start).
   This guard stays regardless: deploy/ holds a from-scratch rebuild template with
   no filter node in it, so this is what stops the cost silently returning. */
function _doPostNeedsLock(action, payload) {
  if (DOPOST_LOCK_FREE[action]) return false;

  if (action === 'zohoSalesOrder' && payload && payload.salesorder) {
    var ch = String(payload.salesorder.sales_channel || '').trim().toLowerCase();
    if (ch && ch !== 'direct_sales') return false;   // eBay/Amazon → discarded anyway
  }

  return true;
}

function doPost(e) {
  // Peek at the action before deciding on the lock. Parsing touches no shared
  // state, so it is safe to do outside. Both doors are checked: the body (n8n)
  // and the query string (the Zoho proxies), matching the dispatch below.
  var _peekAction = '';
  try {
    var _peek = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    _peekAction = _peek.action || (e && e.parameter && e.parameter.action) || '';
  } catch (_peekErr) {
    _peekAction = (e && e.parameter && e.parameter.action) || '';
  }

  var lock = null;
  if (_doPostNeedsLock(_peekAction, _peek)) {
    lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
    } catch (lockErr) {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Server Busy"})).setMimeType(ContentService.MimeType.JSON);
    }
  }

  var MAX_RETRIES = 3;
  var lastErr = null;

  try {
    for (var attempt = 1; attempt <= MAX_RETRIES; attempt++) {
      try {
    var payload = JSON.parse(e.postData.contents);

    // --- NORMALIZE ACTION DISPATCH ---
    // Most callers (n8n's own workflows, the sidebar's UrlFetchApp calls) put
    // `action` in the BODY. The Zoho proxy (added 2026-05-15) is the first to
    // put it in the URL query string, because Zoho's "Default Payload" body
    // is rigidly `{"item":{...}}` with no room to inject an action field.
    //
    // Promote URL-side action into the payload so every downstream dispatch
    // check (`payload.action === 'xyz'`) works regardless of where the caller
    // chose to put it. Body wins if both are present (defensive — the caller
    // had a stronger reason to put it in the body).
    if (!payload.action && e.parameter && e.parameter.action) {
      payload.action = e.parameter.action;
    }

    // --- AUTHENTICATION ---
    // Telegram callbacks are verified by their structure (callback_query with valid bot data)
    // All other requests must include the shared secret token
    if (!payload.callback_query) {
      var token = String(payload.token || (e.parameter && e.parameter.token) || "").trim();
      var expected = APP_SECRET_TOKEN;
      if (token !== expected) {
        console.log("AUTH FAILED - received: [" + token + "] expected: [" + expected + "]");
        return ContentService.createTextOutput(JSON.stringify({
          "status": "error", "message": "Unauthorized"
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // --- TELEGRAM ACTIONS ---
    if (payload.action === 'storeMessageId') return storeMessageId(payload.orderId, payload.messageId, payload.chatId);
    if (payload.action === 'notifyShipped') return notifyTelegramShipped(payload.orderId);
    if (payload.callback_query) return handleTelegramCallback(payload);

    // Telegram TEXT commands (/part, /status, …) — the command router, phase 1
    // of the Telegram Layer. The n8n webhook already receives these updates; it
    // just forwards any `message` carrying text here. Accepts the update either
    // wrapped as {update:{…}} or merged alongside action/token, so the n8n node
    // can be wired either way without a code change.
    // NOTE: this path is token-authenticated like every other non-callback
    // action above; the per-CHAT allowlist lives in handleTelegramCommand.
    if (payload.action === 'telegramCommand') {
      var tgResult = handleTelegramCommand(payload.update || payload);
      return ContentService.createTextOutput(JSON.stringify(tgResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // --- FLOOR BOARD (hosted off-platform) -------------------------------------
    // The board used to live on /exec and reach these three functions through
    // google.script.run, which only exists inside Google's HtmlService frame.
    // Once the page is served from our own domain that channel is gone, so the
    // same three calls are exposed here and reached over HTTP through the n8n
    // proxy (which holds the token — the browser never sees it).
    //
    // EXPOSURE IS UNCHANGED: the board was already public on /exec, and
    // boardSetStatus is still narrowed SERVER-SIDE to PENDING / PREPARING only.
    // It cannot ship, cancel or delete no matter who calls it. That allow-list
    // is the security boundary, exactly as it was before the move.
    if (payload.action === 'boardTick') {
      return ContentService.createTextOutput(JSON.stringify(getDashboardTick()))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === 'boardAdjust') {
      return ContentService.createTextOutput(JSON.stringify(
        boardAdjustStock(payload.sku, payload.target, payload.orderId)
      )).setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === 'boardLeft') {
      return ContentService.createTextOutput(JSON.stringify(
        boardSetLeft(payload.orderId, payload.sku, payload.count)
      )).setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === 'boardStatus') {
      return ContentService.createTextOutput(JSON.stringify(
        boardSetStatus(payload.orderId, payload.status, payload.sku)
      )).setMimeType(ContentService.MimeType.JSON);
    }
    // ⚠ ACKNOWLEDGING A HOLD — the tablet's half of the 2026-08-21 fix. It is a
    // WRITE, so it is deliberately absent from DOPOST_LOCK_FREE and keeps the
    // script lock like every other write. It is NOT gated on the Pick ID: an
    // urgent alert is the worst possible moment to force someone through a
    // dropdown, and boardAckHold records "the floor" rather than refusing.
    if (payload.action === 'boardAckHold') {
      return ContentService.createTextOutput(JSON.stringify(
        boardAckHold(payload.orderId, 'board')
      )).setMimeType(ContentService.MimeType.JSON);
    }
    // ⚠ THE FLOOR'S DOOR FOR A MISSING / REPLACEMENT LINE (2026-08-29).
    // A WRITE, so deliberately ABSENT from DOPOST_LOCK_FREE — it keeps the script
    // lock like every other write, and the allowlist being deny-by-default means
    // simply not naming it here is the correct and safe outcome.
    //
    // Unlike boardAckHold above, this one IS gated on the Pick ID
    // (boardAddReplacementLine → _boardRequirePicker). A hold ack is an emergency
    // where a dropdown is the wrong ask; adding a pick line is deliberate data
    // entry, and it should carry a name. `needsPicker` comes back to the board,
    // which opens the picker drawer and replays the action.
    //
    // This exists because once All Orders is column-locked, an employee cannot
    // type into col D at all — and doPost runs as the OWNER, so this path writes
    // straight past that protection. That is the point.
    if (payload.action === 'boardMissingLine') {
      return ContentService.createTextOutput(JSON.stringify(
        boardAddReplacementLine(payload.kind, payload.originalOrder,
                                payload.sku, payload.qty, payload.note)
      )).setMimeType(ContentService.MimeType.JSON);
    }
    // --- HOSTED KIT EXPANSION (web-app slice 2) --------------------------------
    // Two endpoints only: read the queue, commit one kit. The client sends
    // IDENTIFIERS and CHOICES; commitKitFromWeb re-derives the kit from a fresh
    // scan, so a stale page or tampered payload fails loudly instead of writing
    // something wrong.
    //
    // ⚠ AUTH: unlike the board endpoints above, these WRITE inventory rows and
    // must NOT be open. The n8n `hq-kits` workflow verifies the Telegram
    // initData HMAC before forwarding — that check is the boundary, and it is
    // deliberately a SEPARATE workflow from hq-board so the difference is
    // structural rather than a conditional someone can get wrong later.
    if (payload.action === 'kitsQueue' || payload.action === 'kitsCommit' ||
        payload.action === 'kitsLookup') {
      // The shared token proves the call came through our proxy; it does NOT
      // say who is asking, and every browser hitting the page would present the
      // same one. initData is what identifies the human.
      var who = authorizeWebAppUser(payload.initData);
      if (!who.ok) {
        try { console.log("web-app auth refused (" + payload.action + "): " + who.reason); } catch (_) {}
        return ContentService.createTextOutput(JSON.stringify({
          ok: false, authError: true, message: who.reason
        })).setMimeType(ContentService.MimeType.JSON);
      }

      if (payload.action === 'kitsQueue') {
        return ContentService.createTextOutput(JSON.stringify(getKitQueueForWeb()))
          .setMimeType(ContentService.MimeType.JSON);
      }
      // Component SWAP / custom ADD validation. The modal reaches
      // lookupSkuForKitAlter over google.script.run, which the hosted page has
      // no access to — without this endpoint the ⇄ swap and "+ add component"
      // controls would look alive and silently do nothing. commitKitFromWeb
      // already accepts `alterations`, so this is only the READ half.
      // Read-only (name / shelf / on-hand) and behind the same initData gate.
      if (payload.action === 'kitsLookup') {
        return ContentService.createTextOutput(JSON.stringify(
          lookupSkuForKitAlter(payload.sku)
        )).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify(commitKitFromWeb(
        payload.kitSku, payload.salesOrder, payload.excludedSkus,
        payload.extras, payload.force, payload.alterations
      ))).setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'boardRadio') {
      // getRadioNowPlaying returns a bare string; wrap it so every board
      // endpoint answers with an object and the client can parse uniformly.
      return ContentService.createTextOutput(JSON.stringify({
        ok: true, nowPlaying: getRadioNowPlaying(payload.stationId)
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- BOARD: PRINT + PICKER ---------------------------------------------
    // The pick list as a string, so a tablet can print it — the Sheets mobile
    // app cannot open the Apps Script modal the sidebar uses. The picker guard
    // is unchanged; boardSetPicker is allow-listed to the dropdown's own values.
    if (payload.action === 'boardPrint') {
      return ContentService.createTextOutput(JSON.stringify(
        getPrintHtmlForBoard()
      )).setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === 'boardPickers') {
      return ContentService.createTextOutput(JSON.stringify(
        getBoardPickers()
      )).setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === 'boardSetPicker') {
      return ContentService.createTextOutput(JSON.stringify(
        setBoardPicker(payload.picker)
      )).setMimeType(ContentService.MimeType.JSON);
    }

    // --- TAP-TO-INSPECT (board drawer) -----------------------------------------
    // Thin wrappers over dossiers that ALREADY EXIST and are already reachable
    // from the sidebar consoles — the board simply had no way to ask for them.
    // No new data, no new logic; the whole feature is a doorway.
    //
    // READ-ONLY, and deliberately ON-DEMAND (a tap), never folded into the 20s
    // tick — putting this detail in the poll would undo the tick optimisation
    // that got a board cycle from 26.5s down to 5.5s.
    //
    // ⚠ These inherit the board's OPEN posture. /exec is public by necessity
    // (n8n posts to it unauthenticated) and the board has always been open, so
    // this is the same exposure class as the pick list it sits behind — order
    // ids, SKUs and shelf locations are already on that screen. They add stock
    // figures, prices and an order's timeline. If that ever needs closing, the
    // fix is a Caddy basic_auth on /api/board, which would also have to go on
    // the tablet — a deliberate trade, not an obvious win.
    // Stage 2 of the drawer's progressive load — one MI row + one Zoho row.
    // The heavy dossier (boardPart) follows in the background.
    if (payload.action === 'boardPartLite') {
      return ContentService.createTextOutput(JSON.stringify(
        getPartBasics(payload.sku)
      )).setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === 'boardPart') {
      return ContentService.createTextOutput(JSON.stringify(
        getPartData(payload.sku)
      )).setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === 'boardOrder') {
      return ContentService.createTextOutput(JSON.stringify(
        getOrderCaseData(payload.orderId)
      )).setMimeType(ContentService.MimeType.JSON);
    }

    // --- STATUS UPDATES ---
    if (payload.action === 'updateOrderStatus') {
      var allowedSources = { "n8n": 1, "n8n-verify": 1, "n8n-direct": 1 };
      var sourceTag = allowedSources[payload.source] ? payload.source : "n8n";
      // force: only honored for n8n-verify (the SHIPPED→PENDING revert path).
      // Other sources stay subject to the terminal-state guard so an accidental
      // payload can't roll back a legitimately-shipped order.
      var forceFlag = (sourceTag === "n8n-verify" && payload.force === true);
      var statusResult = updateOrderStatus(payload.orderId, payload.newStatus, {
        source:       sourceTag,
        syncTelegram: false,
        force:        forceFlag
      });
      // Verify-revert: edit the existing Telegram message instead of letting it
      // sit stale at "SHIPPED". Best-effort — if Telegram fails, the sheet
      // revert still stands.
      if (sourceTag === "n8n-verify" && statusResult.success && statusResult.count > 0) {
        try {
          syncStatusToTelegram(payload.orderId, payload.newStatus, {
            revertReason: payload.ebayStatus || "not fulfilled"
          });
        } catch (telegramErr) {
          console.log("verify-revert Telegram edit failed for " + payload.orderId + ": " + telegramErr);
        }
      }
      var responsePayload = {
        found:         statusResult.success && (statusResult.count > 0 || statusResult.blockedCount > 0),
        count:         statusResult.count || 0,
        currentStatus: statusResult.currentStatus || ""
      };
      return ContentService.createTextOutput(JSON.stringify(responsePayload)).setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === 'updateStatus') {
      // Validate row number is within data range
      var rowNum = parseInt(payload.rowNumber);
      if (isNaN(rowNum) || rowNum < Schema.dataStartRow) {
        return ContentService.createTextOutput(JSON.stringify({
          "status": "error", "message": "Invalid row number"
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return updateStatus(rowNum, payload.status);
    }

    // --- MASTER INVENTORY ROW REFRESH (called by n8n eBay-orders workflow per batch) ---
    // Updates MI rows for specific itemIds with fresh qty/quantitySold/qtyLastSync,
    // BEFORE the doPost insert reads MI for HAND computation. Solves the staleness
    // gap between an eBay sale (eBay auto-decrements its qty) and the next bulk
    // sync. Only refreshes rows that ALREADY exist in MI — silently skips unknowns
    // so a malformed payload can't pollute the sheet.
    if (payload.action === 'updateMiRows') {
      try {
        var miInputRows = Array.isArray(payload.rows) ? payload.rows : [];
        var miUpdateResult = updateMiRows(miInputRows);
        return ContentService.createTextOutput(JSON.stringify({
          status:   "success",
          updated:  miUpdateResult.updated,
          notFound: miUpdateResult.notFound,
          rows:     miInputRows.length
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (miErr) {
        Logger.log("updateMiRows action error: " + miErr.toString());
        return ContentService.createTextOutput(JSON.stringify({
          status:  "error",
          message: "updateMiRows failed: " + miErr.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // --- HAND RECOMPUTE (called by n8n Inventory Lite Sync after upserting changes) ---
    // Fires within ~1s of MI being updated, so All Orders HAND values reflect the
    // new inventory immediately instead of waiting for the 15-min time-based trigger.
    if (payload.action === 'recomputeHand') {
      try {
        var handResult = recomputeHand();
        return ContentService.createTextOutput(JSON.stringify({
          status: "success",
          message: "HAND recomputed",
          result: handResult || null
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (handErr) {
        Logger.log("recomputeHand action error: " + handErr.toString());
        return ContentService.createTextOutput(JSON.stringify({
          status: "error",
          message: "recomputeHand failed: " + handErr.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // --- ZOHO STOCK PUSH (called by the n8n "Zoho Stock Sync (Scheduled)" workflow) ---
    // n8n owns the schedule AND the ~60-90s Zoho pagination, then PUSHES the items
    // array here. Apps Script only writes the Zoho Stock sheet + rewrites HAND on
    // All Orders + Prep Queue (~5s) — so this can fire every 2 min without burning
    // the Apps Script daily-runtime quota the way an Apps-Script-side pull would.
    if (payload.action === 'writeZohoStock') {
      try {
        var zsItems = Array.isArray(payload.items) ? payload.items : [];
        var zsWrite = _writeZohoStockRows(zsItems);
        // ⚠ ONE read of Master Inventory + Zoho Stock, handed to BOTH recomputes.
        // They used to build their own copies, so this path read the same two
        // sheets twice every 2 minutes: 2,287 ms of a 5,789 ms cycle, ~9 min/day
        // of the ~90-minute budget spent re-reading what it already had.
        // Measured 2026-08-19 — see recomputeHand's note, and note that
        // narrowing the read instead was tested and does NOT help.
        var zsHand = "", zsPrep = "", zsMaps = null, zsZoho = null;
        try { zsMaps = buildLocationAndInventoryMaps(); zsZoho = buildZohoStockMap(); }
        catch (e0) { zsMaps = null; zsZoho = null; }   // fall back: each rebuilds its own
        try { zsHand = recomputeHand(zsMaps, zsZoho); }        catch (e1) { zsHand = "HAND error: " + e1; }
        try { zsPrep = refreshPrepQueueHand(zsMaps, zsZoho); } catch (e2) { zsPrep = "Prep error: " + e2; }
        return ContentService.createTextOutput(JSON.stringify({
          status:      "success",
          written:     zsWrite.count,
          skipped:     zsWrite.skipped,
          handMessage: zsHand,
          prepMessage: zsPrep
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (zsErr) {
        Logger.log("writeZohoStock action error: " + zsErr.toString());
        return ContentService.createTextOutput(JSON.stringify({
          status:  "error",
          message: "writeZohoStock failed: " + zsErr.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // --- PRICE AUDIT (scheduled run by n8n cron) ---
    // Fires from an n8n Schedule Trigger (twice daily). Same orchestration as
    // the sidebar's "Run Audit" button — internally calls n8n's Zoho bulk-items
    // proxy to fetch active Zoho items, joins to MI's active SKUs via the
    // listingStatus filter, writes the Price Audit sheet with three-bucket
    // direction (drift / INACTIVE CANDIDATE / OOS).
    //
    // Returns the full runPriceAudit() result object so n8n can branch on
    // `ok === false` to fire a Telegram alert via the Global Error Handler.
    // Also lets n8n surface the daily INACTIVE CANDIDATE count to admin if
    // wanted (e.g., "audit ran clean — 3 drifts, 7 INACTIVE CANDIDATES today").
    if (payload.action === 'runPriceAudit') {
      try {
        var auditResult = runPriceAudit();
        return ContentService.createTextOutput(JSON.stringify(auditResult))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (auditErr) {
        Logger.log("runPriceAudit action error: " + auditErr.toString());
        return ContentService.createTextOutput(JSON.stringify({
          ok: false,
          message: "runPriceAudit failed: " + auditErr.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // --- ZOHO ITEM WEBHOOK: KIT REGISTRY REFRESH ---
    // Fired by Zoho Workflow Rule when an Item's Purchase Description is updated.
    // Zoho sends its "Default Payload" shape:
    //     { "item": { "item_id": "...", "sku": "...", "name": "...",
    //                 "purchase_description": "...", ... } }
    // We re-parse this single kit using the same regex logic as the bulk CSV
    // importer, and upsert its rows in the Kit Registry sheet.
    //
    // SEMANTICS
    //   - Item is a kit + has parseable PD     → registry rows added/updated
    //   - Item name no longer matches pattern  → existing rows REMOVED
    //   - PD cleared / no parseable lines      → existing rows REMOVED
    //   - TEMP-* SKU                           → skipped (cleanup if previously had rows)
    //
    // Keeps the registry in sync with Zoho's current state — kit removals
    // propagate, not just additions/updates.
    if (payload.action === 'zohoKitUpdated') {
      try {
        var zohoItem = payload.item || {};
        var zohoResult = refreshOneKitFromZohoPayload(zohoItem);

        // A Zoho edit produced PD line(s) the parser couldn't read → ping the
        // admin chat so the editor can fix the line while still in the item.
        // Best-effort: the alert never fails the webhook response.
        if (zohoResult.unparsedLines && zohoResult.unparsedLines.length > 0) {
          try {
            _sendKitParseAlert(zohoResult.kitSku, zohoResult.kitName || "",
                               zohoResult.unparsedLines);
          } catch (alertErr) {
            console.log("zohoKitUpdated: parse alert error: " + alertErr);
          }
        }

        return ContentService.createTextOutput(JSON.stringify({
          status:            "success",
          kitSku:            zohoResult.kitSku || "",
          actionTaken:       zohoResult.actionTaken || "none",
          componentsWritten: zohoResult.componentsWritten || 0,
          unparsedLines:     (zohoResult.unparsedLines || []).length,
          reason:            zohoResult.reason || ""
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (zohoErr) {
        Logger.log("zohoKitUpdated error: " + zohoErr.toString());
        return ContentService.createTextOutput(JSON.stringify({
          status:  "error",
          message: "zohoKitUpdated failed: " + zohoErr.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // --- ZOHO SALES ORDER + INVOICE WEBHOOK: PENDING SHEET MIRROR ---
    // Fired by TWO Zoho Workflow Rules pointed at the same n8n proxy:
    //   (1) Sales Order module (Created + Edited) → payload wrapped as { "salesorder": {...} }
    //   (2) Invoice module     (Created + Edited) → payload wrapped as { "invoice":    {...} }
    // The n8n proxy hardcodes action=zohoSalesOrder for both; we sniff by
    // wrapper shape here to route to the right handler.
    //
    // SO handler: upserts a row in "Pending Sales Orders" (current snapshot of
    // Zoho state). If the SO was already pulled to DIRECT, applies asymmetric
    // propagation rule:
    //   - new line items     → auto-insert as new DIRECT rows
    //   - removed line items → FLAG existing DIRECT rows (no auto-delete)
    //   - qty changes        → FLAG (no in-place mutation)
    //   - status: void       → flip DIRECT rows to CANCELED
    //   - shipment status    → ignored (employee handles manually)
    //
    // Invoice handler: stamps the invoice_number onto the matching Pending row's
    // INVOICE column (linked via salesorder_id, with reference_number fallback).
    // Enables the sidebar to look up an SO by invoice number (e.g. INV-022496)
    // for customer-service flows.
    //
    // Webhook-level filter (both handlers): sales_channel === "direct_sales".
    if (payload.action === 'zohoSalesOrder') {
      // Sniff: invoice wrapper present → invoice handler. Otherwise (SO wrapper
      // or unwrapped payload) → existing SO handler.
      if (payload.invoice && typeof payload.invoice === 'object') {
        try {
          var invoiceResult = upsertInvoiceFromZoho(payload.invoice);
          return ContentService.createTextOutput(JSON.stringify({
            status:        invoiceResult.status || "unknown",
            invoiceNumber: invoiceResult.invoiceNumber || "",
            soNumber:      invoiceResult.soNumber || "",
            matchedBy:     invoiceResult.matchedBy || "",
            actionTaken:   invoiceResult.actionTaken || "",
            reason:        invoiceResult.reason || ""
          })).setMimeType(ContentService.MimeType.JSON);
        } catch (invErr) {
          Logger.log("zohoInvoice error: " + invErr.toString());
          return ContentService.createTextOutput(JSON.stringify({
            status:  "error",
            message: "zohoInvoice failed: " + invErr.toString()
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }
      try {
        var salesorderPayload = payload.salesorder || payload;
        var soResult = upsertPendingSalesOrder(salesorderPayload);
        return ContentService.createTextOutput(JSON.stringify({
          status:      soResult.status || "unknown",
          soNumber:    soResult.soNumber || "",
          actionTaken: soResult.actionTaken || "",
          wasPulled:   soResult.wasPulled || false,
          propagated:  soResult.propagated || null,
          reason:      soResult.reason || ""
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (soErr) {
        Logger.log("zohoSalesOrder error: " + soErr.toString());
        return ContentService.createTextOutput(JSON.stringify({
          status:  "error",
          message: "zohoSalesOrder failed: " + soErr.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ZOHO BACKFILL — called by the Zoho SO Backfill Proxy n8n workflow.
    // Same shape as zohoSalesOrder (payload carries {salesorder: <full SO>})
    // but routes through the on-demand sidebar Fetch button, NOT the webhook.
    // Single source of truth: just delegates to upsertPendingSalesOrder() so
    // the Pending row gets created/updated using the EXACT same logic as the
    // automatic webhook path. Adding any new field handling to one path
    // automatically benefits the other.
    //
    // Optional extra fields the proxy passes:
    //   matched_via            — "invoice" | "salesorder" (which Zoho endpoint
    //                            resolved the query). Surfaced in the response
    //                            so the sidebar can confirm to the picker how
    //                            we found their query.
    //   matched_invoice_number — invoice # if the picker typed an INV-* lookup
    //                            (sidebar can show "Found via invoice X").
    if (payload.action === 'zohoBackfillSalesOrder') {
      try {
        var backfillPayload = payload.salesorder || payload;
        var backfillResult  = upsertPendingSalesOrder(backfillPayload);

        // STAMP INVOICE COLUMN — when backfill was triggered via INV-* lookup,
        // n8n forwards `matched_invoice_number` alongside the SO payload. The
        // SO upsert handler doesn't touch the INVOICE column (that's normally
        // populated by the separate Invoice webhook handler). For backfill we
        // know the invoice number AND the SO that was just upserted, so we can
        // stamp it directly — otherwise the Pending row would have empty
        // INVOICE col and the picker couldn't later Pull by INV# (Gotcha
        // identified during first INV-path test 2026-05-20).
        var invoiceStamped = false;
        var matchedInvoice = String(payload.matched_invoice_number || "").trim();
        if (matchedInvoice && backfillResult.soNumber) {
          try {
            var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
            var sheet = ss.getSheetByName(PENDING_SO.sheetName);
            if (sheet) {
              var rowNum = _findPendingRow(sheet, backfillResult.soNumber);
              if (rowNum > 0) {
                sheet.getRange(rowNum, PENDING_SO.cols.INVOICE).setValue(matchedInvoice);
                invoiceStamped = true;
              }
            }
          } catch (stampErr) {
            // Don't fail the whole backfill if invoice stamping fails —
            // SO row is already inserted, just log and continue.
            Logger.log("backfill invoice stamp failed: " + stampErr.toString());
          }
        }

        return ContentService.createTextOutput(JSON.stringify({
          status:               backfillResult.status || "unknown",
          soNumber:             backfillResult.soNumber || "",
          actionTaken:          backfillResult.actionTaken || "",
          wasPulled:            backfillResult.wasPulled || false,
          matchedVia:           payload.matched_via || "",
          matchedInvoiceNumber: matchedInvoice,
          invoiceStamped:       invoiceStamped,
          source:               "backfill",
          reason:               backfillResult.reason || ""
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (bfErr) {
        Logger.log("zohoBackfillSalesOrder error: " + bfErr.toString());
        return ContentService.createTextOutput(JSON.stringify({
          status:  "error",
          message: "zohoBackfillSalesOrder failed: " + bfErr.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // --- IMPROVED BATCH ORDER INSERTION ---
    // --- UNKNOWN ACTION: fail LOUDLY ------------------------------------------
    // Anything with an `action` we don't recognise used to fall through to the
    // orders-insert branch below and, with no `orders` array, return
    // {status:"success", added:0}. That is a SILENT failure, and it bit us on
    // 2026-08-05: the newly hosted Floor Board called `boardTick` against an
    // /exec that predated the action, received a success-shaped object with no
    // cockpit in it, and painted all-zeros while reporting LIVE. Nothing looked
    // broken at any layer — Caddy fine, n8n execution green, Apps Script "ok".
    //
    // A caller that names an action it expects to exist deserves an error when
    // it doesn't.
    //
    // ⚠ 2026-08-06 — THIS GUARD BROKE ORDER INSERTS FOR A DAY. Read before
    // touching it. The original version of this comment asserted "the n8n
    // order-insert path sends `orders` with NO action, so it is unaffected."
    // That was WRONG and never verified against the workflow: node
    // "13. Insert Rows via App Script" sends `action: 'insertOrders'`. The
    // guard therefore rejected every batch of new orders.
    //
    // It stayed invisible for the usual two reasons stacked:
    //   1. Apps Script answers HTTP 200 even when the BODY says error, and
    //      node 13 runs onError:continueRegularOutput — so n8n went green.
    //   2. The guard only reached /exec when a New Version was cut (Gotcha
    //      #12), so "worked yesterday, broke today" with no code change in
    //      between.
    //
    // LESSON: a guard that rejects unknown callers must be built from the
    // ACTUAL list of callers, read out of the workflow JSON — not from an
    // assumption about them. The four actions the workflows really send are
    // insertOrders / updateMiRows / updateOrderStatus / writeZohoStock.
    var INSERT_ACTION = 'insertOrders';   // what n8n node 13 actually sends
    if (payload.action && payload.action !== INSERT_ACTION) {
      console.log("doPost: unknown action '" + payload.action + "'");
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Unknown action: " + payload.action +
                 " — this deployment may predate it; cut a New Version."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var orders = payload.orders || [];
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);

    // 1. Gather Existing Orders to prevent duplicates
    // Scan ALL data rows (SKU + SALES_ORDER are enough to dedup)
    var lastRow = sheet.getLastRow();
    var scanRows = Math.max(0, lastRow - Schema.dataStartRow + 1);
    if (scanRows === 0) scanRows = 1;
    var existingData = sheet.getRange(Schema.dataStartRow, 1, scanRows, Schema.cols.SALES_ORDER).getValues();
    var existingSignatures = new Set();
    existingData.forEach(function(row) {
      var so  = row[Schema.idx("SALES_ORDER")];
      var sku = row[Schema.idx("SKU")];
      if (so && sku) existingSignatures.add(String(so).trim() + "|" + String(sku).trim().toUpperCase());
    });

    var newRows = [];
    var results = [];
    var activityLogBatch = [];   // collected during the orders loop, written after successful insert

    // 2. Build location & inventory maps ONCE (not per-item)
    // This reads Master Inventory LIVE - fresh data every doPost call
    // but avoids re-reading the entire sheet for EACH item
    var maps = buildLocationAndInventoryMaps();
    var locationMap = maps.locationMap;
    var inventoryMap = maps.inventoryMap;

    // Track available stock decrements within this batch
    var batchStock = {};

    // Build committed orders map to subtract existing PENDING/PREPARING from available
    var committedMap = getCommittedQuantities();

    // 3. Build the data array first (Don't touch the sheet yet)
    orders.forEach(function(item) {
      var sku = String(item.SKU || "").trim().toUpperCase();
      var skuLower = sku.toLowerCase();
      var salesOrder = String(item["SALES ORDER"] || "").trim();

      if (!salesOrder || salesOrder.length < 3) {
        results.push("Skipped: Invalid ID");
        return;
      }

      // Duplicate Check
      if (existingSignatures.has(salesOrder + "|" + sku)) {
        results.push("Skipped: Duplicate " + salesOrder);
        return;
      }

      // LIVE location lookup from map (built fresh this request)
      var location = locationMap.get(skuLower) || "NOT FOUND";

      // HAND = MI.available for every non-terminal row of this SKU. Same value
      // across rows; no per-row or per-batch decrement.
      //
      // Why no decrement: with the per-order GetItem refresh (n8n nodes 8.5/8.6),
      // MI is fresh at the moment of insert — eBay's QuantitySold already counts
      // every PENDING/PREPARING order's qty (the sale was registered the instant
      // the buyer hit Buy). Subtracting `alreadyCommitted` would re-subtract the
      // same units that QuantitySold already excluded. Decrementing per-row in
      // the same batch (`-= itemQty`) does the same thing for batch-mate rows
      // that eBay's QuantitySold already accounts for. Both = double-counting.
      //
      // Operational meaning: HAND tells the picker "what eBay's listing
      // currently shows as available for new buyers." Physical stock in the
      // warehouse is HAND + sum of all open committed qty for this SKU, but
      // for "is there enough to ship this order?" HAND alone is the right
      // number.
      if (!(skuLower in batchStock)) {
        var invData = inventoryMap.get(skuLower);
        batchStock[skuLower] = invData ? invData.available : 0;
      }
      var itemQty = parseInt(item.QTY) || 1;
      var handValue = batchStock[skuLower];

      // Note from eBay's BuyerCheckoutMessage — prefix it so it visually
      // separates from supervisor notes added by hand later. The supervisor
      // can prefix their own with "Supervisor:" or leave plain — that's their
      // call. We only auto-prefix the auto-source (n8n).
      var rawNote = String(item.NOTE || "").trim();
      var noteForCell = rawNote ? ("Buyer Note: " + rawNote) : "";

      // SHIPPING service + SHIP COST come from n8n now (added 2026-05-01).
      // Defensive: accept multiple field-name variants so this works regardless
      // of how the workflow author spelled them. First non-empty wins.
      var shipping = String(
        item.SHIPPING ||
        item["SHIPPING SERVICE"] ||
        item.SHIPPING_SERVICE ||
        item.shippingService ||
        ""
      ).trim();
      var shipCost = (
        item["SHIP COST"] !== undefined ? item["SHIP COST"] :
        item.SHIP_COST   !== undefined ? item.SHIP_COST   :
        item.SHIPPING_COST !== undefined ? item.SHIPPING_COST :
        item["SHIPPING COST"] !== undefined ? item["SHIPPING COST"] :
        item.shippingCost !== undefined ? item.shippingCost :
        ""
      );

      // Store row data: [SKU, Qty, Loc, OrderID, Note, Status, Hand, Left, Shipping, ShipCost]
      // Full schema width = 10 (Schema.dataWidth). LEFT (col H) stays blank
      // for the picker to fill in manually after counting at the shelf.
      newRows.push([
        sku,
        item.QTY || 1,
        location,
        salesOrder,
        noteForCell,
        Schema.status.PENDING,
        handValue,
        "",            // LEFT — picker fills this in after counting
        shipping,      // SHIPPING service
        shipCost       // SHIP COST
      ]);

      // No batchStock decrement: HAND is MI.available for every row of this
      // SKU. eBay's QuantitySold already accounts for every committed qty.

      existingSignatures.add(salesOrder + "|" + sku);
      results.push("Added: " + salesOrder);

      // Stage Activity Log entry — written after successful insert.
      // Slots: [event, orderId, sku, qty, source, detail, picker, note]
      // Picker is "" (n8n is automation, no warehouse staff involved).
      // Note carries the same "Buyer Note: ..." prefix we wrote to the cell.
      activityLogBatch.push([
        "RECEIVED",
        salesOrder,
        sku,
        itemQty,
        "n8n",
        location ? "Loc: " + location : "",
        "",                // picker (n8n source — no human)
        noteForCell        // note (Buyer Note: ..., or empty)
      ]);
    });

    // 3. Insert and Format in ONE GO (Much Faster)
    // Width = 10 (Schema.dataWidth) — n8n now sends SHIPPING + SHIP_COST as
    // of 2026-05-01, so we write the full row width. LEFT (col H) still stays
    // blank for the picker to fill in after counting at the shelf.
    if (newRows.length > 0) {
      var INSERT_WIDTH = Schema.dataWidth;  // = 10

      // A. Save headers before insertion (guards against Google Sheets filter corruption bug)
      var savedHeaders = sheet.getRange(Schema.headerRow, 1, 1, INSERT_WIDTH).getValues()[0];

      // B. Insert blank rows at the top
      sheet.insertRowsBefore(Schema.dataStartRow, newRows.length);

      // C. Get the target range
      var range = sheet.getRange(Schema.dataStartRow, 1, newRows.length, INSERT_WIDTH);

      // D. Paste Data
      range.setValues(newRows);

      // E. Restore headers if Google Sheets corrupted them during insertion
      verifyAndRestoreHeaders(sheet, savedHeaders);

      // F. Clean Formatting (Fixes "Format Persistence")
      // We copy format from the row *below* the insertion to ensure borders/fonts match
      var templateRow = Schema.dataStartRow + newRows.length;
      sheet.getRange(templateRow, 1, 1, INSERT_WIDTH).copyFormatToRange(
        sheet, 1, INSERT_WIDTH,
        Schema.dataStartRow, Schema.dataStartRow + newRows.length - 1
      );

      // F2. Neutralize inherited SO-badge formats (2026-07-14). Step F copies
      // EVERY format from the template row — including the col-D badge number
      // format ('"1️⃣ "@') whenever the template row belongs to a multi-item
      // group, which put a false badge on every row of two consecutive syncs.
      // New rows must start badge-free NO MATTER WHAT; the painter call below
      // re-derives real badges afterwards. This makes inserts immune even if
      // that repaint fails — same fix class as the setBackground(null)
      // anti-bleed on Zoho inserts (2026-05-23).
      sheet.getRange(Schema.dataStartRow, Schema.cols.SALES_ORDER, newRows.length, 1)
        .setNumberFormat('@');

      updateOrderStatsInSheet();
      updateLastSyncTimestamp();

      // Activity Log: one RECEIVED event per inserted row (best-effort; if the
      // log sheet doesn't exist yet, the call returns silently).
      try { logActivityBatch(activityLogBatch); } catch (logErr) {
        console.log("doPost: activity log error: " + logErr);
      }

      // New orders just landed — drop the Floor Board's cached tick so the next
      // poll shows them instead of waiting out the cache window. This is the
      // whole point of the board: a picker seeing an arrival promptly.
      try { _dashBustTickCache(); } catch (e) {
        console.log("doPost: tick cache bust failed: " + e);
      }

      // Kit SKU markers — apply the ▣ glyph to any newly-inserted row whose
      // SKU is in the Kit Registry. Programmatic insert + setValues doesn't
      // fire onEdit, so kitSkuOnEdit (the user-edit handler) never runs for
      // these rows — explicit batched refresh fills the gap. Also corrects
      // any wrong markers that copyFormatToRange (step F above) may have
      // inherited from the template row.
      try { refreshKitSkuMarkers(); } catch (kitErr) {
        console.log("doPost: kit marker refresh error: " + kitErr);
      }

      // SKU links + order links — programmatic setValues doesn't fire onEdit,
      // so enrich the newly-inserted rows here (col A → listing, col D → order).
      try { refreshAllOrdersEnrichment(); } catch (enrErr) {
        console.log("doPost: enrichment refresh error: " + enrErr);
      }

      // SO badges + duplicate grouping — CRITICAL after inserts (added
      // 2026-07-14 after the incident where every synced batch wore a wrong
      // badge): insertRowsBefore + copyFormatToRange makes new rows INHERIT
      // the col-D badge number format from the template row, and nothing
      // else repaints on this path (onEdit doesn't fire for setValues).
      // Without this call, every n8n sync mints false badges that persist
      // until some editor-side repaint.
      try { setupDuplicateSalesOrderHighlighting(); } catch (dupErr) {
        console.log("doPost: dup-SO badge refresh error: " + dupErr);
      }

      // PUBLISH THE TICK NOW — do not wait for the 5-minute trigger.
      //
      // A new order landing is the one change the board exists to show, and the
      // only one where minutes of lag is actually felt on the floor. Everything
      // else the timer covers fine.
      //
      // Inline is affordable HERE and nowhere else: this path is machine-facing
      // (n8n is waiting, no human is), so the ~4s costs nobody anything. The
      // same call on updateOrderStatus would add those seconds to every Telegram
      // PREP tap, which a human very much does feel — which is why the dirty
      // flag exists for that path instead.
      //
      // Cost scales with real order volume, not with viewers: ~50-100 arrivals a
      // day, not 1,440 timer runs.
      //
      // LAST in the block on purpose — after enrichment and the badge repaint —
      // so the published copy reflects fully-settled rows. Best-effort: a failed
      // publish must never fail an insert, and the trigger will catch it anyway
      // because _dashBustTickCache already marked it dirty.
      try {
        if (typeof publishBoardTick === 'function') publishBoardTick('arrival');
      } catch (pubErr) {
        console.log("doPost: tick publish failed (trigger will retry): " + pubErr);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      "status": "success",
      "added": newRows.length,
      "details": results
    })).setMimeType(ContentService.MimeType.JSON);

      } catch (err) {
        lastErr = err;
        if (attempt < MAX_RETRIES) {
          Utilities.sleep(1000 * attempt);
          continue;
        }
      }
    }
    // All retries exhausted
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": lastErr.toString()})).setMimeType(ContentService.MimeType.JSON);

  } finally {
    if (lock) lock.releaseLock();   // null for the lock-free read actions above
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// 📊 INVENTORY LOOKUP - Single definition is at the bottom of this file
// ═══════════════════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════════════════
// ✨ HIDDEN SHEET MANAGEMENT - Store Message IDs
// ═══════════════════════════════════════════════════════════════════════════════════════

/**
 * Get or create the hidden sheet for storing message IDs
 */
function getHiddenSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  
  if (!sheet) {
    // Create the sheet if it doesn't exist
    sheet = ss.insertSheet(HIDDEN_SHEET_NAME);
    
    // Set headers
    sheet.getRange("A1:D1").setValues([["Order ID", "Message ID", "Chat ID", "Timestamp"]]);
    
    // Format headers
    sheet.getRange("A1:D1")
      .setFontWeight("bold")
      .setBackground("#4285F4")
      .setFontColor("#FFFFFF");
    
    // Set column widths
    sheet.setColumnWidth(1, 150); // Order ID
    sheet.setColumnWidth(2, 120); // Message ID
    sheet.setColumnWidth(3, 150); // Chat ID
    sheet.setColumnWidth(4, 180); // Timestamp
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Hide the sheet
    sheet.hideSheet();
    
    console.log("Created hidden sheet: " + HIDDEN_SHEET_NAME);
  }
  
  return sheet;
}

// storeMessageId() and getMessageId() - single definitions are at the bottom of this file

/**
 * Delete message ID entry after order is deleted or completed
 */
function deleteMessageIdEntry(orderId) {
  try {
    var sheet = getHiddenSheet();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(orderId).trim()) {
        sheet.deleteRow(i + 1);
        console.log("Deleted message ID entry for order: " + orderId);
        return true;
      }
    }
    
    return false;
    
  } catch (e) {
    console.log("deleteMessageIdEntry error for " + orderId + ": " + e);
    return false;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// ✨ AUTO-CLEANUP: Remove old message IDs (older than 7 days)
// ═══════════════════════════════════════════════════════════════════════════════════════

/**
 * Clean up message IDs older than specified days
 * Run this weekly via Apps Script trigger
 */
function cleanupOldMessageIds(daysToKeep) {
  try {
    daysToKeep = daysToKeep || 7; // Default 7 days
    
    var sheet = getHiddenSheet();
    var data = sheet.getDataRange().getValues();
    var today = new Date();
    var deletedCount = 0;

    // Loop backwards to avoid row index shifting
    for (var i = data.length - 1; i > 0; i--) {
      var timestamp = new Date(data[i][3]);
      var ageInDays = (today - timestamp) / (1000 * 60 * 60 * 24);

      if (ageInDays > daysToKeep) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }

    // Single audit-worthy log line for the weekly cleanup run.
    logDebug("cleanupOldMessageIds: deleted " + deletedCount + " entries older than " + daysToKeep + " days");
    
    return {
      success: true,
      deletedCount: deletedCount,
      daysToKeep: daysToKeep
    };
    
  } catch (e) {
    console.log("cleanupOldMessageIds error: " + e);
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * Setup weekly cleanup trigger (run this once manually)
 */
function setupWeeklyCleanup() {
  // Delete existing triggers first
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'cleanupOldMessageIds') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new weekly trigger (runs every Monday at 2 AM)
  ScriptApp.newTrigger('cleanupOldMessageIds')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(2)
    .create();
  
  Logger.log("✅ Weekly cleanup trigger created - runs every Monday at 2 AM");
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// ✨ NOTIFY TELEGRAM WHEN SHIPPED - Uses Stored Message ID
// ═══════════════════════════════════════════════════════════════════════════════════════

/**
 * Called by n8n when an order is marked SHIPPED
 * Updates the existing Telegram message using stored message_id
 */
function notifyTelegramShipped(orderId) {
  // 1. Retrieve the Chat ID and Message ID from the hidden sheet
  var msgData = getMessageId(orderId);

  if (!msgData) {
    console.log("⚠️ Order " + orderId + " not found in Telegram_Messages sheet.");
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "skipped", 
      reason: "Message ID not found for Order " + orderId 
    }));
  }

  // 2. Define the new "Shipped" message text
  var newText = "📦 <b>Order " + orderId + "</b>\n" +
                "────────────────────\n" +
                "✅ <b>STATUS: SHIPPED</b>\n\n" +
                "<i>This order has been processed and shipped.</i>";

  // 3. Send the edit request to Telegram (CORRECTED ENDPOINT)
  var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/editMessageText";

  var payload = {
    'chat_id': String(msgData.chatId),      // Ensure it's a string
    'message_id': parseInt(msgData.messageId), // Must be integer for Telegram API
    'text': newText,
    'parse_mode': 'HTML',
    'reply_markup': { 'inline_keyboard': [] } // Removes the buttons
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true // Prevents crash so we can read the error
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseBody = JSON.parse(response.getContentText());

    if (responseBody.ok) {
      return ContentService.createTextOutput(JSON.stringify({ status: "success", action: "updated_telegram" }));
    } else {
      // Log the exact error from Telegram
      console.error("Telegram Error for Order " + orderId + ": " + responseBody.description);
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        telegram_error: responseBody.description 
      }));
    }

  } catch (e) {
    console.error("Script Error: " + e.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "error", error: e.toString() }));
  }
}

/**
 * Update Telegram message to show SHIPPED status with no buttons
 */
function updateTelegramMessageToShipped(chatId, messageId, orderId) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/editMessageText";
  
  try {
    // Build new message text for shipped status
    var newText = buildShippedMessageText(orderId);
    
    // Build payload with NO buttons
    var payload = {
      "chat_id": chatId,
      "message_id": messageId,
      "text": newText,
      "parse_mode": "HTML",
      "reply_markup": {
        "inline_keyboard": []  // Empty = no buttons
      }
    };
    
    var response = UrlFetchApp.fetch(url, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });

    var result = JSON.parse(response.getContentText());

    if (result.ok) {
      return true;
    } else {
      console.log("notifyTelegramShipped failed: " + result.description);
      return false;
    }

  } catch (e) {
    console.log("notifyTelegramShipped error: " + e);
    return false;
  }
}

/**
 * Build the SHIPPED message text
 * Preserves original order details, just updates status
 */
function buildShippedMessageText(orderId) {
  // This is a simplified version - you may want to fetch full order details from sheet
  var text = "📦 <b>Order Complete</b>\n\n";
  text += "Order: <code>" + orderId + "</code>\n\n";
  text += "━━━━━━━━━━━━━━━\n";
  text += "📋 Status: ✅ <b>SHIPPED - Order Complete!</b>\n";
  text += "🟢 Ready for carrier pickup";
  
  return text;
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// TELEGRAM INTERACTION HANDLER - Updated for SHIPPED
// ═══════════════════════════════════════════════════════════════════════════════════════

function handleTelegramCallback(payload) {
  var callback = payload.callback_query;
  var data = callback.data;
  var chatId = callback.message.chat.id;
  var messageId = callback.message.message_id;
  var originalText = callback.message.text || "";
  
  var action = "";
  var orderId = "";

  if (data.startsWith("PREP_")) {
    action = Schema.status.PREPARING;
    orderId = data.replace("PREP_", "");
  } else if (data.startsWith("PEND_")) {
    action = Schema.status.PENDING;
    orderId = data.replace("PEND_", "");
  } else {
    answerCallbackQuery(callback.id, "❓ Unknown action", true);
    return ContentService.createTextOutput("OK");
  }

  // 1. UPDATE THE SHEET
  var result = findAndUpdateOrder(orderId, action);

  if (!result.found) {
    answerCallbackQuery(callback.id, "⚠️ Order not found", true);
    return ContentService.createTextOutput("OK");
  }

  // ✨ FIX: Check if order was already SHIPPED
  if (result.currentStatus === Schema.status.SHIPPED) {
    answerCallbackQuery(callback.id, "✅ Order already shipped!", true);
    updateMessageStatus(chatId, messageId, originalText, orderId, Schema.status.SHIPPED);
    return ContentService.createTextOutput("OK");
  }

  // 2. SHOW TOAST
  var toastText = action === Schema.status.PREPARING ? "✅ Preparing" : "🔄 Pending";
  answerCallbackQuery(callback.id, toastText, false);

  // 3. EDIT MESSAGE
  updateMessageStatus(chatId, messageId, originalText, orderId, action);

  return ContentService.createTextOutput("OK");
}

/**
 * Back-compat wrapper around updateOrderStatus().
 * Preserves the original return shape { found, count, currentStatus } that
 * doPost (action: updateOrderStatus) and handleTelegramCallback expect.
 *
 * IMPORTANT: syncTelegram is FALSE here on purpose.
 * Original behavior: findAndUpdateOrder only touched the sheet. The two callers
 * each handle Telegram their own way:
 *   - handleTelegramCallback → calls updateMessageStatus afterward (simpler
 *     "Status: emoji STATUS" append, NOT a full rich-message re-render)
 *   - doPost action=updateOrderStatus → n8n manages its own Telegram flow
 * Letting updateOrderStatus sync would double-edit the Telegram message and
 * stomp on the simpler format the callback handler produces.
 *
 * For new code, call updateOrderStatus(orderId, newStatus, options) directly.
 */
function findAndUpdateOrder(orderId, newStatus) {
  var result = updateOrderStatus(orderId, newStatus, {
    source:       "findAndUpdateOrder",
    syncTelegram: false   // see comment above — caller handles Telegram
  });

  return {
    found: result.success && (result.count > 0 || result.blockedCount > 0),
    count: result.count || 0,
    currentStatus: result.currentStatus || ""
  };
}

function answerCallbackQuery(callbackQueryId, text, showAlert) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/answerCallbackQuery";
  var payload = {
    "callback_query_id": callbackQueryId,
    "text": text,
    "show_alert": showAlert || false
  };
  
  try {
    UrlFetchApp.fetch(url, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });
  } catch (e) {
    console.log("answerCallbackQuery error: " + e);
  }
}

function updateMessageStatus(chatId, messageId, originalText, orderId, newStatus) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/editMessageText";
  
  // Remove old status line if exists
  var cleanText = originalText.replace(/\n\n📋 Status:.*$/s, "");
  
  // Different formatting for each status
  var statusEmoji = "";
  var statusText = "";

  if (newStatus === Schema.status.SHIPPED) {
    statusEmoji = "✅";
    statusText = "SHIPPED - Order Complete!";
  } else if (newStatus === Schema.status.CANCELED) {
    statusEmoji = "❌";
    statusText = "CANCELED - Order Cancelled";
  } else if (newStatus === Schema.status.PREPARING) {
    statusEmoji = "🟡";
    statusText = Schema.status.PREPARING;
  } else {
    statusEmoji = "🔴";
    statusText = Schema.status.PENDING;
  }

  var newText = cleanText + "\n\n📋 Status: " + statusEmoji + " " + statusText;

  // No buttons for terminal states (SHIPPED / CANCELED)
  var keyboard = null;

  if (Schema.isTerminal(newStatus)) {
    keyboard = { "inline_keyboard": [] };
  } else if (newStatus === Schema.status.PENDING) {
    keyboard = {
      "inline_keyboard": [
        [{ "text": "🚀 Mark as Preparing", "callback_data": "PREP_" + orderId }]
      ]
    };
  } else if (newStatus === Schema.status.PREPARING) {
    keyboard = {
      "inline_keyboard": [
        [{ "text": "🔄 Revert to Pending", "callback_data": "PEND_" + orderId }]
      ]
    };
  }
  
  var payload = {
    "chat_id": chatId,
    "message_id": messageId,
    "text": newText,
    "reply_markup": keyboard
  };
  
  try {
    var response = UrlFetchApp.fetch(url, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });

    var result = JSON.parse(response.getContentText());

    if (result.ok) {
      return "SUCCESS";
    } else {
      console.log("updateMessageStatus failed: " + result.description);
      return "FAILED: " + result.description;
    }

  } catch (e) {
    console.log("updateMessageStatus error: " + e);
    return "ERROR: " + e.toString();
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// MANUAL DIRECT-TABLE RECEIVE HOOK (Activity Log)
// ═══════════════════════════════════════════════════════════════════════════════════════

/**
 * onEdit handler for MANUAL order entry in the All Orders sheet.
 *
 * Fires for the two human-entry paths:
 *   - DIRECT-table sales orders: typed by hand below the boundary divider.
 *   - eBay-table replacement orders: occasionally typed by hand above the
 *     boundary (e.g., a replacement we're sending for a damaged item that
 *     never went through eBay's normal flow).
 *
 * The doPost RECEIVED hook only fires for n8n-pushed eBay orders. n8n's
 * inserts use programmatic setValues and don't trigger onEdit, so this hook
 * captures only true human entries — no double-logging risk.
 *
 * Trigger condition:
 *   - Edit is on the All Orders sheet
 *   - SKU (col A) OR SALES_ORDER (col D) intersects the edited range
 *   - Row is in a data area (≥ dataStartRow, not the divider or DIRECT header)
 *   - The row's identity is COMPLETE — both SKU and SALES_ORDER non-empty
 *
 * ⚠⚠ IT LOGS ONE OF TWO EVENTS, AND THE DIFFERENCE IS LOAD-BEARING (2026-08-30).
 *   `_mrClassify` decides:
 *
 *     RECEIVED       the identity just became complete, or the new value EXTENDS
 *                    the old one ("Replacement for #:" → "Replacement for #: 26-…").
 *                    An order arrived.
 *     IDENTITY_EDIT  a complete identity was REPLACED with a different one.
 *                    An alteration. Kept for the audit trail — the old value goes
 *                    in DETAIL, so it is recoverable — but NEVER treated as
 *                    evidence that the new identity is legitimate.
 *
 *   Before this split, every non-empty change logged RECEIVED, which is what let
 *   the 2026-08-28 incident write its own alibi: the identity guard reads RECEIVED
 *   events as the record of what a row's identity may be, so overwriting a live
 *   SALES_ORDER legitimised the corrupted value a minute before the guard looked
 *   at it. See _mrClassify and IdentityGuard.js.
 *
 * A typo correction still produces two log rows — that's the audit trail working,
 * showing the correction happened.
 *
 * The DETAIL string distinguishes the table the entry came from
 * ("eBay manual (replacement)" vs "DIRECT manual") for log readability, and
 * carries the alteration note when there is one.
 *
 * Dispatched from Main.js onEditInstallable. Defensive — never blocks other
 * handlers.
 */
function manualReceiveOnEdit(e) {
  try {
    if (!e || !e.range) return;

    var sheet = e.range.getSheet();
    if (sheet.getName() !== MAIN_SHEET_NAME) return;

    var startRow = e.range.getRow();
    var startCol = e.range.getColumn();
    var numRows  = e.range.getNumRows();
    var numCols  = e.range.getNumColumns();
    var lastCol  = startCol + numCols - 1;

    // ⚠⚠ EITHER identity column, not SALES_ORDER alone. Until 2026-08-30 this
    //   tested col D only, so typing the SO FIRST and the SKU SECOND logged
    //   NOTHING: the SO edit bailed on the still-missing SKU (below), and the
    //   later SKU edit never reached this handler at all. That row then carried
    //   no RECEIVED event anywhere — invisible to the archive, the metrics and
    //   the digest, and permanently flagged by the identity guard.
    var skuTouched = (Schema.cols.SKU >= startCol && Schema.cols.SKU <= lastCol);
    var soTouched  = (Schema.cols.SALES_ORDER >= startCol && Schema.cols.SALES_ORDER <= lastCol);
    if (!skuTouched && !soTouched) return;

    // Diagnostic breadcrumb — useful when debugging "didn't land in log" reports.
    console.log("manualReceiveOnEdit: range=" + e.range.getA1Notation() +
                " rows=" + numRows + " cols=" + numCols);

    var boundary = getBoundaryRow();
    var isSingleCell = (numRows === 1 && numCols === 1);

    // Cols A..G per row. Both identity halves are read from the ROW, never from
    // the edited range — the edit may have touched only one of the two.
    var contextValues = sheet.getRange(startRow, 1, numRows, Schema.cols.HAND).getValues();

    var batch = [];

    for (var i = 0; i < numRows; i++) {
      var row = startRow + i;
      if (row < Schema.dataStartRow) continue;
      // Skip boundary divider + DIRECT header row
      if (boundary > 0 && (row === boundary || row === boundary + 1)) continue;

      var rowCtx = contextValues[i];
      var sku = String(rowCtx[Schema.idx("SKU")] || "").trim();
      var so  = String(rowCtx[Schema.idx("SALES_ORDER")] || "").trim();

      // An identity is only worth recording once BOTH halves exist. A half-typed
      // row is someone mid-entry, not a receipt.
      if (!sku || !so) continue;

      var kind = _mrClassify(e, isSingleCell, startCol, sku, so);
      if (!kind) continue;                       // no real change on this row

      var qty  = parseInt(rowCtx[Schema.idx("QTY")]) || 1;
      var note = String(rowCtx[Schema.idx("NOTE")] || "").trim();

      // Tag DETAIL with the originating table so logs are readable at a glance.
      var inEbayTable = (boundary <= 0) || (row < boundary);
      var origin = inEbayTable ? "eBay manual (replacement)" : "DIRECT manual";

      // Slots: [event, orderId, sku, qty, source, detail, picker?, note]
      batch.push([
        kind.event,
        so,
        sku,
        qty,
        "manual",
        kind.detail ? (origin + " · " + kind.detail) : origin,
        undefined,    // picker — let logActivityBatch resolve from G2
        note
      ]);
    }

    if (batch.length > 0) logActivityBatch(batch);
  } catch (err) {
    try { Logger.log("manualReceiveOnEdit error: " + err); } catch (_) {}
  }
}


/**
 * Is this edit a RECEIPT or an ALTERATION?
 *
 * ⚠⚠ WHY THIS EXISTS — 2026-08-30. RECEIVED must mean "an order arrived". It had
 *   come to mean "somebody typed in column D", because the 2026-05-01 fix removed
 *   the empty→filled guard so EVERY non-empty change logged RECEIVED.
 *
 *   That is what made the 2026-08-28 incident undetectable. The identity guard
 *   builds its known-good set from RECEIVED events, so overwriting a live
 *   SALES_ORDER wrote the very evidence that legitimised the new value — the
 *   corrupted pair was in the set before the guard ever looked, and the verdict
 *   came back "ok". The corruption wrote its own alibi.
 *
 * @param {Event}   e             the onEdit event
 * @param {boolean} isSingleCell
 * @param {number}  startCol      the edited column (meaningful only when single-cell)
 * @param {string}  sku           the row's CURRENT sku  (already non-empty)
 * @param {string}  so            the row's CURRENT order id (already non-empty)
 * @returns {{event: string, detail: string}|null}  null → nothing worth logging
 */
function _mrClassify(e, isSingleCell, startCol, sku, so) {
  // ⚠ A PASTE CARRIES NO e.oldValue, so creation and mutation are
  //   indistinguishable from the event alone — and the row cannot be read back,
  //   because onEdit fires AFTER the write. Stay optimistic, preserving the
  //   2026-05-01 liberal-logging ruling; the DUPLICATE rule catches the
  //   consequence of a pasted row in both directions (into an empty row, and
  //   over an occupied one — each creates a second row with the same pair).
  if (!isSingleCell) return { event: "RECEIVED", detail: "" };

  var oldVal = String(e && e.oldValue != null ? e.oldValue : "").trim();
  var editedSku = (startCol === Schema.cols.SKU);
  var newVal = editedSku ? sku : so;

  if (newVal === oldVal) return null;            // no real change

  // Both halves are non-empty by the time we get here, so an EMPTY old value on
  // the edited cell means the identity was incomplete before and is complete now.
  // A real new row — whichever of the two cells finished it.
  if (!oldVal) return { event: "RECEIVED", detail: "" };

  // ⭐ A COMPLETION, NOT A REPLACEMENT. The templated fill the 2026-05-01 fix
  //   exists for — "Replacement for #:" → "Replacement for #: 26-14551-63163" —
  //   EXTENDS the old value rather than discarding it. Still a receipt.
  if (newVal.indexOf(oldVal) === 0) return { event: "RECEIVED", detail: "" };

  // ⚠⚠ AN ALTERATION OF A LIVE IDENTITY — the 2026-08-28 shape. Recorded for the
  //   audit trail (the old value is now recoverable, which it never was before),
  //   and NEVER counted as evidence by the identity guard.
  return {
    event: "IDENTITY_EDIT",
    detail: "identity changed · " + (editedSku ? "SKU" : "SALES_ORDER") + " was " + oldVal
  };
}


// ═══════════════════════════════════════════════════════════════════════════════════════
// NOTE-EDIT HOOK (Activity Log) — captures ANY edit to the NOTE column
// ═══════════════════════════════════════════════════════════════════════════════════════

/**
 * onEdit handler for note edits in the All Orders sheet.
 *
 * Why: a buyer note arriving with the order is one event. A supervisor or
 * picker adding context mid-prep is a different event. The Activity Log
 * needs to capture both so the audit trail tells the full story.
 *
 * Trigger condition:
 *   - Edit is on the All Orders sheet
 *   - Range intersects the NOTE column (col E)
 *   - Row is in a data area (not the divider or DIRECT header)
 *   - For single-cell edits: newVal !== oldVal (real change)
 *   - For multi-cell edits: log optimistically (paste/autofill)
 *
 * Logs a NOTE event. The DETAIL column captures the prior text via
 * "Was: <old>" / "Note added" / "Note removed" — so reading the log shows
 * the diff at a glance. The NOTE column carries the new value.
 *
 * Picker auto-captures from G2 (warehouse-side source). When G2 is unset,
 * the empty PICKER column is the strong signal that this was a non-picker
 * (likely supervisor) edit.
 *
 * Dispatched from Main.js onEditInstallable. Defensive — never blocks others.
 */
function noteEditOnEdit(e) {
  try {
    if (!e || !e.range) return;

    var sheet = e.range.getSheet();
    if (sheet.getName() !== MAIN_SHEET_NAME) return;

    var startRow = e.range.getRow();
    var startCol = e.range.getColumn();
    var numRows  = e.range.getNumRows();
    var numCols  = e.range.getNumColumns();

    // NOTE col (E = 5) must intersect the edited range
    var noteColInRange = Schema.cols.NOTE - startCol;
    if (noteColInRange < 0 || noteColInRange >= numCols) return;

    var boundary = getBoundaryRow();
    var isSingleCell = (numRows === 1 && numCols === 1);

    var rangeValues = e.range.getValues();
    // Need cols A..G per affected row for context (SKU, QTY, ORDER_ID).
    var contextValues = sheet.getRange(startRow, 1, numRows, Schema.cols.HAND).getValues();

    var batch = [];

    for (var i = 0; i < numRows; i++) {
      var row = startRow + i;
      if (row < Schema.dataStartRow) continue;
      if (boundary > 0 && (row === boundary || row === boundary + 1)) continue;

      var newNote = String(rangeValues[i][noteColInRange] || "").trim();

      var oldNote = "";
      if (isSingleCell) {
        oldNote = String(e.oldValue || "").trim();
        if (newNote === oldNote) continue;  // no real change
      }
      // Multi-cell: oldNote stays empty; we'll log a "Note added" or "Note set"
      // event without the old → new diff. Liberal logging accepted.

      // Auto-clear Zoho-flag highlight (soft-red bg + strikethrough) when the
      // picker removes the warning prefix from a flagged row. Without this,
      // clearing the NOTE cell content leaves the bg behind — Sheets keeps
      // background separate from content. Single-cell only (we need e.oldValue
      // to confidently detect the prefix was removed; multi-cell pastes lose
      // that signal, picker can use Format → Clear formatting if needed).
      // Shipped 2026-05-23 alongside the format-bleed fix in _insertAddedItemsToDirect.
      if (isSingleCell) {
        var FLAG_PREFIX_RE = /^⚠️\s+(ZOHO QTY|REMOVED IN ZOHO)/;
        var oldHadFlag = FLAG_PREFIX_RE.test(oldNote);
        var newHasFlag = FLAG_PREFIX_RE.test(newNote);
        if (oldHadFlag && !newHasFlag) {
          try {
            var rowRange = sheet.getRange(startRow + i, 1, 1, Schema.dataWidth);
            rowRange.setBackground(null);
            rowRange.setFontLine('none');
          } catch (clearErr) {
            try { Logger.log("noteEditOnEdit clear-flag error: " + clearErr); } catch (_) {}
          }
        }
      }

      var rowCtx = contextValues[i];
      var sku     = String(rowCtx[Schema.idx("SKU")] || "").trim();
      var qty     = parseInt(rowCtx[Schema.idx("QTY")]) || 0;
      var orderId = String(rowCtx[Schema.idx("SALES_ORDER")] || "").trim();

      // Skip rows that aren't real orders (no SKU AND no SO → just a stray edit)
      if (!sku && !orderId) continue;

      var detail;
      if (!oldNote && newNote)        detail = "Note added";
      else if (oldNote && !newNote)   detail = "Note removed (was: " + _truncate(oldNote, 80) + ")";
      else if (oldNote && newNote)    detail = "Was: " + _truncate(oldNote, 80);
      else                            detail = "Note edit";

      batch.push([
        "NOTE",
        orderId,
        sku,
        qty,
        "manual",
        detail,
        undefined,    // picker auto-captured from G2 (or blank if unset)
        newNote
      ]);
    }

    if (batch.length > 0) logActivityBatch(batch);
  } catch (err) {
    try { Logger.log("noteEditOnEdit error: " + err); } catch (_) {}
  }
}

/** Tiny helper for keeping DETAIL text manageable. */
function _truncate(s, max) {
  s = String(s || "");
  return (s.length > max) ? (s.substring(0, max - 1) + "…") : s;
}


// ═══════════════════════════════════════════════════════════════════════════════════════
// DEBUG LOGGING
// ═══════════════════════════════════════════════════════════════════════════════════════

function logDebug(message) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName("Debug Log");
    
    if (!logSheet) {
      logSheet = ss.insertSheet("Debug Log");
      logSheet.getRange("A1").setValue("Timestamp");
      logSheet.getRange("B1").setValue("Message");
      logSheet.setFrozenRows(1);
    }
    
    var timestamp = new Date().toLocaleString();
    logSheet.appendRow([timestamp, message]);
    
    var lastRow = logSheet.getLastRow();
    if (lastRow > 101) {
      logSheet.deleteRows(2, lastRow - 101);
    }
  } catch (e) {
    // Silently fail
  }
}

function clearDebugLog() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var logSheet = ss.getSheetByName("Debug Log");
  if (logSheet) {
    logSheet.clear();
    logSheet.getRange("A1").setValue("Timestamp");
    logSheet.getRange("B1").setValue("Message");
    logSheet.setFrozenRows(1);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// WEBHOOK MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════════════

function setWebhook() {
  // The bot's webhook MUST point at the n8n Telegram Button Handler workflow
  // (N8N_TELEGRAM_CALLBACK_WEBHOOK_URL in Secrets.js), NOT at WEB_APP_URL.
  // Apps Script /exec always answers with a 302 redirect, which Telegram
  // rejects ("Wrong response from the webhook: 302 Moved Temporarily") — so
  // pointing the webhook at Apps Script silently kills every button click.
  // That exact mistake broke the Telegram buttons after the 2026-05-31 VPS
  // migration (fixed 2026-06-10). Flow: Telegram → n8n → doPost.
  var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/setWebhook?url=" + N8N_TELEGRAM_CALLBACK_WEBHOOK_URL);
  Logger.log(response.getContentText());
}

function deleteWebhook() {
  var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/deleteWebhook");
  Logger.log(response.getContentText());
}

function getWebhookInfo() {
  var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/getWebhookInfo");
  Logger.log(response.getContentText());
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// ORDER MANAGEMENT FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════════════

// addOrderFromN8N — REMOVED 2026-04-29.
// Was a single-row insert path with zero callers in either Apps Script or n8n
// (verified by grep). Carried the same partial-width pattern as the sort bug
// (used 7 cols, would have detached SHIPPING/SHIP_COST). Resurrecting it
// later should go through doPost's batch-insert path, not this code.

function getOrderStats() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) throw new Error("Main sheet not found");
  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return { pending: 0, preparing: 0, shipped: 0, canceled: 0 };
  var statuses = sheet.getRange(
    Schema.dataStartRow, Schema.cols.STATUS,
    lastRow - Schema.dataStartRow + 1, 1
  ).getValues().flat();
  var stats = { pending: 0, preparing: 0, shipped: 0, canceled: 0 };
  statuses.forEach(function(s) {
    s = String(s).trim().toUpperCase();
    if (s === Schema.status.PENDING)        stats.pending++;
    else if (s === Schema.status.PREPARING) stats.preparing++;
    else if (s === Schema.status.SHIPPED)   stats.shipped++;
    else if (s === Schema.status.CANCELED)  stats.canceled++;
  });
  return stats;
}

/**
 * NO-OP since 2026-05-17 (Service Bay v6 port).
 *
 * The G1 stats cell is now driven by a LIVE FORMULA installed by
 * BrandTheme._setSystemPulseBannerFormulas():
 *   ="🔴 "&COUNTIF(F:F,"PENDING")&"   🟡 "&COUNTIF(F:F,"PREPARING")&...
 *
 * That formula re-evaluates on every spreadsheet recalc — accurate, live,
 * cheaper than a function call. If we wrote static text to G1 here, we'd
 * clobber the formula after every n8n batch insert.
 *
 * Kept as a named function so existing callers (doPost batch finalize, etc.)
 * don't break. Body is intentionally empty.
 */
function updateOrderStatsInSheet() {
  // Live formula in G1 handles this now — see BrandTheme._setSystemPulseBannerFormulas.
}

/**
 * Detects and restores filter header corruption caused by insertRowsBefore().
 * Google Sheets has a known bug where inserting rows inside a filtered range
 * replaces header text with "Column 1", "Column 2", etc.
 * Call this after every insertRowsBefore/insertRowBefore on the main sheet.
 * @param {Sheet} sheet - The sheet to check
 * @param {Array} savedHeaders - The correct header values saved before insertion
 */
function verifyAndRestoreHeaders(sheet, savedHeaders) {
  var headerRow = Schema.headerRow;
  var current = sheet.getRange(headerRow, 1, 1, savedHeaders.length).getValues()[0];
  var corrupted = current.some(function(val) {
    return /^Column\s*\d+$/i.test(String(val).trim());
  });
  if (corrupted) {
    sheet.getRange(headerRow, 1, 1, savedHeaders.length).setValues([savedHeaders]);
    Logger.log("⚠️ Headers were corrupted by row insertion — restored successfully.");
  }
}

/**
 * NO-OP since 2026-05-17 (Service Bay v6 port).
 *
 * E1 is now the System Pulse cell, driven by a live formula installed by
 * BrandTheme._setSystemPulseBannerFormulas():
 *   - Sync time derived from MAX('Activity Log'!A:A) — reflects ALL system
 *     activity, not just n8n's idea of "last sync." More accurate.
 *   - Color-coded 🟢/🟡/🔴 indicator based on minutes-since.
 *   - Self-updating "Nm ago" freshness counter.
 *
 * If we wrote static "⏱ Last eBay sync · X:XX PM" here, it would clobber the
 * formula on every n8n batch insert. The formula reads the audit log (the
 * source of truth) instead of being told what time it is.
 *
 * Kept as a named function so existing callers don't break. Body intentionally empty.
 */
function updateLastSyncTimestamp() {
  // Live formula in E1 handles this now — see BrandTheme._setSystemPulseBannerFormulas.
}

/**
 * Status → sort rank. PENDING first, CANCELED last, unknown parked with CANCELED,
 * blank after everything. Shared by both tables so they can never disagree.
 */
function _statusSortRank(status) {
  var s = String(status == null ? '' : status).trim().toUpperCase();
  if (s === Schema.status.PENDING)   return 1;
  if (s === Schema.status.PREPARING) return 2;
  if (s === Schema.status.SHIPPED)   return 3;
  if (s === Schema.status.CANCELED)  return 4;
  if (s === '')                      return 5;
  return 4;   // unknown / typo'd status parks with CANCELED, never above live work
}

/**
 * Rank every SALES ORDER group in the DIRECT table by its MOST URGENT row.
 *
 * WHY A GROUP RANK AND NOT A PER-ROW SORT (2026-08-07): the DIRECT table is
 * order-centric — the picker fulfils one order at a time and the gold box painted
 * by _paintDirectOrderDividers derives its group starts from consecutive col-D
 * values. Sorting rows by status directly would split a HALF-SHIPPED order across
 * the table and render it as TWO boxes for one order, which is the exact thing the
 * divider exists to prevent. So the group moves as a unit, and it moves according
 * to the most urgent row inside it:
 *
 *   any row PENDING    → group ranks 1 → top       (still needs picking)
 *   all rows PREPARING → group ranks 2 → middle    (in progress)
 *   all rows SHIPPED   → group ranks 3 → bottom    (done, out of the way)
 *   all rows CANCELED  → group ranks 4 → below that
 *
 * Consequence worth knowing: a PARTIALLY shipped order stays near the top rather
 * than sinking, because it still contains work. Its shipped lines sink to the
 * bottom of its OWN box (see the within-group rule in _compareDirectRows).
 *
 * @param {Array<{values:Array}>} indexed  paired row objects (values + formats + rich text)
 * @returns {Object} map of trimmed SO → best (lowest) status rank in that group
 */
function _buildDirectSoRank(indexed) {
  var rank = {};
  for (var i = 0; i < indexed.length; i++) {
    var so = String(indexed[i].values[Schema.idx("SALES_ORDER")] || '').trim();
    if (!so) continue;   // blank/buffer rows are sunk by the comparator, not ranked
    var r = _statusSortRank(indexed[i].values[Schema.idx("STATUS")]);
    if (rank[so] === undefined || r < rank[so]) rank[so] = r;
  }
  return rank;
}

/**
 * DIRECT row comparator: group by SALES ORDER (groups ordered by urgency), then
 * status, then aisle within the order.
 *
 * @param {Object} a       paired row object
 * @param {Object} b       paired row object
 * @param {Object} soRank  output of _buildDirectSoRank
 */
function _compareDirectRows(a, b, soRank) {
  var soA = String(a.values[Schema.idx("SALES_ORDER")] || '').trim();
  var soB = String(b.values[Schema.idx("SALES_ORDER")] || '').trim();

  // Empty-SO blank/buffer rows always sink, whatever their status says.
  if (!soA && !soB) return 0;
  if (!soA) return 1;
  if (!soB) return -1;

  if (soA !== soB) {
    // Different orders → the more urgent GROUP wins, so shipped orders sit below
    // preparing ones, which sit below anything still pending.
    var gA = soRank[soA] || 5;
    var gB = soRank[soB] || 5;
    if (gA !== gB) return gA - gB;
    // Same urgency → oldest order first, and deterministic either way.
    return soA.localeCompare(soB);
  }

  // Same order → picked lines drop to the bottom of their own box, and what's
  // left to pick reads top-down in aisle order.
  var cmp = _statusSortRank(a.values[Schema.idx("STATUS")]) -
            _statusSortRank(b.values[Schema.idx("STATUS")]);
  if (cmp !== 0) return cmp;
  return compareLocations(a.values[Schema.idx("LOCATION")],
                          b.values[Schema.idx("LOCATION")]);
}

function sortTableByStatusAndLocation(tableNumber) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  var startRow = (tableNumber === 1) ? Schema.dataStartRow : boundary + 2;
  var endRow = (tableNumber === 1) ? boundary - 1 : sheet.getLastRow();
  var lastDataRow = findLastDataRowInSegment(startRow, endRow);
  if (lastDataRow < startRow) return "Table is empty.";
  var numRows = lastDataRow - startRow + 1;
  // Sort the FULL row width (Schema.dataWidth = 10) so SHIPPING and SHIP_COST
  // travel with their owning row. Reading only 8 cols was the I/J detach bug.
  var range = sheet.getRange(startRow, 1, numRows, Schema.dataWidth);
  var data = range.getValues();
  // CRITICAL: read number formats alongside values and sort them together.
  // Per-cell number formats (like the kit-row "▣ " glyph prefix written by
  // refreshKitSkuMarkers) are NOT moved by range.setValues — they stay glued
  // to their original cell positions. If we sort only values, the ▣ marker
  // stays behind in whatever row it WAS in, and the SKU that gets sorted
  // INTO that row inherits a wrong marker. Bug surfaced 2026-05-19 when a
  // kit row moved from PENDING→PREPARING and dropped to the bottom — its
  // ▣ stayed at the top, and a non-kit SKU above showed the marker.
  // Fix: capture formats[i] alongside data[i], sort as paired rows,
  // setValues + setNumberFormats both. Same applies to ANY per-cell static
  // format anyone adds in the future.
  var formats = range.getNumberFormats();
  // Same class of "doesn't move with setValues" problem as number formats:
  // col-A RICH TEXT (the SKU → eBay listing link) AND col-D RICH TEXT (the
  // SALES ORDER → eBay/Zoho order link) are both stripped by setValues. Capture
  // them, move them WITH their rows through the sort, then re-write after
  // setValues so the links land on the right rows.
  var skuRange = sheet.getRange(startRow, Schema.cols.SKU, numRows, 1);
  var skuRich = skuRange.getRichTextValues();   // [[RichTextValue], ...]
  var soRange = sheet.getRange(startRow, Schema.cols.SALES_ORDER, numRows, 1);
  var soRich = soRange.getRichTextValues();
  // Pair values + formats + col-A & col-D rich text so they travel together.
  var indexed = data.map(function(row, i) {
    return { values: row, formats: formats[i], rich: skuRich[i][0], soRich: soRich[i][0] };
  });
  if (tableNumber === 2) {
    // DIRECT is ORDER-CENTRIC: whole SALES ORDER groups move as units, ordered by
    // the most urgent row inside each (so shipped orders sink to the bottom the way
    // shipped rows do on eBay), then status + aisle within the order. The group rank
    // is what keeps a half-shipped order from splitting into two gold boxes — see
    // _buildDirectSoRank for the full reasoning.
    var soRank = _buildDirectSoRank(indexed);
    indexed.sort(function(a, b) { return _compareDirectRows(a, b, soRank); });
  } else {
    // eBay is AISLE-CENTRIC: one line per order, so there is no grouping to protect
    // — sort rows straight by status then location.
    indexed.sort(function(a, b) {
      var cmp = _statusSortRank(a.values[Schema.idx("STATUS")]) -
                _statusSortRank(b.values[Schema.idx("STATUS")]);
      if (cmp !== 0) return cmp;
      return compareLocations(a.values[Schema.idx("LOCATION")],
                              b.values[Schema.idx("LOCATION")]);
    });
  }
  var sortedData    = indexed.map(function(x) { return x.values; });
  var sortedFormats = indexed.map(function(x) { return x.formats; });
  var sortedRich    = indexed.map(function(x) { return [x.rich]; });
  var sortedSoRich  = indexed.map(function(x) { return [x.soRich]; });
  range.setValues(sortedData);
  range.setNumberFormats(sortedFormats);
  // Re-apply col-A + col-D links AFTER setValues (which wrote plain text +
  // stripped the links). The rich text carries the same cell text, so the
  // columns end correct with their links re-attached.
  skuRange.setRichTextValues(sortedRich);
  soRange.setRichTextValues(sortedSoRich);

  // Safety net — re-derive kit ▣ markers from current SKU values against the
  // Kit Registry. Catches any drift introduced by edits/inserts that happened
  // between sorts (e.g. a SKU got changed but its row format never updated).
  try { refreshKitSkuMarkers(); }
  catch (e) { console.log("sortTableByStatusAndLocation: kit marker refresh error: " + e); }

  // Same class of bug as ▣ markers: per-cell left borders on duplicate SOs
  // don't travel with setValues — they stay glued to the original row position.
  // Without this refresh, post-sort borders point at whoever LANDED on the
  // old duplicate row, not the actual duplicate group.
  try { setupDuplicateSalesOrderHighlighting(); }
  catch (e) { console.log("sortTableByStatusAndLocation: dup-SO refresh error: " + e); }

  return "✅ Sorted";
}

function sortEbayTable() { return sortTableByStatusAndLocation(1); }
function sortDirectTable() { return sortTableByStatusAndLocation(2); }

function refreshProDashboard() {
  // Stats banner (G1) refresh. The date in B1 auto-updates via the
  // =TEXT(TODAY(),...) formula installed by _ensureDateFormula(). The
  // last-sync timestamp in E1 only updates on actual sync events — manual
  // refresh shouldn't fake that, so we don't touch it here.
  updateOrderStatsInSheet();
  return "✅ Dashboard refreshed";
}

/**
 * Back-compat wrapper around updateOrderStatus().
 * Called from doPost (action: updateStatus) — n8n posts a row number + new status.
 *
 * Original behavior: blindly setValue + refresh stats. Did NOT guard terminal
 * states, did NOT sync Telegram. We keep that behavior (syncTelegram: false)
 * because n8n's own webhook flow handles Telegram separately.
 */
function updateStatus(rowNumber, status) {
  var result = updateOrderStatus(rowNumber, status, {
    source:       "n8n-direct",
    syncTelegram: false,
    sortAfter:    false
  });
  return ContentService.createTextOutput(JSON.stringify({
    success: result.success,
    count:   result.count
  })).setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════════════════════════════════════════
// 💾 MESSAGE STORAGE HELPERS (The Missing Link)
// ═══════════════════════════════════════════════════════════════════════════════════════

/**
 * Stores the Telegram Message ID and Chat ID linked to an Order ID
 * Called by doPost when action === 'storeMessageId'
 */
function storeMessageId(orderId, messageId, chatId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  
  if (!sheet) {
    // Safety: Create the sheet if it doesn't exist
    sheet = ss.insertSheet(HIDDEN_SHEET_NAME);
    sheet.appendRow(["Order ID", "Message ID", "Chat ID", "Timestamp"]);
  }
  
  // Append the new log entry
  sheet.appendRow([orderId, messageId, chatId, new Date()]);
  
  return ContentService.createTextOutput("Stored");
}

/**
 * Retrieves the Message ID and Chat ID for a specific Order ID
 * Called by notifyTelegramShipped
 */
function getMessageId(orderId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  
  if (!sheet) {
    console.log("getMessageId: " + HIDDEN_SHEET_NAME + " sheet not found.");
    return null;
  }
  
  var data = sheet.getDataRange().getValues();
  var targetId = String(orderId).trim().toLowerCase();
  
  // Loop backwards (bottom to top) to find the most recent entry for this order
  // Start at i >= 1 to skip the header row
  for (var i = data.length - 1; i >= 1; i--) {
    var rowId = String(data[i][0]).trim().toLowerCase();
    
    if (rowId === targetId) {
      return {
        messageId: data[i][1],
        chatId: data[i][2]
      };
    }
  }
  
  return null; // Not found
}


// ═══════════════════════════════════════════════════════════════════════════════════════
// 📊 INVENTORY LOOKUP FUNCTIONS - For Auto-Populating HAND Column
// ═══════════════════════════════════════════════════════════════════════════════════════

/**
 * Get available inventory for a SKU
 * Used when inserting new orders via n8n
 */
function getInventoryForSKU(sku) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  
  if (!dbSheet || !sku) {
    return 0;
  }
  
  var skuLower = String(sku).trim().toLowerCase();
  var data = dbSheet.getDataRange().getValues();
  var headers = data[0];
  
  var skuCol = headers.indexOf(DB_SKU_HEADER);
  var qtyCol = headers.indexOf(DB_QUANTITY_HEADER);
  var soldCol = headers.indexOf(DB_QUANTITY_SOLD_HEADER);
  
  if (skuCol === -1 || qtyCol === -1 || soldCol === -1) {
    return 0;
  }
  
  // Search for SKU
  for (var i = 1; i < data.length; i++) {
    var dbSku = String(data[i][skuCol] || "").trim().toLowerCase();
    
    if (dbSku === skuLower) {
      var qty = parseInt(data[i][qtyCol]) || 0;
      var sold = parseInt(data[i][soldCol]) || 0;
      return qty - sold; // Return available quantity
    }
  }
  
  return 0; // SKU not found
}


/**
 * Synchronizes a status change from the sheet to Telegram.
 * REPLICATES THE "ELEGANT MOBILE DESIGN" FROM N8N WORKFLOW
 *
 * options.revertReason — optional string. When present, an "Auto-reverted"
 *   banner is rendered above the status line. Used by the verify workflow
 *   (source=n8n-verify) to explain why a SHIPPED row was rolled back —
 *   without it, the picker just sees PENDING again with no context.
 */
function syncStatusToTelegram(orderId, newStatus, options) {
  options = options || {};
  var revertReason = options.revertReason || "";
  var msgData = getMessageId(orderId);
  if (!msgData) {
    console.log("syncStatusToTelegram: no Telegram message id for " + orderId + " — skipping.");
    return;
  }

  // 1. GATHER DATA & CALCULATE TOTALS
  // We grab all rows for this order to build the "Pick List"
  var items = getItemsFromSheet(orderId);
  var totalUnits = 0;
  for (var k = 0; k < items.length; k++) totalUnits += parseInt(items[k].qty) || 0;

  // 2. DEFINE STATUS EMOJI & BUTTONS
  var statusEmoji = "🔴";
  var buttons = []; 

  if (newStatus === Schema.status.PREPARING) {
    statusEmoji = "🟡";
    buttons.push([{ "text": "🔄 Revert to Pending", "callback_data": "PEND_" + orderId }]);
  } else if (newStatus === Schema.status.PENDING) {
    statusEmoji = "🔴";
    buttons.push([{ "text": "🚀 Mark as Preparing", "callback_data": "PREP_" + orderId }]);
  } else if (newStatus === Schema.status.SHIPPED) {
    statusEmoji = "✅"; // No buttons
  } else if (newStatus === Schema.status.CANCELED) {
    statusEmoji = "❌"; // No buttons
  }

  // 3. BUILD MESSAGE (Exact N8N Template Port)
  // Pre-build inventory map once instead of calling getInventoryForSKU per item
  var inventoryCache = {};
  try {
    var maps = buildLocationAndInventoryMaps();
    var invMap = maps.inventoryMap;
  } catch (e) {
    invMap = new Map();
  }

  var timestamp = Utilities.formatDate(new Date(), "America/Chicago", "EEE, MMM d, h:mm a"); // Houston Time

  var msg = "══════════════════════\n";
  msg += "       📦  ORDER\n";
  msg += "══════════════════════\n";
  msg += "🕐  " + timestamp + " CST\n\n";

  msg += "🔖  " + orderId + "\n";
  msg += "📦  " + totalUnits + " total units\n\n";

  // -- PICK LIST SECTION --
  msg += "┌─────────────────────\n";
  msg += "│ PICK LIST\n";
  msg += "├─────────────────────\n";

  for (var j = 0; j < items.length; j++) {
    var item = items[j];
    var isLast = j === items.length - 1;
    var prefix = isLast ? '└' : '├';
    var linePrefix = isLast ? ' ' : '│';

    // Get Live Inventory from cached map (single read)
    var invData = invMap.get(String(item.sku).trim().toLowerCase());
    var availableStock = invData ? invData.available : 0;
    var stockStatus = availableStock <= 20 ? '⚠️' : '✅';

    msg += "│\n";
    msg += prefix + "─ " + (j + 1) + ". SKU: " + item.sku + "\n";
    msg += linePrefix + "      ├─ 📦 " + item.sku + "\n"; // Added explicit SKU line to match n8n
    msg += linePrefix + "      ├─ 📍 Loc: " + item.loc + "\n";
    msg += linePrefix + "      ├─ 🔢 Qty: " + item.qty + "\n";
    msg += linePrefix + "      └─ 📊 Stock: " + stockStatus + " " + availableStock + " units\n";
  }
  msg += "\n";

  // -- NOTE SECTION --
  // If any item has a note, we display the first one (common for order-level notes)
  var orderNote = items.find(i => i.note !== "")?.note || "";
  if (orderNote) {
    msg += "┌─────────────────────\n";
    msg += "│ 💬 BUYER NOTE\n";
    msg += "├─────────────────────\n";
    msg += "│ " + orderNote + "\n";
    msg += "└─────────────────────\n\n";
  }

  if (revertReason) {
    msg += "⚠️  AUTO-REVERTED — eBay: " + revertReason + "\n";
  }
  msg += "📋 Status: " + statusEmoji + " " + newStatus;

  // 4. SEND UPDATE
  // NOTE: No parse_mode set — message uses Unicode box-drawing, not HTML.
  // Setting parse_mode:"HTML" caused silent failures when eBay titles contain & < > chars.
  var payload = {
    "chat_id": String(msgData.chatId),
    "message_id": parseInt(msgData.messageId),
    "text": msg,
    "reply_markup": { "inline_keyboard": buttons }
  };

  try {
    var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/editMessageText", {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });
    var respCode = response.getResponseCode();
    if (respCode !== 200) {
      console.log("syncStatusToTelegram editMessage failed for " + orderId + " (" + respCode + "): " + response.getContentText());
    }
  } catch (e) {
    console.log("syncStatusToTelegram error for " + orderId + ": " + e);
  }
}

/**
 * Helper: Gets all items for a specific order ID from the sheet
 */
function getItemsFromSheet(orderId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var lastRow = sheet.getLastRow();

  if (lastRow < Schema.dataStartRow) return [];

  // Search ALL rows (both eBay and Direct tables) for the order ID.
  // Read SKU through NOTE — covers everything we extract per item.
  var data = sheet.getRange(
    Schema.dataStartRow, 1,
    lastRow - Schema.dataStartRow + 1,
    Schema.cols.NOTE
  ).getValues();
  var items = [];
  var cleanOrderId = String(orderId).trim();

  for (var i = 0; i < data.length; i++) {
    // Skip the DIRECT boundary row
    if (String(data[i][Schema.idx("SKU")]).trim().toUpperCase() === Schema.boundaryMarker) continue;

    if (String(data[i][Schema.idx("SALES_ORDER")]).trim() === cleanOrderId) {
      items.push({
        sku:  data[i][Schema.idx("SKU")],
        qty:  data[i][Schema.idx("QTY")],
        loc:  data[i][Schema.idx("LOCATION")],
        note: data[i][Schema.idx("NOTE")]
      });
    }
  }
  return items;
}

/**
 * HANDLES MANUAL EDITS
 */
function handleManualStatusChange(e) {
  if (!e || !e.range) return;
  var range = e.range;
  var sheet = range.getSheet();

  // Only process STATUS column edits on the main sheet
  if (sheet.getName() !== MAIN_SHEET_NAME) return;
  if (range.getColumn() !== Schema.cols.STATUS) return;
  if (range.getRow() < Schema.dataStartRow) return;


  // The user has ALREADY written the new status to the sheet via their edit.
  // We need to:
  //   - run our canonical update sequence (stats refresh + Telegram sync)
  //   - NOT re-sort (user is mid-edit, don't disturb)
  //   - group rows by their new status (user might have edited multiple rows
  //     with different values via paste — rare but possible)
  //
  // Don't take our own lock here — updateOrderStatus acquires its own.
  var numRows = range.getHeight();
  var statuses = sheet.getRange(range.getRow(), Schema.cols.STATUS, numRows, 1).getValues();

  // Group rows by their new status value (handles paste of mixed values)
  var rowsByStatus = {};
  for (var i = 0; i < numRows; i++) {
    var newStatus = String(statuses[i][0]).trim().toUpperCase();
    if (!Schema.isValidStatus(newStatus)) continue;
    if (!rowsByStatus[newStatus]) rowsByStatus[newStatus] = [];
    rowsByStatus[newStatus].push(range.getRow() + i);
  }

  // One updateOrderStatus call per status group
  Object.keys(rowsByStatus).forEach(function(status) {
    try {
      updateOrderStatus(rowsByStatus[status], status, {
        source:    "manual-edit",
        sortAfter: false,  // user is mid-edit, don't disturb their cursor
        force:     true    // user already typed it — sync sheet+Telegram even if old status was terminal
      });
    } catch (err) {
      console.log("handleManualStatusChange error for status " + status + ": " + err);
    }
  });
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// 🔍 DIAGNOSTIC - Run from Script Editor to debug Telegram 404///
// ═══════════════════════════════════════════════════════════════════════════════════════

function diagnoseTelegram() {
  logDebug("=== TELEGRAM DIAGNOSTIC START ===");

  // 1. Check bot token
  var tokenPreview = TELEGRAM_BOT_TOKEN
    ? (TELEGRAM_BOT_TOKEN.substring(0, 6) + "..." + TELEGRAM_BOT_TOKEN.substring(TELEGRAM_BOT_TOKEN.length - 4))
    : "UNDEFINED";
  logDebug("Token preview: " + tokenPreview);
  logDebug("Token length: " + (TELEGRAM_BOT_TOKEN ? TELEGRAM_BOT_TOKEN.length : 0));
  logDebug("Token type: " + typeof TELEGRAM_BOT_TOKEN);

  // 2. Test getMe (verifies token is valid)
  try {
    var getMeUrl = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/getMe";
    logDebug("Calling getMe...");
    var getMeResp = UrlFetchApp.fetch(getMeUrl, { "muteHttpExceptions": true });
    logDebug("getMe status: " + getMeResp.getResponseCode());
    logDebug("getMe response: " + getMeResp.getContentText());
  } catch (e) {
    logDebug("getMe ERROR: " + e.toString());
  }

  // 3. Test with the most recent message in Telegram_Messages sheet
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var msgSheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  if (!msgSheet || msgSheet.getLastRow() < 2) {
    logDebug("No messages stored in Telegram_Messages sheet");
    logDebug("=== DIAGNOSTIC END ===");
    return;
  }

  var lastRow = msgSheet.getLastRow();
  var testData = msgSheet.getRange(lastRow, 1, 1, 3).getValues()[0];
  var testOrderId = testData[0];
  var testMsgId = testData[1];
  var testChatId = testData[2];

  logDebug("Test order: " + testOrderId);
  logDebug("Test message_id: " + testMsgId + " (type: " + typeof testMsgId + ")");
  logDebug("Test chat_id: " + testChatId + " (type: " + typeof testChatId + ")");
  logDebug("parseInt(message_id): " + parseInt(testMsgId));

  // 4. Try editMessageText with minimal payload
  try {
    var editUrl = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/editMessageText";
    var testPayload = {
      "chat_id": String(testChatId),
      "message_id": parseInt(testMsgId),
      "text": "Diagnostic test - " + new Date().toLocaleString()
    };
    logDebug("Edit payload: " + JSON.stringify(testPayload));

    var editResp = UrlFetchApp.fetch(editUrl, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(testPayload),
      "muteHttpExceptions": true
    });
    logDebug("Edit status: " + editResp.getResponseCode());
    logDebug("Edit response: " + editResp.getContentText());
  } catch (e) {
    logDebug("Edit ERROR: " + e.toString());
  }

  logDebug("=== DIAGNOSTIC END ===");
}