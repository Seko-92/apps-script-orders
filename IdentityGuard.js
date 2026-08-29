// =======================================================================================
// IDENTITYGUARD.JS — does every open row's identity match something we actually received?
// =======================================================================================
//
// WHY IT EXISTS
//   2026-08-28: a SALES_ORDER cell was overwritten by hand on a row already picked and
//   shelf-counted. doPost dedupes on SALES_ORDER + "|" + SKU, so the signature stopped
//   existing and n8n re-inserted the line four minutes later. Nothing alerted; the only
//   trace was two rows that looked like twins.
//
// ⚠⚠ THIS IS THE SECOND DESIGN. The first compared e.oldValue against e.value on every
//   edit, and it FAILED IN PRODUCTION on 2026-08-29 in four ways — all reported from
//   real use, all traceable to the same root cause:
//
//     · Ctrl+Z could not win. Undo is itself a user edit, so it re-entered the handler;
//       the guard put its value back, the operator undid it again, and round it went.
//     · Restoring the CORRECT value still alarmed. Old-vs-new cannot tell "putting it
//       right" from "breaking it" — the alert named a restored value as the offender,
//       with that same value listed underneath as the recovery candidate.
//     · An event-based flag CAN NEVER CLEAR ITSELF, because no edit event means "it is
//       fine now". A corrected row stayed amber forever.
//     · Reverting a SKU left a MIXED ROW: liveUpdateTrigger had already written LOCATION
//       and HAND for the new sku, so the old sku sat beside the wrong shelf and stock.
//
//   ⭐ THE REFRAME: ask about STATE, not the EVENT. Not "what changed?" but "does this
//   row's identity match something we actually received?" Every failure above is a
//   symptom of event-comparison, and every one disappears when the question changes.
//
// ⭐ THE ANSWER IS ALREADY ON THE SHEET. Every row that legitimately exists has a
//   RECEIVED entry in the Activity Log carrying its order id and SKU — written by
//   doPost, Zoho pull, kit expansion, manual entry, and /missing. That is the SAME
//   SALES_ORDER|SKU signature doPost dedupes on, so the log IS the record of what a
//   row's identity is allowed to be.
//
// HOW IT RUNS
//   onEdit does almost nothing: an identity column was touched, so set a dirty flag.
//   One property write. No sheet reads, no alarm, and — critically — NOTHING IS EVER
//   WRITTEN BACK TO AN IDENTITY CELL, so it cannot fight the operator or half-correct
//   a row.
//
//   The reconcile rides runPublishTick, once a minute, exactly like hold escalation.
//   A clean run costs ONE property read. A burst of twenty edits collapses into one
//   check, and that check sees the SETTLED row instead of a transient mid-undo state.
//
// ⚠ NO AUTOMATIC REVERT, deliberately. Revert is what was fighting the operator. This
//   flags, clears itself, and says what the row should be. Teeth can go on top of a
//   detector that has been proven calm — not before.
// =======================================================================================


var IDENTITY_GUARD = {
  // The two columns that say WHICH ROW THIS IS. QTY is deliberately unwatched: a qty
  // correction on a live row is ordinary work, and flagging it would spend the alert's
  // credibility on the normal case — the ruling that killed the "not counted" marker.
  watch: ["SKU", "SALES_ORDER"],

  flagBg:   "#ffe0b2",          // amber — a fact to check, not the red that means act now
  flagNote: "⚠ IDENTITY UNKNOWN",

  dirtyKey:   "IDENTITY_GUARD_DIRTY",
  alertedKey: "IDENTITY_GUARD_ALERTED",   // alert once per crossing, watchdog-style
  toggleKey:  "IDENTITY_GUARD_ALERTS",    // "off" silences Telegram, never the flag

  // How much Activity Log tail to trust as the known-good record. Open rows are recent
  // by construction — n8n sweeps shipped ones at ~1 AM — so a tail is enough, and
  // reading 90 days every minute would not be.
  logTailRows: 4000,

  alertedMax: 60                // prune; a live-issue list, not a history
};


// =======================================================================================
// PURE CORE — no Sheets, no network. Node-testable.
// =======================================================================================

/** Identity signature. Matches doPost's dedupe key so the two can never disagree. */
function _igSig(orderId, sku) {
  return String(orderId == null ? "" : orderId).trim().toLowerCase() + "|" +
         String(sku == null ? "" : sku).trim().toLowerCase();
}


/** Is this row live? A recognised STATUS — a blank one means it is being typed now. */
function _igIsEstablished(statusValue) {
  var s = String(statusValue == null ? "" : statusValue).trim();
  if (!s) return false;
  try { return Schema.isValidStatus(s); } catch (e) { return false; }
}


/**
 * THE VERDICT for one row. Pure — the part that decides whether to cry wolf.
 *
 * @returns {{verdict: string, reason: string}}  verdict: "ok" | "mismatch" | "skip"
 */
function _igVerdict(row, known) {
  var orderId = String(row.orderId == null ? "" : row.orderId).trim();
  var sku     = String(row.sku == null ? "" : row.sku).trim();

  // A half-typed row is not a mismatch — it is someone mid-entry.
  if (!orderId || !sku) return { verdict: "skip", reason: "incomplete row" };

  // ⚠ Only judge LIVE rows. A blank status means the row is being created right now
  //   and nothing has been received for it yet.
  if (!_igIsEstablished(row.status)) return { verdict: "skip", reason: "not established" };

  if (known.pairs[_igSig(orderId, sku)]) {
    return { verdict: "ok", reason: "matches a RECEIVED event" };
  }

  // ⚠⚠ ONLY FLAG WHERE THERE IS EVIDENCE. If neither the order nor the SKU appears in
  //   the tail at all, this row simply predates what we can see. Flagging it would be
  //   an accusation built on an absence — the same "mistook 'I could not read it' for
  //   'there is nothing there'" bug already fixed once in the stock audit.
  var haveOrder = !!known.orders[orderId.toLowerCase()];
  var haveSku   = !!known.skus[sku.toLowerCase()];
  if (!haveOrder && !haveSku) {
    return { verdict: "skip", reason: "no log evidence either way — older than the tail" };
  }

  return { verdict: "mismatch", reason: "this order/SKU pair was never received" };
}


/**
 * Compose the alert. Pure, so the wording is pinned by tests.
 * ⚠ It never claims to know the OLD value. The previous design's habit of asserting one
 *   is exactly what produced a message naming a correct value as the offender.
 */
function _igComposeAlert(hits) {
  var L = ["⚠ ROW IDENTITY DOES NOT MATCH ANY RECEIVED ORDER", ""];
  hits.forEach(function (h) {
    L.push("Row " + h.row + " · " + h.sku + " on " + h.orderId);
    if (h.suggest && h.suggest.length) {
      L.push("  this SKU was received on: " + h.suggest.join(" · "));
    } else {
      L.push("  no received order carries this SKU");
    }
    L.push("");
  });
  L.push("Nothing was changed. Fix the row and the flag clears within a minute.");
  return L.join("\n");
}


// =======================================================================================
// THE EDIT MARKER — deliberately almost nothing
// =======================================================================================

/**
 * Dispatched from onEditInstallable. Its whole job is to say "worth a look".
 *
 * ⚠ IT STORES NO ROW NUMBERS. A minute from now row 17 may be a different row — n8n
 *   inserts at the top all day. Storing coordinates captured before a delay is the
 *   2026-05-08 / 2026-08-21 row-shift bug class, twice bitten here. The reconcile
 *   re-derives everything from the sheet instead.
 */
function identityEditGuard(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== MAIN_SHEET_NAME) return;

  var firstCol = e.range.getColumn();
  var lastCol  = firstCol + e.range.getNumColumns() - 1;

  var touched = IDENTITY_GUARD.watch.some(function (name) {
    var c = Schema.cols[name];
    return c >= firstCol && c <= lastCol;
  });
  if (!touched) return;

  try {
    PropertiesService.getScriptProperties()
      .setProperty(IDENTITY_GUARD.dirtyKey, String(new Date().getTime()));
  } catch (err) {
    console.log("identityEditGuard: could not mark dirty: " + err);
  }
}


// =======================================================================================
// THE RECONCILE — rides runPublishTick, once a minute
// =======================================================================================

/**
 * Compare every live row's identity against the Activity Log, flag what does not match,
 * and CLEAR what does. Returns a short string for the publish log.
 *
 * ⭐ THE SELF-CLEARING HALF IS THE POINT. An event-based flag can never retire itself,
 *   because no edit means "it is fine now" — which is why the first design left rows
 *   amber after they had been corrected. A state check answers fresh every time, so
 *   fixing the row removes the flag on its own.
 */
function runIdentityReconcile() {
  var props = PropertiesService.getScriptProperties();

  // ⭐ CHEAP ON EVERY CLEAN RUN — one property read and out, the same shape as hold
  //   escalation. The sheet is touched only after an identity column was really edited.
  var dirty;
  try { dirty = props.getProperty(IDENTITY_GUARD.dirtyKey); } catch (e) { return "skip"; }
  if (!dirty) return "clean";

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "no sheet";

  var known = _igKnownFromLog(ss);
  if (!known) return "no log";           // cannot judge without the record — say so

  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) {
    props.deleteProperty(IDENTITY_GUARD.dirtyKey);
    return "empty";
  }

  var n = lastRow - Schema.dataStartRow + 1;
  var range = sheet.getRange(Schema.dataStartRow, 1, n, Schema.dataWidth);
  var data  = range.getValues();
  var bgs   = range.getBackgrounds();

  var boundary = -1;
  try { boundary = getBoundaryRow(); } catch (e) { boundary = -1; }

  var flagged = [], cleared = 0;

  for (var i = 0; i < n; i++) {
    var rowNum = Schema.dataStartRow + i;
    if (boundary > 0 && (rowNum === boundary || rowNum === boundary + 1)) continue;

    var r = data[i];
    var v = _igVerdict({
      orderId: r[Schema.idx("SALES_ORDER")],
      sku:     r[Schema.idx("SKU")],
      status:  r[Schema.idx("STATUS")]
    }, known);

    var isFlagged = String(bgs[i][0] || "").toLowerCase() === IDENTITY_GUARD.flagBg.toLowerCase();

    if (v.verdict === "mismatch") {
      if (!isFlagged) _igPaint(sheet, rowNum, true);
      flagged.push({
        row: rowNum,
        orderId: String(r[Schema.idx("SALES_ORDER")] || "").trim(),
        sku:     String(r[Schema.idx("SKU")] || "").trim(),
        suggest: _igOrdersForSku(known, r[Schema.idx("SKU")])
      });
    } else if (isFlagged) {
      // ⭐ SELF-CLEARING. It was flagged and now reconciles — or the row moved out of
      //   scope entirely. Either way the amber has done its job.
      _igPaint(sheet, rowNum, false);
      cleared++;
    }
  }

  props.deleteProperty(IDENTITY_GUARD.dirtyKey);

  // Alert once per crossing on a STABLE key. A row number is not stable, so the
  // identity signature is — the same rule the straggler watchdog runs on.
  var fresh = _igNewAlerts(flagged);
  if (fresh.length) {
    try { _igAlert(_igComposeAlert(fresh)); } catch (e) { console.log("identityEditGuard alert: " + e); }
  }

  return flagged.length + " flagged · " + cleared + " cleared" +
         (fresh.length ? " · " + fresh.length + " new" : "");
}


/** Build the known-good record: RECEIVED signatures from the Activity Log tail. */
function _igKnownFromLog(ss) {
  try {
    var log = ss.getSheetByName(ACTIVITY_LOG.sheetName);
    if (!log) return null;
    var last = log.getLastRow();
    if (last < ACTIVITY_LOG.dataStartRow) return null;

    var n = Math.min(IDENTITY_GUARD.logTailRows, last - ACTIVITY_LOG.dataStartRow + 1);
    var rows = log.getRange(last - n + 1, 1, n, ACTIVITY_LOG.dataWidth).getValues();

    var pairs = {}, orders = {}, skus = {}, bySku = {};
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][ACTIVITY_LOG.idx("EVENT")] || "").trim().toUpperCase() !== "RECEIVED") continue;
      var oid = String(rows[i][ACTIVITY_LOG.idx("ORDER_ID")] || "").trim();
      var sku = String(rows[i][ACTIVITY_LOG.idx("SKU")] || "").trim();
      if (!oid || !sku) continue;
      pairs[_igSig(oid, sku)] = 1;
      orders[oid.toLowerCase()] = 1;
      skus[sku.toLowerCase()] = 1;
      var k = sku.toLowerCase();
      if (!bySku[k]) bySku[k] = [];
      if (bySku[k].indexOf(oid) === -1 && bySku[k].length < 4) bySku[k].push(oid);
    }
    return { pairs: pairs, orders: orders, skus: skus, bySku: bySku };
  } catch (e) {
    console.log("_igKnownFromLog: " + e);
    return null;
  }
}


/** Which orders legitimately carry this SKU — the recovery hint. */
function _igOrdersForSku(known, sku) {
  var k = String(sku == null ? "" : sku).trim().toLowerCase();
  return (known && known.bySku && known.bySku[k]) ? known.bySku[k] : [];
}


/** Paint or unpaint one row. Cosmetic only — no value is ever touched. */
function _igPaint(sheet, row, on) {
  var band = sheet.getRange(row, 1, 1, Schema.dataWidth);
  var cell = sheet.getRange(row, Schema.cols.SALES_ORDER);
  if (on) {
    band.setBackground(IDENTITY_GUARD.flagBg);
    cell.setNote(IDENTITY_GUARD.flagNote +
      "\nThis order/SKU pair was never received. Nothing was changed." +
      "\nFix the row and this clears within a minute.");
  } else {
    // ⚠ null, not white — restores the row to the BANDING underneath rather than
    //   painting over it. Same anti-bleed the Zoho flag clear uses.
    band.setBackground(null);
    cell.clearNote();
  }
}


/** Alert-once-per-crossing, keyed on the identity signature rather than a row number. */
function _igNewAlerts(flagged) {
  var props = PropertiesService.getScriptProperties();
  var seen = {};
  try {
    var raw = props.getProperty(IDENTITY_GUARD.alertedKey);
    if (raw) seen = JSON.parse(raw) || {};
  } catch (e) { seen = {}; }

  var now = new Date().getTime();
  var fresh = [], live = {};

  flagged.forEach(function (h) {
    var k = _igSig(h.orderId, h.sku);
    live[k] = seen[k] || now;
    if (!seen[k]) fresh.push(h);
  });

  // ⚠ Keep only keys still flagged, so a fixed row can alert again if it recurs.
  var keys = Object.keys(live);
  if (keys.length > IDENTITY_GUARD.alertedMax) {
    keys.sort(function (a, b) { return live[a] - live[b]; });
    keys.slice(0, keys.length - IDENTITY_GUARD.alertedMax).forEach(function (k) { delete live[k]; });
  }
  try { props.setProperty(IDENTITY_GUARD.alertedKey, JSON.stringify(live)); } catch (e) {}
  return fresh;
}


/** Telegram the admin chat. Silenced by Script Property; the flag is never silenced. */
function _igAlert(text) {
  var off = "";
  try {
    off = PropertiesService.getScriptProperties().getProperty(IDENTITY_GUARD.toggleKey) || "";
  } catch (e) {}
  console.log("identityEditGuard:\n" + text);
  if (String(off).trim().toLowerCase() === "off") return;
  if (typeof TELEGRAM_ADMIN_CHAT_ID === "undefined" || !TELEGRAM_ADMIN_CHAT_ID) return;
  UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ chat_id: TELEGRAM_ADMIN_CHAT_ID, text: text }),
    muteHttpExceptions: true
  });
}


/**
 * Clear every identity flag by hand. Rarely needed now the reconcile clears its own,
 * but it is the escape hatch if a flag ever outlives its cause.
 */
function clearIdentityFlags() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";
  var last = sheet.getLastRow();
  if (last < Schema.dataStartRow) return "Nothing to clear.";

  var n = last - Schema.dataStartRow + 1;
  var bgs = sheet.getRange(Schema.dataStartRow, 1, n, Schema.dataWidth).getBackgrounds();
  var cleared = 0;
  for (var i = 0; i < bgs.length; i++) {
    if (String(bgs[i][0] || "").toLowerCase() !== IDENTITY_GUARD.flagBg.toLowerCase()) continue;
    _igPaint(sheet, Schema.dataStartRow + i, false);
    cleared++;
  }
  try { PropertiesService.getScriptProperties().deleteProperty(IDENTITY_GUARD.alertedKey); } catch (e) {}
  return "✅ Cleared " + cleared + " identity flag(s).";
}


/** Editor wrapper — the Run button shows no return value, so it logs. */
function checkIdentityNow() {
  try {
    PropertiesService.getScriptProperties()
      .setProperty(IDENTITY_GUARD.dirtyKey, String(new Date().getTime()));
  } catch (e) {}
  var out = runIdentityReconcile();
  console.log("identity reconcile: " + out);
  return out;
}
