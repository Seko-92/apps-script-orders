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

  // ⚠⚠ #ffe0b2 WAS TOO CLOSE TO THE BANDING. The row banding is #fff8e7 cream, and a
  //   pale amber beside it is indistinguishable by eye — so a CLEARED row and a FLAGGED
  //   row looked the same, and 2026-08-29 was spent chasing a highlight that was just
  //   banding. A signal you cannot tell from the background is not a signal.
  //   #ffab40 is unmistakable and still not the red that means "act now".
  flagBg:   "#ffab40",
  // ⚠ Any colour this flag has EVER worn. A row still wearing an old one would
  //   otherwise be unclearable — the clear looks for the current colour only, so a
  //   rename would strand every flag painted before it. Never remove entries.
  legacyBg: ["#ffe0b2"],
  flagNote: "⚠ IDENTITY UNKNOWN",

  dirtyKey:   "IDENTITY_GUARD_DIRTY",
  alertedKey: "IDENTITY_GUARD_ALERTED",   // alert once per crossing, watchdog-style
  toggleKey:  "IDENTITY_GUARD_ALERTS",    // "off" silences Telegram, never the flag

  // How much Activity Log tail to trust as the known-good record. Open rows are recent
  // by construction — n8n sweeps shipped ones at ~1 AM — so a tail is enough, and
  // reading 90 days every minute would not be.
  logTailRows: 4000,

  alertedMax: 60,               // prune; a live-issue list, not a history

  // ⚠⚠ A FLAG MUST NEVER OUTLIVE ITS CAUSE. Clearing used to happen only on a run that
  //   an EDIT had queued — and that assumption broke in production: a Ctrl+Z after a SKU
  //   edit pops liveUpdateTrigger's LOCATION/HAND write rather than the SKU itself, so
  //   nothing queued a re-check and the amber simply sat there. Rather than chase which
  //   operation undo pops, remove the dependency: while any flag is live, re-check on a
  //   slow heartbeat regardless. Rare and short-lived, so a few extra reads cost nothing.
  sweepEveryMin: 5
};


// =======================================================================================
// PURE CORE — no Sheets, no network. Node-testable.
// =======================================================================================

/** Is this background one of OURS — current colour or any it has worn before? */
function _igIsOurFlag(bg) {
  var c = String(bg || "").toLowerCase();
  if (c === IDENTITY_GUARD.flagBg.toLowerCase()) return true;
  return IDENTITY_GUARD.legacyBg.some(function (o) { return c === o.toLowerCase(); });
}


/** Identity signature. Matches doPost's dedupe key so the two can never disagree. */
function _igSig(orderId, sku) {
  return String(orderId == null ? "" : orderId).trim().toLowerCase() + "|" +
         String(sku == null ? "" : sku).trim().toLowerCase();
}


/**
 * Should the reconcile do real work this minute? PURE, so the thing that decides how
 * often the sheet is touched is testable on its own.
 *
 * ⚠ TWO REASONS TO RUN, and the second is the one added after production:
 *   · something was edited            → check now
 *   · a flag is live                  → keep checking until it clears, because the fix
 *                                        may not arrive as an edit we can see
 */
function _igShouldReconcile(dirty, hasLiveFlags, minute) {
  if (dirty) return true;
  if (!hasLiveFlags) return false;
  var every = IDENTITY_GUARD.sweepEveryMin;
  return (Number(minute) % every) === 0;
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
  var dirty = null, alerted = null;
  try {
    dirty   = props.getProperty(IDENTITY_GUARD.dirtyKey);
    alerted = props.getProperty(IDENTITY_GUARD.alertedKey);
  } catch (e) { return "skip"; }

  // ⭐ A live flag keeps the check alive. See sweepEveryMin — clearing must not depend
  //   on the correction arriving as an edit, because in practice it does not.
  var hasLiveFlags = !!(alerted && alerted !== "{}" && alerted.length > 2);
  if (!_igShouldReconcile(dirty, hasLiveFlags, new Date().getMinutes())) {
    return hasLiveFlags ? "waiting" : "clean";
  }

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

    var isFlagged = _igIsOurFlag(bgs[i][0]);

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
    if (!_igIsOurFlag(bgs[i][0])) continue;
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


/**
 * Explain ONE row in full: what it holds, whether it is actually flagged by US, what
 * verdict it gets and why, and what the Activity Log says about its order and its SKU.
 *
 * ⚠ WHY THIS EXISTS. 2026-08-29: a row stayed amber after being corrected and there was
 *   no way to tell, from looking, whether it was our flag, the row banding (#fff8e7),
 *   the duplicate-SO highlight (#fff3b0) or a Zoho flag — nor whether the reconcile had
 *   judged it at all. Three plausible causes, one appearance. This answers it in one Run
 *   instead of another round of inference.
 *
 * ⚠ Output goes to the EXECUTION LOG — the Run button shows no return value, a lesson
 *   that has now cost this project two evenings.
 */
function diagnoseIdentityRow(rowNumber) {
  var row = Number(rowNumber);
  if (!row || row < Schema.dataStartRow) {
    var msg = "Pass a data row number, e.g. diagnoseIdentityRow(12)";
    console.log(msg); return msg;
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) { console.log("no sheet"); return "no sheet"; }

  var vals = sheet.getRange(row, 1, 1, Schema.dataWidth).getValues()[0];
  var bg   = sheet.getRange(row, 1).getBackground();
  var note = sheet.getRange(row, Schema.cols.SALES_ORDER).getNote();

  var orderId = String(vals[Schema.idx("SALES_ORDER")] || "").trim();
  var sku     = String(vals[Schema.idx("SKU")] || "").trim();
  var status  = String(vals[Schema.idx("STATUS")] || "").trim();

  var L = ["── ROW " + row + " ──"];
  L.push("SKU:      " + (sku || "(blank)"));
  L.push("ORDER:    " + (orderId || "(blank)"));
  L.push("STATUS:   " + (status || "(blank)") + (_igIsEstablished(status) ? "  [established]" : "  [not established — skipped]"));
  L.push("");

  // ⭐ THE DISCRIMINATOR. Several things paint a row tan here; only ours writes the note.
  var ours = String(bg || "").toLowerCase() === IDENTITY_GUARD.flagBg.toLowerCase();
  L.push("background:  " + bg + (ours ? "   ← OUR flag" : "   ← NOT our flag"));
  if (!ours) {
    L.push("             (" + IDENTITY_GUARD.flagBg + " is ours; #fff8e7 is the row banding,");
    L.push("              #fff3b0 is the duplicate-SO highlight)");
  }
  L.push("note:        " + (note ? JSON.stringify(note.slice(0, 60)) : "(none)") +
         (note && note.indexOf(IDENTITY_GUARD.flagNote) === 0 ? "   ← ours" : ""));
  L.push("");

  var known = _igKnownFromLog(ss);
  if (!known) {
    L.push("Activity Log: UNREADABLE — no verdict is possible.");
    console.log(L.join("\n")); return L.join("\n");
  }

  var v = _igVerdict({ orderId: orderId, sku: sku, status: status }, known);
  L.push("VERDICT:  " + v.verdict.toUpperCase() + "  —  " + v.reason);
  L.push("");
  L.push("evidence in the last " + IDENTITY_GUARD.logTailRows + " log rows:");
  L.push("  this exact order+SKU pair received?  " + (known.pairs[_igSig(orderId, sku)] ? "YES" : "no"));
  L.push("  this ORDER seen at all?              " + (known.orders[orderId.toLowerCase()] ? "YES" : "no"));
  L.push("  this SKU seen at all?                " + (known.skus[sku.toLowerCase()] ? "YES" : "no"));
  var sug = _igOrdersForSku(known, sku);
  L.push("  orders that received this SKU:       " + (sug.length ? sug.join(" · ") : "(none)"));
  L.push("");

  var dirty = null;
  try { dirty = PropertiesService.getScriptProperties().getProperty(IDENTITY_GUARD.dirtyKey); } catch (e) {}
  L.push("pending check queued: " + (dirty ? "YES — the next publish tick will act" : "no"));
  if (ours && v.verdict !== "mismatch" && !dirty) {
    L.push("");
    L.push("⚠ FLAGGED BUT RECONCILES. Nothing has queued a re-check, so the flag is");
    L.push("  stale — run checkIdentityNow() and it will clear.");
  }

  var out = L.join("\n");
  console.log(out);
  return out;
}


/**
 * Explain EVERY flagged row. No argument, because the editor Run button cannot pass one —
 * a trap this project has now walked into three times, most recently with
 * diagnoseIdentityRow, which answered "Pass a data row number" and nothing else.
 *
 * ⚠ Output goes to the EXECUTION LOG.
 */
function diagnoseIdentityFlags() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) { console.log("no sheet"); return "no sheet"; }

  var last = sheet.getLastRow();
  if (last < Schema.dataStartRow) { console.log("no data"); return "no data"; }

  var n = last - Schema.dataStartRow + 1;
  var bgs = sheet.getRange(Schema.dataStartRow, 1, n, Schema.dataWidth).getBackgrounds();
  var rows = [];
  for (var i = 0; i < n; i++) {
    if (String(bgs[i][0] || "").toLowerCase() === IDENTITY_GUARD.flagBg.toLowerCase()) {
      rows.push(Schema.dataStartRow + i);
    }
  }

  var head = "── IDENTITY FLAGS ──\n" + rows.length + " row(s) currently wearing OUR amber (" +
             IDENTITY_GUARD.flagBg + ")";

  // ⚠⚠ ALWAYS REPORT WHAT IS ACTUALLY THERE. 2026-08-29: a row was visibly tan while the
  //   reconcile reported "0 flagged · 0 cleared" — it could not clear paint it had not
  //   applied, and there was no way to learn the real colour except by asking the sheet.
  //   A diagnostic that only looks for its OWN colour cannot answer "then whose is it?",
  //   which was the entire question.
  var seen = {};
  for (var j = 0; j < n; j++) {
    var c = String(bgs[j][0] || "").toLowerCase();
    if (!seen[c]) seen[c] = { count: 0, rows: [] };
    seen[c].count++;
    if (seen[c].rows.length < 6) seen[c].rows.push(Schema.dataStartRow + j);
  }
  var WHOSE = {
    "#ffab40": "the identity guard  ← OURS",
    "#ffe0b2": "the identity guard (OLD colour, pre-2026-08-29)  ← OURS",
    "#fff8e7": "row banding (normal, alternating)",
    "#ffffff": "NO static fill — what you SEE here is the banding or a CF rule",
    "#ffd400": "the DIRECT divider band",
    "#1d1d1b": "a header row (banner row 3, or the DIRECT header)",
    "#fff3b0": "duplicate SALES ORDER highlight — setupDuplicateSalesOrderHighlighting",
    "#ffe5e5": "Zoho removed/qty-changed flag — _flagDirectRow",
    "#ffe5e5": "Zoho removed/qty-changed flag — _flagDirectRow"
  };
  head += "\n\nEVERY background actually present in column A:";
  Object.keys(seen).sort(function (a, b) { return seen[b].count - seen[a].count; })
    .forEach(function (c) {
      head += "\n  " + (c || "(none)") + "  ×" + seen[c].count +
              "   rows " + seen[c].rows.join(", ") + (seen[c].count > 6 ? " …" : "") +
              "\n      " + (WHOSE[c] || "⚠ UNRECOGNISED — not painted by anything we know of");
    });

  if (!rows.length) {
    head += "\n\nNothing is flagged by the guard, so a tan row you can see was painted by\n" +
            "something else — find its colour above. clearIdentityFlags() only removes\n" +
            IDENTITY_GUARD.flagBg + " and will not touch it.";
    console.log(head);
    return head;
  }

  var out = [head, ""];
  rows.slice(0, 10).forEach(function (r) { out.push(diagnoseIdentityRow(r)); out.push(""); });
  if (rows.length > 10) out.push("… and " + (rows.length - 10) + " more");
  var txt = out.join("\n");
  console.log(txt);
  return txt;
}
