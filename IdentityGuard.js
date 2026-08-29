// =======================================================================================
// IDENTITYGUARD.JS — does every row's identity match something we actually received?
// =======================================================================================
//
// WHY IT EXISTS
//   2026-08-28: a SALES_ORDER cell was overwritten by hand on a row already picked and
//   shelf-counted. doPost dedupes on SALES_ORDER + "|" + SKU, so the signature stopped
//   existing and n8n re-inserted the line four minutes later. Nothing alerted; the only
//   trace was two rows that looked like twins.
//
// ⚠⚠ THIS IS THE THIRD DESIGN, AND THE FIRST ONE THAT PUTS ANYTHING ON THE SHEET.
//
//   v1 compared e.oldValue against e.value on every edit and reverted. It fought the
//   operator: Ctrl+Z re-entered the handler, restoring the CORRECT value still alarmed,
//   and a reverted SKU left a mixed row because liveUpdateTrigger had already written
//   LOCATION and HAND for the new one.
//
//   v2 asked about STATE instead of the EVENT — the right question — but still PAINTED
//   the row, and every remaining problem came from that one write: the paint could not be
//   told from the #fff8e7 banding, clearing depended on a re-check undo did not queue,
//   and a stale mark had no way to retire itself. v2 ended read-only: a Telegram message
//   and nothing on the sheet at all.
//
//   ⭐⭐ v3 — THE MARK IS CONDITIONAL FORMATTING, NOT A STATIC FILL. CF is a DISPLAY
//   LAYER: Sheets recomputes it live from a formula and stores NOTHING on the row. Fix
//   the cell and the red vanishes in the same keystroke — no trigger, no property, no
//   clearing pass, nothing that can go stale. Ctrl+Z is just another value change.
//   Every failure listed above is structurally impossible.
//
//   ⚠ WHAT THIS FILE WRITES, AND WHERE. It never touches a data row. It writes the LIST
//   OF MISMATCHED PAIRS to a hidden __Identity sheet, and the CF rule does nothing but
//   look a row's own pair up in that list. Inverting the helper that way is what makes
//   the timing safe — see _igWriteMismatchList.
//
// ⭐ THE EVIDENCE IS ALREADY ON THE SHEET. Every row that legitimately exists has a
//   RECEIVED entry in the Activity Log carrying its order id and SKU — written by doPost,
//   Zoho pull, kit expansion, manual entry and /missing. That is the SAME SALES_ORDER|SKU
//   signature doPost dedupes on, so the log IS the record of what a row's identity is
//   allowed to be.
//
//   ⚠⚠ AND UNTIL 2026-08-30 A HAND-EDIT WROTE ITS OWN ALIBI INTO THAT RECORD.
//   manualReceiveOnEdit logged RECEIVED for EVERY non-empty change to column D, so
//   overwriting a live SALES_ORDER legitimised the corrupted value a minute before this
//   file ever looked at it — and the verdict came back "ok". The 08-28 incident was
//   undetectable by construction. Fixed in _mrClassify (OrderService.js): a replacement
//   of a complete identity now logs IDENTITY_EDIT, which is audit only and never evidence.
//
// HOW IT RUNS
//   onEdit does almost nothing: an identity column was touched, so set a dirty flag. One
//   property write, no sheet reads, nothing written back to a cell.
//
//   The reconcile rides runPublishTick, once a minute, exactly like hold escalation.
//   A clean run costs ONE property read.
//
//   ⭐ THE ONCE-A-MINUTE CADENCE IS THE GRACE PERIOD, AND IT IS LOAD-BEARING. doPost
//   inserts rows (OrderService.js:854) BEFORE writing their RECEIVED events (:889), and
//   insertRowsBefore forces a flush — so for a moment a brand-new order exists with no
//   evidence. A live formula reading the log would paint it red. Listing a pair only
//   AFTER the reconcile has judged it removes that window entirely.
//
//   ⚠ SO THE ASYMMETRY IS DELIBERATE: red APPEARS within a minute, and VANISHES on the
//   keystroke. Same shape as every override on the board — retire on evidence, expire on
//   a timer.
//
// WHAT IT CATCHES (three rules, painted by BrandTheme.js on cols A + D)
//   GONE        an established row missing its SKU or its SALES_ORDER   · local, instant
//   UNKNOWN     the pair was never received                             · via the list
//   DUPLICATED  the same pair on two rows — the copied-row case         · local, instant
//
// ⚠ NO AUTOMATIC REVERT, deliberately. Revert is what fought the operator in v1. This
//   observes, marks and reports. Teeth can go on top of a detector proven calm.
// =======================================================================================


var IDENTITY_GUARD = {
  // The columns that make a row judgeable. SKU + SALES_ORDER are the identity itself;
  // STATUS is watched because a row only becomes judgeable ONCE IT HAS ONE — type an
  // identity, get skipped as "not established", then set the status a minute later and
  // nothing would ever look again. QTY is deliberately unwatched: a qty correction on a
  // live row is ordinary work, and flagging it would spend the alert's credibility on
  // the normal case — the ruling that killed the "not counted" marker.
  watch: ["SKU", "SALES_ORDER", "STATUS"],

  // The hidden sheet holding the mismatch list the CF rule reads. ⚠ Its name is also
  // hardcoded inside _buildIdentityRules (BrandTheme.js) because a CF formula cannot
  // take a variable — change it in both places or the rule silently stops matching.
  sheetName: "__Identity",
  listMax:   500,               // matches the CF rules' $A$1:$A$500 / $B$1:$B$500 windows
  // Two published columns, because the two questions need different keys:
  //   A · order|sku       — is this row's IDENTITY one we received?
  //   B · order|sku|qty   — is this row's QTY the one we received for that identity?
  listCols:  { identity: 1, qty: 2 },

  dirtyKey:   "IDENTITY_GUARD_DIRTY",
  alertedKey: "IDENTITY_GUARD_ALERTED",   // alert once per crossing, watchdog-style
  toggleKey:  "IDENTITY_GUARD_ALERTS",    // "off" silences Telegram, never the mark

  // How much Activity Log tail to trust as the known-good record.
  // ⚠ THIS IS THE CONSTANT THAT BOUNDS THE ONE REAL FALSE POSITIVE. A row on the sheet
  //   whose RECEIVED has aged out of this window, but whose SKU still sells, reads as a
  //   mismatch. Live rows are recent by construction and n8n sweeps shipped ones at
  //   ~1 AM, so in practice that means a long-lived CANCELED row. If one ever flags,
  //   raise this — it is not a redesign.
  logTailRows: 8000,

  // The note Zoho Pull writes on a legitimate second row for the same pair. ⚠ MUST stay
  // in step with ZohoPull.js's _noteOverride and with the COUNTIFS criterion inside
  // _buildIdentityRules. There is a drift test.
  deltaNoteToken: "delta from Zoho",

  alertedMax: 60,               // prune; a live-issue list, not a history

  // GONE · UNKNOWN · DUPLICATED · QTY. Kept here so the diagnostics cannot drift from
  // what _buildIdentityRules actually installs.
  ruleCount: 4,

  // ⚠ LEGACY ONLY — the colours the PAINTING versions wore, kept so clearIdentityFlags()
  //   can still remove marks they left. Nothing paints a static fill any more.
  legacyBg: ["#ffab40", "#ffe0b2"],
  legacyNote: "⚠ IDENTITY UNKNOWN"
};


// =======================================================================================
// PURE CORE — no Sheets, no network. Node-testable.
// =======================================================================================

/** Identity signature. Matches doPost's dedupe key so the two can never disagree. */
function _igSig(orderId, sku) {
  return String(orderId == null ? "" : orderId).trim().toLowerCase() + "|" +
         String(sku == null ? "" : sku).trim().toLowerCase();
}


/**
 * Identity + quantity. A row can carry a perfectly legitimate identity and still have had
 * its QTY typed over, so the qty question needs its own key.
 *
 * ⚠ Nothing in this codebase writes Schema.cols.QTY after insert — verified by grep, and
 *   Zoho Pull's insert_delta creates a NEW row with its own RECEIVED event rather than
 *   editing an existing qty. So a qty that does not match what was received is always a
 *   hand-edit, which is what makes this safe to flag.
 */
function _igQtySig(orderId, sku, qty) {
  return _igSig(orderId, sku) + "|" + String(qty == null ? "" : qty).trim();
}


/** Is this background one the PAINTING versions left behind? Cleanup only. */
function _igIsLegacyFlag(bg) {
  var c = String(bg || "").toLowerCase();
  return IDENTITY_GUARD.legacyBg.some(function (o) { return c === o.toLowerCase(); });
}


/**
 * Should the reconcile do real work this minute? PURE, so the thing that decides how
 * often the sheet is touched is testable on its own.
 *
 * ⭐ ONE REASON ONLY, now. v2 also swept on a heartbeat while any flag was live, because
 *   a painted flag could not clear itself. The CF clears itself, so there is nothing to
 *   sweep for.
 */
function _igShouldReconcile(dirty) {
  return !!dirty;
}


/** Is this row live? A recognised STATUS — a blank one means it is being typed now. */
function _igIsEstablished(statusValue) {
  var s = String(statusValue == null ? "" : statusValue).trim();
  if (!s) return false;
  try { return Schema.isValidStatus(s); } catch (e) { return false; }
}


/** Does this row's NOTE mark it as Zoho Pull's legitimate delta twin? */
function _igIsDeltaRow(note) {
  return String(note == null ? "" : note)
           .toLowerCase()
           .indexOf(IDENTITY_GUARD.deltaNoteToken.toLowerCase()) !== -1;
}


/**
 * Reduce a block of All Orders values to the rows worth judging, and count how many
 * times each identity pair appears. PURE.
 *
 * ⚠ DELTA ROWS ARE EXCLUDED FROM THE COUNT, NOT FROM THE SCAN. Zoho Pull's insert_delta
 *   legitimately creates a second row carrying the same pair, and the note it writes
 *   exists precisely to tell that apart from a duplicate. Counting only the non-delta
 *   rows means a delta twin reads 1 and stays quiet, while a copied row reads 2.
 *
 * @returns {{rows: Array, pairCounts: Object}}
 */
function _igScanRows(data, boundary) {
  var rows = [], pairCounts = {};

  for (var i = 0; i < data.length; i++) {
    var rowNum = Schema.dataStartRow + i;
    if (boundary > 0 && (rowNum === boundary || rowNum === boundary + 1)) continue;

    var r = data[i];
    var sku    = String(r[Schema.idx("SKU")] || "").trim();
    var so     = String(r[Schema.idx("SALES_ORDER")] || "").trim();
    var note   = String(r[Schema.idx("NOTE")] || "").trim();
    var status = r[Schema.idx("STATUS")];

    if (!_igIsEstablished(status)) continue;

    var qty = r[Schema.idx("QTY")];
    var entry = { row: rowNum, sku: sku, orderId: so, note: note, qty: qty,
                  status: status, sig: _igSig(so, sku),
                  qtySig: _igQtySig(so, sku, qty) };
    rows.push(entry);

    if (sku && so && !_igIsDeltaRow(note)) {
      pairCounts[entry.sig] = (pairCounts[entry.sig] || 0) + 1;
    }
  }

  return { rows: rows, pairCounts: pairCounts };
}


/**
 * THE VERDICT for one row. Pure — the part that decides whether to cry wolf.
 *
 * ⚠ ORDER MATTERS. "gone" comes first because an incomplete row cannot be looked up at
 *   all; "mismatch" outranks "duplicate" because a wrong identity is more informative
 *   than a repeated one when a row is somehow both.
 *
 * @param {Object} row         from _igScanRows
 * @param {Object} known       from _igKnownFromLog
 * @param {Object} pairCounts  from _igScanRows
 * @returns {{verdict: string, reason: string}}
 *          verdict: "ok" | "gone" | "mismatch" | "duplicate" | "skip"
 */
function _igVerdict(row, known, pairCounts) {
  var orderId = String(row.orderId == null ? "" : row.orderId).trim();
  var sku     = String(row.sku == null ? "" : row.sku).trim();

  // ⚠ THE STATUS GATE COMES FIRST, and it must — the CF rule checks it first too, and a
  //   verdict function that trusted its caller to have filtered would silently disagree
  //   with the red on the sheet. A row with no recognised status is being typed right
  //   now, or is a header; either way it is not ours to judge.
  if (!_igIsEstablished(row.status)) return { verdict: "skip", reason: "not established" };

  // ⚠ AN ESTABLISHED ROW WITH HALF AN IDENTITY IS BROKEN, NOT "mid-entry". v2 returned
  //   skip here and was therefore completely blind to the likeliest slip of all —
  //   hitting Delete on the wrong cell.
  if (!orderId || !sku) {
    return { verdict: "gone", reason: (!orderId && !sku) ? "row has no identity at all"
                                    : (!orderId ? "SALES ORDER is missing" : "SKU is missing") };
  }

  var sig = _igSig(orderId, sku);

  if (!known || !known.pairs) {
    // Cannot judge without the record. Duplication is still knowable from the sheet
    // alone, so fall through to that rather than returning nothing useful.
    if (pairCounts && pairCounts[sig] > 1) {
      return { verdict: "duplicate", reason: "this exact pair is on more than one row" };
    }
    return { verdict: "skip", reason: "Activity Log unreadable — no verdict possible" };
  }

  if (!known.pairs[sig]) {
    // ⚠⚠ ONLY FLAG WHERE THERE IS EVIDENCE. If neither the order nor the SKU appears in
    //   the tail at all, this row simply predates what we can see. Flagging it would be
    //   an accusation built on an absence — the same "mistook 'I could not read it' for
    //   'there is nothing there'" bug already fixed once in the stock audit.
    var haveOrder = !!known.orders[orderId.toLowerCase()];
    var haveSku   = !!known.skus[sku.toLowerCase()];
    if (haveOrder || haveSku) {
      return { verdict: "mismatch", reason: "this order/SKU pair was never received" };
    }
    return { verdict: "skip", reason: "no log evidence either way — older than the tail" };
  }

  if (pairCounts && pairCounts[sig] > 1) {
    return { verdict: "duplicate", reason: "this exact pair is on more than one row" };
  }

  // ⚠ THE IDENTITY IS RIGHT AND THE QUANTITY IS NOT. Judged only once the pair is known,
  //   so a row we cannot vouch for at all is never also accused of a wrong qty — one
  //   problem per row, and the more fundamental one wins.
  var qtyRaw = String(row.qty == null ? "" : row.qty).trim();
  if (qtyRaw && known.qtyByPair) {
    var seen = known.qtyByPair[sig];
    if (seen && !seen[qtyRaw]) {
      return { verdict: "qty",
               reason: "we received " + Object.keys(seen).join(" or ") + " of this, not " + qtyRaw };
    }
  }

  return { verdict: "ok", reason: "matches a RECEIVED event" };
}


/**
 * Compose the alert. Pure, so the wording is pinned by tests.
 * ⚠ It never claims to know the OLD value. v1's habit of asserting one is exactly what
 *   produced a message naming a correct value as the offender.
 */
function _igComposeAlert(hits) {
  var byKind = { gone: [], mismatch: [], duplicate: [], qty: [] };
  hits.forEach(function (h) { if (byKind[h.verdict]) byKind[h.verdict].push(h); });

  var L = ["⚠ ROW IDENTITY PROBLEM ON THE SHEET", ""];

  if (byKind.gone.length) {
    L.push("── HALF AN IDENTITY ──");
    byKind.gone.forEach(function (h) {
      L.push("Row " + h.row + " · " + (h.sku || "(no SKU)") + " on " + (h.orderId || "(no order)"));
      L.push("  " + h.reason);
    });
    L.push("");
  }

  if (byKind.mismatch.length) {
    L.push("── NEVER RECEIVED ──");
    byKind.mismatch.forEach(function (h) {
      L.push("Row " + h.row + " · " + h.sku + " on " + h.orderId);
      if (h.suggest && h.suggest.length) {
        L.push("  this SKU was received on: " + h.suggest.join(" · "));
      } else {
        L.push("  no received order carries this SKU");
      }
    });
    L.push("");
  }

  if (byKind.duplicate.length) {
    L.push("── THE SAME LINE TWICE ──");
    byKind.duplicate.forEach(function (h) {
      L.push("Row " + h.row + " · " + h.sku + " on " + h.orderId);
    });
    L.push("  A duplicate depresses HAND for that SKU and doubles the pick line.");
    L.push("");
  }

  if (byKind.qty.length) {
    L.push("── QUANTITY CHANGED ──");
    byKind.qty.forEach(function (h) {
      L.push("Row " + h.row + " · " + h.sku + " on " + h.orderId);
      L.push("  " + h.reason);
    });
    L.push("");
  }

  L.push("The cells are marked red on the sheet. Nothing was changed — fix the row and");
  L.push("the mark clears itself immediately.");
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
 * Judge every established row, publish the mismatch list the CF rule reads, and Telegram
 * anything new. Returns a short string for the publish log.
 *
 * ⭐ IT WRITES TO A HIDDEN SHEET AND NOWHERE ELSE. No data row is ever touched, which is
 *   what removes every failure mode v1 and v2 had.
 */
function runIdentityReconcile() {
  var props = PropertiesService.getScriptProperties();

  // ⭐ CHEAP ON EVERY CLEAN RUN — one property read and out, the same shape as hold
  //   escalation. The sheet is touched only after a watched column was really edited.
  var dirty = null;
  try {
    dirty = props.getProperty(IDENTITY_GUARD.dirtyKey);
  } catch (e) { return "skip"; }

  if (!_igShouldReconcile(dirty)) return "clean";

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "no sheet";

  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) {
    props.deleteProperty(IDENTITY_GUARD.dirtyKey);
    return "empty";
  }

  var n = lastRow - Schema.dataStartRow + 1;
  var data = sheet.getRange(Schema.dataStartRow, 1, n, Schema.dataWidth).getValues();

  var boundary = -1;
  try { boundary = getBoundaryRow(); } catch (e) { boundary = -1; }

  var scan  = _igScanRows(data, boundary);
  var known = _igKnownFromLog(ss);

  var flagged = [], mismatchPairs = [], qtyKeys = [];

  for (var i = 0; i < scan.rows.length; i++) {
    var entry = scan.rows[i];
    var v = _igVerdict(entry, known, scan.pairCounts);
    if (v.verdict === "ok" || v.verdict === "skip") continue;

    flagged.push({
      row:     entry.row,
      orderId: entry.orderId,
      sku:     entry.sku,
      verdict: v.verdict,
      reason:  v.reason,
      suggest: _igOrdersForSku(known, entry.sku)
    });

    // ⚠ ONLY "mismatch" NEEDS THE LIST. "gone" and "duplicate" are decidable from the
    //   row alone, so their CF rules are local formulas — instant in both directions,
    //   with no dependency on this function having run.
    if (v.verdict === "mismatch" && mismatchPairs.indexOf(entry.sig) === -1) {
      mismatchPairs.push(entry.sig);
    }
    if (v.verdict === "qty" && qtyKeys.indexOf(entry.qtySig) === -1) {
      qtyKeys.push(entry.qtySig);
    }
  }

  try { _igWriteMismatchList(ss, mismatchPairs, qtyKeys); }
  catch (e) { console.log("identityEditGuard list write: " + e); }

  props.deleteProperty(IDENTITY_GUARD.dirtyKey);

  // Alert once per crossing on a STABLE key. A row number is not stable, so the identity
  // signature is — the same rule the straggler watchdog runs on.
  var fresh = _igNewAlerts(flagged);
  if (fresh.length) {
    try { _igAlert(_igComposeAlert(fresh)); }
    catch (e) { console.log("identityEditGuard alert: " + e); }
  }

  return flagged.length + " flagged" + (fresh.length ? " · " + fresh.length + " reported" : "");
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

    var pairs = {}, orders = {}, skus = {}, bySku = {}, qtyByPair = {};
    for (var i = 0; i < rows.length; i++) {
      // ⚠ RECEIVED ONLY. IDENTITY_EDIT rows are in this same log and are deliberately
      //   NOT counted — an alteration is not a receipt. See _mrClassify.
      if (String(rows[i][ACTIVITY_LOG.idx("EVENT")] || "").trim().toUpperCase() !== "RECEIVED") continue;
      var oid = String(rows[i][ACTIVITY_LOG.idx("ORDER_ID")] || "").trim();
      var sku = String(rows[i][ACTIVITY_LOG.idx("SKU")] || "").trim();
      if (!oid || !sku) continue;
      var psig = _igSig(oid, sku);
      pairs[psig] = 1;
      // ⚠ EVERY qty ever received for this pair, not just the latest. A pair can be
      //   received more than once — a re-entry, or Zoho Pull's delta row — and each is
      //   legitimate, so the row matches if it equals ANY of them.
      var q = String(rows[i][ACTIVITY_LOG.idx("QTY")] == null ? "" : rows[i][ACTIVITY_LOG.idx("QTY")]).trim();
      if (q) {
        if (!qtyByPair[psig]) qtyByPair[psig] = {};
        qtyByPair[psig][q] = 1;
      }
      orders[oid.toLowerCase()] = 1;
      skus[sku.toLowerCase()] = 1;
      var k = sku.toLowerCase();
      if (!bySku[k]) bySku[k] = [];
      if (bySku[k].indexOf(oid) === -1 && bySku[k].length < 4) bySku[k].push(oid);
    }
    return { pairs: pairs, orders: orders, skus: skus, bySku: bySku, qtyByPair: qtyByPair };
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


/**
 * ⭐⭐ THE HELPER HOLDS THE VERDICT, NOT THE EVIDENCE — and that is the whole reason the
 * on-sheet mark is safe.
 *
 * The obvious design was a formula mirror of the Activity Log for the CF to search. It
 * is wrong: doPost inserts rows (OrderService.js:854) BEFORE writing their RECEIVED
 * events (:889), and insertRowsBefore forces a flush, so every legitimate arrival would
 * flash red in the gap — and a single failed logActivityBatch, which is best-effort by
 * design, would leave a real order red permanently.
 *
 * Listing only pairs the reconcile has ALREADY judged removes that window: the once-a-
 * minute cadence is the grace period. A stale entry is harmless, because no row matches
 * it any more — and if someone re-types that exact bad value it goes red instantly.
 *
 * ⚠ Rewritten only when the set CHANGES. Usually empty, so usually zero writes.
 * ⚠ '@' plain text on the column — Gotcha #16. An order id like 24-14979-87359 left on a
 *   default format would be coerced to a Date and could never match the row again.
 */
function _igWriteMismatchList(ss, pairs, qtyKeys) {
  var sheet = ss.getSheetByName(IDENTITY_GUARD.sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(IDENTITY_GUARD.sheetName);
    sheet.getRange(1, 1, IDENTITY_GUARD.listMax, 2).setNumberFormat('@');
  }
  try { sheet.hideSheet(); } catch (e) { /* already hidden — fine */ }

  var wrote = 0;
  wrote += _igWriteOneList(sheet, ss, IDENTITY_GUARD.listCols.identity, pairs);
  wrote += _igWriteOneList(sheet, ss, IDENTITY_GUARD.listCols.qty, qtyKeys);
  return wrote;
}


/** One published column. Idempotent: a steady state costs nothing. */
function _igWriteOneList(sheet, ss, col, keys) {
  var capped = (keys || []).slice(0, IDENTITY_GUARD.listMax);
  var existing = _igReadMismatchList(ss, col);
  if (existing.length === capped.length &&
      existing.every(function (v, i) { return v === capped[i]; })) {
    return 0;
  }
  sheet.getRange(1, col, IDENTITY_GUARD.listMax, 1).clearContent();
  if (capped.length) {
    sheet.getRange(1, col, capped.length, 1)
         .setNumberFormat('@')
         .setValues(capped.map(function (p) { return [p]; }));
  }
  return capped.length;
}


/** Read one published column back. Returns [] when the sheet is absent. */
function _igReadMismatchList(ss, col) {
  try {
    var c = col || IDENTITY_GUARD.listCols.identity;
    var target = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = target.getSheetByName(IDENTITY_GUARD.sheetName);
    if (!sheet) return [];
    var last = Math.min(sheet.getLastRow(), IDENTITY_GUARD.listMax);
    if (last < 1) return [];
    return sheet.getRange(1, c, last, 1).getValues()
      .map(function (r) { return String(r[0] || "").trim(); })
      .filter(function (v) { return !!v; });
  } catch (e) {
    console.log("_igReadMismatchList: " + e);
    return [];
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
    // ⚠ The verdict is part of the key: a row that goes from duplicate to mismatch is a
    //   different fact and deserves to be said once more.
    var k = h.verdict + ":" + _igSig(h.orderId, h.sku);
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


/** Telegram the admin chat. Silenced by Script Property; the sheet mark is never silenced. */
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


// =======================================================================================
// READERS — used by the sidebar Alerts card
// =======================================================================================

/**
 * Classify rows the way the CF rules do — from the SHEET and the published list, never
 * from the Activity Log.
 *
 * ⭐ THIS IS WHY THE SIDEBAR COUNT CANNOT GO STALE. It is recomputed from the same two
 *   inputs the CF reads, so the badge and the red cells always agree. A snapshot count
 *   parked in a Script Property would have drifted the moment a row was fixed — which is
 *   the exact complaint this whole feature exists to answer.
 *
 * @param {Array}  data      All Orders values, from Schema.dataStartRow
 * @param {number} boundary
 * @param {Array}  list      published mismatch pairs
 * @returns {Array} row numbers with a problem
 */
function _igIssueRows(data, boundary, list, qtyList) {
  var set = {}, qset = {};
  (list || []).forEach(function (p) { set[p] = 1; });
  (qtyList || []).forEach(function (p) { qset[p] = 1; });

  var scan = _igScanRows(data, boundary);
  var out = [];

  scan.rows.forEach(function (entry) {
    var bad = (!entry.sku || !entry.orderId) ||          // GONE
              (set[entry.sig] === 1) ||                   // UNKNOWN
              (qset[entry.qtySig] === 1) ||               // WRONG QTY
              (scan.pairCounts[entry.sig] > 1);           // DUPLICATED
    if (bad) out.push(entry.row);
  });

  return out;
}


// =======================================================================================
// DIAGNOSTICS + LEGACY CLEANUP
// =======================================================================================

/**
 * ⚠ CLEANUP ONLY. Nothing paints a static fill any more — the mark is conditional
 *   formatting. This survives solely so clearIdentityFlags() can remove marks left by
 *   the versions that did, in either colour they wore.
 */
function _igPaint(sheet, row, on) {
  var band = sheet.getRange(row, 1, 1, Schema.dataWidth);
  var cell = sheet.getRange(row, Schema.cols.SALES_ORDER);
  if (on) throw new Error("_igPaint: painting was removed 2026-08-29 — see the file header.");
  // ⚠ null, not white — restores the row to the BANDING underneath rather than painting
  //   over it. Same anti-bleed the Zoho flag clear uses.
  band.setBackground(null);
  cell.clearNote();
}


/**
 * Remove every mark left by the painting versions of this guard — both colours, plus the
 * note. Run once after 2026-08-29; after that there is nothing to clear, because the
 * guard writes nothing to a data row.
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
    if (!_igIsLegacyFlag(bgs[i][0])) continue;
    _igPaint(sheet, Schema.dataStartRow + i, false);
    cleared++;
  }
  try { PropertiesService.getScriptProperties().deleteProperty(IDENTITY_GUARD.alertedKey); } catch (e) {}
  return "✅ Cleared " + cleared + " legacy identity flag(s).";
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
 * How many of the three identity CF rules are actually installed on the sheet?
 *
 * ⚠⚠ WHY THIS EXISTS. 2026-08-30: the whole feature can be perfectly correct and mark
 *   NOTHING, for one boring reason — setupIdentityHighlighting() was never run, so there
 *   are no rules to evaluate. And because the mark is a DISPLAY LAYER, an uninstalled
 *   feature and a clean sheet look exactly the same. There is no residue to notice.
 *   Every diagnostic here must answer this FIRST.
 *
 * Identified by range signature, the same way _stripIdentityRules finds them.
 */
function _igCountInstalledRules(sheet) {
  try {
    var rules = sheet.getConditionalFormatRules();
    var n = 0;
    rules.forEach(function (rule) {
      var ranges = rule.getRanges();
      if (!ranges) return;
      // GONE · UNKNOWN · DUPLICATED — two single-column ranges, on SKU and SALES ORDER
      if (ranges.length === 2) {
        var cols = ranges.map(function (r) {
          return (r.getNumColumns() === 1) ? r.getColumn() : -1;
        }).sort(function (x, y) { return x - y; });
        if (cols[0] === Schema.cols.SKU && cols[1] === Schema.cols.SALES_ORDER) { n++; return; }
      }
      // QTY — column B alone, naming our helper sheet
      if (ranges.length === 1 && ranges[0].getNumColumns() === 1 &&
          ranges[0].getColumn() === Schema.cols.QTY) {
        var bc = rule.getBooleanCondition();
        var f = bc ? String((bc.getCriteriaValues() || [''])[0] || '') : '';
        if (f.indexOf(IDENTITY_GUARD.sheetName) !== -1) n++;
      }
    });
    return n;
  } catch (e) {
    console.log("_igCountInstalledRules: " + e);
    return -1;
  }
}


/**
 * Explain EVERY flagged row: what it holds, what verdict it gets and why, and what the
 * Activity Log says about its order and its SKU.
 *
 * ⚠⚠ IT NEVER LOOKS AT CELL BACKGROUNDS, AND THAT IS THE POINT. getBackgrounds() returns
 *   STATIC fills ONLY — conditional formatting and row banding are DISPLAY LAYERS and
 *   never appear in it. The v2 version of this function hunted for its own paint, which
 *   is both why it could not answer "then whose colour is that?" and why it would report
 *   nothing at all now. It computes the verdict instead.
 *
 * ⚠ Zero-arg — the editor Run button cannot pass one, a trap this project has walked
 *   into three times. Output goes to the EXECUTION LOG.
 */
function diagnoseIdentityFlags() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) { console.log("no sheet"); return "no sheet"; }

  var last = sheet.getLastRow();
  if (last < Schema.dataStartRow) { console.log("no data"); return "no data"; }

  var n = last - Schema.dataStartRow + 1;
  var data = sheet.getRange(Schema.dataStartRow, 1, n, Schema.dataWidth).getValues();

  var boundary = -1;
  try { boundary = getBoundaryRow(); } catch (e) { boundary = -1; }

  var scan  = _igScanRows(data, boundary);
  var known = _igKnownFromLog(ss);
  var list  = _igReadMismatchList(ss, IDENTITY_GUARD.listCols.identity);
  var qlist = _igReadMismatchList(ss, IDENTITY_GUARD.listCols.qty);

  var installed = _igCountInstalledRules(sheet);

  var L = ["── IDENTITY ──"];
  L.push("CF rules installed:       " + installed + " of " + IDENTITY_GUARD.ruleCount +
         (installed === IDENTITY_GUARD.ruleCount ? "   ✓" : "   ⚠⚠ THE SHEET CANNOT MARK ANYTHING"));
  if (installed !== IDENTITY_GUARD.ruleCount) {
    L.push("");
    L.push("  ⚠⚠ NOTHING WILL EVER TURN RED UNTIL THIS READS " +
           IDENTITY_GUARD.ruleCount + " OF " + IDENTITY_GUARD.ruleCount + ".");
    L.push("     Run  setupIdentityHighlighting()  once, from the editor.");
    L.push("     The mark is conditional formatting — a display layer — so an");
    L.push("     UNINSTALLED feature and a clean sheet look identical. There is no");
    L.push("     residue to notice, which is why this line comes first.");
    L.push("");
  }
  L.push("established rows judged:  " + scan.rows.length);
  L.push("published identity list:   " + list.length + " pair(s)   [" +
         IDENTITY_GUARD.sheetName + "!A]");
  L.push("published qty list:        " + qlist.length + " key(s)    [" +
         IDENTITY_GUARD.sheetName + "!B]");
  L.push("Activity Log:             " + (known ? "readable (tail " + IDENTITY_GUARD.logTailRows + ")"
                                               : "UNREADABLE — no verdict possible"));
  var dirty = null;
  try { dirty = PropertiesService.getScriptProperties().getProperty(IDENTITY_GUARD.dirtyKey); } catch (e) {}
  L.push("pending check queued:     " + (dirty ? "YES — the next publish tick will act" : "no"));
  L.push("");
  L.push("⚠ The mark is CONDITIONAL FORMATTING, so getBackgrounds() cannot see it and");
  L.push("  neither can this. What follows is the verdict, recomputed from scratch.");
  L.push("");

  var bad = 0;
  scan.rows.forEach(function (entry) {
    var v = _igVerdict(entry, known, scan.pairCounts);
    if (v.verdict === "ok" || v.verdict === "skip") return;
    bad++;
    if (bad > 15) return;
    L.push("Row " + entry.row + "  " + v.verdict.toUpperCase() + "  —  " + v.reason);
    L.push("   SKU " + (entry.sku || "(blank)") + "   ORDER " + (entry.orderId || "(blank)"));
    if (v.verdict === "mismatch") {
      L.push("   in the published list? " + (list.indexOf(entry.sig) !== -1 ? "YES — cells are red" :
             "no — the reconcile has not run since this appeared; red lands within a minute"));
      var sug = _igOrdersForSku(known, entry.sku);
      L.push("   orders that received this SKU: " + (sug.length ? sug.join(" · ") : "(none)"));
    }
    if (v.verdict === "duplicate") {
      L.push("   times this pair appears (delta rows excluded): " + scan.pairCounts[entry.sig]);
    }
    L.push("");
  });

  if (!bad) L.push("Nothing is flagged. Every established row's identity matches a RECEIVED event.");
  else if (bad > 15) L.push("… and " + (bad - 15) + " more");

  var out = L.join("\n");
  console.log(out);
  return out;
}


/**
 * ⭐⭐ THE SELF-TEST. Reports what is ACTUALLY on the sheet, and — the part that matters —
 * EVALUATES THE REAL FORMULAS IN REAL CELLS and reads back what Sheets returns.
 *
 * ⚠⚠ WHY IT EXISTS. 2026-08-30: this feature was reported doing nothing twice running, and
 *   both rounds of reasoning about why were wrong (first the strippers, then a theory about
 *   key formats). Reasoning about a CF rule from outside Sheets is guessing. This asks
 *   Sheets.
 *
 *   It writes the SAME strings _buildIdentityRules installs — from the one shared
 *   _identityFormulas builder — into a scratch column, reads the booleans back, then clears
 *   them. If a rule says FALSE here, the formula is wrong. If it says TRUE and the cell is
 *   not red, the rule is missing or something precedes it. Those are different bugs and
 *   this is what tells them apart.
 *
 * ⚠ Zero-arg — the editor Run button cannot pass one. Output goes to the EXECUTION LOG.
 */
function diagnoseIdentityCF() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) { console.log("no sheet"); return "no sheet"; }

  var L = ["══ IDENTITY · CF SELF-TEST ══", ""];

  // ── 1 · what rules are actually on the sheet, and in what order ────────────────────
  var all = sheet.getConditionalFormatRules();
  var onIdentityCols = [];
  all.forEach(function (r, i) {
    var rs = r.getRanges(), hit = false;
    rs.forEach(function (g) {
      if (g.getColumn() === Schema.cols.SKU || g.getColumn() === Schema.cols.SALES_ORDER) hit = true;
    });
    if (hit) onIdentityCols.push({ i: i, rule: r, ranges: rs });
  });

  var installed = _igCountInstalledRules(sheet);
  L.push("CF rules on the sheet:        " + all.length);
  L.push("…touching col A or col D:     " + onIdentityCols.length);
  L.push("…that are OURS:               " + installed + " of " + IDENTITY_GUARD.ruleCount +
         (installed === IDENTITY_GUARD.ruleCount ? "   ✓" : "   ⚠⚠ NOTHING CAN MARK"));
  L.push("");
  L.push("EVERY rule touching A or D, IN ORDER (Sheets applies the FIRST match per cell):");
  if (!onIdentityCols.length) L.push("  (none — run setupIdentityHighlighting())");
  onIdentityCols.forEach(function (e) {
    var bc = e.rule.getBooleanCondition();
    var f = bc ? (bc.getCriteriaValues() || [''])[0] : '(not a formula rule)';
    var cols = e.ranges.map(function (g) {
      return g.getA1Notation();
    }).join(", ");
    L.push("  [" + e.i + "] " + cols);
    L.push("       " + String(f).slice(0, 150));
  });
  L.push("");

  // ── 2 · the published list ─────────────────────────────────────────────────────────
  var identSheet = ss.getSheetByName(IDENTITY_GUARD.sheetName);
  var list  = _igReadMismatchList(ss, IDENTITY_GUARD.listCols.identity);
  var qlist = _igReadMismatchList(ss, IDENTITY_GUARD.listCols.qty);
  L.push("__Identity sheet exists:      " + (identSheet ? "YES" : "NO  ⚠ neither lookup rule can match"));
  L.push("published identity pairs (A): " + list.length);
  list.slice(0, 5).forEach(function (p) { L.push("     " + JSON.stringify(p)); });
  L.push("published qty keys (B):       " + qlist.length);
  qlist.slice(0, 5).forEach(function (p) { L.push("     " + JSON.stringify(p)); });
  L.push("");

  // ── 3 · pick a row to test: prefer one the reconcile flagged ───────────────────────
  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) { L.push("no data rows"); console.log(L.join("\n")); return L.join("\n"); }
  var n = lastRow - Schema.dataStartRow + 1;
  var data = sheet.getRange(Schema.dataStartRow, 1, n, Schema.dataWidth).getValues();
  var boundary = -1;
  try { boundary = getBoundaryRow(); } catch (e) {}
  var scan = _igScanRows(data, boundary);

  var target = null;
  for (var i = 0; i < scan.rows.length; i++) {
    if (list.indexOf(scan.rows[i].sig) !== -1) { target = scan.rows[i]; break; }
  }
  if (!target && scan.rows.length) target = scan.rows[0];
  if (!target) { L.push("no established rows to test"); console.log(L.join("\n")); return L.join("\n"); }

  L.push("── TEST ROW " + target.row + (list.indexOf(target.sig) !== -1 ?
         "  (the reconcile flagged this one — it SHOULD be red)" :
         "  (not flagged — expect FALSE everywhere, which is correct)") + " ──");
  var raw = sheet.getRange(target.row, 1, 1, Schema.dataWidth).getValues()[0];
  L.push("  A (SKU)          " + JSON.stringify(raw[Schema.idx("SKU")]) +
         "   typeof " + (typeof raw[Schema.idx("SKU")]));
  L.push("  D (SALES ORDER)  " + JSON.stringify(raw[Schema.idx("SALES_ORDER")]) +
         "   typeof " + (typeof raw[Schema.idx("SALES_ORDER")]));
  L.push("  F (STATUS)       " + JSON.stringify(raw[Schema.idx("STATUS")]) +
         "   typeof " + (typeof raw[Schema.idx("STATUS")]));
  L.push("  sig JS builds:   " + JSON.stringify(target.sig));
  L.push("");

  // ── 4 · ⭐ EVALUATE THE REAL FORMULAS IN REAL CELLS ────────────────────────────────
  var scratchCol = Math.max(Schema.dataWidth + 6, 16);
  if (sheet.getMaxColumns() < scratchCol) {
    L.push("⚠ not enough columns for a scratch evaluation (need col " + scratchCol + ")");
    console.log(L.join("\n")); return L.join("\n");
  }

  var F = _identityFormulas(target.row);
  var probes = [
    ["status guard", F.established],
    ["GONE",         F.gone],
    ["UNKNOWN",      F.unknown],
    ["DUPLICATED",   F.duplicated],
    ["QTY",          F.qty],
    ["sig the FORMULA builds",
      '=LOWER(TRIM($D' + target.row + '))&"|"&LOWER(TRIM($A' + target.row + '))'],
    ["MATCH position in the list",
      '=IFERROR(MATCH(LOWER(TRIM($D' + target.row + '))&"|"&LOWER(TRIM($A' + target.row +
      ')), INDIRECT("\'' + IDENTITY_GUARD.sheetName + '\'!$A$1:$A$' + IDENTITY_GUARD.listMax +
      '"), 0), "NO MATCH")'],
    ["the list, seen through INDIRECT",
      '=IFERROR(COUNTA(INDIRECT("\'' + IDENTITY_GUARD.sheetName + '\'!$A$1:$A$' +
      IDENTITY_GUARD.listMax + '")), "INDIRECT FAILED")']
  ];

  var cell = sheet.getRange(target.row, scratchCol);
  L.push("── WHAT SHEETS ACTUALLY RETURNS (evaluated live, then cleared) ──");
  try {
    probes.forEach(function (p) {
      cell.setFormula(p[1]);
      SpreadsheetApp.flush();
      var v = cell.getValue();
      L.push("  " + (p[0] + "                          ").slice(0, 26) + JSON.stringify(v));
    });
  } catch (e) {
    L.push("  ⚠ evaluation threw: " + e);
  } finally {
    try { cell.clearContent(); } catch (e) {}
  }

  L.push("");
  L.push("── HOW TO READ THIS ──");
  L.push("  UNKNOWN=true  + rule present + cell not red  → something PRECEDES our rule above.");
  L.push("  UNKNOWN=false + a MATCH position             → the guard clauses are the problem.");
  L.push("  'NO MATCH' but the sigs are equal            → the list is not where INDIRECT looks.");
  L.push("  INDIRECT FAILED                              → the __Identity sheet is missing.");
  L.push("  rules OURS < 3                               → run setupIdentityHighlighting().");

  var out = L.join("\n");
  console.log(out);
  return out;
}
