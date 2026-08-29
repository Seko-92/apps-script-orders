// =======================================================================================
// IDENTITYGUARD.JS — the backstop for a row's identity columns
// =======================================================================================
//
// WHY THIS EXISTS, AFTER THE LOCK
//   `protectAllOrdersSheet()` (BrandTheme.js) hard-locks SKU / QTY / SALES ORDER against
//   everyone who is not the owner. **Protection does not restrict the owner.** So the
//   exact 2026-08-28 slip — `09-15094-35132` overwritten with `Missing #: 05-15052-93025`
//   on a row already picked and shelf-counted — is still available to you, and to every
//   script path that runs as you.
//
//   The lock stops the floor. This stops you.
//
// ⚠ FLAG, NEVER AUTO-REVERT. This codebase's standing rule is "annotate, never override"
//   — the same call the Zoho removed-line flag made. An auto-revert would fight someone
//   legitimately correcting a typo'd order id, and it would do so silently, which is the
//   worse failure. Paint it, say it, and let a human decide.
//
// ⚠⚠ THE SHAPE THAT MAKES THIS HARD, AND IT IS THE ONE THAT ACTUALLY HAPPENED.
//   `e.oldValue` is populated for SINGLE-CELL edits ONLY. The 08-28 event at 12:31:43 was
//   a MULTI-ROW PASTE, where it is `undefined` — so a guard keyed on the old value misses
//   one of the two real shapes. For multi-cell edits this keys on whether the row looks
//   ESTABLISHED instead: a recognised STATUS. Nothing auto-fills STATUS on manual entry
//   (checked — liveUpdateTrigger writes LOCATION and HAND, never STATUS), so a row being
//   typed fresh has a blank one while a live row does not.
//
// ⭐ THE WINDOW IS REAL. On 08-28 there were FOUR MINUTES between the overwrite and n8n's
//   re-insert. An alert lands well inside it, which is why carrying the OLD VALUE matters:
//   it makes the row recoverable by hand rather than merely mourned.
//
// RUNS AS THE OWNER (installable trigger), so it can still write its flag onto the very
// columns the lock protects.
// =======================================================================================


var IDENTITY_GUARD = {
  // Only these two carry a row's identity. QTY is locked by the sheet protection but is
  // deliberately NOT watched here: a qty correction on a live row is ordinary work, and
  // flagging it would spend the alert's credibility on the normal case — the standing
  // ruling that killed the "not counted" marker and the always-on dashboard.
  watch: ["SKU", "SALES_ORDER"],

  flagBg:   "#ffe0b2",   // amber — a fact to check, NOT the red that means "act now"
  flagNote: "⚠ IDENTITY CHANGED",

  // Cap the rows a single event may flag. A hundred-row paste is a bulk operation, not a
  // slip, and painting all of it would bury the signal in its own noise.
  maxRowsPerEvent: 25,

  // Script Property: set to "off" to silence the Telegram half without touching code.
  alertToggleKey: "IDENTITY_GUARD_ALERTS"
};


// =======================================================================================
// PURE CORE — no Sheets, no network. Node-testable.
// =======================================================================================

/**
 * Is this row an ESTABLISHED one, i.e. was it live before the edit?
 *
 * Signal: a recognised STATUS. A row being typed fresh has a blank one until the person
 * picks from the dropdown; an n8n insert writes PENDING immediately. Deliberately does
 * NOT use LOCATION or HAND — `liveUpdateTrigger` fills both the moment a SKU is typed,
 * so a brand-new row would look established within a second and every fresh manual entry
 * would false-positive.
 */
function _igIsEstablished(statusValue) {
  var s = String(statusValue == null ? "" : statusValue).trim();
  if (!s) return false;
  try { return Schema.isValidStatus(s); } catch (e) { return false; }
}


/**
 * THE DECISION. Pure, so the part that decides whether to cry wolf is testable alone.
 *
 * @param {{singleCell: boolean, oldValue: *, newValue: *, status: *}} o
 * @returns {{flag: boolean, oldKnown: boolean, reason: string}}
 */
function _igDecide(o) {
  o = o || {};
  var oldV = String(o.oldValue == null ? "" : o.oldValue).trim();
  var newV = String(o.newValue == null ? "" : o.newValue).trim();

  if (o.singleCell) {
    // A no-op edit (retype the same value, or Escape) must never alarm.
    if (oldV === newV) return { flag: false, oldKnown: true, reason: "unchanged" };
    // Filling a BLANK identity cell is ordinary — that is how a manual row is created.
    if (!oldV) return { flag: false, oldKnown: true, reason: "was blank — new row" };
    return { flag: true, oldKnown: true, reason: "overwrote an existing value" };
  }

  // MULTI-CELL: e.oldValue is undefined, so fall back to the row's own state.
  // ⚠ This is the branch the 12:31:43 paste took. A guard without it misses the
  //   incident it was built for.
  if (_igIsEstablished(o.status)) {
    return { flag: true, oldKnown: false, reason: "multi-cell edit on an established row" };
  }
  return { flag: false, oldKnown: false, reason: "row not established — looks like new entry" };
}


/** Compose the alert. Pure, so the wording is pinned by tests. */
function _igComposeAlert(hits) {
  var L = ["⚠ ROW IDENTITY CHANGED", ""];
  hits.forEach(function (h) {
    L.push("Row " + h.row + " · " + h.column);
    L.push("  now:  " + (h.newValue || "(blank)"));
    if (h.oldKnown) {
      L.push("  was:  " + (h.oldValue || "(blank)"));
    } else {
      // Honest about the limit rather than silently omitting it — the recovery
      // candidates below are what make this branch still actionable.
      L.push("  was:  unknown (multi-cell edit — Sheets does not report the old value)");
    }
    if (h.recovery && h.recovery.length) {
      L.push("  recently logged for this SKU: " + h.recovery.join(" · "));
    }
    L.push("");
  });
  L.push("The row was NOT changed back — this is a flag, not a revert.");
  L.push("Undo in the sheet (Ctrl+Z) if it was a slip.");
  return L.join("\n");
}


// =======================================================================================
// THE HANDLER
// =======================================================================================

/**
 * Dispatched from `onEditInstallable` (Main.js) in its own try/catch, like every other
 * handler there — defense in depth, so a throw here can never take the others down.
 */
function identityEditGuard(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== MAIN_SHEET_NAME) return;

  var firstRow = e.range.getRow();
  var numRows  = e.range.getNumRows();
  var firstCol = e.range.getColumn();
  var numCols  = e.range.getNumColumns();
  var lastCol  = firstCol + numCols - 1;

  // Which watched columns does this edit actually touch?
  var touched = [];
  IDENTITY_GUARD.watch.forEach(function (name) {
    var c = Schema.cols[name];
    if (c >= firstCol && c <= lastCol) touched.push({ name: name, col: c });
  });
  if (!touched.length) return;

  var singleCell = (numRows === 1 && numCols === 1);
  if (numRows > IDENTITY_GUARD.maxRowsPerEvent) return;   // a bulk op, not a slip

  // Read the whole affected block ONCE — status for the established test, plus the
  // current values of the watched columns and the SKU for the recovery hint.
  var block = sheet.getRange(firstRow, 1, numRows, Schema.dataWidth).getValues();

  var boundary = -1;
  try { boundary = getBoundaryRow(); } catch (err) { boundary = -1; }

  var hits = [];
  for (var i = 0; i < numRows; i++) {
    var row = firstRow + i;
    if (row < Schema.dataStartRow) continue;                       // banner / headers
    if (boundary > 0 && (row === boundary || row === boundary + 1)) continue;  // divider

    var rowVals = block[i];
    var status  = rowVals[Schema.idx("STATUS")];

    for (var j = 0; j < touched.length; j++) {
      var t = touched[j];
      var newValue = rowVals[t.col - 1];
      var d = _igDecide({
        singleCell: singleCell,
        oldValue:   singleCell ? e.oldValue : undefined,
        newValue:   newValue,
        status:     status
      });
      if (!d.flag) continue;

      hits.push({
        row: row,
        column: t.name,
        newValue: String(newValue == null ? "" : newValue),
        oldValue: singleCell ? String(e.oldValue == null ? "" : e.oldValue) : "",
        oldKnown: d.oldKnown,
        sku: String(rowVals[Schema.idx("SKU")] || ""),
        recovery: []
      });
    }
  }
  if (!hits.length) return;

  // Recovery candidates — only on the branch that needs them, and only for a handful
  // of rows, because it costs an Activity Log read.
  hits.forEach(function (h) {
    if (h.oldKnown || !h.sku) return;
    try { h.recovery = _igRecentOrdersForSku(h.sku); } catch (err) { h.recovery = []; }
  });

  // Paint first — it is the durable half and must survive a Telegram failure.
  hits.forEach(function (h) {
    try { _igFlagRow(sheet, h.row); } catch (err) {
      console.log("identityEditGuard: paint failed on row " + h.row + ": " + err);
    }
  });

  try { _igAlert(_igComposeAlert(hits)); } catch (err) {
    console.log("identityEditGuard: alert failed: " + err);
  }
}


/** Amber wash + a cell note naming what happened. Cosmetic only — no values change. */
function _igFlagRow(sheet, row) {
  sheet.getRange(row, 1, 1, Schema.dataWidth).setBackground(IDENTITY_GUARD.flagBg);
  sheet.getRange(row, Schema.cols.SALES_ORDER)
       .setNote(IDENTITY_GUARD.flagNote + " · " +
                Utilities.formatDate(new Date(), "America/Chicago", "M/d h:mm a") +
                "\nThe row was not changed back. Ctrl+Z if this was a slip.");
}


/**
 * Order ids recently logged against this SKU — the recovery candidates for a
 * multi-cell edit, where Sheets gives us no old value. Best-effort by contract.
 */
function _igRecentOrdersForSku(sku) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var log = ss.getSheetByName(ACTIVITY_LOG.sheetName);
  if (!log) return [];
  var last = log.getLastRow();
  if (last < ACTIVITY_LOG.dataStartRow) return [];

  var n = Math.min(400, last - ACTIVITY_LOG.dataStartRow + 1);   // recent tail only
  var start = last - n + 1;
  var rows = log.getRange(start, 1, n, ACTIVITY_LOG.dataWidth).getValues();
  var want = String(sku).trim().toLowerCase();
  var seen = {}, out = [];
  for (var i = rows.length - 1; i >= 0 && out.length < 3; i--) {
    if (String(rows[i][ACTIVITY_LOG.idx("SKU")] || "").trim().toLowerCase() !== want) continue;
    var oid = String(rows[i][ACTIVITY_LOG.idx("ORDER_ID")] || "").trim();
    if (!oid || seen[oid]) continue;
    seen[oid] = 1;
    out.push(oid);
  }
  return out;
}


/** Telegram the admin chat. Silenced by Script Property, best-effort otherwise. */
function _igAlert(text) {
  var off = "";
  try { off = PropertiesService.getScriptProperties().getProperty(IDENTITY_GUARD.alertToggleKey) || ""; } catch (e) {}
  if (String(off).trim().toLowerCase() === "off") { console.log("identityEditGuard (alerts off):\n" + text); return; }

  console.log("identityEditGuard:\n" + text);
  if (typeof TELEGRAM_ADMIN_CHAT_ID === "undefined" || !TELEGRAM_ADMIN_CHAT_ID) return;
  UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ chat_id: TELEGRAM_ADMIN_CHAT_ID, text: text }),
    muteHttpExceptions: true
  });
}


/**
 * Clear every identity flag. The companion that makes the paint safe to live with —
 * without it the sheet slowly accumulates amber rows nobody can clear.
 */
function clearIdentityFlags() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";

  var last = sheet.getLastRow();
  if (last < Schema.dataStartRow) return "Nothing to clear.";

  var n = last - Schema.dataStartRow + 1;
  var range = sheet.getRange(Schema.dataStartRow, 1, n, Schema.dataWidth);
  var bgs = range.getBackgrounds();
  var cleared = 0;

  for (var i = 0; i < bgs.length; i++) {
    if (String(bgs[i][0] || "").toLowerCase() !== IDENTITY_GUARD.flagBg.toLowerCase()) continue;
    // ⚠ null, not white — restores the row to the BANDING underneath rather than
    //   painting over it, the same anti-bleed the Zoho flag clear uses.
    sheet.getRange(Schema.dataStartRow + i, 1, 1, Schema.dataWidth).setBackground(null);
    sheet.getRange(Schema.dataStartRow + i, Schema.cols.SALES_ORDER).clearNote();
    cleared++;
  }
  return "✅ Cleared " + cleared + " identity flag(s).";
}
