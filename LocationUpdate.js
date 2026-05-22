// =======================================================================================
// LOCATION_UPDATE.gs — "Location Update" sheet for tracking SKU→Location changes////
// =======================================================================================
//
// PURPOSE
//   Warehouse employees use this sheet to record SKUs whose physical location
//   changed. They type an SKU in column B; the system auto-fills:
//     - Col A: sequential counter (for the eye — "how many have I done today")
//     - Col C: LOCATION from Master Inventory
//     - Col D: timestamp in Houston time
//
// ARCHITECTURE — INSTALLABLE TRIGGER ONLY
//   The original implementation ran in the SIMPLE onEdit trigger
//   (locationUpdateTimestamp in Timestampfeature.js, now orphaned). Simple
//   triggers can fail silently when openById is called and permissions aren't
//   granted — that's why the location lookup and timestamp would sometimes
//   "fail to appear." This version runs in onEditInstallable, which has full
//   permissions. Same pattern as prepQueueOnEdit and outOfStockOnEdit.
//
// PUBLIC API
//   setupLocationUpdateSheet()       — one-time: brand-theme, headers, banding
//   openLocationUpdate()             — sidebar: switch active sheet
//   refreshLocationUpdateSheet()     — sidebar: sweep all rows, fill any blanks
//                                       (manual escape hatch when auto-fill missed)
//   sortLocationUpdateByLocation()   — sidebar: sort data rows by LOCATION A→Z
//   locationUpdateOnEdit(e)          — installable trigger dispatcher
// =======================================================================================

// ---------- LOCAL SCHEMA (kept here, not in Schema.js — different sheet) ----------
//
// Schema v3 (2026-05-14): user-requested simplification — LOCATION is now
// MANUALLY edited by the picker (no auto-fill from Master Inventory). The
// auto-fill verification cue was removed in favor of simplicity. Only
// TIMESTAMP is still auto-stamped (on SKU edit).
var LOCATION_UPDATE = {
  sheetName: "Location Update",

  // 1-based column positions
  cols: {
    COUNTER:        1,   // A — formula-driven row label (=ROW()-1), always visible
    SKU:            2,   // B — manually typed
    LOCATION:       3,   // C — manually typed (no auto-fill)
    TIMESTAMP:      4,   // D — auto-stamped at SKU edit time
    EMPLOYEE:       5,   // E — dropdown, user-selected
    FINAL_CHECK_BY: 6    // F — dropdown, user-selected
  },

  idx: function(name) { return LOCATION_UPDATE.cols[name] - 1; },

  dataWidth:    6,
  headerRow:    1,
  dataStartRow: 2,

  headers: ["#", "◈ SKU", "LOCATION", "⏱ TIMESTAMP", "👤 EMPLOYEE", "✓ FINAL CHECK BY"]
};


// =======================================================================================
// PUBLIC API
// =======================================================================================

/**
 * One-time setup: creates "Location Update" sheet if missing, applies brand
 * styling matching Prep Queue and Out of Stock. Idempotent — safe to re-run.
 *
 * NOTE: previous versions of this sheet may have used a 2-row header (title
 * + column labels). This function writes the single-row brand header to row 1
 * and styles row 2+ as the data area. Existing data in row 3+ stays where it
 * is; if a user wants to compact it up, that's a manual fix.
 */
function setupLocationUpdateSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(LOCATION_UPDATE.sheetName);
  }

  // --- HEADERS ---
  sheet.getRange(LOCATION_UPDATE.headerRow, 1, 1, LOCATION_UPDATE.dataWidth)
    .setValues([LOCATION_UPDATE.headers])
    .setBackground('#1d1d1b')   // brand black
    .setFontColor('#ffd966')    // brand yellow
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // Thick yellow underline below header
  sheet.getRange(LOCATION_UPDATE.headerRow, 1, 1, LOCATION_UPDATE.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(LOCATION_UPDATE.cols.COUNTER,         55);
  sheet.setColumnWidth(LOCATION_UPDATE.cols.SKU,            130);
  sheet.setColumnWidth(LOCATION_UPDATE.cols.LOCATION,       130);
  sheet.setColumnWidth(LOCATION_UPDATE.cols.TIMESTAMP,      170);
  sheet.setColumnWidth(LOCATION_UPDATE.cols.EMPLOYEE,       130);
  sheet.setColumnWidth(LOCATION_UPDATE.cols.FINAL_CHECK_BY, 150);

  // --- DATA AREA: column-level format (so new rows inherit) ---
  var maxDataRow = 1000;
  var dataRows = maxDataRow - LOCATION_UPDATE.dataStartRow + 1;

  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.COUNTER, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(11)
    .setFontColor('#5f5f5f').setHorizontalAlignment('center');
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.SKU, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.LOCATION, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.TIMESTAMP, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('center');
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.EMPLOYEE, dataRows, 1)
    .setFontFamily('Roboto').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.FINAL_CHECK_BY, dataRows, 1)
    .setFontFamily('Roboto').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');

  sheet.getRange(LOCATION_UPDATE.dataStartRow, 1, dataRows, LOCATION_UPDATE.dataWidth)
    .setVerticalAlignment('middle');

  // --- COUNTER FORMULA (col A) ---
  // =ROW()-1 in every data row. Always reflects current row position; survives
  // SKU clear (never disappears); auto-corrects on row insert/delete. Per-cell
  // formula (not ArrayFormula) so deleting one cell doesn't blank the column.
  // The refresh button re-paints these in case a cell got cleared accidentally.
  var counterFormulas = [];
  for (var cf = 0; cf < dataRows; cf++) counterFormulas.push(["=ROW()-1"]);
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.COUNTER, dataRows, 1)
    .setFormulas(counterFormulas);

  // --- DATA VALIDATION on Employee + Final Check By (dropdowns) ---
  // Placeholder list so the dropdown widget renders immediately. User edits
  // the list via Data → Data validation, OR calls setLocationUpdateDropdowns()
  // from the Apps Script editor with the real staff names.
  //
  // setAllowInvalid(true) means the cell accepts any typed value while the
  // list is still the placeholder — so warehouse staff aren't blocked before
  // the lists are configured.
  //
  // IMPORTANT: only install the placeholder if validation isn't already set on
  // these columns. Otherwise re-running setup would wipe out the user's
  // configured staff lists. We check the first data cell as a proxy for
  // "has this column been configured yet."
  var placeholderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["—"], true)
    .setAllowInvalid(true)
    .setHelpText("Edit this dropdown's list via Data → Data validation, or run setLocationUpdateDropdowns() from the script editor.")
    .build();
  var empProbe = sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.EMPLOYEE).getDataValidation();
  if (!empProbe) {
    sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.EMPLOYEE, dataRows, 1)
      .setDataValidation(placeholderRule);
  }
  var fcProbe = sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.FINAL_CHECK_BY).getDataValidation();
  if (!fcProbe) {
    sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.FINAL_CHECK_BY, dataRows, 1)
      .setDataValidation(placeholderRule);
  }

  // --- BANDING (cream alternation, brand-consistent) ---
  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  var bandRange = sheet.getRange(1, 1, maxDataRow, LOCATION_UPDATE.dataWidth);
  var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b')
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#fff8e7');

  // --- FREEZE HEADER ROW ---
  sheet.setFrozenRows(1);

  // Paint any existing duplicate SKUs (idempotent — clears stale highlights too)
  _refreshLocationUpdateDuplicates(sheet);

  return "✅ Location Update sheet ready.";
}


/**
 * Sidebar: switch the user's active view to the Location Update sheet.
 *
 * Uses SpreadsheetApp.getActive() (BOUND spreadsheet), not openById().
 * setActiveSheet() only changes the visible tab when called on the active
 * spreadsheet reference. See PrepQueue.openPrepQueue for the same pattern
 * and the v1 bug history that established it.
 */
function openLocationUpdate() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet (open this from the spreadsheet UI).";

  var sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);
  if (!sheet) {
    setupLocationUpdateSheet();
    sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);
  }

  ss.setActiveSheet(sheet);
  return "✅ Opened " + LOCATION_UPDATE.sheetName;
}


/**
 * Sidebar: sweep all data rows, fill any blank TIMESTAMPs where SKU exists,
 * and ensure the COUNTER formula is intact on every row. The manual escape
 * hatch — used when the installable trigger missed an edit, OR when a
 * counter cell got accidentally cleared.
 *
 * Behavior per row:
 *   - SKU empty → blank out LOCATION + TIMESTAMP. COUNTER formula stays.
 *   - SKU present + TIMESTAMP blank → stamp with "now" (best we can do — the
 *     original edit time isn't recoverable; an approximate stamp beats blank)
 *   - LOCATION → NEVER touched by refresh (manually edited per 2026-05-14 decision)
 *   - COUNTER formula always re-written (idempotent: if formula was deleted
 *     accidentally, it's restored)
 *   - EMPLOYEE + FINAL_CHECK_BY untouched (out of read range)
 *
 * Returns a status string the sidebar can show.
 */
function refreshLocationUpdateSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);
  if (!sheet) return "❌ Location Update sheet doesn't exist — run Re-style first.";

  var lastRow = sheet.getLastRow();
  if (lastRow < LOCATION_UPDATE.dataStartRow) {
    return "ℹ️ No data rows to refresh.";
  }

  var nRows = lastRow - LOCATION_UPDATE.dataStartRow + 1;
  // Read columns B-D (SKU, LOCATION, TIMESTAMP).
  // Col A is formula-driven (handled separately).
  // Cols E/F are user input (EMPLOYEE, FINAL_CHECK_BY) — never touched here.
  // LOCATION is read so we can clear it when SKU is empty, but never written
  // when SKU is present (it's manually edited).
  var workRange = sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.SKU, nRows, 3);
  var workData = workRange.getValues();

  var nowStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yyyy h:mm a");

  var timestampsFilled = 0;

  // Local indices into workData (cols B/C/D → 0/1/2 in this slice)
  var SKU_I = 0, LOC_I = 1, TS_I = 2;

  for (var i = 0; i < workData.length; i++) {
    var sku = String(workData[i][SKU_I] || "").trim();
    var existingTimestamp = String(workData[i][TS_I] || "").trim();

    if (!sku) {
      // Empty SKU — clear LOCATION + TIMESTAMP. SKU/LOCATION/TIMESTAMP are
      // paired; losing the SKU loses the row's meaning.
      workData[i][LOC_I] = "";
      workData[i][TS_I] = "";
      continue;
    }

    // SKU present — only restore the timestamp if blank. LOCATION stays as
    // whatever the picker typed (or blank if they haven't typed it yet).
    if (!existingTimestamp) {
      workData[i][TS_I] = nowStr;
      timestampsFilled++;
    }
  }

  // One batched write for SKU + LOCATION + TIMESTAMP columns
  workRange.setValues(workData);

  // Re-paint COUNTER formula on every data row. Idempotent — if formula was
  // already there, this is a no-op visually; if it was deleted accidentally,
  // it's restored. Use setFormulas so each cell gets =ROW()-1 individually.
  var counterFormulas = [];
  for (var cf = 0; cf < nRows; cf++) counterFormulas.push(["=ROW()-1"]);
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.COUNTER, nRows, 1)
    .setFormulas(counterFormulas);

  // Refresh duplicate highlighting too — defensive (no-op if nothing changed)
  _refreshLocationUpdateDuplicates(sheet);

  return "✅ Refreshed: " + timestampsFilled + " timestamp(s) restored. Counter formulas re-applied.";
}


/**
 * Sidebar: sort the data rows by LOCATION column A→Z.
 *
 * Sorts columns B–F (SKU, LOCATION, TIMESTAMP, EMPLOYEE, FINAL_CHECK_BY) only.
 * Column A (COUNTER = ROW()-1 formula) is intentionally NOT included in the
 * sort range — its per-row formula always evaluates against its own row, so
 * the numbering 1, 2, 3, ... stays in order while the data shuffles beneath it.
 *
 * Empty-LOCATION rows naturally land at the bottom — useful cue that those
 * entries are still in-progress (picker typed SKU but hasn't filled LOCATION).
 *
 * After the sort, row positions have changed — re-paint the duplicate-SKU
 * highlight so amber/yellow lands on the right cells.
 */
function sortLocationUpdateByLocation() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);
  if (!sheet) return "❌ Location Update sheet doesn't exist — run Re-style first.";

  var lastRow = sheet.getLastRow();
  if (lastRow < LOCATION_UPDATE.dataStartRow) {
    return "ℹ️ No data rows to sort.";
  }

  var nRows = lastRow - LOCATION_UPDATE.dataStartRow + 1;

  var sortRange = sheet.getRange(
    LOCATION_UPDATE.dataStartRow,
    LOCATION_UPDATE.cols.SKU,             // B (col 2)
    nRows,
    LOCATION_UPDATE.dataWidth - 1         // 5 cols → B through F
  );

  sortRange.sort({ column: LOCATION_UPDATE.cols.LOCATION, ascending: true });

  _refreshLocationUpdateDuplicates(sheet);
  SpreadsheetApp.flush();

  return "✅ Sorted Location Update by LOCATION A→Z.";
}


/**
 * Sidebar/editor helper: install real dropdown values for Employee and Final
 * Check By columns. Called once after setup, or any time the staff list changes.
 *
 * Usage from the script editor:
 *   setLocationUpdateDropdowns(["Alice", "Bob"], ["Carol", "Dan"]);
 *
 * Passing an empty array for either argument leaves that column's dropdown
 * alone (so you can update just one list).
 *
 * Strict mode: once you call this with real names, the dropdown rejects invalid
 * values (setAllowInvalid(false)) — typos get flagged. Different from the
 * placeholder rule installed by setupLocationUpdateSheet, which allows any text.
 */
function setLocationUpdateDropdowns(employees, finalCheckers) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);
  if (!sheet) return "❌ Location Update sheet doesn't exist — run Re-style first.";

  var maxDataRow = 1000;
  var dataRows = maxDataRow - LOCATION_UPDATE.dataStartRow + 1;
  var updates = [];

  if (Array.isArray(employees) && employees.length > 0) {
    var empRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(employees, true)
      .setAllowInvalid(false)
      .setHelpText("Re-run setLocationUpdateDropdowns() to update this list.")
      .build();
    sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.EMPLOYEE, dataRows, 1)
      .setDataValidation(empRule);
    updates.push("Employee (" + employees.length + ")");
  }

  if (Array.isArray(finalCheckers) && finalCheckers.length > 0) {
    var fcRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(finalCheckers, true)
      .setAllowInvalid(false)
      .setHelpText("Re-run setLocationUpdateDropdowns() to update this list.")
      .build();
    sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.FINAL_CHECK_BY, dataRows, 1)
      .setDataValidation(fcRule);
    updates.push("Final Check By (" + finalCheckers.length + ")");
  }

  if (updates.length === 0) {
    return "ℹ️ Nothing to update — pass non-empty arrays for employees and/or finalCheckers.";
  }

  return "✅ Dropdowns updated: " + updates.join(", ");
}


// =======================================================================================
// onEdit DISPATCHER — called from Main.js's onEditInstallable(e)
// =======================================================================================

/**
 * SKU edit on Location Update → auto-stamp TIMESTAMP only (LOCATION is now
 * manual per user request 2026-05-14).
 *
 * Runs in the INSTALLABLE trigger because the duplicate-highlight refresh
 * uses openById via SpreadsheetApp.flush() context — simple triggers can
 * fail silently. (The original location lookup that needed openById was
 * removed; trigger now exists primarily for timestamp + dup-highlight.)
 *
 * Defensive: any error is logged and swallowed so this never blocks other
 * onEditInstallable handlers (Prep Queue, Out of Stock, manual receive, etc.).
 */
function locationUpdateOnEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== LOCATION_UPDATE.sheetName) return;
    if (e.range.getColumn() !== LOCATION_UPDATE.cols.SKU) return;
    if (e.range.getRow() < LOCATION_UPDATE.dataStartRow) return;

    var edits = e.range.getValues();
    var startRow = e.range.getRow();

    var nowStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yyyy h:mm a");

    for (var i = 0; i < edits.length; i++) {
      var row = startRow + i;
      var skuLower = String(edits[i][0]).trim().toLowerCase();

      if (skuLower === "") {
        // SKU cleared — wipe LOCATION + TIMESTAMP so the row reads as empty.
        // SKU and LOCATION are a paired audit entry; clearing the SKU clears
        // the location record with it.
        //
        // COUNTER is a formula (=ROW()-1) that's intentionally left alone — the
        // user wanted # to PERSIST across SKU clears (fixed 2026-05-13).
        // EMPLOYEE + FINAL_CHECK_BY are user-input only; never touched here.
        sheet.getRange(row, LOCATION_UPDATE.cols.LOCATION).setValue("");
        sheet.getRange(row, LOCATION_UPDATE.cols.TIMESTAMP).setValue("");
        continue;
      }

      // Non-empty SKU — stamp TIMESTAMP only.
      // LOCATION is manually typed by the picker (intentional, per 2026-05-14
      // user decision — they preferred KISS over auto-fill verification).
      // COUNTER self-derives from the formula. EMPLOYEE + FINAL_CHECK_BY are
      // dropdown-selected — never touched.
      sheet.getRange(row, LOCATION_UPDATE.cols.TIMESTAMP).setValue(nowStr);
    }

    // Refresh duplicate-SKU highlight after every edit batch — surfaces dupes
    // the moment they're typed, clears highlights when one of a pair is changed.
    _refreshLocationUpdateDuplicates(sheet);
  } catch (err) {
    try { Logger.log("locationUpdateOnEdit error: " + err); } catch (_) {}
  }
}


// =======================================================================================
// PRIVATE: duplicate-SKU highlight (mirrors the PrepQueue + OutOfStock pattern)
// =======================================================================================

/**
 * Scans column B (SKU), identifies SKUs that appear two or more times
 * (case-insensitive, trimmed), and paints those cells with soft amber
 * background + thick yellow border. Removing a duplicate clears the highlight
 * on its formerly-duped counterpart in the same pass.
 *
 * Only touches background + border — never font, alignment, number format,
 * banding, etc. When a duplicate is removed, the highlight clears cleanly
 * and the row's original look is preserved.
 *
 * Scans the full data band (Math.min(maxRows, 1000)) — not just to lastRow —
 * so previously-highlighted cells whose SKU has been deleted still get their
 * background explicitly cleared. See _refreshPrepQueueDuplicates docstring
 * for the full rationale.
 *
 * Run after any change that could affect col B:
 *   - setupLocationUpdateSheet (initial paint)
 *   - locationUpdateOnEdit (live, on every SKU edit)
 *   - refreshLocationUpdateSheet (post-sweep)
 */
function _refreshLocationUpdateDuplicates(sheet) {
  if (!sheet) return;

  var maxScanRow = Math.min(sheet.getMaxRows(), 1000);
  if (maxScanRow < LOCATION_UPDATE.dataStartRow) return;

  var totalRows = maxScanRow - LOCATION_UPDATE.dataStartRow + 1;
  var skuRange  = sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.SKU, totalRows, 1);
  var skus      = skuRange.getValues();

  // Count occurrences (case-insensitive, trimmed)
  var counts = {};
  for (var i = 0; i < skus.length; i++) {
    var k = String(skus[i][0]).trim().toLowerCase();
    if (!k) continue;
    counts[k] = (counts[k] || 0) + 1;
  }

  // Build the FULL backgrounds array — dupes get amber, everything else null.
  var bgs = [];
  var dupeIndexes = [];
  for (var j = 0; j < skus.length; j++) {
    var key = String(skus[j][0]).trim().toLowerCase();
    if (key && counts[key] >= 2) {
      bgs.push(['#fff3b0']);
      dupeIndexes.push(j);
    } else {
      bgs.push([null]);
    }
  }
  skuRange.setBackgrounds(bgs);

  // Borders: clear the full range (single batched call), then add thick yellow
  // borders per dupe (small N, typically 0-4 calls).
  skuRange.setBorder(false, false, false, false, false, false, null, null);
  for (var k = 0; k < dupeIndexes.length; k++) {
    var rowIdx = dupeIndexes[k];
    sheet.getRange(LOCATION_UPDATE.dataStartRow + rowIdx, LOCATION_UPDATE.cols.SKU)
      .setBorder(true, true, true, true, false, false,
                 '#ffb800', SpreadsheetApp.BorderStyle.SOLID_THICK);
  }

  // Force pending writes to land before the user's next interaction.
  SpreadsheetApp.flush();
}
