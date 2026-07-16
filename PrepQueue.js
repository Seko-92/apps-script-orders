// =======================================================================================
// PREP_QUEUE.gs — "Prep Queue" sheet for the warehouse employee's personal todo list
// =======================================================================================
//
// PURPOSE — IMPORTANT BOUNDARY
//   This sheet is a PERSONAL TODO LIST for the warehouse employee. They add
//   SKUs they want to work on LATER — restock, repackage, recheck, organize.
//
//   This is NOT for deferred customer orders. The All Orders sheet is for
//   real customer orders that MUST be prepped and shipped immediately. The
//   sidebar deliberately does NOT have a "move from All Orders to Prep Queue"
//   button — the workflow rule is one-way: orders go OUT, never PARKED.
//
// ARCHITECTURE (two-table since 2026-07-16 — mirrors All Orders' DIRECT split)
//   Standalone sheet ("Prep Queue") holding TWO stacked tables with identical
//   band + header pairs, split by a dynamic divider (symmetric since the
//   2026-07-16 visual pass — each table opens with a yellow ▌ band):
//     Row 1        : ▌ CURRENT title band (frozen)
//     Row 2        : CURRENT column headers (frozen)
//     Rows 3..B-1  : CURRENT prep items (today's work)
//     Row B        : ▌ INCOMING band — col A value is EXACTLY "INCOMING"
//                    (Gotcha #1 class: the ▌ dressing is a number-format
//                    prefix; decorative text in the VALUE breaks
//                    _getPrepBoundaryRow)
//     Row B+1      : INCOMING column headers (same headers)
//     Rows B+2+    : INCOMING / future prep (arriving + to-order items)
//   The CURRENT table is DYNAMIC: _ensurePrepBuffer keeps blank typing rows
//   above the divider and slides it down as entries approach it.
//   Columns: SKU · QTY · LOCATION · HAND · NOTE · DATE ADDED · ✔ DONE
//
//   ✔ DONE (col G) is a native checkbox; a CF rule strikes the row through +
//   grays it the instant the box is ticked (display-only, instant, zero
//   trigger latency). Checkboxes are planted per-row when a SKU lands (onEdit
//   + Quick Add + setup backfill) so empty rows stay clean.
//
//   onEdit hook (prepQueueOnEdit, dispatched from Main.js onEdit):
//     - When SKU is entered/changed in column A, auto-fill LOCATION + HAND
//       from Master Inventory (same lookup as LiveSync uses for All Orders).
//     - When a SKU is first written to a previously-empty row, stamp DATE ADDED
//       and plant the DONE checkbox.
//     - Tops up the blank-row buffer above the divider (typing space).
//
// PUBLIC API
//   setupPrepQueueSheet()         — idempotent: create/migrate sheet, both tables,
//                                   divider, checkboxes, strike CF, pulse chip
//   openPrepQueue()               — switch the user's active sheet to Prep Queue (sidebar button)
//   clearPrepQueue()              — wipe data rows in BOTH tables (sidebar danger button)
//   clearDonePrepItems()          — delete every checked-off row in both tables (sidebar)
//   prepQueueOnEdit(e)            — onEdit dispatcher (called from Main.js)
//   addPrepQueueItem(sku,qty,note,incoming) — sidebar Quick Add; incoming=true
//                                   targets the INCOMING table
//   refreshPrepQueueHand()        — rewrite HAND for every row (Zoho-first; runs every 2 min
//                                   via the n8n writeZohoStock push + sidebar recompute)
//   refreshPrepQueueLocations()   — re-mirror LOCATION from MI (hourly via
//                                   runHourlyHousekeeping in Housekeeping.js + sidebar button);
//                                   never overwrites with NOT FOUND, so hand-typed
//                                   locations for non-eBay items survive
// =======================================================================================

// ---------- LOCAL SCHEMA (kept here, not in Schema.js — different sheet) ----------
var PREP_QUEUE = {
  sheetName: "Prep Queue",

  // 1-based column positions
  cols: {
    SKU:        1,   // A
    QTY:        2,   // B
    LOCATION:   3,   // C
    HAND:       4,   // D
    NOTE:       5,   // E
    DATE_ADDED: 6,   // F
    DONE:       7    // G — native checkbox; TRUE strikes the row via CF (2026-07-16)
  },

  // 0-based array indices for getValues() iteration
  idx: function(name) { return PREP_QUEUE.cols[name] - 1; },

  dataWidth:    7,
  titleRow:     1,   // ▌ CURRENT band (visual twin of the INCOMING divider)
  headerRow:    2,
  dataStartRow: 3,

  // TWO-TABLE CONTRACT (2026-07-16): CURRENT table on top, INCOMING (future
  // prep) below, split by a divider row whose col-A VALUE is exactly this
  // marker. The ▌ dressing is a number-format display prefix — the underlying
  // value must stay exact, same load-bearing rule as All Orders' "DIRECT".
  boundaryMarker: "INCOMING",

  // Blank typable rows kept between the CURRENT table's last entry and the
  // divider; _ensurePrepBuffer tops this up from onEdit + Quick Add.
  bufferRows: 4,

  // Headers — written by setupPrepQueueSheet on BOTH header rows
  headers: ["◈ SKU", "# QTY", "LOCATION", "◫ HAND", "NOTE", "DATE ADDED", "✔ DONE"]
};


// =======================================================================================
// PUBLIC API
// =======================================================================================

/**
 * One-time setup: creates "Prep Queue" sheet if missing, applies headers,
 * brand styling, banding, and HAND low-stock conditional formatting.
 * Idempotent — safe to re-run.
 */
function setupPrepQueueSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(PREP_QUEUE.sheetName);
  }

  // --- TITLE-BAND MIGRATION (2026-07-16 visual pass) ---
  // Older layouts have the column headers on row 1. Shift everything down one
  // so row 1 becomes the ▌ CURRENT title band (symmetric with the INCOMING
  // band).
  var a1 = String(sheet.getRange(1, 1).getValue()).trim();
  if (a1.charAt(0) === '◈') {
    sheet.insertRowsBefore(1, 1);
  }

  // --- CHIP MIGRATION (consolidated across the 2026-07-16 layout hops:
  //     G1/H1 era → H1/I1 → H2/I2 → H1 dark → FINAL: F1 IN-BAND chip,
  //     stamp I1 — the sync time is part of the ▌ CURRENT band now). ---
  // Runs AFTER the title insert so every era's leftovers sit at known spots.
  // Carry the newest stamp Date to I1, wipe every former chip home, and undo
  // any hand-merge on row 1 (the user merged it to tame the old floating
  // dark chip — the in-band design replaces that entirely).
  try {
    var stampCell = sheet.getRange(SHEET_PULSE.prepQueue.stamp);   // I1
    if (!(stampCell.getValue() instanceof Date)) {
      var carriers = ["I2", "H2"];   // prior stamp homes, newest era first
      for (var c = 0; c < carriers.length; c++) {
        var carried = sheet.getRange(carriers[c]).getValue();
        if (carried instanceof Date) { stampCell.setValue(carried); break; }
      }
    }
    try { sheet.getRange(1, 1, 2, 26).breakApart(); } catch (bm) {}
    sheet.getRange("I2").clearContent();
    ["H1", "H2"].forEach(function(oldHome) {   // former dark-chip cells
      sheet.getRange(oldHome).clearContent()
           .setBackground(null).setFontColor(null)
           .setBorder(false, false, false, false, false, false);
    });
    sheet.setColumnWidth(8, 100);              // was 180 for the dark chip
    sheet.getRange(1, 6, 2, 5).clearNote();    // chip notes, rows 1-2 × F..J
    sheet.showColumns(8);                      // G1-era hidden stamp column
    var cfMigrated = sheet.getConditionalFormatRules().filter(function (r) {
      return !r.getRanges().some(function (rg) {
        return rg.getRow() <= 2 && rg.getNumRows() === 1 &&
               rg.getColumn() >= 7 && rg.getColumn() <= 10;
      });
    });
    sheet.setConditionalFormatRules(cfMigrated);
  } catch (mErr) { try { Logger.log("setupPrepQueueSheet: chip migration: " + mErr); } catch (_) {} }

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(PREP_QUEUE.cols.SKU,        110);
  sheet.setColumnWidth(PREP_QUEUE.cols.QTY,         60);
  sheet.setColumnWidth(PREP_QUEUE.cols.LOCATION,    95);
  sheet.setColumnWidth(PREP_QUEUE.cols.HAND,        80);
  sheet.setColumnWidth(PREP_QUEUE.cols.NOTE,       260);
  sheet.setColumnWidth(PREP_QUEUE.cols.DATE_ADDED, 130);
  sheet.setColumnWidth(PREP_QUEUE.cols.DONE,        55);

  // --- DATA AREA: column-level format (so new rows inherit) ---
  var maxDataRow = 1000;
  var dataRows = maxDataRow - PREP_QUEUE.dataStartRow + 1;

  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.SKU, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.QTY, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.LOCATION, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.HAND, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.NOTE, dataRows, 1)
    .setFontFamily('Roboto').setFontStyle('italic').setFontSize(10)
    .setFontColor('#434343').setHorizontalAlignment('left');
  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.DATE_ADDED, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('center');
  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.DONE, dataRows, 1)
    .setHorizontalAlignment('center');

  sheet.getRange(PREP_QUEUE.dataStartRow, 1, dataRows, PREP_QUEUE.dataWidth)
    .setVerticalAlignment('middle');

  // --- BANDING (cream alternation) ---
  // Range starts at the HEADER row (2), NOT row 1: the banding header slot
  // paints over manual fills, and blacking out the ▌ CURRENT band on row 1
  // was the "restyle works then gets messed up" bug (2026-07-16). Row 1 is
  // deliberately OUTSIDE the banding.
  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  var bandRange = sheet.getRange(PREP_QUEUE.headerRow, 1,
                                 maxDataRow - PREP_QUEUE.headerRow + 1, PREP_QUEUE.dataWidth);
  var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b')
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#fff8e7');

  // --- CURRENT TITLE BAND + HEADERS ---
  // Styled AFTER the banding (same reason the INCOMING pair is): manual
  // styling must land on top, never under a banding re-apply.
  _stylePrepBand(sheet, PREP_QUEUE.titleRow, "CURRENT", "TODAY'S PREP · WORK LIST");
  _stylePrepHeaderRow(sheet, PREP_QUEUE.headerRow);

  // --- TWO-TABLE DIVIDER + INCOMING HEADER ---
  // Create the divider below existing data (with typing buffer) on first run;
  // re-style it in place on every re-run.
  var boundary = _getPrepBoundaryRow(sheet);
  if (boundary < 0) {
    var lastContent = Math.max(sheet.getLastRow(), PREP_QUEUE.headerRow);
    boundary = lastContent + PREP_QUEUE.bufferRows + 1;
    sheet.getRange(boundary, PREP_QUEUE.cols.SKU).setValue(PREP_QUEUE.boundaryMarker);
  }
  _stylePrepBand(sheet, boundary, PREP_QUEUE.boundaryMarker, "FUTURE PREP · ARRIVING / TO ORDER");
  _stylePrepHeaderRow(sheet, boundary + 1);

  // --- DONE CHECKBOXES: plant on SKU rows, sweep off empty/structural rows ---
  _normalizePrepCheckboxes(sheet, boundary);

  // --- CONDITIONAL FORMATTING ---
  // Order is load-bearing: Sheets applies the FIRST matching rule per cell,
  // so the done-strike rule comes before HAND low-stock — a checked-off row
  // reads muted gray even when its HAND is in the low-stock band.
  var existingRules = sheet.getConditionalFormatRules();
  var keptRules = existingRules.filter(function(r) {
    var isStrike = false;
    var bc = r.getBooleanCondition();
    if (bc && bc.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
      var f = String(bc.getCriteriaValues()[0] || "");
      isStrike = f.indexOf("=$G") === 0 && f.indexOf("=TRUE") > 0;
    }
    var isHand = r.getRanges().some(function(rg) {
      return rg.getColumn() === PREP_QUEUE.cols.HAND && rg.getNumColumns() === 1;
    });
    return !isStrike && !isHand;
  });

  var strikeRange = sheet.getRange(PREP_QUEUE.dataStartRow, 1, dataRows, PREP_QUEUE.dataWidth - 1);
  var strikeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$G" + PREP_QUEUE.dataStartRow + "=TRUE")
    .setStrikethrough(true)
    .setFontColor('#9e9e9e')
    .setRanges([strikeRange])
    .build();

  var handRange = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.HAND, dataRows, 1);
  var handRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      "=AND(ISNUMBER(D" + PREP_QUEUE.dataStartRow + "), D" + PREP_QUEUE.dataStartRow + "<=20)"
    )
    .setBackground('#ffd966').setFontColor('#1d1d1b').setBold(true)
    .setRanges([handRange])
    .build();

  sheet.setConditionalFormatRules([strikeRule, handRule].concat(keptRules));

  // --- FREEZE TITLE BAND + HEADER (both stay visible while scrolling) ---
  sheet.setFrozenRows(2);

  // --- FRESHNESS PULSE CHIP (H1 chip / I1 stamp — top row, beside the
  //     ▌ CURRENT band, per the user's placement call 2026-07-16) ---
  try { _installPulseChip(sheet, SHEET_PULSE.prepQueue); }
  catch (e) { try { Logger.log("setupPrepQueueSheet: pulse chip error: " + e); } catch (_) {} }

  // Paint any existing duplicate SKUs (idempotent — clears stale highlights too)
  _refreshPrepQueueDuplicates(sheet);

  // Backfill SKU → eBay listing links across existing rows (same as the All
  // Orders "Link SKUs" backfill). Best-effort.
  try { refreshPrepQueueSkuLinks(); }
  catch (e) { try { Logger.log("setupPrepQueueSheet: SKU link backfill error: " + e); } catch (_) {} }

  return "✅ Prep Queue sheet ready (two tables: CURRENT + INCOMING).";
}


// =======================================================================================
// TWO-TABLE STRUCTURE HELPERS (2026-07-16)
// =======================================================================================

/**
 * Find the INCOMING divider row — exact-match contract on col A, same
 * discipline as All Orders' getBoundaryRow ("DIRECT", Gotcha #1). Returns the
 * 1-based row number, or -1 when the sheet has no divider yet (legacy single-
 * table layout — every helper treats that as "everything is CURRENT").
 */
function _getPrepBoundaryRow(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < PREP_QUEUE.dataStartRow) return -1;
  var vals = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.SKU,
                            lastRow - PREP_QUEUE.dataStartRow + 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() === PREP_QUEUE.boundaryMarker) {
      return PREP_QUEUE.dataStartRow + i;
    }
  }
  return -1;
}

/** True for the two structural rows every walker must skip: divider + INCOMING header. */
function _isPrepStructureRow(row, boundary) {
  return boundary > 0 && (row === boundary || row === boundary + 1);
}

/**
 * Style one header row — dark brand band, yellow Oswald text, thick yellow
 * underline. Shared by row 1 and the INCOMING table's header so both tables
 * wear the identical band.
 */
function _stylePrepHeaderRow(sheet, row) {
  sheet.getRange(row, 1, 1, PREP_QUEUE.dataWidth)
    .setValues([PREP_QUEUE.headers])
    .setBackground('#1d1d1b')   // brand black
    .setFontColor('#ffd966')    // brand yellow
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontLine('none')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.getRange(row, 1, 1, PREP_QUEUE.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);
}

/**
 * Style a brand-yellow ▌ title band (mirrors All Orders' DIRECT band). Used
 * for BOTH the row-1 CURRENT band and the INCOMING divider so the two tables
 * wear the same identity. For the divider, markerText MUST be exactly
 * PREP_QUEUE.boundaryMarker — the ▌ dressing is a number-format display
 * prefix, so _getPrepBoundaryRow's strict match keeps working (Gotcha #1).
 */
function _stylePrepBand(sheet, row, markerText, rightLabel) {
  sheet.getRange(row, 1, 1, PREP_QUEUE.dataWidth)
    .setBackground('#ffd400')   // brand action yellow
    .setFontColor('#1d1d1b')
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontLine('none')
    .setFontStyle('normal')
    .setVerticalAlignment('middle')
    .setBorder(true, null, true, null, null, null,
               '#1d1d1b', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange(row, PREP_QUEUE.cols.SKU)
    .setValue(markerText)
    .setNumberFormat('"▌  "@')
    .setHorizontalAlignment('left');
  sheet.getRange(row, PREP_QUEUE.cols.NOTE)
    .setValue(rightLabel)
    .setHorizontalAlignment('right')
    .setFontSize(9);
  sheet.setRowHeight(row, 36);
}

/**
 * Ensure the DONE column matches the sheet: every row holding a SKU (either
 * table) wears a checkbox; empty + structural rows carry none. NEVER re-plants
 * an existing checkbox — insertCheckboxes() resets the cell to unchecked, so
 * it only runs where checkbox validation is missing.
 */
function _normalizePrepCheckboxes(sheet, boundary) {
  var lastRow = sheet.getLastRow();
  if (lastRow < PREP_QUEUE.dataStartRow) return;
  var n = lastRow - PREP_QUEUE.dataStartRow + 1;
  var skus = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.SKU, n, 1).getValues();
  var validations = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.DONE, n, 1)
                         .getDataValidations();

  for (var i = 0; i < n; i++) {
    var row = PREP_QUEUE.dataStartRow + i;
    var wantsBox = !_isPrepStructureRow(row, boundary) && String(skus[i][0]).trim() !== "";
    var dv = validations[i][0];
    var hasBox = !!(dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX);
    if (wantsBox === hasBox) continue;   // steady state — zero calls
    var cell = sheet.getRange(row, PREP_QUEUE.cols.DONE);
    if (wantsBox) cell.insertCheckboxes();
    else { try { cell.removeCheckboxes(); } catch (e) {} cell.clearContent(); }
  }
}

/**
 * Keep typable blank rows between the CURRENT table's last entry and the
 * divider (the All Orders DIRECT-buffer idea). Inserts PREP_QUEUE.bufferRows
 * rows above the divider when fewer than 2 blanks remain. Returns the
 * (possibly shifted) boundary row.
 */
function _ensurePrepBuffer(sheet, boundary) {
  if (boundary <= 0) return boundary;

  var want = 2;
  var from = Math.max(PREP_QUEUE.dataStartRow, boundary - want);
  var nCheck = boundary - from;
  if (nCheck >= want) {
    var vals = sheet.getRange(from, PREP_QUEUE.cols.SKU, nCheck, 1).getValues();
    var allBlank = vals.every(function(v) { return String(v[0]).trim() === ""; });
    if (allBlank) return boundary;   // enough space — nothing to do
  }

  // insertRowsAfter a DATA row (not Before the divider) so the new blanks
  // inherit data-row formatting, not the yellow divider band.
  sheet.insertRowsAfter(boundary - 1, PREP_QUEUE.bufferRows);

  if (boundary - 1 === PREP_QUEUE.headerRow) {
    // Degenerate case: CURRENT table was empty, so the source row was the
    // header — strip the inherited dark band from the fresh rows.
    sheet.getRange(boundary, 1, PREP_QUEUE.bufferRows, PREP_QUEUE.dataWidth)
      .setBackground(null).setFontColor(null).setFontWeight('normal').setFontLine('none')
      .setBorder(false, false, false, false, false, false);
  }
  // Fresh buffer rows must not inherit a checkbox from the row above.
  try {
    sheet.getRange(boundary, PREP_QUEUE.cols.DONE, PREP_QUEUE.bufferRows, 1)
      .removeCheckboxes().clearContent();
  } catch (e) {}

  return boundary + PREP_QUEUE.bufferRows;
}

/**
 * First free row for a new entry inside one table segment: the row after the
 * segment's last non-empty SKU (or the segment's first row when empty).
 * segEndExclusive bounds the scan (the divider row for CURRENT; maxRows+1 for
 * INCOMING).
 */
function _findPrepAppendRow(sheet, segStart, segEndExclusive) {
  var lastRow = Math.min(sheet.getLastRow(), segEndExclusive - 1);
  if (lastRow < segStart) return segStart;
  var vals = sheet.getRange(segStart, PREP_QUEUE.cols.SKU, lastRow - segStart + 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== "") return segStart + i + 1;
  }
  return segStart;
}


/**
 * Shared LOCATION + HAND resolution for the three Prep Queue entry paths
 * (sidebar preview, quick-add, on-sheet edit). HAND prefers Zoho's
 * available_stock — Prep items are restock/personal, so Zoho is the unified
 * stock truth and covers items not listed on eBay — with MI as fallback. No
 * committed subtraction: HAND = available, matching the All Orders HAND
 * semantics (2026-05-09) so the Prep number equals the All Orders number.
 *
 * Pass pre-built maps for batch callers; omit them (null) for single lookups.
 * Returns { location, hand, found }. `found` is true when the SKU is known to
 * either MI (eBay shelf) or Zoho.
 */
function _prepResolveStock(skuLower, locationMap, inventoryMap, zohoMap) {
  var location = locationMap ? (locationMap.get(skuLower) || "NOT FOUND")
                             : getSingleLocation(skuLower);
  var inMi = (location !== "NOT FOUND");

  var zo = zohoMap ? zohoMap.get(skuLower) : getSingleZohoStock(skuLower);
  var hand = "";
  if (zo) {
    hand = zo.available;
  } else if (inMi) {
    var inv = inventoryMap ? inventoryMap.get(skuLower) : getSingleInventory(skuLower);
    hand = inv ? inv.available : "";
  }

  return { location: location, hand: hand, found: (inMi || zo != null) };
}


/**
 * Quick lookup for the sidebar's live preview as the user types a SKU.
 * Returns { sku, location, hand, found } — `found` is false when the SKU is
 * unknown to both MI and Zoho (sidebar uses this to flag NOT FOUND visually).
 *
 * HAND matches the All Orders sheet's HAND column (Zoho-first for direct/non-
 * eBay items, MI for eBay items; no committed subtraction).
 */
function lookupSkuForPrepQueue(sku) {
  var skuLower = String(sku || "").trim().toLowerCase();
  if (!skuLower) return { sku: "", location: "", hand: "", found: false };

  var r = _prepResolveStock(skuLower, null, null, null);
  return { sku: skuLower, location: r.location, hand: r.hand, found: r.found };
}


/**
 * Rewrite HAND for EVERY existing Prep Queue row from the current MI + Zoho
 * maps. This is the Prep-Queue counterpart to recomputeHand (which only walks
 * the All Orders sheet) — without it, Prep rows stayed frozen at their entry
 * value unless the picker re-typed the SKU. Zoho-first, MI fallback, no
 * committed subtraction. Only HAND is rewritten; LOCATION/QTY/NOTE/DATE are
 * left untouched. Returns a status string.
 */
function refreshPrepQueueHand() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) return "ℹ️ Prep Queue sheet not found.";

  var lastRow = sheet.getLastRow();
  if (lastRow < PREP_QUEUE.dataStartRow) return "ℹ️ Prep Queue empty.";

  var nRows = lastRow - PREP_QUEUE.dataStartRow + 1;
  var skuVals  = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.SKU,  nRows, 1).getValues();
  var handVals = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.HAND, nRows, 1).getValues();

  var maps = buildLocationAndInventoryMaps();
  var locationMap  = maps.locationMap;
  var inventoryMap = maps.inventoryMap;
  var zohoMap      = buildZohoStockMap();

  var boundary = _getPrepBoundaryRow(sheet);

  var out = [];
  var updated = 0;
  for (var i = 0; i < nRows; i++) {
    var row = PREP_QUEUE.dataStartRow + i;
    var sku = String(skuVals[i][0] || "").trim();
    // Preserve blank rows, the divider/INCOMING-header rows, and any ◈ header
    // glyph row (belt-and-suspenders — header rows are never SKUs).
    if (!sku || sku.charAt(0) === '◈' || _isPrepStructureRow(row, boundary)) {
      out.push([handVals[i][0]]);
      continue;
    }
    var r = _prepResolveStock(sku.toLowerCase(), locationMap, inventoryMap, zohoMap);
    out.push([r.hand]);
    updated++;
  }

  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.HAND, nRows, 1).setValues(out);

  // Freshness chip. NOTE: when this runs via the n8n writeZohoStock doPost
  // action (every 2 min), the pinned /exec version must be new enough to
  // carry this line for the stamp to fire from that path — the editor,
  // sidebar, and time-trigger paths stamp regardless (Gotcha #12).
  stampSheetPulse(sheet, SHEET_PULSE.prepQueue.stamp);

  return "✅ Prep Queue HAND refreshed for " + updated + " row(s).";
}


/**
 * Re-mirror LOCATION for every existing Prep Queue row from Master Inventory.
 * The onEdit auto-fill stamps the location once at entry time; if the item
 * moves shelves later, the row goes stale. This is the fix — run hourly by
 * runHourlyHousekeeping (Housekeeping.js; MI's own location data only updates
 * hourly via MAIN's Smart Sync, so refreshing more often would re-read the
 * same data), plus a sidebar "Refresh Locations" escape hatch.
 *
 * OVERWRITE POLICY (user's call, 2026-07-13): mirror MI, but NEVER clobber a
 * cell with NOT FOUND — rows whose SKU is unknown to MI (or known but with a
 * blank location, which buildLocationAndInventoryMaps stores as "NOT FOUND")
 * keep whatever the cell already says, so hand-typed locations for non-eBay
 * items survive the hourly pass.
 *
 * Pass pre-built maps from a batch caller (housekeeping) or omit for a
 * standalone run. Only LOCATION is written — QTY/HAND/NOTE/DATE untouched.
 */
function refreshPrepQueueLocations(maps) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) return "ℹ️ Prep Queue sheet not found.";

  var lastRow = sheet.getLastRow();
  if (lastRow < PREP_QUEUE.dataStartRow) {
    stampSheetPulse(sheet, SHEET_PULSE.prepQueue.stamp);
    return "ℹ️ Prep Queue empty — locations up to date.";
  }

  // Accept pre-built maps (shared MI read from housekeeping). Detect by
  // shape, not presence — defensive against ever being wired to a trigger,
  // which passes an event object as the first argument.
  var locationMap = (maps && maps.locationMap) ? maps.locationMap
                                               : buildLocationAndInventoryMaps().locationMap;

  var nRows = lastRow - PREP_QUEUE.dataStartRow + 1;
  var skuVals = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.SKU, nRows, 1).getValues();
  var locVals = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.LOCATION, nRows, 1).getValues();

  var boundary = _getPrepBoundaryRow(sheet);

  var out = [];
  var updated = 0;
  var kept = 0;
  for (var i = 0; i < nRows; i++) {
    var row = PREP_QUEUE.dataStartRow + i;
    var sku = String(skuVals[i][0] || "").trim();
    var current = locVals[i][0];
    // Blank rows, divider/INCOMING-header rows, and ◈ header-glyph rows pass
    // through untouched.
    if (!sku || sku.charAt(0) === '◈' || _isPrepStructureRow(row, boundary)) {
      out.push([current]);
      continue;
    }

    var miLoc = locationMap.get(sku.toLowerCase());
    if (miLoc && miLoc !== "NOT FOUND") {
      if (String(current).trim() !== String(miLoc).trim()) updated++;
      out.push([miLoc]);
    } else {
      out.push([current]);   // unknown to MI → keep the picker's value
      kept++;
    }
  }

  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.LOCATION, nRows, 1).setValues(out);
  stampSheetPulse(sheet, SHEET_PULSE.prepQueue.stamp);

  return "✅ Prep Queue locations refreshed — " + updated + " updated" +
         (kept > 0 ? ", " + kept + " kept (not in MI)" : "") + ".";
}


/**
 * Quick-add path: append a new Prep Queue row with the given SKU, QTY, and
 * NOTE. LOCATION + HAND + DATE ADDED are auto-filled server-side (programmatic
 * setValues doesn't fire onEdit, so we inline the lookup here).
 *
 * Segment-aware (2026-07-16): by default the row lands in the CURRENT table
 * (above the INCOMING divider, topping up the typing buffer first so the
 * divider slides down rather than being overwritten). Pass incoming=true to
 * land it in the INCOMING table instead (sidebar toggle).
 *
 * Returns { success, row, sku, location, hand, found, incoming, error? } so
 * the sidebar can show what just landed.
 */
function addPrepQueueItem(sku, qty, note, incoming) {
  sku = String(sku || "").trim();
  if (!sku) return { success: false, error: "SKU required" };

  var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) {
    setupPrepQueueSheet();
    sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  }

  qty  = parseInt(qty) || 1;
  note = String(note || "").trim();

  // Inline lookup (programmatic writes don't trigger prepQueueOnEdit).
  // HAND is Zoho-first (matches All Orders' HAND semantics, no committed sub).
  var skuLower = sku.toLowerCase();
  var r = _prepResolveStock(skuLower, null, null, null);
  var location = r.location;
  var hand = r.hand;
  var found = r.found;

  var nowStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy h:mm a");

  var boundary = _getPrepBoundaryRow(sheet);
  var insertAt;
  var landedIncoming = false;
  if (boundary > 0 && incoming) {
    insertAt = _findPrepAppendRow(sheet, boundary + 2, sheet.getMaxRows() + 1);
    landedIncoming = true;
  } else if (boundary > 0) {
    boundary = _ensurePrepBuffer(sheet, boundary);   // may shift the divider down
    insertAt = _findPrepAppendRow(sheet, PREP_QUEUE.dataStartRow, boundary);
  } else {
    // Legacy single-table layout (no divider yet) — old append behavior.
    insertAt = Math.max(sheet.getLastRow() + 1, PREP_QUEUE.dataStartRow);
  }
  if (insertAt > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 8);

  sheet.getRange(insertAt, 1, 1, PREP_QUEUE.dataWidth - 1).setValues([[
    sku, qty, location, hand, note, nowStr
  ]]);

  // DONE checkbox for the new row. Guard on existing validation —
  // insertCheckboxes() resets a cell to unchecked, so never re-plant.
  try {
    var doneCell = sheet.getRange(insertAt, PREP_QUEUE.cols.DONE);
    var dv = doneCell.getDataValidation();
    if (!dv || dv.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      doneCell.insertCheckboxes();
    }
  } catch (cbErr) { try { Logger.log("addPrepQueueItem: checkbox error: " + cbErr); } catch (_) {} }

  // Programmatic writes don't fire prepQueueOnEdit, so refresh duplicate
  // highlighting here. If this SKU is already in the queue, both rows get
  // marked; if it's new, no-op.
  _refreshPrepQueueDuplicates(sheet);

  // Link the new SKU cell (programmatic setValues doesn't fire prepQueueOnEdit).
  try {
    applySkuLinksToColumn(sheet, PREP_QUEUE.cols.SKU, insertAt, insertAt, buildSkuEnrichmentMap());
  } catch (enrErr) { try { Logger.log("addPrepQueueItem: SKU link error: " + enrErr); } catch (_) {} }

  return {
    success:  true,
    row:      insertAt,
    sku:      sku,
    location: location,
    hand:     hand,
    found:    found,
    incoming: landedIncoming
  };
}


/**
 * One-shot migration: realigns existing data to the 6-column schema.
 *
 * Symptom this fixes: when the sheet was originally built (or hand-edited)
 * with only 5 columns (SKU · LOCATION · HAND · NOTE · DATE), new entries
 * from `prepQueueOnEdit` start landing in the wrong columns because the
 * code uses the 6-column schema (QTY at B). Location appears under "HAND",
 * date drifts to a hidden col F, etc.
 *
 * Per-row detection:
 *   - If col B looks like a location code (e.g. "A-43", "L-233"), the row is
 *     in 5-col user layout → shift values right by one (insert blank QTY at B).
 *   - Otherwise the row is already in 6-col schema layout — leave it alone.
 *
 * Safe to re-run: rows already in schema layout are detected and not shifted.
 */
function repairPrepQueueLayout() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) return "❌ Prep Queue sheet not found.";

  var lastRow = sheet.getLastRow();

  if (lastRow < PREP_QUEUE.dataStartRow) {
    setupPrepQueueSheet();
    return "✅ Headers fixed — no data to migrate.";
  }

  // Read up to schema width (6) so we can detect both layouts in one shot.
  var nRows = lastRow - PREP_QUEUE.dataStartRow + 1;
  var data = sheet.getRange(PREP_QUEUE.dataStartRow, 1, nRows, PREP_QUEUE.dataWidth).getValues();

  // Location-code heuristic — Master Inventory uses formats like "A-43",
  // "L-233", "I-29". A column-B value matching that pattern is a strong
  // signal the row is in legacy 5-col user layout.
  var locPattern = /^[A-Z]+\s*-\s*\d+/i;

  var migrated = 0;
  var preserved = 0;
  var out = data.map(function(row) {
    var sku = row[0];
    if (!String(sku || "").trim()) {
      return row;  // empty row, leave as-is
    }
    var bStr = String(row[1] || "").trim();
    if (locPattern.test(bStr)) {
      // 5-col user layout: SKU | LOCATION | HAND | NOTE | DATE | (extra)
      // → schema layout:   SKU | QTY      | LOCATION | HAND | NOTE | DATE | (DONE blank)
      migrated++;
      return [sku, "", row[1], row[2], row[3], row[4], ""];
    }
    preserved++;
    return row;   // already schema-aligned (7-wide read) — incl. DONE state
  });

  // Write the realigned data back
  sheet.getRange(PREP_QUEUE.dataStartRow, 1, out.length, PREP_QUEUE.dataWidth).setValues(out);

  // Clean up any orphan column-G data left over from old broken inserts.
  if (sheet.getLastColumn() > PREP_QUEUE.dataWidth) {
    var extraCols = sheet.getLastColumn() - PREP_QUEUE.dataWidth;
    sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.dataWidth + 1, nRows, extraCols).clearContent();
  }

  // Re-apply headers + styling so the visual layout matches the schema.
  setupPrepQueueSheet();

  return "✅ Prep Queue repaired — " + migrated + " row(s) migrated, " +
         preserved + " already aligned.";
}


/**
 * Sidebar: switch the user's active view to the Prep Queue sheet.
 *
 * IMPORTANT: must use SpreadsheetApp.getActive() (the BOUND spreadsheet),
 * not openById(). setActiveSheet() only changes the user's visible tab when
 * called on the active spreadsheet reference. Using openById would silently
 * succeed but the view wouldn't switch — that was the v1 bug.
 */
function openPrepQueue() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet (open this from the spreadsheet UI).";

  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) {
    setupPrepQueueSheet();
    sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  }

  ss.setActiveSheet(sheet);
  return "✅ Opened " + PREP_QUEUE.sheetName;
}


/**
 * Sidebar: clear all data rows in BOTH tables (keeps headers + the INCOMING
 * divider). Checkboxes are removed from the wiped rows so empty rows read
 * clean. The sidebar button asks for confirmation before calling this.
 */
function clearPrepQueue() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) return "ℹ️ Prep Queue sheet doesn't exist yet.";

  var lastRow = sheet.getLastRow();
  if (lastRow < PREP_QUEUE.dataStartRow) return "ℹ️ Queue already empty.";

  var boundary = _getPrepBoundaryRow(sheet);
  var cleared = 0;

  function wipeSegment(startRow, endRow) {
    if (endRow < startRow) return;
    var n = endRow - startRow + 1;
    sheet.getRange(startRow, 1, n, PREP_QUEUE.dataWidth).clearContent();
    try { sheet.getRange(startRow, PREP_QUEUE.cols.DONE, n, 1).removeCheckboxes(); } catch (e) {}
    cleared += n;
  }

  if (boundary > 0) {
    wipeSegment(PREP_QUEUE.dataStartRow, boundary - 1);   // CURRENT table
    wipeSegment(boundary + 2, lastRow);                   // INCOMING table
  } else {
    wipeSegment(PREP_QUEUE.dataStartRow, lastRow);        // legacy single table
  }

  // Clear duplicate-highlight backgrounds/borders on the now-empty rows so
  // the sheet visually resets clean (not just empty cells with stale highlights).
  _refreshPrepQueueDuplicates(sheet);

  return "✅ Cleared " + cleared + " row(s) from " + PREP_QUEUE.sheetName + ".";
}


/**
 * Sidebar "Clear Done": delete every checked-off (struck-through) row in BOTH
 * tables. The divider + headers are untouched; deletes run bottom-up in
 * contiguous runs so row numbers stay valid mid-sweep. Re-tops the CURRENT
 * table's typing buffer afterwards (deleting rows above the divider shrinks it).
 */
function clearDonePrepItems() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) return "ℹ️ Prep Queue sheet doesn't exist yet.";

  var lastRow = sheet.getLastRow();
  if (lastRow < PREP_QUEUE.dataStartRow) return "ℹ️ Queue is empty.";

  var boundary = _getPrepBoundaryRow(sheet);
  var n = lastRow - PREP_QUEUE.dataStartRow + 1;
  var done = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.DONE, n, 1).getValues();

  var rows = [];
  for (var i = 0; i < n; i++) {
    var row = PREP_QUEUE.dataStartRow + i;
    if (_isPrepStructureRow(row, boundary)) continue;
    if (done[i][0] === true) rows.push(row);
  }
  if (rows.length === 0) return "ℹ️ No checked-off rows to clear.";

  // Delete bottom-up in contiguous runs (fewer calls, stable row numbers).
  var runEnd = rows[rows.length - 1];
  var runStart = runEnd;
  for (var j = rows.length - 2; j >= -1; j--) {
    if (j >= 0 && rows[j] === runStart - 1) { runStart = rows[j]; continue; }
    sheet.deleteRows(runStart, runEnd - runStart + 1);
    if (j >= 0) { runEnd = rows[j]; runStart = runEnd; }
  }

  // Deleting CURRENT rows pulls the divider up — restore the typing buffer.
  var newBoundary = _getPrepBoundaryRow(sheet);
  if (newBoundary > 0) { try { _ensurePrepBuffer(sheet, newBoundary); } catch (e) {} }

  _refreshPrepQueueDuplicates(sheet);
  return "✅ Cleared " + rows.length + " done item(s).";
}


// =======================================================================================
// onEdit DISPATCHER — called from Main.js's onEdit(e)
// =======================================================================================

/**
 * SKU edit on Prep Queue → auto-fill LOCATION + HAND + DATE ADDED.
 * Mirrors LiveSync's location-lookup pattern but for this dedicated sheet.
 *
 * Called from Main.js onEdit(e). Defensive (try/catch swallowed) — never
 * blocks other edit handlers.
 */
function prepQueueOnEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== PREP_QUEUE.sheetName) return;
    if (e.range.getColumn() !== PREP_QUEUE.cols.SKU) return;
    if (e.range.getRow() < PREP_QUEUE.dataStartRow) return;

    var edits = e.range.getValues();
    var startRow = e.range.getRow();

    var boundary = _getPrepBoundaryRow(sheet);

    // Pre-build lookup maps once for the whole batch (cheaper than per-row):
    // MI (location + eBay stock) + Zoho (direct/non-eBay stock). HAND is
    // Zoho-first, matching the All Orders HAND semantics.
    var useMap = edits.length > 3;
    var locationMap = null;
    var inventoryMap = null;
    var zohoMap = null;

    if (useMap) {
      var maps = buildLocationAndInventoryMaps();
      locationMap = maps.locationMap;
      inventoryMap = maps.inventoryMap;
      zohoMap = buildZohoStockMap();
    }

    var nowStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy h:mm a");

    for (var i = 0; i < edits.length; i++) {
      var row = startRow + i;
      if (_isPrepStructureRow(row, boundary)) continue;   // divider / INCOMING header
      var rawSku = String(edits[i][0]).trim();
      var skuLower = rawSku.toLowerCase();

      if (skuLower === "") {
        // SKU cleared — wipe LOCATION/HAND/DATE + the DONE checkbox so the
        // row reads as fully empty.
        sheet.getRange(row, PREP_QUEUE.cols.LOCATION).setValue("");
        sheet.getRange(row, PREP_QUEUE.cols.HAND).setValue("");
        sheet.getRange(row, PREP_QUEUE.cols.DATE_ADDED).setValue("");
        try {
          sheet.getRange(row, PREP_QUEUE.cols.DONE).removeCheckboxes().clearContent();
        } catch (cbClr) {}
        continue;
      }

      // Resolve LOCATION + HAND (Zoho-first HAND, MI fallback; no committed sub).
      var r = _prepResolveStock(skuLower, locationMap, inventoryMap, zohoMap);
      sheet.getRange(row, PREP_QUEUE.cols.LOCATION).setValue(r.location);
      sheet.getRange(row, PREP_QUEUE.cols.HAND).setValue(r.hand);

      // Stamp DATE ADDED. Re-stamp on EVERY SKU edit (the row's "added" date
      // is when this SKU landed, not when the row was first created — if you
      // overwrite the SKU, that's a new todo item).
      sheet.getRange(row, PREP_QUEUE.cols.DATE_ADDED).setValue(nowStr);

      // Plant the DONE checkbox once. Guard on existing validation —
      // insertCheckboxes() resets a cell to unchecked, so never re-plant on a
      // row that already has one (re-typing a SKU must not un-tick it).
      try {
        var doneCell = sheet.getRange(row, PREP_QUEUE.cols.DONE);
        var dv = doneCell.getDataValidation();
        if (!dv || dv.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
          doneCell.insertCheckboxes();
        }
      } catch (cbErr) { try { Logger.log("prepQueueOnEdit: checkbox error: " + cbErr); } catch (_) {} }
    }

    // Keep typable blank rows above the divider when entries land near it.
    if (boundary > 0) { try { _ensurePrepBuffer(sheet, boundary); } catch (bufErr) {} }

    // SKU → eBay listing link (same enrichment as All Orders). One MI read
    // covers the whole edited batch.
    try {
      applySkuLinksToColumn(sheet, PREP_QUEUE.cols.SKU, startRow,
                            startRow + edits.length - 1, buildSkuEnrichmentMap());
    } catch (enrErr) { try { Logger.log("prepQueueOnEdit: SKU link error: " + enrErr); } catch (_) {} }

    // Refresh duplicate-SKU highlight after every edit batch — surfaces
    // dupes the moment they're typed, clears highlights when one of a
    // dup-pair is deleted/changed.
    _refreshPrepQueueDuplicates(sheet);
  } catch (err) {
    // Swallow — don't break other onEdit handlers
    try { Logger.log("prepQueueOnEdit error: " + err); } catch (_) {}
  }
}


// =======================================================================================
// PRIVATE: duplicate-SKU highlight (mirrors the OutOfStock + All Orders pattern)
// =======================================================================================

/**
 * Scans column A, identifies SKUs that appear two or more times (case-insensitive,
 * trimmed), and paints those cells with a soft amber background + thick yellow
 * border so duplicates are unmistakable.
 *
 * IMPORTANT — only touches background + border. The cell's font, alignment,
 * number format, banding etc. are untouched. When a duplicate is removed (or
 * the SKU is changed/cleared), the next call to this function explicitly
 * clears the background and border on the formerly-duped row, restoring its
 * original look without disturbing other formatting.
 *
 * Per-cell explicit clears (not a range-wide reset) are deliberate: a fast
 * follow-up edit on a previously-duplicate cell could otherwise see stale
 * formatting because Sheets batches background/border writes. The explicit
 * per-cell pass ensures every row's state is set in a single, ordered batch.
 *
 * Run after any change that could affect column A:
 *   - setupPrepQueueSheet (initial paint)
 *   - prepQueueOnEdit (live, on every SKU edit)
 *   - addPrepQueueItem (programmatic adds from sidebar Quick Add)
 *   - clearPrepQueue (removes any lingering highlights from cleared rows)
 *   - repairPrepQueueLayout (transitively, via setupPrepQueueSheet)
 */
function _refreshPrepQueueDuplicates(sheet) {
  if (!sheet) return;

  // CRITICAL: scan the entire data band (not just lastRow), otherwise rows
  // where the SKU was deleted but background formatting persists won't get
  // their highlight cleared. `getLastRow()` returns the last row with
  // CONTENT, so a previously-painted row that had its SKU cleared (and
  // dropped off lastRow) keeps its yellow forever. We walk the full
  // data band so every row's background is set explicitly to either
  // "dupe-yellow" or "default (null)".
  var maxScanRow = Math.min(sheet.getMaxRows(), 1000);
  if (maxScanRow < PREP_QUEUE.dataStartRow) return;

  var totalRows = maxScanRow - PREP_QUEUE.dataStartRow + 1;
  var skuRange  = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.SKU, totalRows, 1);
  var skus      = skuRange.getValues();

  // Two-table awareness: the divider + INCOMING header must keep their own
  // styling (yellow band / dark header) — never counted, never repainted.
  // A SKU present in BOTH tables deliberately DOES count as a duplicate:
  // "you're adding an incoming item that's already in current prep" is signal.
  var boundary = _getPrepBoundaryRow(sheet);
  var existingBgs = skuRange.getBackgrounds();

  // Count occurrences (case-insensitive, trimmed)
  var counts = {};
  for (var i = 0; i < skus.length; i++) {
    if (_isPrepStructureRow(PREP_QUEUE.dataStartRow + i, boundary)) continue;
    var k = String(skus[i][0]).trim().toLowerCase();
    if (!k) continue;
    counts[k] = (counts[k] || 0) + 1;
  }

  // Build the FULL backgrounds array for the entire scan range. Rows that
  // are dupes get yellow; everything else (empty rows, single SKUs, formerly
  // duped rows whose SKU was deleted) gets explicit null; the two structural
  // rows pass their existing background through. One batched write.
  var bgs = [];
  var dupeIndexes = [];
  for (var i = 0; i < skus.length; i++) {
    if (_isPrepStructureRow(PREP_QUEUE.dataStartRow + i, boundary)) {
      bgs.push([existingBgs[i][0]]);
      continue;
    }
    var k = String(skus[i][0]).trim().toLowerCase();
    if (k && counts[k] >= 2) {
      bgs.push(['#fff3b0']);
      dupeIndexes.push(i);
    } else {
      bgs.push([null]);
    }
  }
  skuRange.setBackgrounds(bgs);

  // Borders: clear per SEGMENT (so the divider band's own thick borders and
  // the INCOMING header's underline survive), then add the thick yellow
  // border per dupe (small N, typically 0-4 calls). The clear-then-paint
  // sequence is safe within a single execution because Apps Script applies
  // queued writes in order; we flush at the end so external follow-up edits
  // see the final state, not an intermediate one.
  var clearSpans = [];
  if (boundary > 0) {
    if (boundary - 1 >= PREP_QUEUE.dataStartRow) clearSpans.push([PREP_QUEUE.dataStartRow, boundary - 1]);
    if (maxScanRow >= boundary + 2)              clearSpans.push([boundary + 2, maxScanRow]);
  } else {
    clearSpans.push([PREP_QUEUE.dataStartRow, maxScanRow]);
  }
  clearSpans.forEach(function(span) {
    sheet.getRange(span[0], PREP_QUEUE.cols.SKU, span[1] - span[0] + 1, 1)
      .setBorder(false, false, false, false, false, false, null, null);
  });
  for (var j = 0; j < dupeIndexes.length; j++) {
    var rowIdx = dupeIndexes[j];
    sheet.getRange(PREP_QUEUE.dataStartRow + rowIdx, PREP_QUEUE.cols.SKU)
      .setBorder(true, true, true, true, false, false,
                 '#ffb800', SpreadsheetApp.BorderStyle.SOLID_THICK);
  }

  // Force pending writes to land before the user's next interaction —
  // without this, a fast follow-up edit can see stale formatting because
  // Sheets batches background/border changes.
  SpreadsheetApp.flush();
}
