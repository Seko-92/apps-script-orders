// =======================================================================================
// ROW_MANAGEMENT.gs - v2.5 SIMPLE (Copies from eBay which works!)
// =======================================================================================

/**
 * Deletes empty rows while preserving buffer rows
 * @param {number} t - Table number (1 or 2)
 * @returns {string} - Status message
 */
function deleteEmptyRows(t) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  var b = getBoundaryRow();
  var start = (t === 1) ? Schema.dataStartRow : b + 2;
  var end   = (t === 1) ? b - 1               : sheet.getMaxRows();
  var last  = findLastDataRowInSegment(start, end);

  var delStart = (t === 1) ? last + 4 : last + MAX_EMPTY_ROWS_TO_KEEP + 1;

  if (t === 1 && delStart >= b) return "ℹ️ 3-row buffer already exists.";

  if (delStart < end) {
    sheet.deleteRows(delStart, end - delStart + 1);
    return "✅ Cleanup complete (3-row buffer preserved).";
  }
  return "ℹ️ Already clean.";
}

function runDeleteEmptyRowsTableOne() { return deleteEmptyRows(1); }
function runDeleteEmptyRowsTableTwo() { return deleteEmptyRows(2); }

/**
 * Ensures the DIRECT table always has at least 3 empty buffer rows
 * with proper data formatting (not header formatting).
 * Called automatically via onChange when rows are deleted.
 */
function ensureDirectTableBuffer() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  var boundary = getBoundaryRow();
  if (boundary === -1) return;

  var BUFFER_SIZE = 3;
  var directDataStart = boundary + 2; // First data row after DIRECT header
  var lastRow = sheet.getLastRow();

  // Find last data row in DIRECT table
  var lastDataRow = findLastDataRowInSegment(directDataStart, lastRow);

  // Count empty rows after last data (or after header if no data)
  var emptyStart = (lastDataRow >= directDataStart) ? lastDataRow + 1 : directDataStart;
  var emptyCount = lastRow - emptyStart + 1;
  if (emptyStart > lastRow) emptyCount = 0;

  if (emptyCount >= BUFFER_SIZE) return; // Buffer already exists

  var rowsToAdd = BUFFER_SIZE - emptyCount;

  // Add rows at the end of the sheet
  sheet.insertRowsAfter(lastRow, rowsToAdd);

  // Copy formatting from eBay data row (which always has correct format)
  var sourceRange = sheet.getRange(Schema.dataStartRow, 1, 1, Schema.dataWidth);
  var targetRange = sheet.getRange(lastRow + 1, 1, rowsToAdd, Schema.dataWidth);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sheet.setRowHeights(lastRow + 1, rowsToAdd, 30);
}

/**
 * Adds rows to Table 1 (eBay) - PUSHES DIRECT TABLE DOWN
 * @param {number} n - Number of rows to add
 * @returns {string} - Status message
 */
function addRowsTableOne(n) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  var lastUsedRow = findLastDataRowInSegment(Schema.dataStartRow, boundary - 1);
  var insertionPoint = (lastUsedRow < Schema.dataStartRow) ? Schema.dataStartRow : lastUsedRow + 1;
  var rowsToInsert = parseInt(n);

  sheet.insertRowsAfter(insertionPoint, rowsToInsert);

  return "✅ Inserted " + rowsToInsert + " rows. DIRECT moved to Row " + (boundary + rowsToInsert) + ".";
}

/**
 * SIMPLEST FIX: Copy format from eBay table (which works perfectly!)
 * Since both tables should have the same format anyway
 * @param {number} n - Number of rows to add
 * @returns {string} - Status message
 */
function addRowsTableTwo(n) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  var numRows = parseInt(n);
  var lastRow = sheet.getLastRow();

  // Insert the rows at the end
  sheet.insertRowsAfter(lastRow, numRows);

  // Copy format from eBay table's first data row (which always has correct format)
  var sourceRange = sheet.getRange(Schema.dataStartRow, 1, 1, Schema.dataWidth);
  var targetRange = sheet.getRange(lastRow + 1, 1, numRows, Schema.dataWidth);

  // Copy ONLY the format (not content)
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // Set row heights to match
  sheet.setRowHeights(lastRow + 1, numRows, 30);

  return "✅ Added " + numRows + " rows (format copied from eBay table).";
}

// =======================================================================================
// BOUNDARY PROTECTION FUNCTIONS
// =======================================================================================

function protectBoundaryRow() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  
  removeExistingBoundaryProtection(sheet);

  var boundaryRange = sheet.getRange(boundary, 1, 1, Schema.dataWidth);
  var protection = boundaryRange.protect();
  protection.setDescription('DIRECT_BOUNDARY_PROTECTED');
  protection.setWarningOnly(true);

  var headerRange = sheet.getRange(boundary + 1, 1, 1, Schema.dataWidth);
  var headerProtection = headerRange.protect();
  headerProtection.setDescription('DIRECT_HEADER_PROTECTED');
  headerProtection.setWarningOnly(true);
  
  return "✅ Protected DIRECT boundary (Row " + boundary + ") and header (Row " + (boundary + 1) + ").";
}

function removeExistingBoundaryProtection(sheet) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var desc = protections[i].getDescription();
    if (desc === 'DIRECT_BOUNDARY_PROTECTED' || desc === 'DIRECT_HEADER_PROTECTED') {
      protections[i].remove();
    }
  }
}

function unprotectBoundaryRow() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  removeExistingBoundaryProtection(sheet);
  return "✅ Boundary protection removed.";
}

function validateBoundaryIntegrity() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  
  if (boundary === -1) {
    Logger.log("⚠️ CRITICAL: DIRECT boundary row not found!");
    return false;
  }
  
  var cellValue = sheet.getRange(boundary, Schema.cols.SKU).getValue();
  if (String(cellValue).toUpperCase().indexOf(Schema.boundaryMarker) === -1) {
    Logger.log("⚠️ WARNING: Boundary row " + boundary + " doesn't contain '" + Schema.boundaryMarker + "'. Value: " + cellValue);
    return false;
  }
  
  Logger.log("✅ Boundary integrity OK. DIRECT is at row " + boundary);
  return true;
}

// =======================================================================================
// HIGHLIGHT DUPLICATES - Shared Infrastructure
// =======================================================================================

/**
 * Bright color palette for duplicate SKU groups.
 * Each entry: [background, fontColor] — visually matched pairs.
 * 20 pairs, cycles if more groups exist.
 */
var SKU_DUPE_COLORS = [
  ["#ff6d6d", "#7a0000"],  // Bright Red / Dark Red
  ["#4fc3f7", "#01579b"],  // Sky Blue / Navy
  ["#81c784", "#1b5e20"],  // Bright Green / Forest
  ["#ffb74d", "#e65100"],  // Bright Orange / Burnt Orange
  ["#ba68c8", "#4a148c"],  // Bright Purple / Deep Purple
  ["#4dd0e1", "#006064"],  // Bright Cyan / Dark Cyan
  ["#e57373", "#b71c1c"],  // Vivid Coral / Crimson
  ["#fff176", "#f57f17"],  // Bright Yellow / Amber
  ["#aed581", "#33691e"],  // Lime Green / Olive
  ["#ff8a65", "#bf360c"],  // Tangerine / Mahogany
  ["#7986cb", "#1a237e"],  // Bright Indigo / Deep Indigo
  ["#4db6ac", "#004d40"],  // Bright Teal / Dark Teal
  ["#f06292", "#880e4f"],  // Hot Pink / Wine
  ["#dce775", "#827717"],  // Chartreuse / Olive Gold
  ["#64b5f6", "#0d47a1"],  // Dodger Blue / Royal Blue
  ["#ffab91", "#bf360c"],  // Salmon / Rust
  ["#a1887f", "#3e2723"],  // Mocha / Espresso
  ["#90caf9", "#0d47a1"],  // Cornflower / Dark Blue
  ["#ce93d8", "#6a1b9a"],  // Orchid / Plum
  ["#80cbc4", "#00695c"],  // Aquamarine / Emerald
];

// (ORDER_BORDER_COLORS removed 2026-07-14 — the colored left-border tabs
// were replaced by the SO badge glyphs, which survive B&W printing and
// scattered rows. Git history has the palette if ever wanted back.)

// =======================================================================================
// HIGHLIGHT DUPLICATE SKUs (Per-Group, Bright, Auto-Refresh)
// =======================================================================================

/**
 * Sets up per-group duplicate SKU highlighting with matched font colors.
 * Each duplicate SKU group gets its own bright color + dark complementary font.
 * Skips DIRECT boundary row.
 * Called from onOpen() and auto-refreshed on edits.
 */
function setupDuplicateHighlighting() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return null;

  removeDuplicateHighlightRules(sheet);

  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return null;

  var allData = sheet.getRange(Schema.dataStartRow, Schema.cols.SKU, lastRow - Schema.dataStartRow + 1, 1).getValues();
  var boundary = getBoundaryRow();

  var skuCount = {};
  for (var i = 0; i < allData.length; i++) {
    var currentRow = Schema.dataStartRow + i;
    if (boundary > 0 && (currentRow === boundary || currentRow === boundary + 1)) continue;
    var sku = String(allData[i][0]).trim().toUpperCase();
    if (sku && sku !== Schema.boundaryMarker) {
      skuCount[sku] = (skuCount[sku] || 0) + 1;
    }
  }

  var duplicateSkus = [];
  for (var sku in skuCount) {
    if (skuCount[sku] > 1) duplicateSkus.push(sku);
  }

  if (duplicateSkus.length === 0) return null;

  var rules = sheet.getConditionalFormatRules();
  var skuRange = sheet.getRange(Schema.dataStartRow, Schema.cols.SKU, 1000, 1);
  var ref = "A" + Schema.dataStartRow;

  for (var i = 0; i < duplicateSkus.length; i++) {
    var escapedSku = duplicateSkus[i].replace(/"/g, '""');
    var pair = SKU_DUPE_COLORS[i % SKU_DUPE_COLORS.length];

    var formula = '=UPPER(TRIM(' + ref + '))="' + escapedSku + '"';
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(pair[0])
      .setFontColor(pair[1])
      .setRanges([skuRange])
      .build();

    rules.push(rule);
  }

  sheet.setConditionalFormatRules(rules);
}

function highlightAllDuplicates() {
  setupDuplicateHighlighting();
  return "✅ Duplicate SKU highlighting enabled.";
}

function removeDuplicateHighlightRules(sheet) {
  var rules = sheet.getConditionalFormatRules();
  var filtered = [];
  for (var i = 0; i < rules.length; i++) {
    var bc = rules[i].getBooleanCondition();
    if (bc && bc.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
      var values = bc.getCriteriaValues();
      if (values.length > 0) {
        var formula = values[0];
        if (formula === '=1=1' || formula === '=2=2') continue;
        var ranges = rules[i].getRanges();
        var isSkuColumn = false;
        for (var j = 0; j < ranges.length; j++) {
          if (ranges[j].getColumn() === Schema.cols.SKU && ranges[j].getNumColumns() === 1) {
            isSkuColumn = true;
            break;
          }
        }
        if (isSkuColumn && (formula.indexOf('COUNTIF') !== -1 || formula.indexOf('UPPER(TRIM(') !== -1)) {
          continue;
        }
      }
    }
    filtered.push(rules[i]);
  }
  sheet.setConditionalFormatRules(filtered);
}

function clearAllDuplicateHighlights() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  removeDuplicateHighlightRules(sheet);
  PropertiesService.getScriptProperties().deleteProperty('DUPE_ORIGINAL_BGS');
  return "✅ Duplicate SKU highlights cleared.";
}

// =======================================================================================
// DUPLICATE SALES ORDER BORDERS (Per-Group, Colored Left Border Tabs)
// =======================================================================================

// Keycap digits 1-10 — the SO BADGE glyph set (2026-07-14, final form).
// Rendered via NUMBER FORMAT prefix ONLY (display layer): the cell VALUE
// stays the clean order id, so every machine reader is untouched — n8n's
// four All-Orders reader nodes (pinned to UNFORMATTED_VALUE), the
// updateOrderStatus col-D matcher, dedupe, and the rich-text order links.
// Same device as the ▣ kit marker and the ▌ DIRECT divider.
//
// WHY KEYCAP EMOJI (the glyph saga, condensed): the filled circled digits
// (❶❷❸) are illegible at the table's 10px, and making ONLY the glyph render
// larger is IMPOSSIBLE — Sheets normalizes cell-level font writes and
// rich-text run styles into each other on every write (setFontSizes stomps
// run sizes; setRichTextValue resets the cell default from the runs), so a
// format prefix can never durably out-size the value text beside it. Two
// shipped attempts proved this. Keycap emoji solve it a different way:
// they render as colored squares visually LARGER than letterforms at the
// SAME font size — instantly distinguishable, table stays uniform 10px.
// The PRINT pick list maps the badge to a drawn ink circle-digit instead
// (emoji print as gray mush on B&W; see _badgeFromFormat + .so-badge).
// Numbering restarts per TABLE (eBay / DIRECT) and cycles past 10 — a badge
// only needs to be unique among the currently-visible multi-item orders of
// its own table.
var SO_BADGE_GLYPHS = ["1️⃣","2️⃣","3️⃣","4️⃣","5️⃣","6️⃣","7️⃣","8️⃣","9️⃣","🔟"];

/**
 * Multi-item Sales Order marking on Column D — the SO BADGE:
 * a keycap digit (1️⃣2️⃣3️⃣…) prefixed via number format on every row of the
 * group. Survives the aisle sort's scattering (repeated on every row of the
 * group); the print pick list reads the SAME format and renders it as a
 * drawn ink circle-digit (B&W-crisp). The table stays uniform 10px — the
 * keycap's salience comes from the glyph itself, not font size. (The
 * colored left-border tabs this replaced were dropped 2026-07-14; the
 * band-wide border clear remains so legacy bars self-wipe.)
 * Clears stale borders AND stale badge formats first, then re-applies for
 * current duplicates. Skips DIRECT boundary row and its header.
 * Called from onOpen(), auto-refreshed on edits/inserts, and re-run by
 * sortTableByStatusAndLocation (formats travel with the sort, but the
 * repaint keeps assignment canonical — ▣ lesson).
 */
function setupDuplicateSalesOrderHighlighting() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return null;

  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return null;

  var boundary = getBoundaryRow();

  // Data read is bounded by lastRow (only cells that COULD have a SO value).
  var dataRows = lastRow - Schema.dataStartRow + 1;
  var allData = sheet.getRange(Schema.dataStartRow, Schema.cols.SALES_ORDER, dataRows, 1).getValues();

  // 1. Remove any legacy CF rules on column D (from previous highlight approach)
  removeLegacySalesOrderCFRules(sheet);

  // 2. Clear all left borders on column D — using a WIDER range than lastRow.
  //
  //    Why: getLastRow() returns the position of the last row with CONTENT in
  //    ANY column. When a row's SO is cleared (or n8n removes a shipped order
  //    and empties the whole row), lastRow shrinks. The previously-bordered
  //    row falls outside the lastRow-bounded clear range, so its stale left-
  //    border survives forever. The user sees: "I removed the duplicate but
  //    the highlight stays on the empty cell."
  //
  //    Same class of bug we fixed in _refreshPrepQueueDuplicates 2026-05-06.
  //    Cheap fix: extend the clear band a generous margin past lastRow
  //    (capped at sheet.getMaxRows() to stay in bounds).
  var clearLastRow = Math.min(sheet.getMaxRows(), lastRow + 200);
  var clearRowCount = clearLastRow - Schema.dataStartRow + 1;
  var fullRange = sheet.getRange(Schema.dataStartRow, Schema.cols.SALES_ORDER, clearRowCount, 1);
  fullRange.setBorder(null, false, null, null, null, null);

  // 2. Count occurrences, skipping boundary rows
  var orderCount = {};
  var orderRows = {};  // Map order → [row numbers]
  for (var i = 0; i < allData.length; i++) {
    var currentRow = Schema.dataStartRow + i;
    if (boundary > 0 && (currentRow === boundary || currentRow === boundary + 1)) continue;
    var order = String(allData[i][0]).trim();
    if (order) {
      orderCount[order] = (orderCount[order] || 0) + 1;
      if (!orderRows[order]) orderRows[order] = [];
      orderRows[order].push(currentRow);
    }
  }

  // 3. Identify duplicates and assign border colors
  var duplicateOrders = [];
  for (var order in orderCount) {
    if (orderCount[order] > 1) duplicateOrders.push(order);
  }

  // 3b. SO BADGES — rebuild column D's number formats over the whole clear
  // band in ONE batched write. Default '@' (plain text: clears stale badges
  // when a group dissolves AND guards order ids against date coercion);
  // multi-item groups get '"❶ "@' etc. Boundary + DIRECT-header rows keep
  // whatever format they already carry. Runs even when there are NO
  // duplicates — that's what erases the last badge when a group shrinks to
  // one row.
  var bandFormats = fullRange.getNumberFormats();
  for (var f = 0; f < bandFormats.length; f++) {
    var fRow = Schema.dataStartRow + f;
    if (boundary > 0 && (fRow === boundary || fRow === boundary + 1)) continue;
    bandFormats[f][0] = '@';
  }

  // Per-table sequences, numbered top-to-bottom by each group's first row.
  var ebaySeq = [];
  var directSeq = [];
  for (var d = 0; d < duplicateOrders.length; d++) {
    var dupFirstRow = orderRows[duplicateOrders[d]][0];   // rows collected top-down
    if (boundary > 0 && dupFirstRow > boundary) directSeq.push(duplicateOrders[d]);
    else ebaySeq.push(duplicateOrders[d]);
  }
  var byFirstRow = function(a, b) { return orderRows[a][0] - orderRows[b][0]; };
  ebaySeq.sort(byFirstRow);
  directSeq.sort(byFirstRow);
  [ebaySeq, directSeq].forEach(function(seq) {
    for (var s = 0; s < seq.length; s++) {
      var glyph = SO_BADGE_GLYPHS[s % SO_BADGE_GLYPHS.length];
      var gRows = orderRows[seq[s]];
      for (var g = 0; g < gRows.length; g++) {
        bandFormats[gRows[g] - Schema.dataStartRow][0] = '"' + glyph + ' "@';
      }
    }
  });
  fullRange.setNumberFormats(bandFormats);

  // 3c. Column font: uniform 10px across the band. Heals the 12px/14px
  // residue from the abandoned size-split attempts (see the SO_BADGE_GLYPHS
  // comment for why per-glyph sizing is impossible — the keycap glyphs carry
  // their own visual weight at 10px instead).
  var bandSizes = fullRange.getFontSizes();
  for (var fs = 0; fs < bandSizes.length; fs++) {
    var fsRow = Schema.dataStartRow + fs;
    if (boundary > 0 && (fsRow === boundary || fsRow === boundary + 1)) continue;
    bandSizes[fs][0] = 10;
  }
  fullRange.setFontSizes(bandSizes);

  // COLOR BARS DROPPED (2026-07-14, user's call after living with both):
  // the badge answers "which order" — the bar only ever answered "some
  // group", spent scarce color budget, and died on B&W prints. The band-wide
  // border CLEAR above (step 2) stays so legacy bars wipe themselves on the
  // first repaint. Border-apply loop deleted; git history has it.

  // Force the clear-then-apply sequence to land before any subsequent reads.
  SpreadsheetApp.flush();
}

function highlightAllDuplicateSalesOrders() {
  setupDuplicateSalesOrderHighlighting();
  return "✅ Duplicate Sales Order border tabs applied.";
}

/**
 * Removes leftover CF rules from the old background-highlight approach.
 * Safe to call repeatedly — only strips rules targeting column D with TRIM/COUNTIF formulas.
 */
function removeLegacySalesOrderCFRules(sheet) {
  var rules = sheet.getConditionalFormatRules();
  var filtered = [];
  for (var i = 0; i < rules.length; i++) {
    var bc = rules[i].getBooleanCondition();
    if (bc && bc.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
      var values = bc.getCriteriaValues();
      if (values.length > 0) {
        var formula = values[0];
        var ranges = rules[i].getRanges();
        var isOrderColumn = false;
        for (var j = 0; j < ranges.length; j++) {
          if (ranges[j].getColumn() === Schema.cols.SALES_ORDER && ranges[j].getNumColumns() === 1) {
            isOrderColumn = true;
            break;
          }
        }
        if (isOrderColumn && (formula.indexOf('COUNTIF') !== -1 || formula.indexOf('TRIM(') !== -1)) {
          continue;  // Skip (remove) this legacy rule
        }
      }
    }
    filtered.push(rules[i]);
  }
  if (filtered.length !== rules.length) {
    sheet.setConditionalFormatRules(filtered);
  }
}

function clearAllDuplicateSalesOrderHighlights() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return "✅ Nothing to clear.";

  var fullRange = sheet.getRange(Schema.dataStartRow, Schema.cols.SALES_ORDER, lastRow - Schema.dataStartRow + 1, 1);
  fullRange.setBorder(null, false, null, null, null, null);
  return "✅ Duplicate Sales Order border tabs cleared.";
}

// =======================================================================================
// AUTO-REFRESH: Unified handler for both SKU and Sales Order duplicate highlights
// =======================================================================================

/**
 * Refreshes both duplicate highlight systems on any edit to the main sheet.
 * Called from onEditInstallable(e) — triggers on ANY data-area edit,
 * not just column A or D, so highlights update when you clear/edit any cell.
 */
function refreshDuplicateHighlightsOnEdit(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    if (sheet.getName() !== MAIN_SHEET_NAME) return;
    if (range.getRow() < Schema.dataStartRow) return;
    // Only auto-refresh Sales Order highlights (SKU is manual-only)
    setupDuplicateSalesOrderHighlighting();
  } catch (err) { /* silent */ }
}

// consolidateTable, runMergeEbayDuplicates, runMergeDirectDuplicates —
// REMOVED 2026-04-29.
//
// These were the old "merge duplicate SKU rows" feature. Verified zero callers
// in Apps Script, HTML sidebars, and n8n workflows. The function itself
// carried a self-warning ("May affect n8n duplicate detection") because
// merging rows breaks the SKU+SalesOrder dedup contract that doPost relies on.
//
// Modern duplicate handling lives in:
//   - setupDuplicateSalesOrderHighlighting() above (visual color tabs)
//   - setupDuplicateHighlighting() above (SKU group colors)
// These VISUALIZE duplicates rather than destroying data.