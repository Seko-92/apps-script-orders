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

/**
 * Bold border colors for duplicate Sales Order group indicators.
 * Each group gets a thick colored LEFT border on column D.
 * 20 distinctive colors that cycle if more groups exist.
 */
var ORDER_BORDER_COLORS = [
  "#1a73e8",  // Google Blue
  "#e53935",  // Red
  "#43a047",  // Green
  "#fb8c00",  // Orange
  "#8e24aa",  // Purple
  "#00acc1",  // Cyan
  "#d81b60",  // Pink
  "#6d4c41",  // Brown
  "#3949ab",  // Indigo
  "#00897b",  // Teal
  "#c0ca33",  // Lime
  "#f4511e",  // Deep Orange
  "#5e35b1",  // Deep Purple
  "#039be5",  // Light Blue
  "#7cb342",  // Light Green
  "#ffb300",  // Amber
  "#1e88e5",  // Blue
  "#e91e63",  // Hot Pink
  "#26a69a",  // Medium Teal
  "#546e7a",  // Blue Grey
];

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

/**
 * Applies colored left border tabs on Column D for duplicate Sales Order groups.
 * Each group gets a unique thick left border color — no background fill.
 * Clears all previous borders first, then re-applies for current duplicates.
 * Skips DIRECT boundary row and its header.
 * Called from onOpen() and auto-refreshed on edits.
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

  if (duplicateOrders.length === 0) return null;

  // 4. Apply thick colored left border per group
  for (var i = 0; i < duplicateOrders.length; i++) {
    var color = ORDER_BORDER_COLORS[i % ORDER_BORDER_COLORS.length];
    var rows = orderRows[duplicateOrders[i]];

    for (var j = 0; j < rows.length; j++) {
      var cell = sheet.getRange(rows[j], Schema.cols.SALES_ORDER);
      cell.setBorder(null, true, null, null, null, null, color, SpreadsheetApp.BorderStyle.SOLID_THICK);
    }
  }

  // Force the clear-then-apply sequence to land before any subsequent reads.
  // Without flush(), Sheets can batch the writes and the clear can lose to
  // the re-apply in certain races, leaving stale borders on rows whose SO
  // was just cleared or whose row was just deleted.
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