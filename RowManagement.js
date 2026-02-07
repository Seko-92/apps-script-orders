// =======================================================================================
// ROW_MANAGEMENT.gs - v2.5 SIMPLE (Copies from eBay which works!)
// =======================================================================================

/**
 * Deletes empty rows while preserving buffer rows
 * @param {number} t - Table number (1 or 2)
 * @returns {string} - Status message
 */
function deleteEmptyRows(t) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  var b = getBoundaryRow();
  var start = (t === 1) ? DATA_START_ROW : b + 2;
  var end = (t === 1) ? b - 1 : sheet.getMaxRows();
  var last = findLastDataRowInSegment(start, end);
  
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
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
  var sourceRange = sheet.getRange(DATA_START_ROW, 1, 1, 8);
  var targetRange = sheet.getRange(lastRow + 1, 1, rowsToAdd, 8);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sheet.setRowHeights(lastRow + 1, rowsToAdd, 30);
}

/**
 * Adds rows to Table 1 (eBay) - PUSHES DIRECT TABLE DOWN
 * @param {number} n - Number of rows to add
 * @returns {string} - Status message
 */
function addRowsTableOne(n) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  var lastUsedRow = findLastDataRowInSegment(DATA_START_ROW, boundary - 1);
  var insertionPoint = (lastUsedRow < DATA_START_ROW) ? DATA_START_ROW : lastUsedRow + 1;
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  var numRows = parseInt(n);
  var lastRow = sheet.getLastRow();
  
  // Insert the rows at the end
  sheet.insertRowsAfter(lastRow, numRows);
  
  // Copy format from eBay table's first data row (which works perfectly!)
  var ebaySourceRow = DATA_START_ROW;
  var sourceRange = sheet.getRange(ebaySourceRow, 1, 1, 8);
  var targetRange = sheet.getRange(lastRow + 1, 1, numRows, 8);
  
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  
  removeExistingBoundaryProtection(sheet);
  
  var boundaryRange = sheet.getRange(boundary, 1, 1, 8);
  var protection = boundaryRange.protect();
  protection.setDescription('DIRECT_BOUNDARY_PROTECTED');
  protection.setWarningOnly(true);
  
  var headerRange = sheet.getRange(boundary + 1, 1, 1, 8);
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  removeExistingBoundaryProtection(sheet);
  return "✅ Boundary protection removed.";
}

function validateBoundaryIntegrity() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  
  if (boundary === -1) {
    Logger.log("⚠️ CRITICAL: DIRECT boundary row not found!");
    return false;
  }
  
  var cellValue = sheet.getRange(boundary, 1).getValue();
  if (String(cellValue).toUpperCase().indexOf("DIRECT") === -1) {
    Logger.log("⚠️ WARNING: Boundary row " + boundary + " doesn't contain 'DIRECT'. Value: " + cellValue);
    return false;
  }
  
  Logger.log("✅ Boundary integrity OK. DIRECT is at row " + boundary);
  return true;
}

// =======================================================================================
// HIGHLIGHT DUPLICATES
// =======================================================================================

function highlightAllDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  
  var COLOR_FIRST = "#dd7e6b";
  var COLOR_DUPE = "#e6b8af";
  var skuLocations = {};
  
  var table1Start = DATA_START_ROW;
  var table1End = boundary - 1;
  var table1LastData = findLastDataRowInSegment(table1Start, table1End);
  
  if (table1LastData >= table1Start) {
    var table1Data = sheet.getRange(table1Start, 1, table1LastData - table1Start + 1, 1).getValues();
    for (var i = 0; i < table1Data.length; i++) {
      var sku = String(table1Data[i][0]).trim().toUpperCase();
      if (!sku || sku === TABLE_TWO_IDENTIFIER) continue;
      var actualRow = table1Start + i;
      if (!skuLocations[sku]) skuLocations[sku] = [];
      skuLocations[sku].push(actualRow);
    }
  }
  
  var table2Start = boundary + 2;
  var table2End = sheet.getLastRow();
  var table2LastData = findLastDataRowInSegment(table2Start, table2End);
  
  if (table2LastData >= table2Start) {
    var table2Data = sheet.getRange(table2Start, 1, table2LastData - table2Start + 1, 1).getValues();
    for (var i = 0; i < table2Data.length; i++) {
      var sku = String(table2Data[i][0]).trim().toUpperCase();
      if (!sku || sku === TABLE_TWO_IDENTIFIER) continue;
      var actualRow = table2Start + i;
      if (!skuLocations[sku]) skuLocations[sku] = [];
      skuLocations[sku].push(actualRow);
    }
  }
  
  var duplicateSkus = 0;
  var totalHighlighted = 0;
  
  for (var sku in skuLocations) {
    var rows = skuLocations[sku];
    if (rows.length > 1) {
      duplicateSkus++;
      for (var j = 0; j < rows.length; j++) {
        var row = rows[j];
        var cell = sheet.getRange(row, 1);
        cell.setBackground(j === 0 ? COLOR_FIRST : COLOR_DUPE);
        totalHighlighted++;
      }
    }
  }
  
  if (duplicateSkus === 0) {
    return "✅ No duplicates found. All SKUs are unique!";
  }
  
  return "✅ Found " + duplicateSkus + " duplicate SKUs (" + totalHighlighted + " cells highlighted)";
}

function clearAllDuplicateHighlights() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  var highlightColors = ["#dd7e6b", "#e6b8af"];
  var clearedCount = 0;
  
  var table1Start = DATA_START_ROW;
  var table1End = boundary - 1;
  var table1LastData = findLastDataRowInSegment(table1Start, table1End);
  
  if (table1LastData >= table1Start) {
    var range1 = sheet.getRange(table1Start, 1, table1LastData - table1Start + 1, 1);
    var backgrounds1 = range1.getBackgrounds();
    for (var i = 0; i < backgrounds1.length; i++) {
      var bg = String(backgrounds1[i][0]).toLowerCase();
      if (highlightColors.indexOf(bg) !== -1) {
        backgrounds1[i][0] = null;
        clearedCount++;
      }
    }
    range1.setBackgrounds(backgrounds1);
  }
  
  var table2Start = boundary + 2;
  var table2End = sheet.getLastRow();
  var table2LastData = findLastDataRowInSegment(table2Start, table2End);
  
  if (table2LastData >= table2Start) {
    var range2 = sheet.getRange(table2Start, 1, table2LastData - table2Start + 1, 1);
    var backgrounds2 = range2.getBackgrounds();
    for (var i = 0; i < backgrounds2.length; i++) {
      var bg = String(backgrounds2[i][0]).toLowerCase();
      if (highlightColors.indexOf(bg) !== -1) {
        backgrounds2[i][0] = null;
        clearedCount++;
      }
    }
    range2.setBackgrounds(backgrounds2);
  }
  
  if (clearedCount === 0) {
    return "ℹ️ No highlights to clear.";
  }
  
  return "✅ Cleared " + clearedCount + " highlighted cells.";
}

// =======================================================================================
// LEGACY FUNCTIONS
// =======================================================================================

function consolidateTable(tableNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  var startRow = (tableNumber === 1) ? DATA_START_ROW : boundary + 2;
  var endRow = (tableNumber === 1) ? boundary - 2 : sheet.getLastRow();
  var lastDataRow = findLastDataRowInSegment(startRow, endRow);
  if (lastDataRow < startRow) return "No data found.";

  var range = sheet.getRange(startRow, 1, lastDataRow - startRow + 1, DATA_WIDTH);
  var data = range.getValues();
  var map = new Map();

  data.forEach(row => {
    var sku = String(row[0]).trim().toUpperCase();
    if (!sku || sku === TABLE_TWO_IDENTIFIER) return;

    if (map.has(sku)) {
      var exist = map.get(sku);
      exist[1] = (parseFloat(exist[1]) || 0) + (parseFloat(row[1]) || 0);
      var newOrder = String(row[3]).trim();
      if (newOrder && exist[3].indexOf(newOrder) === -1) {
        exist[3] = exist[3] + " / " + newOrder;
      }
    } else {
      map.set(sku, [...row]);
    }
  });

  var out = Array.from(map.values());
  range.clearContent();
  
  if (out.length > 0) {
    sheet.getRange(startRow, 1, out.length, DATA_WIDTH).setValues(out);
  }
  
  return "⚠️ Merged SKUs (Warning: May affect n8n duplicate detection)";
}

function runMergeEbayDuplicates() { return consolidateTable(1); }
function runMergeDirectDuplicates() { return consolidateTable(2); }