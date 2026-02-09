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
// HIGHLIGHT DUPLICATES (Auto via Conditional Formatting)
// =======================================================================================

/**
 * Sets up automatic COUNTIF-based conditional formatting for duplicate SKUs.
 * Highlights update in real-time as data changes - no triggers needed.
 * First occurrence: dark red (#dd7e6b), subsequent: light red (#e6b8af).
 * Called from onOpen() to ensure rules are always active.
 */
function setupDuplicateHighlighting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  var COLOR_FIRST = "#dd7e6b";
  var COLOR_DUPE = "#e6b8af";

  // Remove any existing duplicate highlight rules first
  removeDuplicateHighlightRules(sheet);

  var rules = sheet.getConditionalFormatRules();
  var skuRange = sheet.getRange(DATA_START_ROW, SKU_COLUMN, 1000, 1);
  var ref = "A" + DATA_START_ROW;

  // Rule 1 (higher priority): First occurrence of a duplicate SKU → dark red
  // COUNTIF(A$1:A4, A4)=1 means this is the first time the value appears (top-down)
  // COUNTIF(A:A, A4)>1 means the value appears more than once overall
  var firstFormula = '=AND(' + ref + '<>"", COUNTIF(A$1:' + ref + ',' + ref + ')=1, COUNTIF(A:A,' + ref + ')>1)';
  var firstRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(firstFormula)
    .setBackground(COLOR_FIRST)
    .setRanges([skuRange])
    .build();

  // Rule 2 (lower priority): All duplicate SKUs → light red
  // First occurrences already matched Rule 1 above, so they stay dark red
  var dupeFormula = '=AND(' + ref + '<>"", COUNTIF(A:A,' + ref + ')>1)';
  var dupeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(dupeFormula)
    .setBackground(COLOR_DUPE)
    .setRanges([skuRange])
    .build();

  rules.push(firstRule);
  rules.push(dupeRule);
  sheet.setConditionalFormatRules(rules);
}

/**
 * Button handler: enables duplicate highlighting and returns a count.
 * Sets up the auto CF rules, then scans data to report how many dupes exist.
 */
function highlightAllDuplicates() {
  setupDuplicateHighlighting();

  // Scan data and return count for UI feedback
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  var skuCount = {};

  var table1LastData = findLastDataRowInSegment(DATA_START_ROW, boundary - 1);
  if (table1LastData >= DATA_START_ROW) {
    var data = sheet.getRange(DATA_START_ROW, 1, table1LastData - DATA_START_ROW + 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      var sku = String(data[i][0]).trim().toUpperCase();
      if (sku && sku !== TABLE_TWO_IDENTIFIER) {
        skuCount[sku] = (skuCount[sku] || 0) + 1;
      }
    }
  }

  var table2Start = boundary + 2;
  var table2LastData = findLastDataRowInSegment(table2Start, sheet.getLastRow());
  if (table2LastData >= table2Start) {
    var data = sheet.getRange(table2Start, 1, table2LastData - table2Start + 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      var sku = String(data[i][0]).trim().toUpperCase();
      if (sku && sku !== TABLE_TWO_IDENTIFIER) {
        skuCount[sku] = (skuCount[sku] || 0) + 1;
      }
    }
  }

  var duplicateSkus = 0;
  var totalCells = 0;
  for (var sku in skuCount) {
    if (skuCount[sku] > 1) {
      duplicateSkus++;
      totalCells += skuCount[sku];
    }
  }

  if (duplicateSkus === 0) {
    return "✅ Auto-highlight enabled. No duplicates found!";
  }

  return "✅ Auto-highlight enabled. " + duplicateSkus + " duplicate SKUs (" + totalCells + " cells)";
}

/**
 * Removes duplicate highlight conditional formatting rules.
 * Identifies rules by: CUSTOM_FORMULA on SKU column containing COUNTIF,
 * or old marker formulas (=1=1, =2=2) from previous version.
 */
function removeDuplicateHighlightRules(sheet) {
  var rules = sheet.getConditionalFormatRules();
  var filtered = [];
  for (var i = 0; i < rules.length; i++) {
    var bc = rules[i].getBooleanCondition();
    if (bc && bc.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
      var values = bc.getCriteriaValues();
      if (values.length > 0) {
        var formula = values[0];
        // Old marker formulas from previous version
        if (formula === '=1=1' || formula === '=2=2') {
          continue;
        }
        // New COUNTIF-based formulas - check it targets SKU column
        var ranges = rules[i].getRanges();
        var isSkuColumn = false;
        for (var j = 0; j < ranges.length; j++) {
          if (ranges[j].getColumn() === SKU_COLUMN && ranges[j].getNumColumns() === 1) {
            isSkuColumn = true;
            break;
          }
        }
        if (isSkuColumn && formula.indexOf('COUNTIF') !== -1) {
          continue;
        }
      }
    }
    filtered.push(rules[i]);
  }
  sheet.setConditionalFormatRules(filtered);
}

/**
 * Button handler: clears duplicate highlights by removing CF rules.
 * Since we never touched actual cell backgrounds, removing rules
 * perfectly restores the original appearance.
 */
function clearAllDuplicateHighlights() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);

  // Check if any duplicate highlight rules exist
  var rules = sheet.getConditionalFormatRules();
  var hasDupeRules = false;
  for (var i = 0; i < rules.length; i++) {
    var bc = rules[i].getBooleanCondition();
    if (bc && bc.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
      var values = bc.getCriteriaValues();
      if (values.length > 0) {
        var formula = values[0];
        if (formula === '=1=1' || formula === '=2=2') { hasDupeRules = true; break; }
        var ranges = rules[i].getRanges();
        for (var j = 0; j < ranges.length; j++) {
          if (ranges[j].getColumn() === SKU_COLUMN && formula.indexOf('COUNTIF') !== -1) {
            hasDupeRules = true; break;
          }
        }
        if (hasDupeRules) break;
      }
    }
  }

  if (!hasDupeRules) {
    return "ℹ️ No highlights to clear.";
  }

  removeDuplicateHighlightRules(sheet);

  // Clean up old PropertiesService data if it exists from previous version
  PropertiesService.getScriptProperties().deleteProperty('DUPE_ORIGINAL_BGS');

  return "✅ Duplicate highlights cleared.";
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