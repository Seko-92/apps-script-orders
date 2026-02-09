// =======================================================================================
// HELPERS.gs - Core Utility Functions//
// =======================================================================================

/**
 * Gets the column index (0-based) for a given header name
 * @param {Sheet} sheet - The sheet to search
 * @param {string} headerName - The header name to find
 * @returns {number} - 0-based column index, or -1 if not found
 */
function getColumnIndexByHeader(sheet, headerName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === String(headerName).trim().toLowerCase()) {
      return i;
    }
  }
  return -1;
}

/**
 * Finds the boundary row (where TABLE_TWO_IDENTIFIER is located)
 * @returns {number} - Row number of the boundary, or -1 if not found
 */
function getBoundaryRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return -1;
  var values = sheet.getRange(1, SKU_COLUMN, sheet.getLastRow(), 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === TABLE_TWO_IDENTIFIER) return i + 1;
  }
  return -1; 
}

/**
 * Finds the last row with data in a segment
 * @param {number} startRow - Start row of the segment
 * @param {number} endRow - End row of the segment
 * @returns {number} - Last row with data
 */
function findLastDataRowInSegment(startRow, endRow) {
  if (endRow < startRow) return startRow - 1;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  var values = sheet.getRange(startRow, SKU_COLUMN, endRow - startRow + 1, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim().length > 0) return startRow + i;
  }
  return startRow - 1;
}

/**
 * Builds a map of SKU → total committed quantity (PENDING + PREPARING orders)
 * Used to subtract from Master Inventory available for accurate HAND values
 * @returns {Map} - Map of SKU (lowercase) → total committed qty
 */
function getCommittedQuantities() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) return new Map();

  // Read SKU (A), QTY (B), STATUS (F)
  var data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 6).getValues();
  var committed = new Map();

  for (var i = 0; i < data.length; i++) {
    var sku = String(data[i][0]).trim().toLowerCase();
    var qty = parseInt(data[i][1]) || 0;
    var status = String(data[i][5]).trim().toUpperCase();

    if (!sku || sku === TABLE_TWO_IDENTIFIER.toLowerCase()) continue;
    if (status !== 'PENDING' && status !== 'PREPARING') continue;

    committed.set(sku, (committed.get(sku) || 0) + qty);
  }

  return committed;
}

/**
 * Sets up conditional formatting on the HAND column (Column G)
 * so low-stock highlighting is ALWAYS accurate and never stale.
 * Replaces all manual setBackground calls.
 */
function setupHandConditionalFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  // Remove any existing HAND highlight rules to avoid duplicates
  var rules = sheet.getConditionalFormatRules();
  var filtered = [];
  for (var i = 0; i < rules.length; i++) {
    var ranges = rules[i].getRanges();
    var isHandRule = false;
    for (var j = 0; j < ranges.length; j++) {
      if (ranges[j].getColumn() === HAND_COLUMN && ranges[j].getNumColumns() === 1) {
        isHandRule = true;
        break;
      }
    }
    if (!isHandRule) filtered.push(rules[i]);
  }

  // Clear stale manual backgrounds ONLY on data rows (skip boundary + header)
  var boundary = getBoundaryRow();
  var lastRow = Math.max(sheet.getLastRow(), DATA_START_ROW);

  // eBay data rows
  if (boundary > DATA_START_ROW) {
    var ebayCount = boundary - 1 - DATA_START_ROW + 1;
    if (ebayCount > 0) {
      sheet.getRange(DATA_START_ROW, HAND_COLUMN, ebayCount, 1).setBackground(null);
    }
  }

  // DIRECT data rows (skip boundary row and DIRECT header row)
  if (boundary > 0 && boundary + 2 <= lastRow) {
    var directStart = boundary + 2;
    var directCount = lastRow - directStart + 1;
    if (directCount > 0) {
      sheet.getRange(directStart, HAND_COLUMN, directCount, 1).setBackground(null);
    }
  }

  // Build conditional formatting rule using custom formula
  // Only triggers on numeric values <= 20, ignores empty cells and text
  var handRange = sheet.getRange(DATA_START_ROW, HAND_COLUMN, 1000, 1);
  var formula = "=AND(ISNUMBER(G" + DATA_START_ROW + "), G" + DATA_START_ROW + "<=20)";
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground("#FF6B6B")
    .setFontColor("#FFFFFF")
    .setRanges([handRange])
    .build();

  filtered.push(rule);
  sheet.setConditionalFormatRules(filtered);
}

/**
 * Gets the live update state from Settings sheet
 * @returns {string} - "ON" or "OFF"
 */
function getLiveUpdateState() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LIVE_UPDATE_SHEET);
  return s ? s.getRange(LIVE_UPDATE_TOGGLE_CELL).getValue() : "OFF";
}

/**
 * Toggles the live update state
 * @param {string} st - "ON" or "OFF"
 * @returns {string} - The new state
 */
function toggleLiveUpdate(st) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(LIVE_UPDATE_SHEET) || ss.insertSheet(LIVE_UPDATE_SHEET).hideSheet();
  s.getRange(LIVE_UPDATE_TOGGLE_CELL).setValue(st);
  return st;
}
