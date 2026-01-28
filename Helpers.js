// =======================================================================================
// HELPERS.gs - Core Utility Functions
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
