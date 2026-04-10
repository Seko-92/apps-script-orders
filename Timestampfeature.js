// =======================================================================================
// TIMESTAMP_FEATURE.gs - v3.0 - Auto-Update Last Order Timestamp
// =======================================================================================

/**
 * ONE-TIME SETUP: Initialize the timestamp cell (F2)
 * Run this ONCE to create the formatted cell
 */
function setupTimestampCell() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  
  if (!sheet) {
    return "❌ Sheet '" + MAIN_SHEET_NAME + "' not found.";
  }
  
  var cell = sheet.getRange("F2");
  
  // Set initial value
  cell.setValue("📦 Last Order: Never");
  
  // Format the cell
  cell.setFontSize(10);
  cell.setFontWeight("bold");
  cell.setFontColor("#1976D2"); // Blue color
  cell.setHorizontalAlignment("left");
  cell.setVerticalAlignment("middle");
  
  // Optional: Light blue background
  cell.setBackground("#E3F2FD");
  
  return "✅ Timestamp cell (F2) initialized!";
}

/**
 * Updates the timestamp in cell F2
 * Called automatically by triggerN8NWebhook() after sync
 * 
 * @param {string} cellAddress - Cell address (default: "F2")
 */
function updateLastOrderTimestamp(cellAddress) {
  cellAddress = cellAddress || "F2";
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  
  if (!sheet) {
    Logger.log("Sheet not found: " + MAIN_SHEET_NAME);
    return;
  }
  
  var cell = sheet.getRange(cellAddress);
  
  // Get current time in Houston timezone
  var now = new Date();
  var houstonTZ = "America/Chicago";
  
  // Format: "📦 Last Order: 1/19/2026 10:23 PM"
  var dateStr = Utilities.formatDate(now, houstonTZ, "M/d/yyyy");
  var timeStr = Utilities.formatDate(now, houstonTZ, "h:mm a");
  
  var timestamp = "📦 Last Order: " + dateStr + " " + timeStr;
  
  // Update cell
  cell.setValue(timestamp);
  
  // Ensure formatting is preserved
  cell.setFontSize(10);
  cell.setFontWeight("bold");
  cell.setFontColor("#1976D2");
  cell.setBackground("#E3F2FD");
  
  Logger.log("Timestamp updated: " + timestamp);
}

/**
 * Alternative: Get formatted timestamp string without updating cell
 * Useful for returning to sidebar
 * 
 * @returns {string} - Formatted timestamp
 */
function getFormattedTimestamp() {
  var now = new Date();
  var houstonTZ = "America/Chicago";
  var dateStr = Utilities.formatDate(now, houstonTZ, "M/d/yyyy");
  var timeStr = Utilities.formatDate(now, houstonTZ, "h:mm a");
  return "📦 Last sync: " + dateStr + " " + timeStr;
}

/**
 * Test function - Updates timestamp immediately
 */
function testTimestampUpdate() {
  updateLastOrderTimestamp("F2");
  return "✅ Timestamp updated in F2!";
}

// =======================================================================================
// LOCATION UPDATE SHEET - Auto-Timestamp on SKU Edit
// =======================================================================================

/**
 * Stamps column D with Houston time when SKU (column B) is added/edited.
 * Clears column D when SKU is removed.
 * Called from onEdit(e) in Main.gs.
 * @param {Event} e - The edit event
 */
function locationUpdateTimestamp(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    var sheetName = sheet.getName();

    Logger.log("locationUpdateTimestamp fired on sheet: " + sheetName + ", cell: " + range.getA1Notation());

    if (sheetName !== LOCATION_UPDATE_SHEET) return;

    // Only trigger on SKU column (B = column 2)
    var col = range.getColumn();
    Logger.log("Edited column: " + col);
    if (col !== 2) return;

    var startRow = range.getRow();
    var numRows = range.getNumRows();

    // Skip header rows (row 1 and 2)
    if (startRow <= 2 && startRow + numRows - 1 <= 2) return;

    var houstonTZ = "America/Chicago";
    var now = new Date();
    var timestamp = Utilities.formatDate(now, houstonTZ, "M/d/yyyy h:mm a");

    for (var i = 0; i < numRows; i++) {
      var row = startRow + i;
      if (row <= 2) continue; // skip header rows

      var skuValue = String(sheet.getRange(row, 2).getValue()).trim();
      var tsCell = sheet.getRange(row, 4); // Column D = Time Stamp

      Logger.log("Row " + row + " SKU: '" + skuValue + "'");

      if (skuValue === "") {
        tsCell.setValue("");
      } else {
        tsCell.setValue(timestamp);
      }
    }
    Logger.log("locationUpdateTimestamp completed successfully");
  } catch (err) {
    Logger.log("locationUpdateTimestamp ERROR: " + err.message);
  }
}
