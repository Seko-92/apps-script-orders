// =======================================================================================
// TIMESTAMP_FEATURE.gs - v3.0 - Auto-Update Last Order Timestamp
// =======================================================================================

/**
 * ONE-TIME SETUP: Initialize the timestamp cell (F2)
 * Run this ONCE to create the formatted cell
 */
function setupTimestampCell() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
