// =======================================================================================
// TIMESTAMP_FEATURE.gs - v3.0 - Auto-Update Last Order Timestamp
// =======================================================================================

/**
 * ⚠️ DEPRECATED 2026-05-19 — DO NOT RUN.
 * F2 is now the Pick ID for Shipping dropdown anchor (validated cell).
 * Running this function would (a) overwrite the picker's current selection,
 * (b) attempt to setBackground on a validated cell, and (c) blow up the
 * dropdown's selectable list. See deprecation note on updateLastOrderTimestamp
 * below for the full story.
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
 * ⚠️ DEPRECATED 2026-05-19 — DO NOT RE-WIRE INTO triggerN8NWebhook() OR ANY
 * RUNTIME PATH WITHOUT FIRST CHECKING THE CALLER CELL FOR DATA VALIDATION.
 *
 * Original purpose: write "📦 Last Order: M/d/yyyy h:mm AM/PM" to F2 after
 * every n8n sync, as a "you have new orders" cue in the banner.
 *
 * Why deprecated:
 *   - The Service Bay v6 System Pulse in E1 already shows live sync time
 *     ("Last sync · 5:20 PM 🟢 ALIVE") driven by the Activity Log. The F2
 *     stamp was a cruder version of the same idea.
 *   - 2026-05-19 layout compaction moved the Pick ID for Shipping dropdown
 *     to F2, putting a data-validation rule on the cell. Writing a timestamp
 *     string into a validation-protected cell throws — see Gotcha #2 about
 *     cells with data validation being read-only for our code.
 *
 * If you ever need this functionality again, pick a different target cell
 * that is NOT validation-protected (E1 is already taken; A1 is the HQ chip;
 * pick something in the unused banner area or add a dedicated row).
 *
 * `setupTimestampCell()` (above) and `testTimestampUpdate()` (below) also
 * target F2 — same deprecation applies.
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
 * ⚠️ DEPRECATED 2026-05-19 — DO NOT RUN.
 * Calls updateLastOrderTimestamp("F2") which will fail with a data-validation
 * error since F2 is now the Pick ID for Shipping dropdown. See deprecation
 * note on updateLastOrderTimestamp above.
 */
function testTimestampUpdate() {
  updateLastOrderTimestamp("F2");
  return "✅ Timestamp updated in F2!";
}

// =======================================================================================
// LOCATION UPDATE SHEET - Auto-Timestamp on SKU Edit
// =======================================================================================

/**
 * ⚠️ ORPHANED (2026-05-13) — superseded by locationUpdateOnEdit() in
 * LocationUpdate.js. This simple-trigger version was the root cause of the
 * "location/timestamp sometimes fails to appear" issue: simple triggers fail
 * silently when openById is needed, so the location lookup never ran. The
 * new handler runs in onEditInstallable (full permissions) and does the
 * complete COUNTER + LOCATION + TIMESTAMP fill in one pass.
 *
 * This function is kept here for manual debugging only — it is no longer
 * called from any trigger. Safe to delete in a future cleanup pass.
 *
 * Stamps column D with Houston time when SKU (column B) is added/edited.
 * Clears column D when SKU is removed.
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
