// =======================================================================================
// UI_SERVICE.gs - Sidebar and UI Functions
// =======================================================================================

/**
 * Shows the control panel sidebar
 */
function showSidebar() {
  var html = HtmlService.createTemplateFromFile('Sidebar').evaluate().setTitle('⚙️ Control Panel');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Updates both the Date and Time cells in the sheet.
 * Best used with a 1-minute time-driven trigger.
 */
function updateSheetClock() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var now = new Date();
  
  // Explicitly use Houston Timezone
  var houstonTZ = "America/Chicago";

  // Update the Date Cell
  var dateStr = Utilities.formatDate(now, houstonTZ, "EEEE, MMMM d, yyyy");
  sheet.getRange("B1").setValue(dateStr);

  // Update the Clock Cell (D1)
  var timeStr = Utilities.formatDate(now, houstonTZ, "h:mm a");
  sheet.getRange(CLOCK_CELL).setValue(timeStr);
  
  return "Houston Time Synced";
}

/**
 * Saves sidebar module order (called from sidebar)
 * @param {Array} order - Array of module IDs in order
 */
function saveSidebarOrder(order) {
  var userProps = PropertiesService.getUserProperties();
  userProps.setProperty('sidebarOrder', JSON.stringify(order));
}

/**
 * Gets saved sidebar module order
 * @returns {Array} - Array of module IDs
 */
function getSidebarOrder() {
  var userProps = PropertiesService.getUserProperties();
  var order = userProps.getProperty('sidebarOrder');
  return order ? JSON.parse(order) : null;
}
