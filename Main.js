// =======================================================================================
// MAIN.gs - Entry Points and Triggers
// =======================================================================================

/**
 * Runs when the spreadsheet opens
 * Creates the menu and updates stats
 */
/**
 * Combined onOpen function: 
 * Creates menus for Control Panel & Arcade, updates stats, 
 * enables live sync, and auto-opens the Command Center.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 1. Create the Control Panel Menu
  ui.createMenu('⚙️ Control Panel')
    .addItem('Open Control Panel', 'showSidebar')
    .addToUi();

  // 2. Create the NEW Arcade Menu
  ui.createMenu('🕹️ HQ ARCADE')
    .addItem('Launch HQ Snake', 'showSnakeSidebar')
    .addToUi();
  
  // 3. Run your existing background logic
  updateOrderStatsInSheet();
  toggleLiveUpdate('ON');  // Auto-enable live sync
  
  // 4. AUTO-LOAD the Control Panel on startup
  showSidebar(); 
}

/**
 * Launches the Snake Game Sidebar
 */
function showSnakeSidebar() {
  // Ensure the file name in quotes matches your HTML file name exactly (Snake)
  const html = HtmlService.createHtmlOutputFromFile('Snake')
    .setTitle('HQ COMMAND CENTER: SNAKE')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function autoEnableLiveSync() {
  toggleLiveUpdate('ON');
}

/**
 * onChange trigger - Updates stats when sheet changes
 * @param {Event} e - The change event
 */
function onChange(e) {
  updateOrderStatsInSheet();
}

/**
 * onEdit trigger - Handles live updates
 * @param {Event} e - The edit event
 */
/**function onEdit(e) {
  liveUpdateTrigger(e);
}

/**
 * Hides or shows rows where status is "SHIPPED"
 * @param {string} state - 'ON' to hide, 'OFF' to show
 */
function toggleFocusMode(state) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  var hide = (state === 'ON');
  
  // Define the segments for both tables
  var segments = [
    { start: DATA_START_ROW, end: boundary - 2 },      // eBay Table
    { start: boundary + 2, end: sheet.getLastRow() }    // Direct Table
  ];

  segments.forEach(function(seg) {
    if (seg.end < seg.start) return;
    
    var range = sheet.getRange(seg.start, STATUS_COLUMN, seg.end - seg.start + 1, 1);
    var values = range.getValues();
    
    for (var i = 0; i < values.length; i++) {
      var status = String(values[i][0]).trim().toUpperCase();
      var currentRow = seg.start + i;
      
      if (hide && status === "SHIPPED") {
        sheet.hideRows(currentRow);
      } else {
        sheet.showRows(currentRow);
      }
    }
  });

  return hide ? "🌑 Focus Mode: ON (Shipped hidden)" : "🌕 Focus Mode: OFF (All rows visible)";
}