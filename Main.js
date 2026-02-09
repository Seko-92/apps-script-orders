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
  setupHandConditionalFormatting();  // Ensure HAND highlight rule is active
  setupDuplicateHighlighting();     // Ensure duplicate SKU highlight rules are active

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
  ensureDirectTableBuffer();
}

/**
 * onEdit trigger (SIMPLE) - Handles local-only operations
 * Simple triggers CANNOT call external APIs (UrlFetchApp).
 * Telegram sync is handled by onEditInstallable() below.
 * @param {Event} e - The edit event
 */
function onEdit(e) {
  liveUpdateTrigger(e);
  // NOTE: handleManualStatusChange moved to installable trigger
  // because it calls UrlFetchApp (Telegram API) which requires
  // elevated permissions that simple triggers don't have.
}

/**
 * INSTALLABLE onEdit trigger - Handles operations that need external API access.
 * This function CAN call UrlFetchApp (Telegram API, etc.)
 *
 * To install: Run setupInstallableEditTrigger() once from the Script Editor.
 * @param {Event} e - The edit event
 */
function onEditInstallable(e) {
  console.log("onEditInstallable fired: " + (e && e.range ? e.range.getA1Notation() : "unknown"));
  handleManualStatusChange(e);
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

    // Batch consecutive rows for hide/show to minimize API calls
    var batchStart = -1;
    var batchIsShipped = false;

    for (var i = 0; i <= values.length; i++) {
      var isShipped = (i < values.length) && String(values[i][0]).trim().toUpperCase() === "SHIPPED";

      if (i === values.length || isShipped !== batchIsShipped) {
        // Flush previous batch
        if (batchStart >= 0) {
          var rowStart = seg.start + batchStart;
          var count = i - batchStart;
          if (hide && batchIsShipped) {
            sheet.hideRows(rowStart, count);
          } else if (!hide || !batchIsShipped) {
            sheet.showRows(rowStart, count);
          }
        }
        batchStart = i;
        batchIsShipped = isShipped;
      }
    }
  });

  return hide ? "🌑 Focus Mode: ON (Shipped hidden)" : "🌕 Focus Mode: OFF (All rows visible)";
}

// ═══════════════════════════════════════════════════════════════════════════════
// INSTALLABLE TRIGGER SETUP
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Run this ONCE from the Apps Script Editor to install the trigger.
 * Go to: Run > setupInstallableEditTrigger
 *
 * This creates an installable onEdit trigger that has full permissions
 * (UrlFetchApp, LockService, etc.) - required for Sheet→Telegram sync.
 */
function setupInstallableEditTrigger() {
  // Remove any existing installable onEdit triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEditInstallable' &&
        triggers[i].getEventType() === ScriptApp.EventType.ON_EDIT) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create the installable trigger
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  Logger.log("Installable onEdit trigger created for onEditInstallable()");

  // Show UI alert only if called from a UI context (menu click, onOpen, etc.)
  // When run from Script Editor's Run button, there is no UI context.
  try {
    SpreadsheetApp.getUi().alert(
      "Trigger Installed",
      "The installable onEdit trigger has been created. Sheet→Telegram sync is now active.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    // Running from Script Editor - no UI available. That's fine.
    Logger.log("Trigger installed successfully. (No UI context for alert)");
  }
}

/**
 * Run this to verify the installable trigger is active.
 * Check the Execution Log for the result.
 */
function verifyTriggerInstalled() {
  var triggers = ScriptApp.getProjectTriggers();
  var found = false;
  for (var i = 0; i < triggers.length; i++) {
    var handler = triggers[i].getHandlerFunction();
    var eventType = triggers[i].getEventType();
    console.log("Trigger found: " + handler + " (" + eventType + ")");
    if (handler === 'onEditInstallable') {
      found = true;
    }
  }
  if (found) {
    console.log("✅ onEditInstallable trigger is ACTIVE");
  } else {
    console.log("❌ onEditInstallable trigger NOT found. Run setupInstallableEditTrigger()");
  }
  return found;
}