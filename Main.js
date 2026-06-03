// =======================================================================================
// MAIN.gs - Entry Points and Triggers/
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

  // 2b. Floor Board — the warehouse monitor. Opens in-sheet for a quick look;
  //     the always-on version is the doGet web-app URL (open in a tablet tab).
  ui.createMenu('📺 Floor Board')
    .addItem('Open Floor Board', 'openFloorBoard')
    .addToUi();
  
  // 3. Run your existing background logic
  updateOrderStatsInSheet();
  toggleLiveUpdate('ON');  // Auto-enable live sync
  setupHandConditionalFormatting();  // Ensure HAND highlight rule is active
  // SKU duplicate highlighting: manual only (use sidebar button to avoid visual clutter)
  setupDuplicateSalesOrderHighlighting(); // Ensure duplicate Sales Order highlight rules are active

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
 * INSTALLABLE onChange trigger - Handles row deletions, paste, structural changes.
 * NOTE: Do NOT call setupDuplicateHighlighting here — modifying CF rules
 * triggers another onChange, causing an infinite loop.
 */
function onChangeInstallable(e) {
  // Only run on structural changes (row insert/delete), not on CF rule edits
  var changeType = e && e.changeType ? e.changeType : "";
  if (changeType === "REMOVE_ROW" || changeType === "INSERT_ROW") {
    try {
      setupDuplicateSalesOrderHighlighting();
    } catch (err) { /* silent */ }
  }
}

/**
 * onEdit trigger (SIMPLE) - Handles local-only operations
 * Simple triggers CANNOT call external APIs (UrlFetchApp).
 * Telegram sync is handled by onEditInstallable() below.
 * @param {Event} e - The edit event
 */
function onEdit(e) {
  // liveUpdateTrigger uses openById which can fail in simple triggers;
  // wrap in try-catch so it doesn't block other handlers.
  // It also runs via installable trigger, so this is just a fallback.
  try {
    liveUpdateTrigger(e);
  } catch (err) {
    // Expected in simple trigger context — installable trigger handles it
  }

  // NOTE: handleManualStatusChange, prepQueueOnEdit, outOfStockOnEdit,
  // locationUpdateOnEdit, manualReceiveOnEdit, and noteEditOnEdit all run
  // in the INSTALLABLE trigger because they need full permissions (openById,
  // UrlFetchApp). Simple triggers can fail silently for those calls — which
  // was the root cause of the "Location Update sometimes fails to fill"
  // issue (the old simple-trigger locationUpdateTimestamp is now orphaned).
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

  // EVERY handler is wrapped — an exception in one must NEVER block the
  // others. Bug history: handleManualStatusChange / refreshDuplicateHighlightsOnEdit
  // were unprotected, and a throw inside refreshDuplicateHighlightsOnEdit on a
  // SALES_ORDER edit silently killed manualReceiveOnEdit so direct manual orders
  // weren't being logged. Defense-in-depth: each handler is its own try block.
  try { handleManualStatusChange(e); }
  catch (err) { Logger.log("handleManualStatusChange (installable) error: " + err); }

  try { refreshDuplicateHighlightsOnEdit(e); }
  catch (err) { Logger.log("refreshDuplicateHighlightsOnEdit (installable) error: " + err); }

  // Prep Queue SKU lookup (auto-fill LOCATION + HAND + DATE ADDED) needs
  // full permissions because it calls openById through getSingleLocation /
  // getSingleInventory / getCommittedQuantities. Defensive try/catch so any
  // error stays contained and doesn't block the other handlers above.
  try {
    prepQueueOnEdit(e);
  } catch (err) {
    Logger.log("prepQueueOnEdit (installable) error: " + err);
  }

  // Out of Stock SKU lookup (auto-fill LOCATION + QTY + SOLD + AVAILABLE +
  // FIRST SEEN + LAST CHECKED). Same Master-Inventory-via-openById pattern,
  // same containment.
  try {
    outOfStockOnEdit(e);
  } catch (err) {
    Logger.log("outOfStockOnEdit (installable) error: " + err);
  }

  // Location Update SKU lookup (auto-fill COUNTER + LOCATION + TIMESTAMP).
  // Same pattern as Prep Queue / Out of Stock — runs in INSTALLABLE because
  // location lookup goes through openById. Replaces the orphaned simple-trigger
  // locationUpdateTimestamp in Timestampfeature.js (which was the root cause
  // of "sometimes location/timestamp fails to appear").
  try {
    locationUpdateOnEdit(e);
  } catch (err) {
    Logger.log("locationUpdateOnEdit (installable) error: " + err);
  }

  // Manual sales-order entry (eBay or DIRECT) → log a RECEIVED event so the
  // Activity Log captures manual orders (not just n8n-pushed eBay ones).
  try {
    manualReceiveOnEdit(e);
  } catch (err) {
    Logger.log("manualReceiveOnEdit (installable) error: " + err);
  }

  // Note-column edit → log a NOTE event so the audit trail captures every
  // supervisor/picker remark added or changed mid-prep, not just the original
  // buyer note that arrived with the order.
  try {
    noteEditOnEdit(e);
  } catch (err) {
    Logger.log("noteEditOnEdit (installable) error: " + err);
  }

  // Kit SKU marker (▣ glyph prefix) — applies/clears the per-cell number-format
  // marker when a SKU is typed into col A of the All Orders sheet. Covers
  // manual-entry cases (picker types a kit SKU directly into a row); n8n /
  // Zoho insert paths call refreshKitSkuMarkers() explicitly since programmatic
  // setValues doesn't fire onEdit.
  try {
    kitSkuOnEdit(e);
  } catch (err) {
    Logger.log("kitSkuOnEdit (installable) error: " + err);
  }
}

/**
 * Hides or shows rows where status is "SHIPPED"
 * @param {string} state - 'ON' to hide, 'OFF' to show
 */
function toggleFocusMode(state) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  var hide = (state === 'ON');
  
  // Define the segments for both tables
  var segments = [
    { start: Schema.dataStartRow, end: boundary - 2 },      // eBay Table
    { start: boundary + 2, end: sheet.getLastRow() }    // Direct Table
  ];

  segments.forEach(function(seg) {
    if (seg.end < seg.start) return;

    var range = sheet.getRange(seg.start, Schema.cols.STATUS, seg.end - seg.start + 1, 1);
    var values = range.getValues();

    // Batch consecutive rows for hide/show to minimize API calls
    var batchStart = -1;
    var batchIsShipped = false;

    for (var i = 0; i <= values.length; i++) {
      var isShipped = (i < values.length) && String(values[i][0]).trim().toUpperCase() === Schema.status.SHIPPED;

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
  var triggers = ScriptApp.getProjectTriggers();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Remove existing installable onEdit and onChange triggers to avoid duplicates
  for (var i = 0; i < triggers.length; i++) {
    var handler = triggers[i].getHandlerFunction();
    if (handler === 'onEditInstallable' || handler === 'onChangeInstallable') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create installable onEdit trigger (Telegram sync + duplicate highlight refresh)
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Create installable onChange trigger (row deletions, paste, structural changes)
  ScriptApp.newTrigger('onChangeInstallable')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  Logger.log("Installable triggers created: onEditInstallable + onChangeInstallable");

  try {
    SpreadsheetApp.getUi().alert(
      "Triggers Installed",
      "Installable onEdit + onChange triggers created. Duplicate highlights will now auto-refresh on edits AND row deletions.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log("Triggers installed successfully. (No UI context for alert)");
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