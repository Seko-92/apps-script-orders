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

  // 2. Arcade menu — a single door into the cabinet (Snake now lives INSIDE it
  //    alongside Tetris/Flappy/Breakout/Invaders/Pac-Man).
  ui.createMenu('🕹️ HQ ARCADE')
    .addItem('Open HQ Arcade', 'openHQArcade')
    .addToUi();

  // 2b. Floor Board — the warehouse monitor. Opens in-sheet for a quick look;
  //     the always-on version is the doGet web-app URL (open in a tablet tab).
  ui.createMenu('📺 Floor Board')
    .addItem('Open Floor Board', 'openFloorBoard')
    .addToUi();
  
  // ⭐⭐ 3. THE SIDEBAR OPENS BEFORE ANY BACKGROUND WORK, AND THAT ORDER IS THE FIX.
  //
  // ⚠⚠ 2026-08-30: with All Orders locked, an EMPLOYEE opening the sheet got no sidebar at
  //   all. onOpen runs as the USER who opened the file, step 4 below writes borders and
  //   number formats to the protected sheet, that throw killed the rest of the function —
  //   and showSidebar() used to sit after it. One refused cosmetic refresh took away the
  //   whole control panel.
  //
  //   Same shape as the 2026-05-01 bug where an unprotected handler in onEditInstallable
  //   killed every handler after it. The ruling that came out of that applies here too:
  //   NOTHING A USER DEPENDS ON MAY SIT DOWNSTREAM OF SOMETHING THAT CAN FAIL.
  try {
    showSidebar();
  } catch (err) {
    console.log("onOpen.showSidebar: " + err);
  }

  // 4. Background maintenance — every step isolated, and none of it load-bearing.
  //
  // ⚠ OWNER ONLY, deliberately. These WRITE to All Orders, which is protected, so for an
  //   employee they would either throw or take a pointless /exec round trip on every single
  //   sheet open. They are only keep-fresh passes: doPost re-runs the duplicate painter on
  //   every insert (OrderService.js), and the HAND rule is idempotent. Skipping them for
  //   staff costs nothing and removes a whole class of open-the-sheet failure.
  var maintainer = true;
  try { maintainer = (typeof _obIsOwner !== "function") || _obIsOwner(); } catch (e) {}

  try { updateOrderStatsInSheet(); } catch (err) { console.log("onOpen.stats: " + err); }
  try { toggleLiveUpdate('ON'); }   catch (err) { console.log("onOpen.liveSync: " + err); }

  if (maintainer) {
    try { setupHandConditionalFormatting(); }
    catch (err) { console.log("onOpen.handCF: " + err); }
    // SKU duplicate highlighting: manual only (use sidebar button to avoid visual clutter)
    try { setupDuplicateSalesOrderHighlighting(); }
    catch (err) { console.log("onOpen.dupSO: " + err); }
  }
}

/**
 * Launches the Snake Game Sidebar
 */
/* showSnakeSidebar() REMOVED 2026-08-19 along with Snake.html. Snake was
   absorbed into HQArcadeModal on 2026-07-29 and the cabinet has been the only
   door since; this served a file nothing else referenced. Git has both. */


/**
 * Opens the HQ Arcade cabinet — a break-room pop-up with Tetris / Flappy /
 * Breakout / Invaders (touch + keyboard, high scores on-device). Pure fun,
 * zero data. Called from the HQ ARCADE menu + the sidebar Arcade card.
 */
function openHQArcade() {
  var html = HtmlService.createHtmlOutputFromFile('HQArcadeModal')
    .setWidth(1100).setHeight(860);   // clamps to viewport; bigger playfield
  SpreadsheetApp.getUi().showModalDialog(html, '🕹️ HQ Arcade');
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
/* How long the badge repaint is suppressed after one runs. Long enough to
   collapse a sync's row-by-row storm (~5s apart), short enough that a human
   working in the sheet never waits — their FIRST change always paints. */
var ONCHANGE_PAINT_KEY     = 'hqOnChangePaint';
var ONCHANGE_PAINT_GAP_SEC = 60;

function onChangeInstallable(e) {
  // Only run on structural changes (row insert/delete), not on CF rule edits
  var changeType = e && e.changeType ? e.changeType : "";
  if (changeType === "REMOVE_ROW" || changeType === "INSERT_ROW") {
    /* ⚠⚠ DEBOUNCE THE REPAINT — A ROW STORM WAS EATING THE SCRIPT (2026-08-17).
       onChange fires for EVERY sheet in the spreadsheet, not just All Orders,
       and it fires per WRITE — so a sync appending rows to Master Inventory or
       Zoho Stock lands one event per append. Each one then repainted the whole
       All Orders badge band, which is expensive and, for a Master Inventory
       write, achieves precisely nothing.

       Caught in the Executions panel during a soak: onChangeInstallable firing
       every ~5 SECONDS for five minutes straight at 1.1–3.7s a run — roughly
       96 seconds of script time burned inside a 5-minute window, doing no work.
       That was the last remaining failure burst on the floor: reads and writes
       queued behind it, and the board's taps timed out. Outside that window the
       same soak was 132/132 clean.

       The first event still paints IMMEDIATELY, so a person inserting a row
       sees badges update at once — only the storm behind it is collapsed. And
       the painter is self-healing: it also runs from onOpen, every sort, and
       every insert path, so a suppressed repaint corrects itself shortly after.

       ⚠ THE FIX IS DELIBERATELY CAUSE-AGNOSTIC. I am inferring which sync
       produces the storm; debouncing works whatever the source is, which is the
       property worth having when the cause is a guess. */
    var _paint = true;
    try {
      var _c = CacheService.getScriptCache();
      if (_c.get(ONCHANGE_PAINT_KEY)) _paint = false;
      else _c.put(ONCHANGE_PAINT_KEY, '1', ONCHANGE_PAINT_GAP_SEC);
    } catch (err) { /* no cache — fall through and paint, as before */ }

    if (_paint) {
      try {
        setupDuplicateSalesOrderHighlighting();
      } catch (err) { /* silent */ }
    }

    // ⚠ A HUMAN DELETING OR INSERTING ROWS IS A CHANGE THE BOARD MUST SEE.
    // Reported 2026-08-14: rows deleted from the sheet stayed on the Floor
    // Board for minutes. Every dirty-flag chokepoint was one of OUR writes —
    // doPost, updateOrderStatus, boardSetStatus, boardSetLeft, boardAdjust —
    // and none of them is a person editing the sheet by hand. So a manual
    // change never invalidated the published tick and the board kept serving
    // the last copy until the 8-minute keep-fresh republish.
    //
    // It barely showed before 2026-08-13, because every poll fell through to
    // Apps Script live behind a 45s cache. Once n8n could actually READ the
    // published cell, this went from ~45s to up to 8 minutes — and the
    // dangerous direction is not a stale delete but a manually TYPED order
    // being invisible to the floor for that long.
    try { _dashBustTickCache(); } catch (err) { /* best-effort */ }
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

  // SKU enrichment — title-on-hover (cell note) + clickable listing link, looked
  // up live from Master Inventory by SKU when a SKU is typed into col A. Same
  // manual-entry coverage as location + ▣; programmatic inserts call
  // refreshSkuEnrichment() at their insert sites.
  try {
    skuEnrichmentOnEdit(e);
  } catch (err) {
    Logger.log("skuEnrichmentOnEdit (installable) error: " + err);
  }

  // Order link — SALES ORDER (col D) → clickable eBay/Zoho order. Same pattern
  // as the SKU link; programmatic inserts call refreshAllOrdersEnrichment().
  try {
    orderLinkOnEdit(e);
  } catch (err) {
    Logger.log("orderLinkOnEdit (installable) error: " + err);
  }

  // ⚠ ROW IDENTITY — the backstop for the 2026-08-28 slip, where a SALES_ORDER
  // cell was overwritten on a row already picked and shelf-counted.
  //
  // Placed AFTER the handlers that legitimately change a row, and BEFORE the
  // cache bust, for two reasons: it only READS and PAINTS (it can never fight
  // another handler's write), and running it late means the row it inspects is
  // the settled one rather than a half-applied edit.
  //
  // ⚠ It still matters after protectAllOrdersSheet() locks cols A/B/D: protection
  // does not restrict the OWNER, so this is what catches your own slips.
  try {
    identityEditGuard(e);
  } catch (err) {
    Logger.log("identityEditGuard (installable) error: " + err);
  }

  // ⚠ LAST, AND ON PURPOSE: tell the board the sheet moved under it.
  // Companion to the same call in onChangeInstallable — see the long note
  // there. That one covers row inserts/deletes; this covers a person typing
  // into All Orders (a manual order, a status, a qty, a location).
  //
  // Scoped to the MAIN sheet so editing Prep Queue, Kit Registry or the audit
  // sheets does not force a tick rebuild they cannot affect. Runs last because
  // the handlers above are what actually change the row — invalidating before
  // them could republish the pre-edit state and pin the stale copy for another
  // whole minute.
  try {
    var _sh = e && e.range && e.range.getSheet();
    if (_sh && _sh.getName() === MAIN_SHEET_NAME) _dashBustTickCache();
  } catch (err) { /* best-effort — never block an edit on the board's freshness */ }
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