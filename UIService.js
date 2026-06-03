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
 * Web app GET handler — serves the warehouse dashboard at the deployed URL.
 *
 * Opens in any browser at the Apps Script web app URL. Auto-refreshes every
 * 30s via the same getSidebarTick() function the sidebar heartbeat uses
 * (single source of truth for both surfaces). Brand-styled fullscreen layout
 * for TV / kiosk display.
 *
 * Deployment: this MUST be a SEPARATE deployment from the n8n doPost endpoint
 * (different access settings — dashboard needs "Anyone with Google account"
 * for floor-display kiosk login; doPost stays "Anyone" for n8n webhook usage).
 * See Gotcha #12 for the broader deployment-discipline rule.
 *
 * Access pattern for the floor display:
 *   1. Deploy this project as a web app, "Execute as me", "Anyone with Google account"
 *   2. Copy the resulting URL
 *   3. On the kiosk device (Fire TV / Pi / old laptop), sign into a dedicated
 *      Google account that has VIEW access to the spreadsheet
 *   4. Open the URL in a browser, F11 fullscreen, leave it running
 *
 * @returns {HtmlOutput} the rendered Dashboard.html
 */
function doGet(e) {
  // The Floor Board (FloorBoard.html) is the single warehouse monitor. The old
  // multi-feature showpiece (Dashboard.html) was retired 2026-06-03.
  return HtmlService.createTemplateFromFile('FloorBoard')
    .evaluate()
    .setTitle('HQ Motor Service · Floor Board')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * RETIRED 2026-04-29.
 *
 * This used to write a wall clock to E1 and a date string to B1 every minute
 * via a time-driven trigger. Both writes now conflict with newer logic:
 *
 *   - E1 is now the LAST-SYNC TIMESTAMP, written by updateLastSyncTimestamp()
 *     in OrderService.js after every successful n8n insert. A wall clock
 *     overwriting it every minute would erase the signal it carries.
 *
 *   - B1 holds the auto-updating date formula installed by _ensureDateFormula()
 *     in BrandTheme.js: =TEXT(TODAY(), "dddd, mmmm d, yyyy"). Writing a
 *     static string would overwrite the formula and freeze the date.
 *
 * The function is kept as a no-op so any pre-existing time-driven trigger
 * stays harmless. SAFE TO DELETE the trigger at your convenience:
 *   Apps Script editor → Triggers → find the row for "updateSheetClock" → trash icon.
 */
function updateSheetClock() {
  return "updateSheetClock is retired — see UIService.js header comment. " +
         "Last-sync timestamp displays in " + Schema.cellSyncTime +
         " (updated by updateLastSyncTimestamp). Date formula owns B1.";
}

/**
 * Consolidated sidebar heartbeat. One server call covers the data five
 * separate polls used to fetch — cockpit snapshot, last-sync banner cell,
 * API quotas, actionable alerts, current picker. Cuts the 30s tick from
 * 5 round-trips to 1.
 *
 * Each piece is wrapped in its own try/catch so a single slow/failing
 * source can't black out the rest of the tick. Failed pieces come back
 * null/empty; the client keeps its last-known display for those.
 *
 * Returned shape mirrors the original individual return values so client-
 * side paint helpers can be reused unchanged. Individual functions
 * (getDashboardSnapshot, getLatestApiMetrics, etc.) stay exported for
 * on-demand callers — manual refresh buttons, post-action re-polls.
 */
function getSidebarTick() {
  var result = { cockpit: null, lastSync: '', api: null, alerts: null, picker: '' };
  try { result.cockpit  = getDashboardSnapshot(); } catch (e) { console.error('getSidebarTick.cockpit: '  + e); }
  try { result.lastSync = getLastSyncFromSheet(); } catch (e) { console.error('getSidebarTick.lastSync: ' + e); }
  try { result.api      = getLatestApiMetrics();  } catch (e) { console.error('getSidebarTick.api: '      + e); }
  try { result.alerts   = getActionableAlerts();  } catch (e) { console.error('getSidebarTick.alerts: '   + e); }
  try { result.picker   = getCurrentPicker();     } catch (e) { console.error('getSidebarTick.picker: '   + e); }
  return result;
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
