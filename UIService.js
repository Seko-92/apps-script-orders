// =======================================================================================
// UI_SERVICE.gs - Sidebar and UI Functions
// =======================================================================================

/**
 * Shows the control panel sidebar
 */
function showSidebar() {
  var t = HtmlService.createTemplateFromFile('Sidebar');

  // ⭐ THE PUBLISHED-TICK ENDPOINT (2026-08-19). Injected rather than hardcoded in
  // the HTML because it lives in Secrets.js, which is gitignored — the panel must
  // still render for a fresh clone that has no Secrets, so an empty string here
  // simply means "use the live google.script.run path", which is the old behaviour.
  t.boardApiUrl = (typeof HQ_BOARD_API_URL === 'string') ? HQ_BOARD_API_URL : '';

  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle('⚙️ Control Panel'));
}

/**
 * The Floor Board's standalone URL (the web-app /exec that doGet serves the
 * board on). The sidebar's "📺 Floor Board" card fetches this so the operator
 * can open OR copy the link from inside the sheet — no more hunting for it to
 * send to the picker's tablet.
 */
function getFloorBoardUrl() {
  return (typeof WEB_APP_URL !== 'undefined') ? WEB_APP_URL : '';
}


/**
 * Where the read-only SCREENS live — the Floor Board and the wall monitor.
 *
 * ⚠ NOT WEB_APP_URL. Both moved onto the VPS on 2026-08-05 and are served by
 * Caddy from /opt/hq-app; the sidebar's link had been pointing at /exec ever
 * since, which still WORKS (doGet serves the board) but is the expensive path —
 * every open costs an Apps Script rebuild instead of reading the published cell.
 * Nobody noticed because a slow board looks exactly like a working one.
 *
 * ⚠ DERIVED FROM ONE CONSTANT, never written twice. HQ_BOARD_API_URL already
 * names the host; the two display URLs are that host plus a path. A second
 * hardcoded copy of the domain is how the 2026-05-18 URL-drift incident began.
 *
 * Falls back to WEB_APP_URL when Secrets.js has no host — a fresh clone must
 * still produce a working link rather than an empty one.
 *
 * @returns {{board:string, wall:string, hosted:boolean}}
 */
function getDisplayUrls() {
  var base = "";
  try {
    if (typeof HQ_BOARD_API_URL === "string" && HQ_BOARD_API_URL) {
      base = HQ_BOARD_API_URL.replace(/\/api\/board\/?$/, "");
    }
  } catch (e) { base = ""; }

  var legacy = (typeof WEB_APP_URL !== "undefined") ? WEB_APP_URL : "";
  return {
    board:  base ? base + "/"     : legacy,
    wall:   base ? base + "/wall" : "",
    hosted: !!base
  };
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
/* ⚠⚠ THE HEARTBEAT WAS STARVING THE WAREHOUSE (2026-08-17).
   This one function reads the Activity Log tail, All Orders, and — through
   getActionableAlerts — five more sheets, EVERY 30 SECONDS, PER OPEN SIDEBAR.
   Measured mid-shift in the Executions panel: 12–29 seconds a call, with three
   overlapping starts inside three seconds. Apps Script serialises executions of
   the same script, so at ~3 open Control Panels the heartbeat alone demanded
   more script time than wall-clock time existed, and every other caller queued
   behind it — runPublishTick at 48.9s, and doPost (a picker tapping ✓ Pick) at
   16–17s, which blew the Floor Board's 25s bound and reported "could not reach
   the sheet" to the floor. The pickers were tapping 3–4 times per item.

   THE COST WAS NEVER THE WORK — IT WAS THE MULTIPLIER. Every sidebar computed
   the same global answer independently: the cockpit, the alert counts, the API
   metrics and the sync time are identical for everyone looking at the sheet.
   So cache the RESULT in the SCRIPT cache (shared across users, deliberately
   not the user cache) and let the second, third and fourth sidebar read it for
   free. N open panels now cost ONE execution per window instead of N.

   ⚠⚠ THE TTL MUST BE LONGER THAN THE POLL, NOT SHORTER — I had this backwards
   on the first cut and it matters. At a 25s TTL against a 30s beat, the entry
   is ALWAYS expired by the time the next poll arrives, so a single open sidebar
   hits it exactly never and pays the full 12–29s rebuild every 30 seconds. The
   cache would then only help when a SECOND panel happened to land inside the
   window. At 60s a lone panel rebuilds once every two beats instead of every
   beat, which is where most of the saving actually comes from. (The board's own
   tick cache has always been 45s against a 20s poll — longer, for this reason.)

   Cost of the choice: badge counts and the cockpit can be up to ~60s behind.
   That is ambient monitoring on a panel that was already up to 30s stale by
   construction, and nothing in it gates an action.

   ⚠ This is a READ cache on a heartbeat, and every number in it is ambient. A
   picker acting on a number sees it through the ACTION's own return value, not
   through this. Nothing here gates a write. */
/* ⚠⚠ AND A PLAIN TTL CACHE STAMPEDES — my own first cut did (2026-08-17).
   Every panel polls on the same 30s beat, so when the entry expires they ALL
   miss in the same instant and ALL rebuild. Average load drops; the BURST does
   not, and the burst is what blocks the floor's writes. Measured after shipping
   the plain cache: clean stretches broken by ~30s stalls on a regular beat —
   the shape of three 13s rebuilds landing together.

   So: STALE-WHILE-REVALIDATE. Keep the payload for hours and treat FRESH as a
   separate, shorter window. Past FRESH, exactly one caller rebuilds and
   everyone else is handed the slightly-old copy IMMEDIATELY rather than queuing
   behind a rebuild they don't need.

   ⚠ THE MUTEX IS CACHE-BASED, NEVER LockService. getScriptLock() is the SAME
   lock updateOrderStatus takes, so holding it here for a 13s rebuild would
   block every ✓ Pick on the floor — that would be worse than the stampede it
   fixes. A cache flag is not atomic, so two rebuilds can still race through the
   gap; that is fine. The goal is turning N simultaneous rebuilds into ~1, not
   proving mutual exclusion. */
var SIDEBAR_TICK_CACHE_KEY = 'hqSidebarTick_v2';   // v2: shape gained _builtAt
var SIDEBAR_TICK_BUILD_KEY = 'hqSidebarTick_building';
/* ⚠ THIS ONE NUMBER IS THE SIDEBAR'S WHOLE COST. Raised 120 -> 300 on
   2026-08-19 after measuring where the daily quota actually goes.

   The panel asks the server for numbers every 30 SECONDS. The server keeps a
   copy of the answer for FRESH_SEC and hands it back for free until it expires;
   the first ask AFTER it expires pays for a full rebuild — about 1.5-2 SECONDS
   of billed execution time (the Executions panel's Duration column, which is
   what Google charges, NOT the ~40ms the function body measures).

       at 120s:  1 ask in 4  rebuilds   ~720 rebuilds/day   ~21   min/day
       at 300s:  1 ask in 10 rebuilds   ~288 rebuilds/day   ~8.6  min/day

   ⚠ IT MUST EXCEED THE 30s POLL, AND COMFORTABLY. Set equal to the poll and
     every single ask arrives just as the copy dies, so nothing is ever reused —
     the same trap the board's own tick cache was built to avoid (45s TTL against
     a 20s poll, for exactly this reason).

   ⚠ AND THE COST IS NOT BOUNDED BY THE TEAM. The sheet is shared by public link
     and the panel auto-opens, so this is paid on behalf of everyone who has it
     open. Proven per-script: with ONE tab open, the Executions panel showed
     getSidebarTick arriving every ~15s against a 30s poll.

   WHAT IT COSTS: the cockpit figures (shipped today, oldest pending, queue
   counts) can be up to 5 minutes old instead of 2. They are ambient awareness,
   not action triggers, and the alerts half already rides its own 300s clock.
   The real fix is to publish this tick to a cell and let n8n serve it, the way
   the Floor Board is served — then viewers cost nothing at all. */
var SIDEBAR_TICK_FRESH_SEC = 300;         // how long a copy counts as current
var SIDEBAR_TICK_KEEP_SEC  = 21600;       // keep as a fallback (6h = cache max)
var SIDEBAR_TICK_BUILD_SEC = 45;          // rebuild-in-progress flag lifetime
var SIDEBAR_TICK_MAX_BYTES = 90000;       // CacheService hard-caps a value at 100KB

/* ⭐ THE EXPENSIVE HALF, ON ITS OWN CLOCK (2026-08-18).
   getActionableAlerts opens FIVE-PLUS SHEETS — Out of Stock, Prep Queue, the
   Price Audit snapshot, Kit Health, Investigations — to paint badge counts.
   getLatestApiMetrics reads the API Usage sheet. Together they are the bulk of
   the 12–29s this heartbeat was costing, and they are answering questions whose
   answers barely move: an OOS count or a photo backlog changes over hours, not
   over seconds.

   The cockpit is different — shipped/received today, oldest pending, the queue
   split — that is the shift's live picture and belongs on the fast beat.

   So: split the cadences. The cheap half refreshes every ~2 min, the expensive
   half every ~5. Nothing here gates an action; the sidebar's own buttons call
   their individual refreshers, which never read this cache.

   ⚠ Same stale-while-revalidate shape as the tick above, and for the same
   reason — a plain TTL would stampede when several panels expire together. */
var SIDEBAR_SLOW_CACHE_KEY = 'hqSidebarSlow_v1';
/* ⚠ 900s, NOT 300s (2026-08-19). MEASURED: with the fast half also at 300s the
   split stopped buying anything — two probe windows both showed tick rebuilds and
   slow rebuilds at IDENTICAL counts (3 and 3, then 3 and 3), i.e. the eight-sheet
   read was riding along with every single cockpit rebuild. The whole point of the
   08-18 split was that these two halves answer questions on different clocks.
   ⚠ ACCEPTED TRADE: paidShipping and newFromZoho can now lag up to 15 minutes.
   That is defensible because the Alerts card is a BACKLOG view, not an arrival
   notifier — arrivals are announced by the Floor Board beacon and by Telegram,
   both of which are unaffected. An OOS count or a photo backlog moves over hours. */
var SIDEBAR_SLOW_FRESH_SEC = 900;        // 15 min
var SIDEBAR_SLOW_KEEP_SEC  = 21600;

function _sidebarSlowParts() {
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) { cache = null; }

  var stale = null;
  if (cache) {
    try {
      var hit = cache.get(SIDEBAR_SLOW_CACHE_KEY);
      if (hit) {
        var c = JSON.parse(hit);
        if (Date.now() - (c._builtAt || 0) < SIDEBAR_SLOW_FRESH_SEC * 1000) return c;
        stale = c;                       // past FRESH but still worth serving
      }
    } catch (e) { /* unreadable is a miss */ }
  }

  var out = { api: null, alerts: null, _builtAt: Date.now() };
  try { out.api    = getLatestApiMetrics(); } catch (e) { console.error('sidebarSlow.api: '    + e); }
  try { out.alerts = getActionableAlerts(); } catch (e) { console.error('sidebarSlow.alerts: ' + e); }

  // ⚠ If BOTH reads failed, keep serving the last good copy rather than blanking
  // the badges — a transient sheet error should not empty the Alerts card.
  if (out.api === null && out.alerts === null && stale) return stale;

  if (cache) {
    try {
      var json = JSON.stringify(out);
      if (json.length <= SIDEBAR_TICK_MAX_BYTES) {
        cache.put(SIDEBAR_SLOW_CACHE_KEY, json, SIDEBAR_SLOW_KEEP_SEC);
      }
    } catch (e) { console.error('sidebarSlow.cache: ' + e); }
  }
  return out;
}

/**
 * One round-trip for the sidebar heartbeat.
 * @param {boolean} [force] - skip the cache (used after an action changes state)
 */
function getSidebarTick(force) {
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) { cache = null; }

  var stale = null;
  if (cache && !force) {
    try {
      var hit = cache.get(SIDEBAR_TICK_CACHE_KEY);
      if (hit) {
        var cached = JSON.parse(hit);
        var ageMs  = Date.now() - (cached._builtAt || 0);
        if (ageMs < SIDEBAR_TICK_FRESH_SEC * 1000) {
          cached._cached = true;          // provenance, same habit as the board tick
          return cached;
        }
        stale = cached;                   // past FRESH, but still perfectly usable
      }
    } catch (e) { /* unreadable cache is a miss, never a failure */ }
  }

  /* Past FRESH with a copy in hand: let ONE caller rebuild and hand everyone
     else the old copy on the spot. A panel showing numbers a couple of minutes
     old costs nothing; a panel queueing behind a rebuild costs the floor its
     writes. */
  if (cache && stale) {
    try {
      if (cache.get(SIDEBAR_TICK_BUILD_KEY)) {
        stale._cached = true;
        stale._stale  = true;             // says WHY it is old — someone is rebuilding
        return stale;
      }
      cache.put(SIDEBAR_TICK_BUILD_KEY, '1', SIDEBAR_TICK_BUILD_SEC);
    } catch (e) { /* no mutex available — rebuild rather than serve nothing */ }
  }

  var result = { cockpit: null, lastSync: '', api: null, alerts: null, picker: '' };
  try { result.cockpit  = getDashboardSnapshot(); } catch (e) { console.error('getSidebarTick.cockpit: '  + e); }
  try { result.lastSync = getLastSyncFromSheet(); } catch (e) { console.error('getSidebarTick.lastSync: ' + e); }
  // ⭐ THE SLOW HALF RUNS ON ITS OWN, SLOWER CLOCK (2026-08-18).
  var slow = _sidebarSlowParts();
  result.api    = slow.api;
  result.alerts = slow.alerts;
  try { result.picker   = getCurrentPicker();     } catch (e) { console.error('getSidebarTick.picker: '   + e); }

  if (cache) {
    try {
      result._builtAt = Date.now();
      var json = JSON.stringify(result);
      // Over the cap, put() throws and would take the whole tick down with it.
      // Skipping the cache costs speed; throwing costs the panel.
      if (json.length <= SIDEBAR_TICK_MAX_BYTES) {
        cache.put(SIDEBAR_TICK_CACHE_KEY, json, SIDEBAR_TICK_KEEP_SEC);
      } else {
        console.log('getSidebarTick: payload ' + json.length + 'B — too big to cache');
      }
    } catch (e) { console.error('getSidebarTick.cache: ' + e); }
    // ⚠ Clear the flag LAST and unconditionally. Leaving it set would make every
    // panel serve stale until it aged out on its own.
    try { cache.remove(SIDEBAR_TICK_BUILD_KEY); } catch (e) {}
  }
  return result;
}

/* ⚠ DELIBERATELY NOT BUSTED ON WRITES. Every other cache here is invalidated at
   the write chokepoints, and doing that would be wrong for this one: picks
   happen constantly on a busy shift, so busting on write would invalidate the
   entry almost every time and hand back the exact N-executions-per-window
   problem this exists to remove. Nothing in this payload gates an action — it
   is badge counts and a cockpit behind a 30s heartbeat, already up to 30s stale
   by construction, so 25s more is inside its existing tolerance. A sidebar
   button that changes state calls its own individual refresher, which does not
   read this cache and is therefore always fresh. */

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
