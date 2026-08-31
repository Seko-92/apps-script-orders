// =======================================================================================
// DashboardService.js — server endpoints for the Floor Board (FloorBoard.html)
// =======================================================================================
//
// The Floor Board is served via doGet() in UIService.js (browser tab) and via
// openFloorBoard() (in-sheet modal). It polls getDashboardTick() every ~15s for a
// single payload that wraps the sidebar tick (cockpit + alerts + picker + lastSync)
// with two Floor-Board signals:
//
//   - paceCar     — current shipping velocity + linear projection to 5pm
//                   ("on pace for N by 5pm")
//   - openOrders  — every OPEN (PENDING/PREPARING) row across both tables, sorted
//                   by aisle, for the "To pick" panel (SKU · qty · location)
//
// Both extras are best-effort: any throw → empty default, no error propagates to
// the client. The board treats missing fields as "no data yet" and keeps painting
// from last-known values.
//
// 2026-06-03: the old multi-feature showpiece (Dashboard.html) was retired — the
// Floor Board is the single board. The showpiece-only tick signals (hourlyBuckets,
// recentPrints, todayEvents) and their helpers were removed with it.
// =======================================================================================


// ---------- DASHBOARD CONSTANTS ----------
var DASHBOARD = {
  sunriseHour: 7,    // start of the workday (pace baseline)
  sunsetHour:  17,   // end of the workday — 5pm (pace projection target)

  // Pick-list cap. Applied AFTER the aisle sort (see _dashOpenOrders) so what
  // gets dropped is the far end of the walk rather than an arbitrary slice of
  // the sheet. The board shows "+N more" whenever it bites.
  pickListCap: 60,

  // Kit-SKU lookup cache — buildKitMap reads ~1,500 rows and this is on the
  // 15-second poll. Kit composition changes rarely; minutes of staleness is fine.
  kitCacheKey: 'dashKitSkus',
  kitCacheSec: 300,
  // KIT THREADS = 2 hues × 2 patterns (solid / dashed) = 4 distinguishable.
  //
  // ⚠ TWO hues, not three, and this was MEASURED not chosen. An exhaustive
  // search of every colour that clears 4.5:1 on the card AND stays clear of
  // yellow / red / green / amber found: 2 hues separate by ΔE 57.8 in the
  // worst case across normal + protanopia + deuteranopia + tritanopia; THREE
  // collapse to ΔE 12.5, i.e. indistinguishable. Under deuteranopia the space
  // folds onto a blue↔yellow axis and yellow already means "do something"
  // here, so the third hue has nowhere to live.
  //
  // Hence PATTERN as the second channel — it is immune to colour vision
  // deficiency entirely, and it is what makes four concurrent kits safe.
  // Past four they cycle; the strip names each one.
  kitHues:     2,
  kitPatterns: 2,

  // WHOLE-TICK cache. Measured 2026-08-05: a cold tick took ~26s, and every
  // device polling pays that independently. Apps Script gives a consumer
  // account ~90 min of runtime a DAY, so uncached this cannot support even one
  // board left open, let alone ten. A cache hit costs ~50ms instead.
  //
  // Lives HERE rather than only in n8n so the protection applies to every
  // caller — the hosted board, the in-sheet modal, anything later.
  // TTL must be COMFORTABLY LONGER than the board's poll interval (20s). Set
  // equal, every poll arrives exactly as the cache expires and a lone board
  // misses almost every time — paying the full build. At 45s roughly two polls
  // in three are served from cache, so the real cost is ~1 build per minute
  // instead of 3.
  //
  // Staleness is not a concern in practice: a WRITE busts the cache immediately
  // (see _dashBustTickCache), so ✓ Pick still feels instant. Only inbound
  // changes — a new order landing — can lag, and the Telegram ping is the
  // real-time alert for those. The board is an ambient display.
  tickCacheKey: 'dashTick',
  tickCacheSec: 45,

  // ⚠ THE COUNT PROMPT ONLY APPEARS ON A SHELF SOMEONE WOULD ACTUALLY COUNT.
  // The floor's own rule (2026-08-11): a deviance is checked on low-quantity
  // items, because counting a hundred pistons mid-pick is not a thing anyone
  // does. Above this the board stays silent rather than asking for a number
  // nobody is going to produce — the same mark-the-exception discipline that
  // took GRAB off all eleven rows.
  // 25 matches what the floor described; the sidebar's low-stock badge uses
  // 20, so they are deliberately close but not coupled.
  countMaxHand: 25
};


// =======================================================================================
// PUBLIC: opener
// =======================================================================================

/**
 * Open the Floor Board in an in-sheet modal — the calm, glanceable warehouse
 * monitor (orders-to-grab + a by-aisle pick list + paid-shipping + a live
 * event feed + pace). Reuses getDashboardTick(). The always-on browser-tab
 * version is served via doGet() in UIService.js. This is now the ONE board —
 * the old multi-feature showpiece (Dashboard.html) was retired 2026-06-03.
 */
function openFloorBoard() {
  var html = HtmlService.createTemplateFromFile('FloorBoard')
    .evaluate()
    .setWidth(1400)
    .setHeight(820)
    .setTitle('HQ · Floor Board');
  SpreadsheetApp.getUi().showModalDialog(html, 'HQ Motor Service · Floor Board');
}


// =======================================================================================
// PUBLIC: dashboard tick
// =======================================================================================

/**
 * Single consolidated poll for the dashboard. Wraps getSidebarTick (cockpit,
 * alerts, api, picker, lastSync) and adds dashboard-only signals.
 *
 * Each extra is wrapped so one failure can't black out the rest of the tick.
 * Client treats undefined fields as "skip this paint pass" and keeps showing
 * last-known values — same contract as the sidebar.
 */
function getDashboardTick() {
  var cache = null;
  try {
    cache = CacheService.getScriptCache();
    var hit = cache.get(DASHBOARD.tickCacheKey);
    if (hit) {
      var cached = JSON.parse(hit);
      cached._cached = true;
      return cached;
    }
  } catch (e) { /* cache unavailable — just build it */ }

  var tick = _buildDashboardTick();

  try {
    if (cache) {
      var s = JSON.stringify(tick);
      // CacheService caps a value at 100KB. A busy tick (60 pick rows + the
      // day's timeline) runs ~20KB, but never let an oversize payload throw.
      if (s.length < 95000) cache.put(DASHBOARD.tickCacheKey, s, DASHBOARD.tickCacheSec);
    }
  } catch (e) { /* not cacheable this time — harmless */ }

  return tick;
}


/**
 * Drop the cached tick so the next caller rebuilds.
 *
 * Called after a board write: without it, ✓ Pick would be followed by up to
 * `tickCacheSec` of the board insisting the row is still PENDING.
 */
function _dashBustTickCache() {
  try { CacheService.getScriptCache().remove(DASHBOARD.tickCacheKey); }
  catch (e) { /* nothing to do */ }

  // Same moment, second job: mark the PUBLISHED tick stale so the next trigger
  // run rewrites it. Hooked HERE rather than at each call site because every
  // write chokepoint already funnels through this function — updateOrderStatus,
  // the doPost insert, and boardSetStatus — so one edit covers every path that
  // exists now or later. Cheap (one property write) and never fatal: a missed
  // flag costs one late publish, not correctness.
  try { if (typeof _pubMarkDirty === 'function') _pubMarkDirty(); }
  catch (e) { /* publishing is best-effort — never block a write on it */ }
}


/** The real work — everything getDashboardTick returns when the cache misses. */
function _buildDashboardTick() {
  // Deliberately NOT getSidebarTick(): that also fetches getLatestApiMetrics(),
  // a whole sheet read the board never displays (zero references to tick.api in
  // FloorBoard.html). The sidebar still uses getSidebarTick unchanged.
  //
  // getActionableAlerts() is ALSO skipped, and that is the bigger saving: it
  // opens EIGHT sheets (All Orders, Prep Queue, Out of Stock, Pending Sales
  // Orders, Price Audit, Kit Health, Investigations, Photo Queue) to build the
  // sidebar's Alerts card — and the board renders exactly ONE number out of it,
  // paidShipping.count, for the amber strip. That count is now produced by
  // _dashOpenOrders inside a row scan that was happening anyway, so seven sheet
  // reads per tick disappear.
  var base = { cockpit: null, lastSync: '', picker: '', pickers: [] };
  try { base.cockpit  = getDashboardSnapshot(); } catch (e) { console.error('tick.cockpit: '  + e); }
  try { base.lastSync = getLastSyncFromSheet(); } catch (e) { console.error('tick.lastSync: ' + e); }
  try { base.picker   = getCurrentPicker();     } catch (e) { console.error('tick.picker: '   + e); }
  // ⚠ THE PICK ID LIST RIDES ALONG (2026-08-14). Tapping the footer picker chip
  // used to fire boardPickers — a 3-6s round trip, of which ~3s is just the
  // fixed cost of reaching Apps Script through n8n, to fetch four strings that
  // change roughly never. Carried on the tick instead, the drawer opens with
  // ZERO server calls. It is a data-validation read on a cell we already touch
  // for `picker`, and the tick is now built about once a minute rather than
  // once per poll, so it is nearly free. boardPickers stays as the fallback for
  // a board whose tick predates this.
  try { base.pickers  = (getBoardPickers() || {}).pickers || []; }
  catch (e) { console.error('tick.pickers: ' + e); }

  var pace = null;
  var openOrders = [];

  try { pace = _dashPaceCarStats(base.cockpit); }
  catch (e) { console.error('getDashboardTick.pace: ' + e); }

  try { openOrders = _dashOpenOrders(); }
  catch (e) { console.error('getDashboardTick.openOrders: ' + e); }

  // Floor notes are read client-side from each open order's NOTE (the "**"
  // marker) — no server store needed; openOrders already carries the note text.
  return {
    cockpit:    base.cockpit  || {},
    // Shaped exactly like getActionableAlerts' paidShipping entry so the board's
    // paintPaid() reads it unchanged. `rows` stays empty — the board never
    // jumps to rows (that's a sidebar affordance), it only shows the count.
    alerts:     { paidShipping: { count: (openOrders && openOrders.paidCount) || 0, rows: [] } },
    api:        null,          // board never renders API quota; key kept for shape

    picker:     base.picker   || '',
    pickers:    base.pickers  || [],   // see the note in _buildDashboardTick
    lastSync:   base.lastSync || '',
    paceCar:    pace,
    openOrders: openOrders,
    // How many were open BEFORE the cap. The array carries `.total` as an own
    // property, but that does NOT survive JSON serialisation of an array — so
    // it has to be lifted onto the tick explicitly or the board can never tell
    // it's showing a truncated walk.
    openOrdersTotal: (openOrders && openOrders.total) || (openOrders || []).length,
    // Per-channel totals before the cap, so each SECTION can report its own
    // "+N more" instead of one pooled number that hides which half was cut.
    // Same array-property lifting problem as `total` above.
    openOrdersBy: (openOrders && openOrders.byChannel) || { EBAY: 0, DIRECT: 0 },
    // Kits in progress — same lifting problem as `total` above.
    kits:       (openOrders && openOrders.kits) || [],
    // Held orders, open AND shipped-but-not-yet-collected. Same array-property
    // lifting problem as `total`. Normally empty — the board's alert strip is
    // drawn ONLY when this has entries, so a quiet day costs nothing and any
    // ink on that strip is a finding.
    held:       (openOrders && openOrders.held) || [],
    serverTime: new Date().toISOString()
  };
}


/**
 * The kit this row is a COMPONENT of, or "" if it is not one.
 *
 * ⚠ MATCHES BOTH TAG SHAPES. Expansion writes `↳ from KIT-<sku>` for registered
 * components and `↳ added to KIT-<sku>` for custom adds. This function used to
 * be an inline `indexOf("↳ from KIT-")`, which saw only the first — so a custom
 * -added part was counted as a loose line: no kit thread, and missing from the
 * `done/total` denominator. That last one is the dangerous half — the counter
 * exists to stop a half-packed box shipping, and a kit reading "5 of 5" with a
 * sixth custom part still on its shelf is exactly the failure it was built to
 * prevent. The rest of the codebase (_findKitComponentRows,
 * _countExistingKitComponents) already matched both; this was the outlier.
 *
 * ⚠ SURVIVES A ZOHO FLAG. _flagDirectRow PREPENDS its warning as its own first
 * LINE, and it CASCADES onto a removed kit's components — so a flagged
 * component's note starts with "⚠️" and the kit tag sits on line 2. Matching the
 * raw string would drop those rows out of their kit entirely, which (because
 * `expandedOf` would fall to zero) would even un-collapse the parent and put a
 * box that no longer exists back on the pick list.
 */
function _dashKitTag(note) {
  // ⚠ DELEGATES — the parser itself lives in Helpers.kitComponentTag, because
  // this tag is read in at least seven places and a second copy is how a fix
  // starts having to land twice (the KitRegistry parser lesson, 2026-07-14).
  return kitComponentTag(note);
}


// =======================================================================================
// THE ATTRIBUTION CHOKEPOINT for board writes
// =======================================================================================

/**
 * Refuse a warehouse-side board write when no Pick ID is set.
 *
 * ⚠ THE GAP THIS CLOSES, reported from the floor 2026-08-14: printing has
 * refused without a Pick ID since 2026-05-01 and stock adjustment since
 * 2026-08-11 — but ✓ Pick and the shelf count never asked, and ✓ Pick is by far
 * the most frequent thing the floor does. So a shift that happened not to print
 * produced an entire day of Activity Log rows with a blank PICKER, and nothing
 * on the board did more than tint a chip amber.
 *
 * ⚠ THIS IS NOT AN OCCASIONAL STATE — IT IS HOW EVERY MORNING STARTS.
 * `resetDailyPickIds` blanks F2 back to the dropdown's placeholder at 4am ON
 * PURPOSE, so yesterday's picker cannot roll forward onto today's work. The
 * first action of a shift is a pick, never a print — so the reset was
 * guaranteeing the very gap it exists to prevent. Printing was the chokepoint
 * on paper; picking is the chokepoint in practice.
 *
 * ⚠ RETURNS `needsPicker` SO THE CLIENT CAN OFFER THE DOOR. The board opens its
 * own Pick ID list (which now rides on the tick, so it costs nothing) and then
 * REPLAYS the action, so the picker never loses the tap they made. A refusal
 * that only says "go set it somewhere else" is the mistake the footer chip
 * already taught us: a capability you can only reach by failing first is not a
 * capability the floor has.
 *
 * ⚠ WHAT THIS DOES **NOT** DO — stated plainly. F2 is ONE cell for the whole
 * sheet, so this guarantees a name is PRESENT, not that it is the RIGHT one:
 * with two people picking at once, whoever set the Pick ID owns both their
 * work. Accepted 2026-08-14 (user's call: one shared Pick ID per shift). The
 * fix if that ever bites is a per-device picker, which means the server taking
 * the name per call instead of reading F2 — a different trust model, not a
 * tweak.
 *
 * @returns {{ok:boolean, picker:string=, needsPicker:boolean=, error:string=}}
 */
function _boardRequirePicker() {
  var picker = '';
  try { picker = getCurrentPicker(); } catch (e) { picker = ''; }
  if (picker) return { ok: true, picker: picker };
  return {
    ok: false,
    needsPicker: true,
    error: 'Set the Pick ID first — every pick is filed under a name.'
  };
}


// =======================================================================================
// PUBLIC: board console — mark an order picked (PENDING ↔ PREPARING ONLY)
// =======================================================================================

/**
 * The Floor Board's interactive "✓ Pick" action. Called from the board via
 * google.script.run. Routes through the canonical updateOrderStatus (lock +
 * Activity Log + Telegram sync inherited).
 *
 * SAFETY — the board is a URL-reachable surface with no per-user PIN, so this
 * function is deliberately NARROW: it will ONLY set PENDING or PREPARING. It
 * cannot ship, cancel, or delete anything, regardless of who opens the link.
 * (PREPARING is reversible/non-terminal/no-customer-impact, and the team
 * already has this exact toggle via the Telegram buttons.)
 */
function boardSetStatus(orderId, status, sku) {
  orderId = String(orderId || '').trim();
  status  = String(status  || '').trim().toUpperCase();
  sku     = String(sku     || '').trim();
  if (!orderId) return { ok: false, error: 'No order' };
  if (status !== 'PENDING' && status !== 'PREPARING') {
    return { ok: false, error: 'Board may only set PENDING or PREPARING' };
  }
  // ⚠ BEFORE THE LOCK, deliberately — this is one cell read and it must not
  // extend the time the script lock is held on the floor's most frequent write.
  // It also gates the UNDO (→ PENDING) on purpose: reverting a pick is a
  // warehouse-side decision and deserves a name for exactly the same reason.
  var gate = _boardRequirePicker();
  if (!gate.ok) return gate;
  try {
    // ⚠ ONE LINE, NOT THE WHOLE ORDER — reported from the floor 2026-08-10.
    // ✓ Pick sits on a ROW, so a picker taps it having grabbed THAT part. It
    // was flipping every line of the sales order, which on a five-line box
    // meant one tap marked four parts picked that were still on their shelves.
    // It also silently broke the kit strip: progress jumped 0 of 5 → 5 of 5 in
    // one tap, so the counter that exists to stop a half-packed box shipping
    // could never show a half-packed box.
    // A SKU is passed whenever the tap came from a row. Omitting it keeps the
    // whole-order behaviour for any caller that genuinely means the order.
    var target = sku ? { orderId: orderId, sku: sku } : orderId;
    var res = updateOrderStatus(target, status, { source: 'board', syncTelegram: true });
    // The board polls straight after a pick — without this it would be served a
    // cached tick still showing PENDING.
    _dashBustTickCache();
    // ⚠ PUBLISH INLINE — ✓ Pick is the most frequent HUMAN action on the board,
    // and until 2026-08-15 it was the only one still riding the dirty flag alone.
    // That cost it the full 1-min trigger wait (45s typical / 95s worst end to
    // end), which outlived the client's optimistic override and un-picked the row
    // in front of the picker. Reported from the floor; see PICK_OVERRIDE_MS.
    //
    // ⚠ AFTER THE LOCK, NOT INSIDE IT. updateOrderStatus takes and releases its
    // own lock; this runs once that is done, so a ~3.5s rebuild can never extend
    // lock contention on the floor's most frequent write. The picker is not
    // waiting on it either — the row flipped optimistically on the tap.
    //
    // ⚠ NOT in updateOrderStatus. Machine paths (n8n sweeps) and the Telegram
    // PREP tap deliberately stay on the flag — the 2026-08-14 ruling. This is the
    // board path only.
    //
    // Best-effort by contract: the cell is already written and keep-fresh is the
    // backstop, so a publish failure must never turn a successful pick into an
    // error on the tablet.
    try {
      if (typeof publishBoardTickInline === 'function') publishBoardTickInline(undefined, 'pick');
    } catch (e) {
      console.log('boardSetStatus inline publish: ' + e);
    }
    return { ok: !!(res && res.count), count: (res && res.count) || 0, status: status };
  } catch (e) {
    console.error('boardSetStatus: ' + e);
    return { ok: false, error: String(e) };
  }
}


// =======================================================================================
// PUBLIC: board console — record a physical count into ◩ LEFT
// =======================================================================================

/**
 * Write the picker's physical shelf count to ◩ LEFT for ONE line.
 *
 * THE WORKFLOW THIS REPLACES (described by the floor, 2026-08-11): the picker
 * pulls the line, glances at the shelf, and if what remains matches what the
 * system said would remain they type NOTHING. Only a DEVIANCE gets written —
 * then today they carry it on paper to a PC and correct Zoho. This kills the
 * paper and the walk; the Zoho correction stays a separate, reviewed step.
 *
 * ⚠ SCOPE IS DELIBERATELY TINY, exactly like boardSetStatus. This writes ONE
 * CELL in ONE COLUMN. It cannot touch status, quantity, price or stock, so the
 * public board URL gains no new power over inventory — the number lands on the
 * sheet where a human already decides what to do with it.
 *
 * ⚠ Rows are resolved by orderId + SKU INSIDE the lock, never by a row number
 * from the client. A polling client's row number can be stale by the time it
 * arrives (the 2026-05-08 shifted-row incident); values cannot be.
 *
 * @param {string} orderId
 * @param {string} sku
 * @param {number|string} count  the physical count; "" clears the cell
 * @returns {{ok:boolean, count:number, cleared:boolean, error:string=}}
 */
function boardSetLeft(orderId, sku, count) {
  orderId = String(orderId || '').trim();
  sku     = String(sku     || '').trim();
  var raw = String(count == null ? '' : count).trim();
  if (!orderId || !sku) return { ok: false, error: 'Need an order and a SKU' };

  var clearing = (raw === '');
  var n = clearing ? null : parseFloat(raw.replace(/[^0-9.\-]/g, ''));
  if (!clearing && (isNaN(n) || n < 0)) {
    return { ok: false, error: 'Count must be zero or more' };
  }

  // Same chokepoint as ✓ Pick — and before the lock for the same reason. In
  // practice this costs the floor nothing: ✓ Pick almost always comes first, so
  // the Pick ID is already set by the time anyone counts a shelf. It closes the
  // one remaining door: counting BEFORE picking anything.
  var gate = _boardRequirePicker();
  if (!gate.ok) return gate;

  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) return { ok: false, error: 'Sheet busy — try again' };

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return { ok: false, error: 'No sheet' };
    var lastRow = sheet.getLastRow();
    if (lastRow < Schema.dataStartRow) return { ok: false, error: 'No rows' };

    var rows = _resolveStatusTargetRows(sheet, { orderId: orderId, sku: sku }, lastRow);
    if (!rows.length) return { ok: false, error: 'Line not found — it may have shipped' };

    for (var i = 0; i < rows.length; i++) {
      sheet.getRange(rows[i], Schema.cols.LEFT).setValue(clearing ? '' : n);
    }
    SpreadsheetApp.flush();

    // The count is a real operational event — who counted what, and when.
    // Best-effort: a logging failure must never lose the number itself.
    try {
      // ⚠ POSITIONAL, not an options object:
      //   (event, orderId, sku, qty, source, detail, picker, note)
      // The picker is left UNDEFINED on purpose — logActivity reads G2 itself
      // for warehouse-side sources, which is the one place that value is
      // authoritative. Passing "" would override it with nothing.
      logActivity('NOTE', orderId, sku, '', 'board',
                  clearing ? 'Shelf count cleared'
                           : 'Shelf count ◩ LEFT = ' + n);
    } catch (e) { console.log('boardSetLeft log: ' + e); }

    _dashBustTickCache();
    return { ok: true, count: clearing ? 0 : n, cleared: clearing, rows: rows.length };
  } catch (e) {
    console.error('boardSetLeft: ' + e);
    return { ok: false, error: String(e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


// =======================================================================================
// PUBLIC: radio now-playing (server-side fetch — bypasses browser CORS)
// =======================================================================================

/**
 * Returns "artist – title" for a SomaFM station's current track, fetched
 * server-side (UrlFetchApp has no CORS restriction, unlike the browser — which
 * is why the client fetch came back blank). Called from the Floor Board radio
 * widget via google.script.run. Non-SomaFM stations (empty id, e.g. the Quran
 * stream) and any failure return '' → the widget just shows the station name.
 */
function getRadioNowPlaying(stationId) {
  try {
    if (!stationId) return '';
    var resp = UrlFetchApp.fetch('https://somafm.com/songs/' + encodeURIComponent(stationId) + '.json', {
      muteHttpExceptions: true,
      followRedirects:    true
    });
    if (resp.getResponseCode() !== 200) return '';
    var data = JSON.parse(resp.getContentText());
    if (data && data.songs && data.songs.length) {
      var s = data.songs[0];
      return ((s.artist || '') + (s.title ? ' – ' + s.title : '')).trim();
    }
  } catch (e) {
    console.error('getRadioNowPlaying: ' + e);
  }
  return '';
}


// =======================================================================================
// PRIVATE: dashboard extras
// =======================================================================================

/**
 * Pace Car projection — current ships/hr × hours-remaining-in-shift.
 * Returns the floor projection used in "ON PACE FOR N BY 5PM."
 */
function _dashPaceCarStats(cockpit) {
  if (!cockpit) return null;
  var shipped = parseFloat(cockpit.shippedToday) || 0;

  var tz = SpreadsheetApp.openById(SPREADSHEET_ID).getSpreadsheetTimeZone() || 'America/Chicago';
  var now = new Date();
  var hr = parseFloat(Utilities.formatDate(now, tz, 'H')) +
           parseFloat(Utilities.formatDate(now, tz, 'm')) / 60;

  var elapsedHrs = Math.max(0.25, hr - DASHBOARD.sunriseHour);
  var remainingHrs = Math.max(0, DASHBOARD.sunsetHour - hr);
  var ratePerHr = shipped / elapsedHrs;
  var projection = shipped + Math.round(ratePerHr * remainingHrs);

  return {
    shipped:      shipped,
    ratePerHr:    Math.round(ratePerHr * 10) / 10,
    remainingHrs: Math.round(remainingHrs * 10) / 10,
    projection:   projection,
    insideWorkday: (hr >= DASHBOARD.sunriseHour && hr <= DASHBOARD.sunsetHour)
  };
}

/**
 * Every OPEN (PENDING / PREPARING) row across BOTH tables — the picker's live
 * worklist. Each = {channel, orderId, sku, qty, location, status, note}. Drives
 * the Floor Board "To pick" panel so a picker can grab items straight off the
 * screen (SKU · qty · location) without opening the sheet — including manually
 * typed eBay replacement rows (Missing:/Replacement #:). Sorted by LOCATION for
 * a natural pick walk (NOT FOUND / blank sink to the end), so it reads in aisle
 * order regardless of where rows physically sit in the (unsorted) sheet.
 * Capped to keep the paint cheap.
 */
function _dashOpenOrders() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return [];

  // FULL width now (was up to HAND) so the paid-shipping count can be computed
  // in this same pass — see the `paidCount` note below.
  var n = lastRow - Schema.dataStartRow + 1;
  var data = sheet.getRange(Schema.dataStartRow, 1, n, Schema.dataWidth).getValues();

  var kitSkus = _dashKitSkuSet();     // cached — see helper

  // ── KIT PARENTS THAT HAVE ALREADY BEEN EXPANDED ───────────────────────────
  // Expansion inserts the components as their own rows and deliberately LEAVES
  // the parent in place; kit-parent auto-follow only fires on a TERMINAL
  // transition, so the parent sits at PENDING for the whole pick. The board was
  // therefore listing it as one more pick line — wearing a KIT badge and a
  // ✓ Pick button — for a box that no longer exists as a physical thing, and
  // counting its qty again in UNITS TO PICK on top of the components.
  //
  // Components are tagged "↳ from KIT-<parent sku>" and carry the parent's
  // SALES ORDER, so the pairing is exact.
  //
  // ⚠ Collapse ONLY when the order has exactly ONE open parent row for that
  // SKU. Two of the same kit on one SO with just one of them expanded cannot be
  // told apart from here, and hiding a kit nobody has decided yet is far worse
  // than showing one extra row — so the ambiguous case stays visible.
  // The SAME pass tallies each expanded kit for the board's KIT STRIP.
  //
  // ⚠ THE DENOMINATOR MUST COME FROM THE SHEET, NEVER FROM THE KIT REGISTRY.
  // Expansion supports EXCLUSIONS and custom adds, so a kit whose registry
  // composition is 8 parts may only have had 6 rows inserted. A board that
  // promised "of 8" would send a picker hunting for two components that were
  // deliberately left out — worse than showing no counter at all. What was
  // actually INSERTED is the only honest total, and it is right here in the
  // rows we have already read.
  var expandedOf = {}, parentsOf = {}, parentNoteOf = {};
  var kitTot = {}, kitDone = {}, kitLeft = {}, kitMeta = {};
  for (var p = 0; p < data.length; p++) {
    var pSku = String(data[p][Schema.idx("SKU")] || "").trim();
    if (pSku.toUpperCase() === Schema.boundaryMarker) continue;
    if (!pSku) continue;
    var pStatus = String(data[p][Schema.idx("STATUS")] || "").trim().toUpperCase();
    var pOpen   = (pStatus === Schema.status.PENDING ||
                   pStatus === Schema.status.PREPARING);
    var pOrder = String(data[p][Schema.idx("SALES_ORDER")] || "").trim();
    var pNote  = String(data[p][Schema.idx("NOTE")] || "").trim();
    var pm = _dashKitTag(pNote);
    if (pm) {
      var ek = pOrder + "|" + pm;

      // TALLY over EVERY status, not just the open ones. A component already
      // shipped still counts toward the kit's SIZE — otherwise the denominator
      // shrinks as the pick proceeds and a half-finished box prints "3 of 3".
      // CANCELED is the exception: that line is no longer part of the box, so
      // it counts toward neither side.
      if (pStatus !== Schema.status.CANCELED) {
        kitTot[ek] = (kitTot[ek] || 0) + 1;
        if (pStatus === Schema.status.PENDING) {
          // still to find — remember WHERE, that is the actionable part
          var pLoc = String(data[p][Schema.idx("LOCATION")] || "").trim();
          if (!kitLeft[ek]) kitLeft[ek] = [];
          if (pLoc && pLoc.toUpperCase() !== "NOT FOUND" &&
              kitLeft[ek].indexOf(pLoc) === -1) kitLeft[ek].push(pLoc);
        } else {
          // ✓ Pick flips PENDING → PREPARING, so PREPARING means grabbed.
          kitDone[ek] = (kitDone[ek] || 0) + 1;
        }
        if (!kitMeta[ek]) kitMeta[ek] = { parent: pm, order: pOrder };
      }

      // The COLLAPSE decision below keeps its ORIGINAL open-only semantics —
      // it asks "is this parent still standing over live components", which is
      // a different question from "how big is this kit".
      if (pOpen) expandedOf[ek] = (expandedOf[ek] || 0) + 1;
    } else if (kitSkus[pSku] === 1 && pOpen) {
      var pk = pOrder + "|" + pSku;
      parentsOf[pk] = (parentsOf[pk] || 0) + 1;
      // ⚠ KEEP THE PARENT'S NOTE ALIVE (2026-08-14). Once a kit is expanded the
      // parent row is collapsed off the board — correctly, it is not pickable —
      // and ANY note written on it after that point had nowhere to go. That is
      // precisely when a hold gets added ("customer called, don't ship this"),
      // so the one path most likely to carry a safety-critical instruction was
      // the one path that went silent. Expansion already copies the parent's
      // note onto every component at WRITE time; this does the same at READ
      // time, so the result is identical whether the note was written before or
      // after. The picker then meets it at EVERY shelf they walk for that box,
      // which is what a hold needs — not one line at the top of a list.
      if (pNote) parentNoteOf[pk] = pNote;
    }
  }

  var out = [];
  var paidCount = 0;
  var inDirect = false;
  for (var i = 0; i < data.length; i++) {
    var sku = String(data[i][Schema.idx("SKU")] || "").trim();
    // Boundary divider (col A == "DIRECT") flips us onto the DIRECT side.
    if (sku.toUpperCase() === Schema.boundaryMarker) { inDirect = true; continue; }
    if (!sku) continue;
    var status = String(data[i][Schema.idx("STATUS")] || "").trim().toUpperCase();

    // PAID SHIPPING — the ONLY alert the Floor Board renders (the amber strip).
    // Computed here, inside a scan that was happening anyway, so the board no
    // longer has to call getActionableAlerts() — which opens EIGHT sheets to
    // produce seven numbers the board discards.
    //
    // Rule copied deliberately from Alerts.js (the 2026-04-30 fix): parse as a
    // dollar amount and flag only when > 0. That rejects "FREE", "", zero AND
    // the DIRECT header row, whose column J literally contains the text
    // "SHIP COST" — the original false-positive. The header-glyph guards below
    // mirror the same belt-and-suspenders check.
    var skuUpper = sku.toUpperCase();
    var headerish = (skuUpper.indexOf('◈') !== -1) || skuUpper === 'SKU' ||
                    skuUpper === '# SKU' || skuUpper === '◈ SKU';
    if (!headerish && !Schema.isTerminal(status)) {
      var scStr = String(data[i][Schema.idx("SHIP_COST")] == null
                           ? "" : data[i][Schema.idx("SHIP_COST")]).trim();
      var scNum = parseFloat(scStr.replace(/[^0-9.\-]/g, ''));
      if (!isNaN(scNum) && scNum > 0) paidCount++;
    }

    if (status !== Schema.status.PENDING && status !== Schema.status.PREPARING) continue;

    var note = String(data[i][Schema.idx("NOTE")] || "").trim();

    // An expanded kit's parent is not a pickable line — its components are.
    var orderId     = String(data[i][Schema.idx("SALES_ORDER")] || "").trim();
    var compTag     = _dashKitTag(note);
    var isComponent = !!compTag;
    var isKitParent = !isComponent && kitSkus[sku] === 1;
    var kitKey      = orderId + "|" + sku;
    if (isKitParent && expandedOf[kitKey] > 0 && parentsOf[kitKey] === 1) continue;

    var row = {
      channel:  inDirect ? "DIRECT" : "EBAY",
      orderId:  orderId,
      sku:      sku,
      qty:      data[i][Schema.idx("QTY")],
      location: String(data[i][Schema.idx("LOCATION")] || "").trim(),
      status:   status,
      note:     note,
      // ◩ HAND / ◩ LEFT — free, because this scan already reads the full row
      // width for the paid-shipping tally.
      // HAND = MI.available / Zoho.available_stock, never decremented per row
      // (the 2026-05-09 ruling). ⚠ It is the figure the picker should find on
      // the shelf BEFORE pulling THIS line — it is NOT "what will be left
      // afterwards". The board subtracting qty from it was the 2026-08-12
      // false-deviance bug: every correct ×1 shelf reported "+1 vs system" and
      // FIX ZOHO would have pushed hand+1 into Zoho. Subtract qty ONLY once the
      // line is PREPARING, i.e. once the units have actually left the shelf.
      // LEFT is the picker's physical count, an EXCEPTION FIELD — blank means
      // the shelf agreed.
      hand:     _dashNumOrNull(data[i][Schema.idx("HAND")]),
      left:     _dashNumOrNull(data[i][Schema.idx("LEFT")]),
      // A kit row that hasn't been expanded yet can't actually be picked from
      // the shelf — its components aren't on the sheet. Flagging it stops a
      // picker walking to a K-* aisle expecting a box. Components themselves
      // (NOTE starts "↳ from KIT-") are normal pickable rows, so exclude them.
      // Already-expanded parents never reach here — they were skipped above —
      // so the badge now means what it says: THIS one still needs a decision.
      isKit: isKitParent
    };
    // Which kit this component belongs to. On the SHEET the "↳ from KIT-x"
    // note says so on every row and survives any sort; the board received that
    // text and threw it away, so a component looked exactly like a standalone
    // line. Set ONLY on components — an absent field costs nothing in the
    // published payload, which is guarded against the 50K cell limit.
    if (isComponent) {
      row.kit = compTag;
      // The collapsed parent's note, carried down so it cannot be lost. Sent
      // RAW — the client's humanNote() strips machine prefixes and dedupes it
      // against whatever this row already carries, so a note written BEFORE
      // expansion (already copied here by KitExpansion) does not print twice.
      var pn = parentNoteOf[orderId + "|" + compTag];
      if (pn) row.kitNote = pn;
    }
    out.push(row);
  }

  // ── KITS IN PROGRESS ──────────────────────────────────────────────────────
  // One entry per expanded kit that still has a component to find. The rows
  // carry membership (a coloured spine); this carries COMPLETENESS, which is
  // one fact about the kit and therefore does NOT belong repeated on every
  // row — the same reason GRAB was removed from all eleven lines.
  var kits = [];
  for (var kk in kitTot) {
    if (!Object.prototype.hasOwnProperty.call(kitTot, kk)) continue;
    var kTot = kitTot[kk], kDone = kitDone[kk] || 0;
    if (kTot < 2) continue;       // a single-component "kit" needs no thread
    // FINISHED KITS STAY IN THIS LIST. Membership is permanent for as long as
    // the rows are open — the picker who grabbed all eight still has to know
    // which eight go in one box at the bench. The board drops only the STRIP
    // when done >= total; the rows keep their thread.
    kits.push({
      key:    kk,
      parent: kitMeta[kk].parent,
      order:  kitMeta[kk].order,
      total:  kTot,
      done:   kDone,
      left:   (kitLeft[kk] || []).slice().sort(compareLocations)
    });
  }
  // TWO orderings, deliberately. Colour is assigned from a STABLE key sort so
  // a kit keeps its hue for its whole life and the list does not flicker as
  // kits complete. Display is by completeness DESCENDING, because a kit at
  // 7 of 8 is the one most likely to be mistaken for finished.
  kits.sort(function (a, b) { return a.key < b.key ? -1 : (a.key > b.key ? 1 : 0); });
  for (var ci = 0; ci < kits.length; ci++) {
    // hue alternates fastest, pattern carries the overflow — so the first two
    // kits differ by COLOUR (read fastest) and only the third and fourth need
    // the picker to notice a dashed spine.
    kits[ci].hue  = ci % DASHBOARD.kitHues;
    kits[ci].dash = Math.floor(ci / DASHBOARD.kitHues) % DASHBOARD.kitPatterns;
  }
  kits.sort(function (a, b) { return (b.done / b.total) - (a.done / a.total); });

  // SORT FIRST, THEN CAP.
  // 2026-08-05: these were the other way round — the cap ran during the scan, so
  // on a busy day the board kept the first 60 rows in SHEET order and sorted
  // only those. The rows it dropped were whichever sat lowest in the sheet, not
  // the last by aisle, so a picker could walk the whole list and still miss
  // items. Capping after the sort makes the omission predictable (the far end of
  // the walk) — and `openOrdersTotal` on the tick lets the board SAY it's
  // truncated instead of looking complete.
  // ⚠ ORDER ANCHORS — so a multi-line order cannot be split by another
  // order's rows. See _dashComparePickRows for the full reasoning.
  _dashSetOrderAnchors(out);
  out.sort(_dashComparePickRows);

  var capped = _dashCapPerChannel(out, DASHBOARD.pickListCap);
  capped.total     = out.length;    // all four read by _buildDashboardTick
  capped.paidCount = paidCount;     // before JSON serialisation drops them
  capped.kits      = kits;
  capped.byChannel = { EBAY:   _dashCountChannel(out, "EBAY"),
                       DIRECT: _dashCountChannel(out, "DIRECT") };
  // ⚠ HELD ORDERS — INCLUDING SHIPPED ONES, which is the point (2026-08-21).
  // Everything above this line filtered to PENDING/PREPARING, so the moment a
  // label was bought the order stopped existing as far as the board was
  // concerned — while the box sat in the building for hours. holdScanRows walks
  // the SAME `data` this scan already read, at full width, so the one surface
  // that was blind costs zero extra reads to open.
  capped.held = holdScanRows(data);
  return capped;
}


/**
 * Aisle order WITHIN one block. Natural, not lexical — compareLocations() in
 * Helpers.js. A-9 must come before A-50 on a list whose entire job is the walk.
 * Rows with no shelf sink to the bottom of their own block rather than the
 * bottom of the whole list, so a box's missing part stays beside its box.
 *
 * @param {Object} a
 * @param {Object} b
 * @returns {number}
 */
function _dashCompareAisle(a, b) {
  var byLoc = _dashCompareShelf(a.location, b.location);
  if (byLoc !== 0) return byLoc;
  return String(a.sku).localeCompare(String(b.sku));
}


/**
 * SHELF ONLY — no SKU tiebreak.
 *
 * ⚠ Split out because the anchor comparison must NOT fall through to a SKU.
 * When two orders anchor on the same shelf, the anchor rows' SKUs would decide
 * which order came first — an arbitrary answer that also flips whenever an
 * unrelated line is added. The order id is the deterministic tiebreak there,
 * and it can only be reached if this stops at the shelf.
 *
 * @param {string} la
 * @param {string} lb
 * @returns {number}
 */
function _dashCompareShelf(la, lb) {
  la = String(la || ""); lb = String(lb || "");
  var am = (!la || la.toUpperCase() === "NOT FOUND") ? 1 : 0;
  var bm = (!lb || lb.toUpperCase() === "NOT FOUND") ? 1 : 0;
  if (am !== bm) return am - bm;
  return compareLocations(la, lb);
}


/**
 * THE PICK-LIST ORDER — channel first, then the rule each channel is actually
 * picked by.
 *
 * ⚠ THIS IS NOT A COSMETIC SPLIT. Every other surface already draws this line
 * and the board was the only one that didn't:
 *
 *   - the SHEET sorts eBay by status→location and DIRECT by SALES ORDER then
 *     location within the order (_compareDirectRows, OrderService.js), with a
 *     gold box painted around each SO group;
 *   - the PRINT renders two separate tables under "eBay Orders" / "Direct
 *     Orders" section heads, and gives DIRECT a per-order band as well.
 *
 * They are two different JOBS. An eBay order is ~1 line: the aisle walk IS the
 * work. A DIRECT order is a multi-line box you assemble, usually with expanded
 * kit components scattered across the floor — you work it one order at a time.
 * Interleaving them by pure aisle order asked the picker to hold both jobs in
 * their head at once and rebuild the boxes mentally as they went.
 *
 * DIRECT groups run in SALES ORDER order (ascending = oldest first, and
 * deterministic), matching the sheet exactly. Aisle order applies WITHIN each
 * order. Age-based triage is the board's own Age toggle, client-side.
 *
 * @param {Object} a
 * @param {Object} b
 * @returns {number}
 */
function _dashComparePickRows(a, b) {
  var ca = (a.channel === "DIRECT") ? 1 : 0;
  var cb = (b.channel === "DIRECT") ? 1 : 0;
  if (ca !== cb) return ca - cb;          // eBay block, then DIRECT block

  var idA = String(a.orderId || "").trim();
  var idB = String(b.orderId || "").trim();

  if (ca === 1) {
    // DIRECT — keep an order's lines together, in SO order (oldest first, and
    // deterministic). The accordion works one box at a time, so age beats
    // walk order here.
    if (!idA && idB) return 1;            // an SO-less row sinks
    if (idA && !idB) return -1;
    if (idA !== idB) return idA.localeCompare(idB);
    return _dashCompareAisle(a, b);
  }

  // ⚠ eBay: ANCHOR the whole order at its EARLIEST shelf, then keep its lines
  // together. Pure aisle order looks right and is quietly wrong the moment an
  // order has more than one line: on 2026-08-10 order 08-15017-44806 had five
  // lines at E-89 / F-6 / F-28 / F-51 / L-208, and another order's rows at
  // J-8 and J-27 sorted between the fourth and the fifth. The board banded the
  // contiguous four as "4 LINES" and stranded the fifth further down under its
  // own id — so a picker who worked the band would have shipped four of five
  // and had no reason to suspect otherwise.
  //
  // That is the SAME defect the sheet hit on 2026-08-07, where an insert split
  // a DIRECT order and the painter drew it as two complete boxes. The ruling
  // then was: fix the source so contiguity holds BY CONSTRUCTION, and make the
  // renderer incapable of claiming a fragment is whole. Both halves apply here
  // — this is the source half; withSections carries the safety net.
  //
  // Anchoring at the earliest shelf (rather than by order id) keeps the walk
  // intact: a single-line order anchors to its own shelf and slots exactly
  // where it always did, so the common eBay case is unchanged.
  if (idA !== idB) {
    var byAnchor = _dashCompareShelf((a._anchor || a).location,
                                     (b._anchor || b).location);
    if (byAnchor !== 0) return byAnchor;
    return idA.localeCompare(idB);        // same anchor shelf → deterministic
  }
  return _dashCompareAisle(a, b);
}


/**
 * Stamp every row with its ORDER's earliest shelf, so the comparator can keep a
 * multi-line order together without needing to see the whole array.
 *
 * Single-line orders anchor to themselves, which is why adding this changes
 * nothing for the ordinary eBay line.
 *
 * @param {Array<Object>} rows
 */
function _dashSetOrderAnchors(rows) {
  var best = {};
  for (var i = 0; i < rows.length; i++) {
    var id = String(rows[i].orderId || "").trim();
    if (!id) continue;
    if (!best[id] ||
        _dashCompareShelf(rows[i].location, best[id].location) < 0) best[id] = rows[i];
  }
  for (var j = 0; j < rows.length; j++) {
    var jid = String(rows[j].orderId || "").trim();
    rows[j]._anchor = { location: (jid && best[jid] ? best[jid] : rows[j]).location };
  }
}


/**
 * Cap EACH channel independently.
 *
 * ⚠ A single global cap could starve one channel to nothing: 55 open eBay
 * lines would leave 5 slots for DIRECT, so a picker looking at the DIRECT
 * section would see a fragment of one box and no sign that seven more orders
 * existed. Each channel gets its own budget, and each reports its own "+N
 * more" — the truncation stays legible per section instead of pooling into
 * one number that hides which half was cut.
 *
 * Payload size is guarded downstream regardless: Published.js trims the
 * pick list to 25 rows rather than exceed the 50K cell limit.
 *
 * @param {Array<Object>} rows  already sorted by _dashComparePickRows
 * @param {number} cap          per-channel maximum
 * @returns {Array<Object>}
 */
function _dashCapPerChannel(rows, cap) {
  var kept = [], seen = { EBAY: 0, DIRECT: 0 };
  for (var i = 0; i < rows.length; i++) {
    var ch = (rows[i].channel === "DIRECT") ? "DIRECT" : "EBAY";
    if (seen[ch] >= cap) continue;
    seen[ch]++;
    kept.push(rows[i]);
  }
  return kept;
}


/**
 * How many rows of one channel, BEFORE the cap — so the board can say how many
 * it is not showing, per section.
 *
 * @param {Array<Object>} rows
 * @param {string} channel
 * @returns {number}
 */
function _dashCountChannel(rows, channel) {
  var n = 0;
  for (var i = 0; i < rows.length; i++) {
    var ch = (rows[i].channel === "DIRECT") ? "DIRECT" : "EBAY";
    if (ch === channel) n++;
  }
  return n;
}


/**
 * A cell as a number, or null when it holds nothing usable.
 *
 * ⚠ NOT `|| null` — that would turn a real 0 into null, and 0 is the single
 * most important value here: an empty shelf. Same falsy-zero trap that made
 * the printed pick list show a blank Hand cell for a genuinely out-of-stock
 * part (2026-07-20).
 *
 * @param {*} v
 * @returns {number|null}
 */
function _dashNumOrNull(v) {
  if (v === null || v === undefined || v === "") return null;
  var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? null : n;
}


/**
 * Kit SKUs as a lookup object, CACHED.
 *
 * buildKitMap() reads the whole Kit Registry (~1,500 rows) and this sits on the
 * board's 15-second poll, so it must not run every tick. Kit composition changes
 * rarely — a few minutes of staleness only means a freshly-added kit isn't
 * flagged for a little while, which is harmless.
 *
 * @returns {Object} { "<sku>": 1 }
 */
function _dashKitSkuSet() {
  try {
    var cache = CacheService.getScriptCache();
    var hit = cache.get(DASHBOARD.kitCacheKey);
    if (hit) return JSON.parse(hit);

    var set = {};
    buildKitMap().forEach(function (_v, sku) { set[String(sku)] = 1; });
    try { cache.put(DASHBOARD.kitCacheKey, JSON.stringify(set), DASHBOARD.kitCacheSec); }
    catch (e) { /* over the 100KB cache-entry cap — just don't cache */ }
    return set;
  } catch (e) {
    try { console.log("_dashKitSkuSet: " + e); } catch (_) {}
    return {};    // degrade to "nothing is a kit" — never break the board
  }
}


// =======================================================================================
// BOARD: PICKER SELECTION
// =======================================================================================
//
// Printing refuses without a real Pick ID in Schema.cellEmployeeId — that guard
// is the single chokepoint for warehouse accountability, and once set, every
// status event for the rest of the shift carries the picker's name into the
// Activity Log.
//
// Which left the tablet in a bind: it could not print, because it could not set
// the picker, because setting it meant walking to a computer — the exact trip
// printing from the board is supposed to remove.
//
// So the board can set it, under the SAME allow-list discipline as
// boardSetStatus: the value must already be one of the dropdown's own options.
// Arbitrary text cannot be written, so the worst a stranger on the URL could do
// is select a different REAL picker — the same thing they could do by walking
// up to the sheet, and visible in the Activity Log either way.

/**
 * The Pick ID options, read from the cell's own data validation so there is
 * exactly one source of truth and adding a picker in the sheet is enough.
 *
 * @returns {{ok:boolean, pickers:Array<string>, current:string}}
 */
function getBoardPickers() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return { ok: false, pickers: [], current: "" };

    var range = sheet.getRange(Schema.pickIdA1());
    var current = String(range.getValue() || "").trim();

    var dv = range.getDataValidation();
    var opts = dv ? (dv.getCriteriaValues()[0] || []) : [];

    // Keep only real pickers. The list also carries the dropdown's own
    // placeholder ("Pick ID for Shipping"), which the print guard rejects — so
    // offering it here would just produce a confusing refusal downstream.
    var pickers = [];
    for (var i = 0; i < opts.length; i++) {
      var v = String(opts[i] || "").trim();
      if (v && /^Shipping\s*-\s*/i.test(v)) pickers.push(v);
    }

    return { ok: true, pickers: pickers, current: /^Shipping\s*-\s*/i.test(current) ? current : "" };
  } catch (err) {
    try { console.log("getBoardPickers: " + err); } catch (_) {}
    return { ok: false, pickers: [], current: "" };
  }
}


/**
 * Set the shift's picker from the board.
 *
 * ⚠ ALLOW-LISTED: the value must already appear in the cell's validation list.
 * This is the security boundary, exactly as the PENDING/PREPARING allow-list is
 * for boardSetStatus — not a PIN, but a capability narrow enough that a public
 * URL cannot do damage with it.
 *
 * @returns {{ok:boolean, picker:string, message:string}}
 */
/**
 * ONE BODY, TWO DOORS. The board and the sidebar both set the shift's picker, and this
 * is the only place that writes it.
 *
 * ⚠⚠ NOT A BARE REUSE, AND NOT A COPY. It takes `source` because SOURCE is the only
 *    thing that distinguishes a tablet action from a desk one in the Activity Log, and
 *    that distinction is the whole point of logging it. But it is not duplicated either:
 *    THE ALLOW-LIST IS THE SECURITY BOUNDARY — the public board URL can call
 *    boardSetPicker, so "only a value the dropdown itself offers" is what stops it
 *    becoming a free-text write into an attributed field. A second copy of that check is
 *    how a fix lands in one place and not the other.
 *
 * @param {string} value  must be a member of the dropdown's own option list
 * @param {string} source Activity Log SOURCE — "board" or "sidebar"
 */
function _setPickerAllowlisted(value, source) {
  var want = String(value || "").trim();
  if (!want) return { ok: false, picker: "", message: "No picker given." };

  try {
    var avail = getBoardPickers();
    if (!avail.ok) return { ok: false, picker: "", message: "Could not read the picker list." };
    if (avail.pickers.indexOf(want) === -1) {
      return { ok: false, picker: "", message: "Not a known picker." };
    }

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    // ⚠ The address is runtime state during the 2026-08-31 grace period — see
    //   Schema.pickIdA1. The cells are HIDDEN now, so this write is the only way the
    //   picker gets set from the sidebar at all.
    sheet.getRange(Schema.pickIdA1()).setValue(want);

    try {
      // "sidebar" is already in ACTIVITY_LOG.warehouseSources, so the picker
      // auto-captures on the entry — no extra plumbing needed.
      logActivity("NOTE", "", "", 0, source || "board", "picker set to " + want);
    } catch (e) { /* best-effort */ }

    // ⚠ THE PICKER IS PART OF THE TICK, so setting it IS a change to the tick
    // (2026-08-14). Without this the board served a cached payload still saying
    // "no picker" for up to a minute after someone had just identified
    // themselves — invisible while the chip was only cosmetic, but the moment
    // ✓ Pick started gating on it, the client's 45s override would expire into
    // a stale tick and ask AGAIN mid-shift. It also fixes cross-device latency:
    // a second tablet now learns the shift's picker within ~1 min rather than
    // waiting on the 8-minute keep-fresh republish.
    // Same lesson as the 2026-08-14 human-edit regression: enumerate EVERY door
    // that changes data the cache holds, not just the obvious ones.
    try { _dashBustTickCache(); } catch (e) { /* best-effort */ }

    // ⚠ AND THE SIDEBAR'S OWN CACHE — a different store with a different TTL. Busting
    //   only the board's left the panel that OWNS this control showing the old name for
    //   up to five minutes, including on a second device where the optimistic override
    //   cannot mask it. Enumerate every cache that holds the value, not just the first.
    try { _sidebarBustTickCache(); } catch (e) { /* best-effort */ }

    return { ok: true, picker: want, message: "" };
  } catch (err) {
    try { console.log("_setPickerAllowlisted(" + source + "): " + err); } catch (_) {}
    return { ok: false, picker: "", message: String(err.message || err) };
  }
}

/** The Floor Board's door. A doPost action — reachable from the public board URL. */
function setBoardPicker(value) {
  return _setPickerAllowlisted(value, "board");
}

/**
 * The sidebar's door. Runs as the INVOKING USER through google.script.run, unlike
 * setBoardPicker which arrives via doPost and therefore runs as the owner.
 *
 * ⚠⚠ THAT DIFFERENCE IS WHY THE LOCK MATTERS HERE AND NOT THERE. Sheets protection
 *    checks WHO, never WHERE FROM, so this write needs the pick-ID cell to be inside
 *    the lock's carve-out. A stale carve-out fails SILENTLY for staff while working
 *    perfectly for the owner — re-run protectAllOrdersSheet() after any address change,
 *    and verify it as a STAFF account.
 *
 * ⚠ Deliberately NOT bridged through OwnerBridge. _obRunAsOwner's dispatch map lives
 *   inside doPost on the pinned /exec, so a name added at HEAD returns "Not an
 *   allowlisted owner action" for every staff tap while working for the owner — a
 *   failure invisible to whoever tests it.
 */
function setSidebarPicker(value) {
  return _setPickerAllowlisted(value, "sidebar");
}
