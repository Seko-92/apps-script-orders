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
  tickCacheSec: 45
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
  var base = { cockpit: null, lastSync: '', picker: '' };
  try { base.cockpit  = getDashboardSnapshot(); } catch (e) { console.error('tick.cockpit: '  + e); }
  try { base.lastSync = getLastSyncFromSheet(); } catch (e) { console.error('tick.lastSync: ' + e); }
  try { base.picker   = getCurrentPicker();     } catch (e) { console.error('tick.picker: '   + e); }

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
    lastSync:   base.lastSync || '',
    paceCar:    pace,
    openOrders: openOrders,
    // How many were open BEFORE the cap. The array carries `.total` as an own
    // property, but that does NOT survive JSON serialisation of an array — so
    // it has to be lifted onto the tick explicitly or the board can never tell
    // it's showing a truncated walk.
    openOrdersTotal: (openOrders && openOrders.total) || (openOrders || []).length,
    serverTime: new Date().toISOString()
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
function boardSetStatus(orderId, status) {
  orderId = String(orderId || '').trim();
  status  = String(status  || '').trim().toUpperCase();
  if (!orderId) return { ok: false, error: 'No order' };
  if (status !== 'PENDING' && status !== 'PREPARING') {
    return { ok: false, error: 'Board may only set PENDING or PREPARING' };
  }
  try {
    var res = updateOrderStatus(orderId, status, { source: 'board', syncTelegram: true });
    // The board polls straight after a pick — without this it would be served a
    // cached tick still showing PENDING.
    _dashBustTickCache();
    return { ok: !!(res && res.count), count: (res && res.count) || 0, status: status };
  } catch (e) {
    console.error('boardSetStatus: ' + e);
    return { ok: false, error: String(e) };
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
  var expandedOf = {}, parentsOf = {};
  for (var p = 0; p < data.length; p++) {
    var pSku = String(data[p][Schema.idx("SKU")] || "").trim();
    if (pSku.toUpperCase() === Schema.boundaryMarker) continue;
    if (!pSku) continue;
    var pStatus = String(data[p][Schema.idx("STATUS")] || "").trim().toUpperCase();
    if (pStatus !== Schema.status.PENDING && pStatus !== Schema.status.PREPARING) continue;
    var pOrder = String(data[p][Schema.idx("SALES_ORDER")] || "").trim();
    var pNote  = String(data[p][Schema.idx("NOTE")] || "").trim();
    var pm = pNote.match(/^↳ from KIT-(\S+)/);
    if (pm) {
      var ek = pOrder + "|" + pm[1];
      expandedOf[ek] = (expandedOf[ek] || 0) + 1;
    } else if (kitSkus[pSku] === 1) {
      var pk = pOrder + "|" + pSku;
      parentsOf[pk] = (parentsOf[pk] || 0) + 1;
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
    var isComponent = note.indexOf("↳ from KIT-") === 0;
    var isKitParent = !isComponent && kitSkus[sku] === 1;
    var kitKey      = orderId + "|" + sku;
    if (isKitParent && expandedOf[kitKey] > 0 && parentsOf[kitKey] === 1) continue;

    out.push({
      channel:  inDirect ? "DIRECT" : "EBAY",
      orderId:  orderId,
      sku:      sku,
      qty:      data[i][Schema.idx("QTY")],
      location: String(data[i][Schema.idx("LOCATION")] || "").trim(),
      status:   status,
      note:     note,
      // A kit row that hasn't been expanded yet can't actually be picked from
      // the shelf — its components aren't on the sheet. Flagging it stops a
      // picker walking to a K-* aisle expecting a box. Components themselves
      // (NOTE starts "↳ from KIT-") are normal pickable rows, so exclude them.
      // Already-expanded parents never reach here — they were skipped above —
      // so the badge now means what it says: THIS one still needs a decision.
      isKit: isKitParent
    });
  }

  // SORT FIRST, THEN CAP.
  // 2026-08-05: these were the other way round — the cap ran during the scan, so
  // on a busy day the board kept the first 60 rows in SHEET order and sorted
  // only those. The rows it dropped were whichever sat lowest in the sheet, not
  // the last by aisle, so a picker could walk the whole list and still miss
  // items. Capping after the sort makes the omission predictable (the far end of
  // the walk) — and `openOrdersTotal` on the tick lets the board SAY it's
  // truncated instead of looking complete.
  //
  // Aisle order is NATURAL, not lexical — compareLocations() in Helpers.js.
  // A-9 must come before A-50 on a list whose entire job is the walk order.
  out.sort(function (a, b) {
    var la = String(a.location || ""), lb = String(b.location || "");
    var am = (!la || la.toUpperCase() === "NOT FOUND") ? 1 : 0;
    var bm = (!lb || lb.toUpperCase() === "NOT FOUND") ? 1 : 0;
    if (am !== bm) return am - bm;
    var byLoc = compareLocations(la, lb);
    if (byLoc !== 0) return byLoc;
    return String(a.sku).localeCompare(String(b.sku));
  });

  var capped = out.length > DASHBOARD.pickListCap
                 ? out.slice(0, DASHBOARD.pickListCap) : out;
  capped.total     = out.length;    // both read by _buildDashboardTick before
  capped.paidCount = paidCount;     // JSON serialisation drops these
  return capped;
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

    var range = sheet.getRange(Schema.cellEmployeeId);
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
function setBoardPicker(value) {
  var want = String(value || "").trim();
  if (!want) return { ok: false, picker: "", message: "No picker given." };

  try {
    var avail = getBoardPickers();
    if (!avail.ok) return { ok: false, picker: "", message: "Could not read the picker list." };
    if (avail.pickers.indexOf(want) === -1) {
      return { ok: false, picker: "", message: "Not a known picker." };
    }

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    sheet.getRange(Schema.cellEmployeeId).setValue(want);

    try {
      logActivity("NOTE", "", "", 0, "board", "picker set to " + want);
    } catch (e) { /* best-effort */ }

    return { ok: true, picker: want, message: "" };
  } catch (err) {
    try { console.log("setBoardPicker: " + err); } catch (_) {}
    return { ok: false, picker: "", message: String(err.message || err) };
  }
}
