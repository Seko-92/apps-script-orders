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
  sunsetHour:  17    // end of the workday — 5pm (pace projection target)
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
  var base;
  try {
    base = getSidebarTick();
  } catch (e) {
    console.error('getDashboardTick.base: ' + e);
    base = { cockpit: null, lastSync: '', api: null, alerts: null, picker: '' };
  }

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
    alerts:     base.alerts   || {},
    api:        base.api      || null,
    picker:     base.picker   || '',
    lastSync:   base.lastSync || '',
    paceCar:    pace,
    openOrders: openOrders,
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

  var n = lastRow - Schema.dataStartRow + 1;
  var data = sheet.getRange(Schema.dataStartRow, 1, n, Schema.cols.HAND).getValues();

  var out = [];
  var inDirect = false;
  for (var i = 0; i < data.length; i++) {
    var sku = String(data[i][Schema.idx("SKU")] || "").trim();
    // Boundary divider (col A == "DIRECT") flips us onto the DIRECT side.
    if (sku.toUpperCase() === Schema.boundaryMarker) { inDirect = true; continue; }
    if (!sku) continue;
    var status = String(data[i][Schema.idx("STATUS")] || "").trim().toUpperCase();
    if (status !== Schema.status.PENDING && status !== Schema.status.PREPARING) continue;
    out.push({
      channel:  inDirect ? "DIRECT" : "EBAY",
      orderId:  String(data[i][Schema.idx("SALES_ORDER")] || "").trim(),
      sku:      sku,
      qty:      data[i][Schema.idx("QTY")],
      location: String(data[i][Schema.idx("LOCATION")] || "").trim(),
      status:   status,
      note:     String(data[i][Schema.idx("NOTE")] || "").trim()
    });
    if (out.length >= 60) break;            // hard cap — keep the paint cheap
  }

  // Sort by LOCATION for a natural pick walk; NOT FOUND / blank sink to the end.
  out.sort(function (a, b) {
    var la = String(a.location || ""), lb = String(b.location || "");
    var am = (!la || la.toUpperCase() === "NOT FOUND") ? 1 : 0;
    var bm = (!lb || lb.toUpperCase() === "NOT FOUND") ? 1 : 0;
    if (am !== bm) return am - bm;
    return la.localeCompare(lb);
  });
  return out;
}
