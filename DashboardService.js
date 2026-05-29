// =======================================================================================
// DashboardService.js — server endpoints for the wall-mounted warehouse dashboard
// =======================================================================================
//
// The dashboard (Dashboard.html) is served via doGet() in UIService.js. It polls
// getDashboardTick() every ~15s for a single payload that wraps the sidebar tick
// (cockpit + alerts + api + picker + lastSync) with dashboard-only enrichments:
//
//   - hourlyBuckets[]   — 11 elements 7am-5pm with {shipped,received,total}
//                         drives the "day as a comic strip" hourly panels
//   - recentPrints[]    — last 10 PRINTED events from Activity Log
//                         drives the Doc Echo rail of FUL chips
//   - todayEvents[]     — every today event mapped onto sunrise→sunset position
//                         drives the Sun Arc dots
//   - paceCarStats      — current shipping velocity + linear projection to 5pm
//                         drives the "ON PACE FOR N BY 5PM" forecast
//
// All extras are best-effort: any throw → empty default, no error propagates to
// the client. The dashboard treats missing fields as "no data yet" and continues
// painting from last-known values.
//
// Also exposes openWarehouseDashboard() — opens the dashboard inside the sheet
// via showModalDialog for at-the-desk viewing without leaving the spreadsheet.
// The wall-mounted TV uses doGet() instead (separate web-app deployment URL).
// =======================================================================================


// ---------- DASHBOARD CONSTANTS ----------
var DASHBOARD = {
  sunriseHour:  7,     // start of the "workday arc"
  sunsetHour:   17,    // end of the workday arc (5pm)
  printRailLimit: 10,  // how many recent prints to surface in Doc Echo
  arcEventCap:    250  // hard cap so the arc paint stays cheap on busy days
};


// =======================================================================================
// PUBLIC: opener
// =======================================================================================

/**
 * Open the warehouse dashboard in an in-sheet modal. Sized for a typical
 * desktop browser. The wall-mounted display uses the doGet web-app URL.
 */
function openWarehouseDashboard() {
  var html = HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setWidth(1400)
    .setHeight(800)
    .setTitle('HQ · Warehouse Dashboard');
  SpreadsheetApp.getUi().showModalDialog(html, 'HQ Motor Service · Warehouse Dashboard');
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

  var hourly = [];
  var prints = [];
  var arc = [];
  var pace = null;

  try { hourly = _dashHourlyBuckets(); }
  catch (e) { console.error('getDashboardTick.hourly: ' + e); }

  try { prints = _dashRecentPrints(DASHBOARD.printRailLimit); }
  catch (e) { console.error('getDashboardTick.prints: ' + e); }

  try { arc = _dashTodayEventsForArc(); }
  catch (e) { console.error('getDashboardTick.arc: ' + e); }

  try { pace = _dashPaceCarStats(base.cockpit); }
  catch (e) { console.error('getDashboardTick.pace: ' + e); }

  return {
    cockpit:       base.cockpit  || {},
    alerts:        base.alerts   || {},
    api:           base.api      || null,
    picker:        base.picker   || '',
    lastSync:      base.lastSync || '',
    hourlyBuckets: hourly,
    recentPrints:  prints,
    todayEvents:   arc,
    paceCar:       pace,
    serverTime:    new Date().toISOString()
  };
}


// =======================================================================================
// PRIVATE: dashboard extras
// =======================================================================================

/**
 * 11 hourly buckets (sunrise → sunset) with {shipped, received, total}
 * counted off the Activity Log. Drives the comic-strip hourly panels.
 */
function _dashHourlyBuckets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var log = ss.getSheetByName(ACTIVITY_LOG.sheetName);
  if (!log) return [];
  var last = log.getLastRow();
  if (last < ACTIVITY_LOG.dataStartRow) return [];

  var tz = ss.getSpreadsheetTimeZone() || 'America/Chicago';
  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var span = DASHBOARD.sunsetHour - DASHBOARD.sunriseHour + 1;
  var buckets = [];
  for (var h = 0; h < span; h++) {
    buckets.push({
      hour:     DASHBOARD.sunriseHour + h,
      shipped:  0,
      received: 0,
      total:    0
    });
  }

  // Read TS + EVENT only — cheap two-column scan
  var nRows = last - ACTIVITY_LOG.dataStartRow + 1;
  var data = log.getRange(
    ACTIVITY_LOG.dataStartRow,
    ACTIVITY_LOG.cols.TIMESTAMP,
    nRows,
    2
  ).getValues();

  for (var i = 0; i < data.length; i++) {
    var ts = data[i][0];
    var evt = String(data[i][1] || '').toUpperCase();
    if (!ts) continue;
    var d = (ts instanceof Date) ? ts : new Date(ts);
    if (isNaN(d.getTime())) continue;
    if (Utilities.formatDate(d, tz, 'yyyy-MM-dd') !== todayStr) continue;
    var hr = parseInt(Utilities.formatDate(d, tz, 'H'), 10);
    if (hr < DASHBOARD.sunriseHour || hr > DASHBOARD.sunsetHour) continue;
    var idx = hr - DASHBOARD.sunriseHour;
    buckets[idx].total++;
    if (evt === 'SHIPPED') buckets[idx].shipped++;
    else if (evt === 'RECEIVED') buckets[idx].received++;
  }
  return buckets;
}

/**
 * Last N PRINTED events. Each event carries the FUL doc number in DETAIL
 * and the picker who fired the print. Drives the Doc Echo rail chips.
 */
function _dashRecentPrints(limit) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var log = ss.getSheetByName(ACTIVITY_LOG.sheetName);
  if (!log) return [];
  var last = log.getLastRow();
  if (last < ACTIVITY_LOG.dataStartRow) return [];

  // Scan up to 500 most-recent rows looking for PRINTED — small N keeps cost flat
  var scanRows = Math.min(500, last - ACTIVITY_LOG.dataStartRow + 1);
  var startRow = last - scanRows + 1;
  var data = log.getRange(
    startRow,
    ACTIVITY_LOG.cols.TIMESTAMP,
    scanRows,
    ACTIVITY_LOG.dataWidth
  ).getValues();

  var tz = ss.getSpreadsheetTimeZone() || 'America/Chicago';
  var prints = [];
  for (var i = data.length - 1; i >= 0 && prints.length < limit; i--) {
    var evt = String(data[i][ACTIVITY_LOG.idx('EVENT')] || '').toUpperCase();
    if (evt !== 'PRINTED') continue;
    var ts = data[i][ACTIVITY_LOG.idx('TIMESTAMP')];
    var d = (ts instanceof Date) ? ts : new Date(ts);
    if (isNaN(d.getTime())) continue;
    prints.push({
      timeLabel: Utilities.formatDate(d, tz, 'h:mm a'),
      epoch:     d.getTime(),
      detail:    String(data[i][ACTIVITY_LOG.idx('DETAIL')] || ''),
      picker:    String(data[i][ACTIVITY_LOG.idx('PICKER')] || '')
    });
  }
  return prints;
}

/**
 * Every today event mapped to its position on the sunrise→sunset arc.
 * Each event = {position: 0..1, type, orderId}. The arc renderer uses
 * `position` directly for x-coordinate. NOTE events excluded (visual noise).
 */
function _dashTodayEventsForArc() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var log = ss.getSheetByName(ACTIVITY_LOG.sheetName);
  if (!log) return [];
  var last = log.getLastRow();
  if (last < ACTIVITY_LOG.dataStartRow) return [];

  var tz = ss.getSpreadsheetTimeZone() || 'America/Chicago';
  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var sr = DASHBOARD.sunriseHour;
  var span = DASHBOARD.sunsetHour - sr;

  // Read TS + EVENT + ORDER_ID
  var nRows = last - ACTIVITY_LOG.dataStartRow + 1;
  var data = log.getRange(
    ACTIVITY_LOG.dataStartRow,
    ACTIVITY_LOG.cols.TIMESTAMP,
    nRows,
    3
  ).getValues();

  var events = [];
  for (var i = 0; i < data.length; i++) {
    var ts = data[i][0];
    var evt = String(data[i][1] || '').toUpperCase();
    if (!ts || evt === 'NOTE') continue;
    var d = (ts instanceof Date) ? ts : new Date(ts);
    if (isNaN(d.getTime())) continue;
    if (Utilities.formatDate(d, tz, 'yyyy-MM-dd') !== todayStr) continue;
    var hr = parseFloat(Utilities.formatDate(d, tz, 'H')) +
             parseFloat(Utilities.formatDate(d, tz, 'm')) / 60;
    if (hr < sr || hr > DASHBOARD.sunsetHour) continue;
    events.push({
      position: (hr - sr) / span,
      type:     evt,
      orderId:  String(data[i][2] || '')
    });
  }
  if (events.length > DASHBOARD.arcEventCap) {
    events = events.slice(events.length - DASHBOARD.arcEventCap);
  }
  return events;
}

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
