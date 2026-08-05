/**
 * KpiHistory.js — weekly KPI snapshot log (shipped 2026-08-05)
 * ============================================================================
 *
 * WHY THIS EXISTS
 * ---------------
 * Every other number in this system is a LIVE READ — Kit Health, OOS, Photo
 * Queue and the price audits all describe "right now." That makes the weekly
 * digest and the PDF report a photograph, never a trend. And a trend is the
 * one thing you CANNOT backfill: every week that passes without a snapshot is
 * a week of history that is gone permanently.
 *
 * So this file does one small job — it appends the digest's own KPI numbers to
 * a "KPI History" sheet each time a digest is sent, building the baseline that
 * "▲ 12 vs last week" arrows are computed from.
 *
 * DESIGN RULES
 * ------------
 *  1. PURE READER + APPENDER. This file never recomputes anything. It records
 *     exactly the object `_gatherWeeklyDigest()` already built, so the history
 *     can never disagree with the digest that was sent alongside it.
 *  2. BEST-EFFORT, ALWAYS. Every entry point is wrapped so a logging failure
 *     can never block the digest or the report. A missing history sheet simply
 *     means "no trend yet," never an error.
 *  3. ONE ROW PER CALENDAR DAY (upsert, not blind append). The Monday trigger
 *     and a manual "Send now" on the same day must not produce two rows. Same
 *     day → the existing row is updated in place.
 *  4. TREND IS COMPARED AGAINST THE PREVIOUS *ROW*, not "7 days ago" — and the
 *     comparison reports how many days back that row actually was. Honest by
 *     construction: if a week was missed, the label says so rather than
 *     silently pretending the gap was seven days.
 *
 * ENTRY POINTS
 * ------------
 *   setupKpiHistorySheet()    create / re-style the sheet (idempotent)
 *   recordKpiSnapshot(d, src) append-or-update today's row  (called by the digest)
 *   recordKpiSnapshotNow()    manual snapshot (sidebar / editor)
 *   getKpiTrend(current)      deltas vs the previous snapshot, or null
 *   getKpiHistoryCount()      how many snapshots we hold
 *   openKpiHistory()          activate the sheet
 */

var KPI_HISTORY = {
  sheetName: "KPI History",

  cols: {
    SNAPSHOT:       1,   // A — date the snapshot was taken (Houston time)
    WEEK:           2,   // B — ISO-ish "yyyy-Www" label, handy for grouping
    TOTAL_KITS:     3,   // C
    UNDERPRICED:    4,   // D
    ON_TABLE:       5,   // E — $ sum of |Δ| across underpriced kits
    BUILDABLE:      6,   // F
    BLOCKED:        7,   // G — in stock but buildable 0
    OUT_OF_STOCK:   8,   // H
    NEEDS_PHOTOS:   9,   // I
    OPEN_CASES:     10,  // J
    PRICE_DRIFT:    11,  // K
    SOURCE:         12   // L — weekly-send / manual / etc.
  },

  idx: function (name) { return KPI_HISTORY.cols[name] - 1; },

  dataWidth:    12,
  headerRow:    1,
  dataStartRow: 2,

  headers: ["📅 SNAPSHOT", "WEEK", "TOTAL KITS", "UNDERPRICED", "$ ON TABLE",
            "BUILDABLE", "BLOCKED", "OUT OF STOCK", "NEEDS PHOTOS",
            "OPEN CASES", "PRICE DRIFT", "SOURCE"],

  // Which digest field feeds which column. Single source of truth for both the
  // writer and the trend reader, so adding a KPI later means editing ONE list.
  fieldMap: [
    { col: "TOTAL_KITS",   field: "totalKits",          label: "kits",         goodWhenUp: true  },
    { col: "UNDERPRICED",  field: "underpriced",        label: "underpriced",  goodWhenUp: false },
    { col: "ON_TABLE",     field: "underpricedDollars", label: "on the table", goodWhenUp: false },
    { col: "BUILDABLE",    field: "buildableNow",       label: "buildable",    goodWhenUp: true  },
    { col: "BLOCKED",      field: "inStockBlocked",     label: "blocked",      goodWhenUp: false },
    { col: "OUT_OF_STOCK", field: "outOfStock",         label: "out of stock", goodWhenUp: false },
    { col: "NEEDS_PHOTOS", field: "needPhotos",         label: "needs photos", goodWhenUp: false },
    { col: "OPEN_CASES",   field: "openCases",          label: "open cases",   goodWhenUp: false },
    { col: "PRICE_DRIFT",  field: "priceDrift",         label: "price drift",  goodWhenUp: false }
  ]
};


// =======================================================================================
// SHEET SETUP
// =======================================================================================

/**
 * Create (or re-style) the KPI History sheet. Idempotent — safe to re-run; it
 * never touches existing data rows, only the header band and column formats.
 */
function setupKpiHistorySheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(KPI_HISTORY.sheetName);
  if (!sheet) sheet = ss.insertSheet(KPI_HISTORY.sheetName);

  var W = KPI_HISTORY.dataWidth;

  // --- header band ---
  var header = sheet.getRange(KPI_HISTORY.headerRow, 1, 1, W);
  header.setValues([KPI_HISTORY.headers]);
  header.setBackground(BRAND.ink)
        .setFontColor("#ffffff")
        .setFontFamily(BRAND.fontDisplay)
        .setFontWeight("bold")
        .setFontSize(10)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
  sheet.setRowHeight(KPI_HISTORY.headerRow, 34);
  sheet.setFrozenRows(KPI_HISTORY.headerRow);

  // --- column formats (applied generously so growth never escapes them) ---
  var lastFmtRow = Math.max(sheet.getMaxRows(), 500);
  var nFmt = lastFmtRow - KPI_HISTORY.dataStartRow + 1;
  if (nFmt > 0) {
    var body = sheet.getRange(KPI_HISTORY.dataStartRow, 1, nFmt, W);
    body.setFontFamily(BRAND.fontMono)
        .setFontSize(10)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");

    sheet.getRange(KPI_HISTORY.dataStartRow, KPI_HISTORY.cols.SNAPSHOT, nFmt, 1)
         .setNumberFormat("m/d/yy");
    sheet.getRange(KPI_HISTORY.dataStartRow, KPI_HISTORY.cols.ON_TABLE, nFmt, 1)
         .setNumberFormat("$#,##0");
    sheet.getRange(KPI_HISTORY.dataStartRow, KPI_HISTORY.cols.SOURCE, nFmt, 1)
         .setFontFamily(BRAND.fontData)
         .setFontColor(BRAND.inkSoft)
         .setHorizontalAlignment("left");
  }

  // --- widths ---
  sheet.setColumnWidth(KPI_HISTORY.cols.SNAPSHOT, 95);
  sheet.setColumnWidth(KPI_HISTORY.cols.WEEK, 85);
  sheet.setColumnWidth(KPI_HISTORY.cols.ON_TABLE, 100);
  sheet.setColumnWidth(KPI_HISTORY.cols.SOURCE, 130);

  return "KPI History sheet ready.";
}


// =======================================================================================
// WRITE
// =======================================================================================

/** Houston-time date key (yyyy-MM-dd) — the upsert identity for a snapshot. */
function _kpiDateKey(date) {
  return Utilities.formatDate(date || new Date(), WEEKLY_DIGEST.timezone, "yyyy-MM-dd");
}

/** Houston-time week label (yyyy-Www) for at-a-glance grouping. */
function _kpiWeekKey(date) {
  return Utilities.formatDate(date || new Date(), WEEKLY_DIGEST.timezone, "yyyy-'W'ww");
}

/**
 * Record one KPI snapshot. Appends a new row, or UPDATES today's row if one
 * already exists (so a manual send + the Monday trigger on the same day give
 * one row, not two).
 *
 * BEST-EFFORT: returns a result object and never throws — the digest must send
 * even if this fails.
 *
 * @param {Object} d       the object returned by _gatherWeeklyDigest()
 * @param {string} source  provenance label, e.g. "weekly-send" / "manual"
 * @returns {{ok:boolean, action:string, row:number, message:string}}
 */
function recordKpiSnapshot(d, source) {
  try {
    if (!d) return { ok: false, action: "none", row: 0, message: "No data to record." };

    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(KPI_HISTORY.sheetName);
    if (!sheet) {
      setupKpiHistorySheet();
      sheet = ss.getSheetByName(KPI_HISTORY.sheetName);
      if (!sheet) return { ok: false, action: "none", row: 0, message: "Could not create KPI History sheet." };
    }

    var now  = new Date();
    var key  = _kpiDateKey(now);
    var W    = KPI_HISTORY.dataWidth;

    // Build the row from the field map so the schema has one definition.
    var row = new Array(W).fill("");
    row[KPI_HISTORY.idx("SNAPSHOT")] = now;
    row[KPI_HISTORY.idx("WEEK")]     = _kpiWeekKey(now);
    row[KPI_HISTORY.idx("SOURCE")]   = String(source || "manual");
    for (var i = 0; i < KPI_HISTORY.fieldMap.length; i++) {
      var m = KPI_HISTORY.fieldMap[i];
      var v = d[m.field];
      row[KPI_HISTORY.idx(m.col)] = (typeof v === 'number' && !isNaN(v)) ? v : 0;
    }

    // --- upsert by date key ---
    var existingRow = _kpiFindRowByDateKey(sheet, key);
    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, W).setValues([row]);
      return { ok: true, action: "updated", row: existingRow,
               message: "KPI snapshot updated for " + key + "." };
    }

    var target = Math.max(sheet.getLastRow() + 1, KPI_HISTORY.dataStartRow);
    sheet.getRange(target, 1, 1, W).setValues([row]);
    return { ok: true, action: "appended", row: target,
             message: "KPI snapshot recorded for " + key + "." };

  } catch (e) {
    try { console.log("recordKpiSnapshot: " + e); } catch (_) {}
    return { ok: false, action: "none", row: 0, message: String(e) };
  }
}

/** Find the row whose SNAPSHOT falls on the given Houston date key, else 0. */
function _kpiFindRowByDateKey(sheet, key) {
  var last = sheet.getLastRow();
  if (last < KPI_HISTORY.dataStartRow) return 0;
  var n = last - KPI_HISTORY.dataStartRow + 1;
  var vals = sheet.getRange(KPI_HISTORY.dataStartRow, KPI_HISTORY.cols.SNAPSHOT, n, 1).getValues();
  for (var i = 0; i < n; i++) {
    var v = vals[i][0];
    if (!v) continue;
    var k = (v instanceof Date) ? _kpiDateKey(v) : String(v).trim();
    if (k === key) return KPI_HISTORY.dataStartRow + i;
  }
  return 0;
}

/**
 * MANUAL snapshot — gathers fresh digest numbers and records them. Safe to run
 * from the editor or a sidebar button; the daily upsert means repeat clicks
 * refresh today's row rather than piling up duplicates.
 */
function recordKpiSnapshotNow() {
  var d = _gatherWeeklyDigest();
  var r = recordKpiSnapshot(d, "manual");
  try { console.log(r.message); } catch (_) {}
  return r;
}


// =======================================================================================
// READ / TREND
// =======================================================================================

/** Number of snapshots currently held. */
function getKpiHistoryCount() {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName(KPI_HISTORY.sheetName);
    if (!sheet) return 0;
    return Math.max(0, sheet.getLastRow() - KPI_HISTORY.dataStartRow + 1);
  } catch (e) { return 0; }
}

/**
 * Deltas between `current` and the most recent snapshot taken on an EARLIER
 * day. Returns null when there is no prior snapshot to compare against (i.e.
 * the first ever run) — callers should treat null as "no trend yet."
 *
 * @param {Object} current  a _gatherWeeklyDigest() object
 * @returns {?{daysAgo:number, deltas:Object, prev:Object}}
 */
function getKpiTrend(current) {
  try {
    if (!current) return null;
    var sheet = SpreadsheetApp.getActive().getSheetByName(KPI_HISTORY.sheetName);
    if (!sheet) return null;

    var last = sheet.getLastRow();
    if (last < KPI_HISTORY.dataStartRow) return null;

    var n = last - KPI_HISTORY.dataStartRow + 1;
    var rows = sheet.getRange(KPI_HISTORY.dataStartRow, 1, n, KPI_HISTORY.dataWidth).getValues();
    var todayKey = _kpiDateKey(new Date());

    // Walk backwards to the newest row that is NOT from today.
    for (var i = rows.length - 1; i >= 0; i--) {
      var stamp = rows[i][KPI_HISTORY.idx("SNAPSHOT")];
      if (!(stamp instanceof Date)) continue;
      if (_kpiDateKey(stamp) === todayKey) continue;

      var deltas = {}, prev = {};
      for (var j = 0; j < KPI_HISTORY.fieldMap.length; j++) {
        var m = KPI_HISTORY.fieldMap[j];
        var was = rows[i][KPI_HISTORY.idx(m.col)];
        was = (typeof was === 'number' && !isNaN(was)) ? was : 0;
        var isNow = current[m.field];
        isNow = (typeof isNow === 'number' && !isNaN(isNow)) ? isNow : 0;
        prev[m.field]   = was;
        deltas[m.field] = isNow - was;
      }

      var days = Math.round((new Date().getTime() - stamp.getTime()) / 86400000);
      return { daysAgo: Math.max(1, days), deltas: deltas, prev: prev };
    }
    return null;
  } catch (e) {
    try { console.log("getKpiTrend: " + e); } catch (_) {}
    return null;
  }
}

/**
 * Render one trend fragment, e.g. "  ▲ 12" / "  ▼ 3" / "" when unchanged.
 * PURE — no I/O, so it stays Node-testable alongside the digest builder.
 *
 * @param {?Object} trend  a getKpiTrend() result (null = no history yet)
 * @param {string}  field  digest field name
 * @param {string}  [pre]  optional prefix for the number (e.g. "$")
 */
function _kpiTrendFragment(trend, field, pre) {
  if (!trend || !trend.deltas) return "";
  var delta = trend.deltas[field];
  if (typeof delta !== 'number' || isNaN(delta) || delta === 0) return "";
  var arrow = delta > 0 ? "▲" : "▼";
  return "  " + arrow + " " + (pre || "") + Math.abs(delta);
}

/** Activate the KPI History sheet. */
function openKpiHistory() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(KPI_HISTORY.sheetName);
  if (!sheet) { setupKpiHistorySheet(); sheet = ss.getSheetByName(KPI_HISTORY.sheetName); }
  if (sheet) ss.setActiveSheet(sheet);
  return "KPI History opened.";
}
