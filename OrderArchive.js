/**
 * OrderArchive.js — THE PERMANENT OPERATIONAL RECORD
 * =======================================================================================
 *
 * WHY THIS EXISTS
 * ---------------
 * purgeOldActivityLog() deletes every log row older than ACTIVITY_LOG.retentionDays
 * (90) and keeps nothing. Some of what it destroys could be painfully rebuilt from
 * eBay or Zoho. The parts that matter cannot:
 *
 *     WHO PICKED IT           — nowhere else, ever
 *     YOUR-CLOCK CYCLE TIME   — eBay's timestamp is when the buyer PAID, not when
 *                               the order hit your floor
 *     NOTE events             — what supervisors actually wrote
 *     HOLD lifecycle          — held → seen → escalated
 *     PRINTED                 — nowhere else
 *
 * So the thing being deleted nightly is the OPERATIONAL record, which is the only
 * part that was ever ours. This module rolls each completed order into one durable
 * row before that happens.
 *
 * ⚠ IT IS ALSO WHY NOTHING HAS EVER MEASURED CYCLE TIME. The only time math in this
 *   codebase is oldestPendingMinutes / pastRedlineCount / orderAgeMin — all
 *   PRESENT-TENSE age of still-open orders. And KPI_HISTORY's nine weekly fields are
 *   all catalogue health, zero throughput, so the digest cannot answer "did we ship
 *   more this week than last."
 *
 *
 * ⭐⭐ THE DESIGN DECISION EVERYTHING ELSE FALLS OUT OF
 * ----------------------------------------------------
 * THE ARCHIVE IS A PURE FUNCTION OF THE ACTIVITY LOG. Every field is already logged:
 *
 *     lines / units    count and sum of the order's RECEIVED events
 *     received         earliest RECEIVED timestamp
 *     first pick       earliest PREPARING timestamp
 *     printed          earliest PRINTED timestamp
 *     terminal         latest SHIPPED | CANCELED timestamp
 *     picker           the PICKER column already on each event
 *     held             a NOTE event carrying the hold word
 *     kit              a RECEIVED whose DETAIL says "kit expansion from …"
 *     channel          _floorOrderIsDirect(orderId) — pure, id-shape only
 *
 * That buys three things at once:
 *
 *   1. NO HOOK IN updateOrderStatus. Zero risk to the floor's hot write path. This
 *      project's own ruling is not to extend lock time on ✓ Pick, and a per-transition
 *      archive write would ALSO mean one Activity Log read per flipped order — n8n's
 *      shipped sweep flips them in batches. A batched sweep reads the log ONCE.
 *   2. NO SCHEMA CHANGE TO All Orders. No dataWidth 10 → 11, nothing to migrate.
 *   3. THE BACKFILL AND THE ONGOING JOB ARE THE SAME CODE over different date ranges,
 *      and the core is pure, so it is Node-testable against fixture rows.
 *
 *
 * ⭐ TWO INTERVALS, NOT ONE — the point of the whole thing
 * -------------------------------------------------------
 *     QUEUE_MIN  received → someone started   (how long before anyone touched it)
 *     EXEC_MIN   started  → terminal          (how long the work itself took)
 *
 * Most warehouses discover the problem is entirely the first, and the two numbers
 * point at completely different fixes — staffing and triggering vs process and
 * layout. Today you cannot tell which you have.
 *
 * ⚠ WHAT "CYCLE TIME" MEANS HERE, so nobody misreads it later: time on OUR SHEET.
 *   It deliberately excludes eBay's polling lag and the gap between purchase and
 *   arrival, because both are upstream of anything the floor controls. It is a
 *   measure of us, not of the marketplace.
 *
 *
 * PUBLIC API
 * ----------
 *   setupOrderArchiveSheet()     — idempotent create + style
 *   runOrderArchiveSweep()       — the daily job (rides _housekeepingPass)
 *   backfillOrderArchive()       — one-shot, resumable, CONSOLE-LOGGING
 *   getOrderArchiveStatus()      — {watermark, rows, lastDay} for the sidebar
 *   openOrderArchive()           — sidebar "Open"
 *   _oaRollUpOrders(events,dayOf)— THE PURE CORE (exported for tests)
 */

// ---------- LOCAL SCHEMA ----------
// ⚠ Duplicate-global scan run before adding this (grep -rhoP '^var \w+' | sort | uniq -d
//   → empty). Apps Script puts every root .js file in ONE global scope, and this
//   project has already been bitten once: Budget.js had to declare `PROBE` because
//   Helpers.js already owned `BUDGET`.
var ORDER_ARCHIVE = {
  sheetName: "Order Archive",

  cols: {
    ORDER_ID:        1,   // A
    CHANNEL:         2,   // B — eBay | DIRECT
    DAY:             3,   // C — America/Chicago yyyy-MM-dd of the TERMINAL event
    RECEIVED:        4,   // D
    FIRST_PICK:      5,   // E
    PRINTED:         6,   // F
    TERMINAL_AT:     7,   // G
    TERMINAL_STATUS: 8,   // H — SHIPPED | CANCELED
    QUEUE_MIN:       9,   // I
    EXEC_MIN:       10,   // J
    TOTAL_MIN:      11,   // K
    LINES:          12,   // L
    UNITS:          13,   // M
    PICKER:         14,   // N
    HELD:           15,   // O — "yes" | ""
    KIT:            16    // P — "yes" | ""
  },

  idx: function (name) { return ORDER_ARCHIVE.cols[name] - 1; },

  dataWidth: 16,
  bannerRow: 1,
  headerRow: 2,
  dataStartRow: 3,

  headers: [
    "ORDER ID", "CHANNEL", "DAY", "⏱ RECEIVED", "FIRST PICK", "PRINTED",
    "TERMINAL AT", "STATUS", "QUEUE MIN", "EXEC MIN", "TOTAL MIN",
    "# LINES", "# UNITS", "👤 PICKER", "HELD", "KIT"
  ],

  // Script Property holding the last FULLY archived Chicago day (yyyy-MM-dd).
  // ⚠ This is an OPTIMISATION, not the correctness mechanism — it lets 11 of the
  //   12 hourly housekeeping runs return without reading anything. Correctness
  //   comes from the ORDER_ID dedupe in _oaExistingOrderIds().
  watermarkKey: "ORDER_ARCHIVE_WATERMARK",

  // Backfill processes at most this many days per invocation, then asks to be
  // re-run. Keeps a 90-day catch-up inside the 6-minute execution ceiling.
  backfillChunkDays: 45
};


// =======================================================================================
// THE PURE CORE  — no SpreadsheetApp, no Utilities, no clock. Node-testable.
// =======================================================================================

/**
 * Minutes between two Dates, or "" when either is missing.
 *
 * ⚠ A NEGATIVE GAP RETURNS "" RATHER THAN A NEGATIVE NUMBER. Events can land out of
 *   order (a manual edit backdated, a revert, clock skew between writers). A blank
 *   cell reads as "unknown", which is true; a negative duration reads as data anyone
 *   downstream will average into nonsense.
 */
function _oaMinutesBetween(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date)) return "";
  var m = Math.round((b.getTime() - a.getTime()) / 60000);
  return (m < 0) ? "" : m;
}

function _oaEarlier(a, b) {
  if (!(a instanceof Date)) return b;
  if (!(b instanceof Date)) return a;
  return (a.getTime() <= b.getTime()) ? a : b;
}

function _oaLater(a, b) {
  if (!(a instanceof Date)) return b;
  if (!(b instanceof Date)) return a;
  return (a.getTime() >= b.getTime()) ? a : b;
}

/**
 * ⭐ THE ROLL-UP. Events in → one row object per COMPLETED order out.
 *
 * @param events  [{ts:Date, event, orderId, sku, qty, detail, note, picker}]
 * @param dayOf   Date -> "yyyy-MM-dd". Injected so this stays pure and testable;
 *                production passes the America/Chicago formatter.
 * @return        [{orderId, channel, day, received, firstPick, printed,
 *                  terminalAt, terminalStatus, queueMin, execMin, totalMin,
 *                  lines, units, picker, held, kit}]
 *
 * ⚠ ONLY ORDERS WITH A TERMINAL EVENT ARE EMITTED. "Completed" is the archive's
 *   entry condition — an order still being worked has no cycle time yet, and
 *   writing a half-row we would have to update later is exactly the reconciler
 *   shape this project deleted from Zoho propagation on 2026-05-23.
 */
function _oaRollUpOrders(events, dayOf) {
  var byOrder = {};
  var order   = [];   // preserves first-seen order so output is deterministic

  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    if (!ev) continue;

    var idRaw = String(ev.orderId == null ? "" : ev.orderId).trim();
    if (!idRaw) continue;                       // NOTE/PRINTED rows can be orderless
    var key = idRaw.toUpperCase();

    if (!byOrder[key]) {
      byOrder[key] = {
        orderId: idRaw, received: null, firstPick: null, printed: null,
        terminalAt: null, terminalStatus: "", lines: 0, units: 0,
        pickerAtTerminal: "", pickerAny: "", held: false, kit: false
      };
      order.push(key);
    }
    var o  = byOrder[key];
    var ts = (ev.ts instanceof Date) ? ev.ts : null;
    var evt = String(ev.event == null ? "" : ev.event).trim().toUpperCase();

    var picker = String(ev.picker == null ? "" : ev.picker).trim();
    if (picker && !o.pickerAny) o.pickerAny = picker;

    if (evt === "RECEIVED") {
      o.received = _oaEarlier(o.received, ts);
      o.lines++;
      var q = parseFloat(ev.qty);
      if (!isNaN(q)) o.units += q;

      // ⚠ KIT DETECTION READS THE DETAIL, NOT THE NOTE. KitExpansion.js:1037 builds
      //   `"kit expansion from " + rowSku` as the Activity Log DETAIL for every
      //   inserted component row. The NOTE carries the "↳ from KIT-" tag, but the
      //   note is also where buyer text, Zoho flags and hold text live — the detail
      //   is written by exactly one place and means exactly one thing.
      if (String(ev.detail || "").toLowerCase().indexOf("kit expansion from") !== -1) {
        o.kit = true;
      }
    }
    else if (evt === "PREPARING") {
      o.firstPick = _oaEarlier(o.firstPick, ts);
    }
    else if (evt === "PRINTED") {
      o.printed = _oaEarlier(o.printed, ts);
    }
    else if (evt === "SHIPPED" || evt === "CANCELED") {
      // LATEST terminal wins — an order reverted by the n8n verify sweep and then
      // re-shipped should report the second one.
      var prev = o.terminalAt;
      o.terminalAt = _oaLater(o.terminalAt, ts);
      if (o.terminalAt !== prev || !o.terminalStatus) {
        o.terminalStatus = evt;
        if (picker) o.pickerAtTerminal = picker;
      }
    }
    else if (evt === "NOTE") {
      // ⚠ Same word-boundary rule as Holds.holdNoteHasHold, deliberately — the hold
      //   is recognised by the whole word HOLD anywhere in the note, any case, so
      //   "household" and "withhold" do not fire.
      var body = String(ev.note || "") + " " + String(ev.detail || "");
      if (/\bHOLD\b/i.test(body)) o.held = true;
    }
  }

  var out = [];
  for (var k = 0; k < order.length; k++) {
    var r = byOrder[order[k]];
    if (!(r.terminalAt instanceof Date)) continue;      // not finished → not archived

    // "Started" is whichever of first-pick / print came first. PRINTED can precede
    // PREPARING (print the list, then walk) or follow it; either marks the moment a
    // human began.
    var started = _oaEarlier(r.firstPick, r.printed);

    // ⚠ NO RECEIVED means the order's arrival was already purged (open >90 days) —
    //   or it was never logged. The row is still WRITTEN, with blank intervals, so
    //   the order is not silently lost. Blank self-documents as "unknown"; a zero
    //   would be a reassuring label on a state we cannot actually see.
    var hasReceived = (r.received instanceof Date);

    out.push({
      orderId:        r.orderId,
      channel:        _oaChannelOf(r.orderId),
      day:            dayOf(r.terminalAt),
      received:       hasReceived ? r.received : "",
      firstPick:      (r.firstPick instanceof Date) ? r.firstPick : "",
      printed:        (r.printed   instanceof Date) ? r.printed   : "",
      terminalAt:     r.terminalAt,
      terminalStatus: r.terminalStatus,
      queueMin:       hasReceived ? _oaMinutesBetween(r.received, started || r.terminalAt) : "",
      execMin:        started ? _oaMinutesBetween(started, r.terminalAt) : "",
      totalMin:       hasReceived ? _oaMinutesBetween(r.received, r.terminalAt) : "",
      lines:          r.lines > 0 ? r.lines : "",
      units:          r.lines > 0 ? r.units : "",
      picker:         r.pickerAtTerminal || r.pickerAny || "",
      held:           r.held ? "yes" : "",
      kit:            r.kit  ? "yes" : ""
    });
  }
  return out;
}

/**
 * eBay | DIRECT from the order-id shape.
 *
 * Delegates to _floorOrderIsDirect (ActivityLog.js) when it is present so the two
 * can never disagree, and carries an identical local fallback for the Node harness,
 * which loads this file alone.
 */
function _oaChannelOf(orderId) {
  if (typeof _floorOrderIsDirect === "function") {
    return _floorOrderIsDirect(orderId) ? "DIRECT" : "eBay";
  }
  var u = String(orderId || "").trim().toUpperCase();
  if (!u) return "eBay";
  if (u.indexOf("SO-") === 0 || u.indexOf("INV-") === 0) return "DIRECT";
  if (/^[0-9][0-9\-]+$/.test(u)) return "eBay";
  return "DIRECT";
}

/** Row object → sheet row array, in schema order. */
function _oaToSheetRow(o) {
  var row = new Array(ORDER_ARCHIVE.dataWidth).fill("");
  row[ORDER_ARCHIVE.idx("ORDER_ID")]        = o.orderId;
  row[ORDER_ARCHIVE.idx("CHANNEL")]         = o.channel;
  row[ORDER_ARCHIVE.idx("DAY")]             = o.day;
  row[ORDER_ARCHIVE.idx("RECEIVED")]        = o.received;
  row[ORDER_ARCHIVE.idx("FIRST_PICK")]      = o.firstPick;
  row[ORDER_ARCHIVE.idx("PRINTED")]         = o.printed;
  row[ORDER_ARCHIVE.idx("TERMINAL_AT")]     = o.terminalAt;
  row[ORDER_ARCHIVE.idx("TERMINAL_STATUS")] = o.terminalStatus;
  row[ORDER_ARCHIVE.idx("QUEUE_MIN")]       = o.queueMin;
  row[ORDER_ARCHIVE.idx("EXEC_MIN")]        = o.execMin;
  row[ORDER_ARCHIVE.idx("TOTAL_MIN")]       = o.totalMin;
  row[ORDER_ARCHIVE.idx("LINES")]           = o.lines;
  row[ORDER_ARCHIVE.idx("UNITS")]           = o.units;
  row[ORDER_ARCHIVE.idx("PICKER")]          = o.picker;
  row[ORDER_ARCHIVE.idx("HELD")]            = o.held;
  row[ORDER_ARCHIVE.idx("KIT")]             = o.kit;
  return row;
}


// =======================================================================================
// DAY KEYS  — America/Chicago, because that is the shop's day
// =======================================================================================

/**
 * ⚠ CHICAGO, VIA Utilities.formatDate — NOT raw millisecond arithmetic.
 *
 * The script timezone is Asia/Amman (appsscript.json), 8 hours from Houston, and
 * mixing the two has already produced one false three-hour outage report in this
 * project. getDashboardSnapshot settled on exactly this pattern for the same reason:
 * comparing formatted yyyy-MM-dd strings is DST-safe, offset arithmetic is not.
 */
function _oaDayKey(d) {
  return Utilities.formatDate(d, "America/Chicago", "yyyy-MM-dd");
}

/** Shift a yyyy-MM-dd label by n days. Noon-UTC anchor so no DST hop can flip it. */
function _oaDayAdd(dayStr, n) {
  var p = String(dayStr).split("-");
  var dt = new Date(Date.UTC(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10), 12, 0, 0));
  dt.setUTCDate(dt.getUTCDate() + n);
  return dt.toISOString().slice(0, 10);
}


// =======================================================================================
// SHEET I/O
// =======================================================================================

/** Read the whole Activity Log into plain event objects for the pure core. */
function _oaReadLogEvents() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ACTIVITY_LOG.sheetName);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < ACTIVITY_LOG.dataStartRow) return [];

  // ⚠ ONE read of the whole log, deliberately. The alternative — a read per order —
  //   is what makes a per-transition hook untenable: n8n's shipped sweep flips orders
  //   in batches. The log is bounded at 90 days by purgeOldActivityLog, so this is a
  //   bounded read, and runOrderArchiveSweep only reaches it once a day.
  var data = sheet.getRange(
    ACTIVITY_LOG.dataStartRow, 1,
    lastRow - ACTIVITY_LOG.dataStartRow + 1,
    ACTIVITY_LOG.dataWidth
  ).getValues();

  var out = [];
  for (var i = 0; i < data.length; i++) {
    var ts = data[i][ACTIVITY_LOG.idx("TIMESTAMP")];
    if (!(ts instanceof Date)) continue;          // header leak / blank row
    out.push({
      ts:      ts,
      event:   data[i][ACTIVITY_LOG.idx("EVENT")],
      orderId: data[i][ACTIVITY_LOG.idx("ORDER_ID")],
      sku:     data[i][ACTIVITY_LOG.idx("SKU")],
      qty:     data[i][ACTIVITY_LOG.idx("QTY")],
      detail:  data[i][ACTIVITY_LOG.idx("DETAIL")],
      note:    data[i][ACTIVITY_LOG.idx("NOTE")],
      picker:  data[i][ACTIVITY_LOG.idx("PICKER")]
    });
  }
  return out;
}

/**
 * Order ids already archived.
 *
 * ⭐ THIS, NOT THE WATERMARK, IS THE CORRECTNESS MECHANISM. The watermark only
 *    decides whether the expensive log read is worth doing; this decides what gets
 *    written. So an interrupted backfill, a re-run, or a hand-cleared property can
 *    never produce duplicate rows.
 */
function _oaExistingOrderIds(sheet) {
  var seen = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < ORDER_ARCHIVE.dataStartRow) return seen;

  var ids = sheet.getRange(
    ORDER_ARCHIVE.dataStartRow, ORDER_ARCHIVE.cols.ORDER_ID,
    lastRow - ORDER_ARCHIVE.dataStartRow + 1, 1
  ).getValues();

  for (var i = 0; i < ids.length; i++) {
    var v = String(ids[i][0] || "").trim().toUpperCase();
    if (v) seen[v] = true;
  }
  return seen;
}


// =======================================================================================
// THE SHARED ENGINE  — the sweep and the backfill are the same code
// =======================================================================================

/**
 * Archive every completed order whose TERMINAL event fell in [fromDay .. toDay].
 *
 * @param fromDay  yyyy-MM-dd, or null for "the earliest day present in the log"
 * @param toDay    yyyy-MM-dd (inclusive)
 * @param maxDays  cap on the span processed in one invocation
 * @return {ok, written, skipped, from, to, more, message}
 */
function _oaProcess(fromDay, toDay, maxDays) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ORDER_ARCHIVE.sheetName);
  if (!sheet) { setupOrderArchiveSheet(); sheet = ss.getSheetByName(ORDER_ARCHIVE.sheetName); }

  var events = _oaReadLogEvents();
  if (!events.length) {
    return { ok: true, written: 0, skipped: 0, from: fromDay, to: toDay, more: false,
             message: "Activity Log is empty — nothing to archive." };
  }

  var all = _oaRollUpOrders(events, _oaDayKey);
  if (!all.length) {
    return { ok: true, written: 0, skipped: 0, from: fromDay, to: toDay, more: false,
             message: "No completed orders in the log." };
  }

  // Earliest completed day present, when we have no watermark to start from.
  if (!fromDay) {
    fromDay = all[0].day;
    for (var a = 1; a < all.length; a++) if (all[a].day < fromDay) fromDay = all[a].day;
  }
  if (fromDay > toDay) {
    return { ok: true, written: 0, skipped: 0, from: fromDay, to: toDay, more: false,
             message: "Nothing new (through " + toDay + ")." };
  }

  // Cap the span so one invocation stays inside the execution ceiling.
  var effTo = toDay, more = false;
  var capped = _oaDayAdd(fromDay, maxDays - 1);
  if (capped < toDay) { effTo = capped; more = true; }

  var seen = _oaExistingOrderIds(sheet);
  var rows = [], skipped = 0;

  for (var i = 0; i < all.length; i++) {
    var o = all[i];
    if (o.day < fromDay || o.day > effTo) continue;
    if (seen[String(o.orderId).trim().toUpperCase()]) { skipped++; continue; }
    rows.push(_oaToSheetRow(o));
  }

  if (rows.length) {
    var start = Math.max(sheet.getLastRow() + 1, ORDER_ARCHIVE.dataStartRow);
    var need  = start + rows.length - 1;
    if (need > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows() + 50);
    var range = sheet.getRange(start, 1, rows.length, ORDER_ARCHIVE.dataWidth);
    range.setValues(rows);
    _oaApplyDataFormats(sheet, start, rows.length);   // ⚠ see Gotcha #16
  }

  PropertiesService.getScriptProperties().setProperty(ORDER_ARCHIVE.watermarkKey, effTo);
  _oaPaintStatus(sheet, effTo);

  return {
    ok: true, written: rows.length, skipped: skipped,
    from: fromDay, to: effTo, more: more,
    message: "Archived " + rows.length + " order(s) for " + fromDay + " → " + effTo +
             (skipped ? " · " + skipped + " already present" : "") +
             (more ? " · MORE REMAINS, run again" : "")
  };
}


// =======================================================================================
// PUBLIC ENTRY POINTS
// =======================================================================================

/**
 * The daily job. Rides _housekeepingPass as an isolated 5th job.
 *
 * ⭐ CHEAP ON 11 OF 12 HOURLY RUNS. Once the watermark reaches yesterday there is
 *    nothing to do, and this returns after ONE Script Property read — no sheet
 *    access at all. The expensive log read happens at most once a day.
 *
 * ⚠ ONLY COMPLETED DAYS ARE ARCHIVED, so an order finishing today is archived
 *   tomorrow. This is analysis data, not live: a day of latency costs nothing and
 *   buys a clean, re-runnable day boundary. The known cost of that choice is that a
 *   terminal event REVERTED after its day was archived keeps the original row — rare
 *   (the n8n verify sweep reverts within hours) and visible, since the order simply
 *   reappears as open on the sheet.
 */
function runOrderArchiveSweep() {
  try {
    var wm = PropertiesService.getScriptProperties().getProperty(ORDER_ARCHIVE.watermarkKey) || "";
    var yesterday = _oaDayAdd(_oaDayKey(new Date()), -1);

    if (wm && wm >= yesterday) return "⏸ Archive current (through " + wm + ")";

    var res = _oaProcess(wm ? _oaDayAdd(wm, 1) : null, yesterday, ORDER_ARCHIVE.backfillChunkDays);
    return "🗄 " + res.message;
  } catch (e) {
    console.log("runOrderArchiveSweep error: " + e);
    return "❌ Order archive: " + e;
  }
}

/**
 * One-shot catch-up over everything still in the Activity Log. Resumable — run it
 * again until it says DONE.
 *
 * ⚠⚠ IT LOGS RATHER THAN RETURNS, AND THAT IS DELIBERATE. The Apps Script editor's
 *    Run button does not display return values. This project has already lost two
 *    separate mornings to that on getPublishedTick, and the second time proved a
 *    note telling the next person to remember does not work — hence checkPublishedTickNow.
 *    Read the EXECUTION LOG, not the return value.
 */
function backfillOrderArchive() {
  var yesterday = _oaDayAdd(_oaDayKey(new Date()), -1);
  var wm = PropertiesService.getScriptProperties().getProperty(ORDER_ARCHIVE.watermarkKey) || "";

  console.log("=== ORDER ARCHIVE BACKFILL ===");
  console.log("watermark : " + (wm || "(none — starting from the oldest log day)"));
  console.log("through   : " + yesterday + "  (completed days only)");

  var res = _oaProcess(wm ? _oaDayAdd(wm, 1) : null, yesterday, ORDER_ARCHIVE.backfillChunkDays);

  console.log("range     : " + res.from + " → " + res.to);
  console.log("written   : " + res.written);
  console.log("skipped   : " + res.skipped + " (already archived)");
  console.log(res.more
    ? "⏭ MORE REMAINS — run backfillOrderArchive() again."
    : "✅ DONE — the archive is caught up through " + res.to + ".");
  return res.message;
}

/** {watermark, rows, lastDay} — for the sidebar card. */
function getOrderArchiveStatus() {
  var out = { watermark: "", rows: 0, exists: false };
  try {
    out.watermark = PropertiesService.getScriptProperties()
                      .getProperty(ORDER_ARCHIVE.watermarkKey) || "";
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDER_ARCHIVE.sheetName);
    if (!sheet) return out;
    out.exists = true;
    out.rows = Math.max(0, sheet.getLastRow() - ORDER_ARCHIVE.dataStartRow + 1);
  } catch (e) {
    console.log("getOrderArchiveStatus error: " + e);
  }
  return out;
}

/** Sidebar "Open" — ⚠ getActive(), never openById (the v1 openPrepQueue trap). */
function openOrderArchive() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(ORDER_ARCHIVE.sheetName);
  if (!sheet) { setupOrderArchiveSheet(); sheet = ss.getSheetByName(ORDER_ARCHIVE.sheetName); }
  ss.setActiveSheet(sheet);
  return "Opened " + ORDER_ARCHIVE.sheetName;
}


// =======================================================================================
// SHEET SETUP + FORMATS
// =======================================================================================

/**
 * ⚠⚠ GOTCHA #16 — A COLUMN THAT HOLDS CODE-WRITTEN NUMBERS MUST *SET* ITS NUMBER
 *    FORMAT. This project has been bitten three times (Zoho Stock SELLING PRICE,
 *    Out of Stock DAYS OUT, Kit Health AT RISK): clearContent PRESERVES formats, so a
 *    column that once held dates renders every later integer as a 1900-era date AND
 *    getValues then returns Date OBJECTS, not numbers. The write side succeeds and
 *    looks fine; only a reader notices, and a reader doing parseFloat scores it ZERO.
 *
 * ⚠ DAY IS PINNED TO '@' (plain text) for exactly that reason — "2026-08-27" left on
 *   a default format is coerced to a Date on write, and comes back as a Date on read,
 *   which would break every string comparison in _oaProcess.
 */
function _oaApplyDataFormats(sheet, startRow, numRows) {
  if (numRows <= 0) return;
  var c = ORDER_ARCHIVE.cols;
  var stamp = "M/d/yy h:mm AM/PM";

  sheet.getRange(startRow, c.DAY,        numRows, 1).setNumberFormat("@");
  sheet.getRange(startRow, c.ORDER_ID,   numRows, 1).setNumberFormat("@");
  sheet.getRange(startRow, c.RECEIVED,   numRows, 1).setNumberFormat(stamp);
  sheet.getRange(startRow, c.FIRST_PICK, numRows, 1).setNumberFormat(stamp);
  sheet.getRange(startRow, c.PRINTED,    numRows, 1).setNumberFormat(stamp);
  sheet.getRange(startRow, c.TERMINAL_AT,numRows, 1).setNumberFormat(stamp);
  sheet.getRange(startRow, c.QUEUE_MIN,  numRows, 1).setNumberFormat("0");
  sheet.getRange(startRow, c.EXEC_MIN,   numRows, 1).setNumberFormat("0");
  sheet.getRange(startRow, c.TOTAL_MIN,  numRows, 1).setNumberFormat("0");
  sheet.getRange(startRow, c.LINES,      numRows, 1).setNumberFormat("0");
  sheet.getRange(startRow, c.UNITS,      numRows, 1).setNumberFormat("0");
}

/**
 * Row-1 status line: how far the archive has been swept.
 *
 * ⚠ DELIBERATELY NOT THE SHARED PULSE CHIP (_installPulseChip / SHEET_PULSE). That
 *   chip's CF tiers are quiet <2h, amber 2–26h, RED >26h — sized for sheets refreshed
 *   hourly. This one is swept ONCE A DAY by design, so a healthy archive would sit
 *   permanently amber and start drifting red before every sweep. Same lesson the hold's
 *   sidebar count taught in round six: a staleness budget is chosen for the group it
 *   sits in, and dropping something with a different cadence into that group inherits
 *   a threshold nobody sized for it. A plain factual line cannot cry wolf.
 */
function _oaPaintStatus(sheet, throughDay) {
  try {
    var rows = Math.max(0, sheet.getLastRow() - ORDER_ARCHIVE.dataStartRow + 1);
    sheet.getRange(ORDER_ARCHIVE.bannerRow, 13)
      .setValue("swept through " + throughDay + "  ·  " + rows + " orders")
      .setHorizontalAlignment("right")
      .setFontFamily("Roboto Mono")
      .setFontSize(9)
      .setFontColor("#5c4a00");
  } catch (e) { /* cosmetic only — never break a sweep over the banner */ }
}

/**
 * Idempotent create + style. Safe to re-run.
 */
function setupOrderArchiveSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ORDER_ARCHIVE.sheetName);
  if (!sheet) sheet = ss.insertSheet(ORDER_ARCHIVE.sheetName);

  if (sheet.getMaxColumns() < ORDER_ARCHIVE.dataWidth) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), ORDER_ARCHIVE.dataWidth - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < 200) sheet.insertRowsAfter(sheet.getMaxRows(), 200 - sheet.getMaxRows());

  // --- BANDING FIRST, then the manual band/header styling on top ---
  // ⚠ ORDER IS LOAD-BEARING. applyRowBanding with showHeader=true paints its own
  //   header colour over row 1; the Prep Queue title band was blacked out mid-restyle
  //   by exactly this. So banding starts at the HEADER row (row 1 sits deliberately
  //   OUTSIDE it) and every manual fill runs AFTER.
  try {
    var existing = sheet.getBandings();
    for (var b = 0; b < existing.length; b++) existing[b].remove();
    sheet.getRange(ORDER_ARCHIVE.headerRow, 1,
                   sheet.getMaxRows() - ORDER_ARCHIVE.headerRow + 1,
                   ORDER_ARCHIVE.dataWidth)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  } catch (e) { console.log("OrderArchive banding: " + e); }

  // --- ROW 1: title band (brand yellow, same grammar as OOS / Prep Queue) ---
  sheet.getRange(ORDER_ARCHIVE.bannerRow, 1, 1, ORDER_ARCHIVE.dataWidth)
    .setBackground("#ffd400")
    .setFontColor("#1d1d1b")
    .setFontFamily("Oswald")
    .setFontWeight("bold")
    .setFontSize(12)
    .setVerticalAlignment("middle");
  sheet.getRange(ORDER_ARCHIVE.bannerRow, 1)
    .setValue("ORDER ARCHIVE")
    .setNumberFormat('"▌  "@')          // display-only glyph; the VALUE stays clean
    .setHorizontalAlignment("left");

  // --- ROW 2: column headers ---
  sheet.getRange(ORDER_ARCHIVE.headerRow, 1, 1, ORDER_ARCHIVE.dataWidth)
    .setValues([ORDER_ARCHIVE.headers])
    .setBackground("#1d1d1b")
    .setFontColor("#ffd966")
    .setFontFamily("Oswald")
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  sheet.setFrozenRows(ORDER_ARCHIVE.headerRow);
  sheet.setRowHeight(ORDER_ARCHIVE.bannerRow, 34);
  sheet.setRowHeight(ORDER_ARCHIVE.headerRow, 34);

  // --- Column widths ---
  var c = ORDER_ARCHIVE.cols, W = {};
  W[c.ORDER_ID] = 140; W[c.CHANNEL] = 74;  W[c.DAY] = 92;   W[c.RECEIVED] = 130;
  W[c.FIRST_PICK] = 130; W[c.PRINTED] = 130; W[c.TERMINAL_AT] = 130;
  W[c.TERMINAL_STATUS] = 90; W[c.QUEUE_MIN] = 84; W[c.EXEC_MIN] = 84;
  W[c.TOTAL_MIN] = 84; W[c.LINES] = 64; W[c.UNITS] = 64; W[c.PICKER] = 120;
  W[c.HELD] = 56; W[c.KIT] = 56;
  for (var col in W) { try { sheet.setColumnWidth(parseInt(col, 10), W[col]); } catch (e) {} }

  // --- Data formats across the whole sheet so future appends inherit them ---
  _oaApplyDataFormats(sheet, ORDER_ARCHIVE.dataStartRow,
                      sheet.getMaxRows() - ORDER_ARCHIVE.dataStartRow + 1);

  sheet.getRange(ORDER_ARCHIVE.dataStartRow, 1,
                 sheet.getMaxRows() - ORDER_ARCHIVE.dataStartRow + 1,
                 ORDER_ARCHIVE.dataWidth)
    .setFontFamily("Roboto Mono")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  var wm = PropertiesService.getScriptProperties().getProperty(ORDER_ARCHIVE.watermarkKey);
  _oaPaintStatus(sheet, wm || "never");

  return "✅ " + ORDER_ARCHIVE.sheetName + " ready.";
}
