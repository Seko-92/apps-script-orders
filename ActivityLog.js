// =======================================================================================
// ACTIVITY_LOG.gs — Append-only event log of order lifecycle events
// =======================================================================================
//
// PURPOSE
//   Durable record of what happened, when, and why. The All Orders sheet only
//   shows current state — it can't answer "how many orders did we ship today?"
//   or "how long has order X been pending?" or "did n8n actually process the
//   last batch?". This sheet does.
//
//   Two write sites only — both already canonical chokepoints:
//     - doPost (n8n insert) → logs RECEIVED for each new row
//     - StatusService.updateOrderStatus → logs PREPARING / SHIPPED / CANCELED
//
//   Daily 3am trigger purges rows older than 90 days. Append-only otherwise.
//
// SCHEMA
//   TIMESTAMP · EVENT · ORDER_ID · SKU · QTY · SOURCE · DETAIL
//
// EVENTS
//   RECEIVED      — order line item written to All Orders (n8n insert OR manual entry)
//   PREPARING     — status transitioned to PREPARING
//   SHIPPED       — status transitioned to SHIPPED
//   CANCELED      — status transitioned to CANCELED
//   PRINTED       — picking list printed (picker accountability)
//   NOTE          — NOTE column edited (added / changed / removed) — captures any
//                   supervisor or picker remark added mid-prep, not just the
//                   original buyer note. DETAIL shows old text, NOTE col shows new.
//   ROW_DELETED   — (future) row manually deleted
//   FAILURE       — (future) updateOrderStatus returned success:false
//
// SOURCES
//   n8n           — eBay → n8n webhook (RECEIVED from automation)
//   manual        — manual entry directly in the sheet (e.g., DIRECT-table sales order)
//   telegram      — Telegram callback button press
//   manual-edit   — direct cell edit on F column
//   sidebar-bulk  — Mark Selection Preparing / similar bulk actions
//   sidebar       — sidebar action (e.g., Print Pick List)
//   n8n-direct    — n8n's status-update endpoint (vs full insert)
//   n8n-verify    — n8n's SHIPPED-verification sweep — reverts a sheet-SHIPPED order
//                   to PENDING when eBay's Fulfillment API says it's not actually
//                   FULFILLED yet. Surfacing this source distinctly is the canary:
//                   non-zero counts mean either the row-shift race recurred or some
//                   other path is producing bad SHIPPED state.
//   system        — automated/internal events
//
// PICKER COLUMN (added 2026-05-01)
//   Captures the warehouse staff identifier (Pick ID for Shipping, cell G2)
//   at the moment of each warehouse-side event. Empty for `n8n` and
//   `n8n-direct` sources (no human involved). The print path is gated on
//   G2 being non-empty (preparePrintSheet refuses if blank), and a daily
//   4am trigger blanks G2 so each shift starts with a fresh selection.
//
// PUBLIC API
//   setupActivityLogSheet()     — one-time: create sheet, brand styling
//   logActivity(...)            — append a single event (defensive — silent on failure)
//   logActivityBatch(rows)      — append many events in one setValues call
//   getTodayMetrics()           — { shippedToday, receivedToday, oldestPendingMinutes }
//   getRecentActivity(n)        — last N events for sidebar feed (newest first)
//   purgeOldActivityLog()       — delete rows older than 90 days (daily trigger)
//   setupActivityLogTrigger()   — install daily 3am purge trigger
//   removeActivityLogTrigger()  — uninstall purge trigger
// =======================================================================================

// ---------- LOCAL SCHEMA ----------
var ACTIVITY_LOG = {
  sheetName: "Activity Log",

  // 1-based column positions
  cols: {
    TIMESTAMP: 1,   // A — real Date
    EVENT:     2,   // B
    ORDER_ID:  3,   // C
    SKU:       4,   // D
    QTY:       5,   // E
    SOURCE:    6,   // F
    DETAIL:    7,   // G — structured info (e.g. "from PENDING", "DIRECT manual")
    NOTE:      8,   // H — free-text NOTE field from the order row at event time
    PICKER:    9    // I — Pick ID for Shipping at time of event (warehouse-side only)
  },

  idx: function(name) { return ACTIVITY_LOG.cols[name] - 1; },

  dataWidth:    9,
  headerRow:    1,
  dataStartRow: 2,

  headers: ["⏱ TIMESTAMP", "EVENT", "ORDER ID", "◈ SKU", "# QTY", "SOURCE", "DETAIL", "NOTE", "👤 PICKER"],

  // Sources that should capture the current picker (G2) on log write.
  // Anything not in this set leaves PICKER blank (e.g., n8n automation events).
  warehouseSources: ["telegram", "manual", "manual-edit", "sidebar-bulk", "sidebar"],

  // Retention window — anything older than this gets purged by the daily trigger
  retentionDays: 90
};


// =======================================================================================
// PUBLIC API
// =======================================================================================

/**
 * Idempotent setup. Creates the sheet if missing, applies brand styling.
 */
function setupActivityLogSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Force the spreadsheet's timezone to America/Chicago (= Houston, CT).
  // Sheets renders Date cells in the SPREADSHEET's TZ, not the script's. If
  // the spreadsheet is on a different zone, Activity Log timestamps display
  // off by hours. This is a spreadsheet-wide setting; safe to re-apply.
  try { ss.setSpreadsheetTimeZone("America/Chicago"); } catch (e) { /* no-op */ }

  var sheet = ss.getSheetByName(ACTIVITY_LOG.sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(ACTIVITY_LOG.sheetName);
  }

  // --- HEADERS ---
  sheet.getRange(ACTIVITY_LOG.headerRow, 1, 1, ACTIVITY_LOG.dataWidth)
    .setValues([ACTIVITY_LOG.headers])
    .setBackground('#1d1d1b')
    .setFontColor('#ffd966')
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(ACTIVITY_LOG.headerRow, 1, 1, ACTIVITY_LOG.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(ACTIVITY_LOG.cols.TIMESTAMP, 160);
  sheet.setColumnWidth(ACTIVITY_LOG.cols.EVENT,     110);
  sheet.setColumnWidth(ACTIVITY_LOG.cols.ORDER_ID,  150);
  sheet.setColumnWidth(ACTIVITY_LOG.cols.SKU,       110);
  sheet.setColumnWidth(ACTIVITY_LOG.cols.QTY,        60);
  sheet.setColumnWidth(ACTIVITY_LOG.cols.SOURCE,    110);
  sheet.setColumnWidth(ACTIVITY_LOG.cols.DETAIL,    220);
  sheet.setColumnWidth(ACTIVITY_LOG.cols.NOTE,      260);
  sheet.setColumnWidth(ACTIVITY_LOG.cols.PICKER,    130);

  // --- DATA AREA: column-level formats ---
  var maxDataRow = 5000;
  var dataRows = maxDataRow - ACTIVITY_LOG.dataStartRow + 1;

  sheet.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.TIMESTAMP, dataRows, 1)
    .setNumberFormat('M/d/yy h:mm:ss a')
    .setFontFamily('Roboto Mono').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('left');
  sheet.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.EVENT, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.ORDER_ID, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.SKU, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.QTY, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.SOURCE, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('center');
  sheet.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.DETAIL, dataRows, 1)
    .setFontFamily('Roboto').setFontStyle('italic').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('left');
  sheet.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.NOTE, dataRows, 1)
    .setFontFamily('Roboto').setFontSize(10)
    .setFontColor('#1d1d1b').setHorizontalAlignment('left').setWrap(true);
  sheet.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.PICKER, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setFontColor('#1d1d1b').setHorizontalAlignment('center');

  sheet.getRange(ACTIVITY_LOG.dataStartRow, 1, dataRows, ACTIVITY_LOG.dataWidth)
    .setVerticalAlignment('middle');

  // --- BANDING (cream alternation) ---
  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  var bandRange = sheet.getRange(1, 1, maxDataRow, ACTIVITY_LOG.dataWidth);
  var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b')
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#fff8e7');

  // --- EVENT-COLUMN CONDITIONAL FORMATTING ---
  // Different colors per event type so the log is scannable at a glance.
  var existingRules = sheet.getConditionalFormatRules();
  var keptRules = existingRules.filter(function(r) {
    var ranges = r.getRanges();
    return !ranges.some(function(rg) {
      return rg.getColumn() === ACTIVITY_LOG.cols.EVENT;
    });
  });
  var eventRange = sheet.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.EVENT, dataRows, 1);
  function eventRule(eventName, bg, fg) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(eventName)
      .setBackground(bg).setFontColor(fg).setBold(true)
      .setRanges([eventRange])
      .build();
  }
  keptRules.push(eventRule('RECEIVED',  '#fff8e7', '#1d1d1b'));   // cream
  keptRules.push(eventRule('PREPARING', '#ffd966', '#1d1d1b'));   // brand yellow
  keptRules.push(eventRule('SHIPPED',   '#e8f5e9', '#1b5e20'));   // green
  keptRules.push(eventRule('CANCELED',  '#f0f0f0', '#5f5f5f'));   // gray
  keptRules.push(eventRule('PRINTED',   '#e3f2fd', '#0d47a1'));   // soft blue — printed/processed
  keptRules.push(eventRule('NOTE',      '#fff4b0', '#1d1d1b'));   // soft yellow — note added/changed
  keptRules.push(eventRule('FAILURE',   '#ff6b6b', '#ffffff'));   // red

  // Also drop any prior PICKER-column rule, then add: soft yellow tint when populated
  keptRules = keptRules.filter(function(r) {
    var ranges = r.getRanges();
    return !ranges.some(function(rg) { return rg.getColumn() === ACTIVITY_LOG.cols.PICKER; });
  });
  var pickerRange = sheet.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.PICKER, dataRows, 1);
  // PICKER is now col I (was col H before NOTE was added at col H). Use a
  // dynamic letter computed from the schema so future reorders don't break.
  var pickerColLetter = String.fromCharCode(64 + ACTIVITY_LOG.cols.PICKER);  // 65='A', 73='I'
  keptRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=LEN(" + pickerColLetter + ACTIVITY_LOG.dataStartRow + ")>0")
      .setBackground('#fff8e7').setRanges([pickerRange]).build()
  );

  // BUYER NOTE highlighting on the NOTE column — mirrors the All Orders rule
  // (italic + muted gold #8a7434 for cells starting with "Buyer Note:").
  // No background change so the row banding + PICKER tint underneath read
  // through unchanged. Edit-to-default flip works the same way: the moment
  // a supervisor removes the "Buyer Note:" prefix, the rule stops matching
  // and the cell returns to its banded baseline appearance.
  // Drop any prior buyer-note rule on the NOTE column before re-adding.
  keptRules = keptRules.filter(function(r) {
    var bc = r.getBooleanCondition();
    if (!bc) return true;
    var formula = (bc.getCriteriaValues() || [''])[0] || '';
    var ranges  = r.getRanges();
    var isNoteRange = ranges.some(function(rg) {
      return rg.getColumn() === ACTIVITY_LOG.cols.NOTE && rg.getNumColumns() === 1;
    });
    return !(isNoteRange && formula.toLowerCase().indexOf('buyer note') !== -1);
  });
  var noteRange = sheet.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.NOTE, dataRows, 1);
  var noteColLetter = String.fromCharCode(64 + ACTIVITY_LOG.cols.NOTE);  // 65='A', 72='H'
  var buyerNoteFormula =
    '=AND($' + noteColLetter + ACTIVITY_LOG.dataStartRow + '<>"", ' +
    'REGEXMATCH(TO_TEXT($' + noteColLetter + ACTIVITY_LOG.dataStartRow + '), "(?i)^buyer note:"))';
  keptRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(buyerNoteFormula)
      .setItalic(true).setFontColor('#8a7434')
      .setRanges([noteRange]).build()
  );

  sheet.setConditionalFormatRules(keptRules);

  // --- FREEZE HEADER ROW ---
  sheet.setFrozenRows(1);

  return "✅ Activity Log sheet ready.";
}


/**
 * Sidebar entry point: returns the current picker name (cleaned) for the
 * top-of-sidebar accountability banner. Empty string when unset — the sidebar
 * uses that to render the "No picker set" warning state.
 */
function getCurrentPicker() {
  return _currentPicker();
}


/**
 * Reads the current Pick ID for Shipping (cell G2) from the All Orders sheet
 * and returns it cleaned via _extractPickIdData (e.g. "Shipping - YAwiss 1" →
 * "YAwiss · 1"). Empty string if blank, sheet unavailable, or value is the
 * dropdown's placeholder header ("Pick ID for Shipping" — must be rejected
 * as "unset," not treated as a valid picker name).
 *
 * Used by warehouse-side log paths to stamp the picker on each event AND by
 * the sidebar banner to show set/unset state.
 */
function _currentPicker() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return "";
    var raw = sheet.getRange(Schema.cellEmployeeId).getValue();
    if (!raw) return "";
    var rawStr = String(raw).trim();
    // Real values look like "Shipping - [Name] [Id]". The dropdown's header
    // label "Pick ID for Shipping" doesn't match this pattern → treat as unset.
    if (!/^Shipping\s*-\s*/i.test(rawStr)) return "";
    // _extractPickIdData lives in FulfillmentService.js — strips prefix +
    // standardizes the trailing ID separator. Returns "—" for blanks; map
    // that back to "" here so the log column is clean-empty when unset.
    var cleaned = (typeof _extractPickIdData === "function")
      ? _extractPickIdData(raw)
      : rawStr;
    return (cleaned === "—") ? "" : cleaned;
  } catch (err) {
    return "";
  }
}

/**
 * Decides whether a given source represents a warehouse-side action where
 * the current picker (G2) should be captured. n8n / system events return
 * empty picker — no human is involved, so attribution would be misleading.
 */
function _shouldCapturePicker(source) {
  return ACTIVITY_LOG.warehouseSources.indexOf(String(source || "")) !== -1;
}


/**
 * Append a single event. DEFENSIVE — silent on any failure (the log is
 * best-effort; a logging error must NEVER block the main operation).
 *
 * If the Activity Log sheet doesn't exist yet, this returns silently. The user
 * runs setupActivityLogSheet() once to opt in.
 *
 * The picker arg is optional; if not passed (or undefined) and the source is
 * warehouse-side, G2 is read automatically. Pass an explicit string (or empty
 * "") to override.
 *
 * The note arg is optional and goes to its own column (separate from DETAIL).
 * Use it for the order's free-text NOTE field — buyer messages, supervisor
 * remarks, etc.
 */
function logActivity(event, orderId, sku, qty, source, detail, picker, note) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ACTIVITY_LOG.sheetName);
    if (!sheet) return;  // not set up yet — silent skip

    var pickerOut = (typeof picker === "string")
      ? picker
      : (_shouldCapturePicker(source) ? _currentPicker() : "");

    sheet.appendRow([
      new Date(),
      String(event || "").toUpperCase(),
      String(orderId || ""),
      String(sku || ""),
      qty || "",
      String(source || ""),
      String(detail || ""),
      String(note || ""),
      pickerOut
    ]);
  } catch (err) {
    try { Logger.log("logActivity error: " + err); } catch (_) {}
  }
}


/**
 * Append many events in one setValues call. Use this when a single operation
 * generates multiple events (e.g., doPost inserting 5 orders, or
 * updateOrderStatus changing 12 rows at once).
 *
 * `rows` is an array of arrays, each:
 *   [event, orderId, sku, qty, source, detail, picker?, note?]
 *
 *   - `picker` (slot 6): if omitted/undefined, G2 is read once per source per
 *     batch (warehouse sources only — n8n/system get blank).
 *   - `note` (slot 7): order's free-text NOTE field; goes to its own column.
 *     Optional; defaults to empty string.
 */
function logActivityBatch(rows) {
  if (!rows || rows.length === 0) return;
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ACTIVITY_LOG.sheetName);
    if (!sheet) return;

    // Cache picker per source — warehouse sources resolve to G2 (read once),
    // others to "". Avoids repeated reads when batch contains mixed sources.
    var pickerCache = {};
    function pickerFor(source) {
      if (pickerCache.hasOwnProperty(source)) return pickerCache[source];
      var p = _shouldCapturePicker(source) ? _currentPicker() : "";
      pickerCache[source] = p;
      return p;
    }

    var now = new Date();
    var withTs = rows.map(function(r) {
      var src = String(r[4] || "");
      var pickerOut = (r.length >= 7 && typeof r[6] === "string") ? r[6] : pickerFor(src);
      var noteOut = (r.length >= 8) ? String(r[7] || "") : "";
      return [
        now,
        String(r[0] || "").toUpperCase(),
        String(r[1] || ""),
        String(r[2] || ""),
        r[3] || "",
        src,
        String(r[5] || ""),
        noteOut,
        pickerOut
      ];
    });

    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, withTs.length, ACTIVITY_LOG.dataWidth).setValues(withTs);
  } catch (err) {
    try { Logger.log("logActivityBatch error: " + err); } catch (_) {}
  }
}


/**
 * Today's operational metrics for the sidebar dashboard.
 *   - shippedToday: # of SHIPPED events with timestamp today
 *   - receivedToday: # of RECEIVED events with timestamp today
 *   - oldestPendingMinutes: age in minutes of the oldest currently-PENDING order,
 *                           computed by joining the All Orders sheet's PENDING
 *                           rows against the log's RECEIVED events.
 *                           null if no pending orders.
 *
 * "Today" is local to America/Chicago (matches the rest of the system).
 */
function getTodayMetrics() {
  var result = {
    shippedToday: 0,
    receivedToday: 0,
    oldestPendingMinutes: null
  };

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName(ACTIVITY_LOG.sheetName);
    var mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
    if (!logSheet) return result;

    // Boundary: midnight today in America/Chicago, expressed as a JS Date.
    // Apps Script JS Dates are ms-since-epoch UTC, so we convert via the
    // session timezone string, then parse back.
    var nowChicago = Utilities.formatDate(new Date(), "America/Chicago", "yyyy-MM-dd");
    var startOfToday = new Date(nowChicago + "T00:00:00-06:00").getTime();
    // (-06:00 is CST; Apps Script's date arithmetic tolerates DST offset drift
    // for our purpose since we're comparing within the same day.)

    var lastRow = logSheet.getLastRow();
    if (lastRow < ACTIVITY_LOG.dataStartRow) return result;

    var data = logSheet.getRange(
      ACTIVITY_LOG.dataStartRow, 1,
      lastRow - ACTIVITY_LOG.dataStartRow + 1,
      ACTIVITY_LOG.dataWidth
    ).getValues();

    // First pass: count today's events + build orderId → earliest RECEIVED ms map
    var receivedMap = {};
    for (var i = 0; i < data.length; i++) {
      var ts = data[i][ACTIVITY_LOG.idx("TIMESTAMP")];
      if (!(ts instanceof Date)) continue;
      var event = String(data[i][ACTIVITY_LOG.idx("EVENT")] || "").toUpperCase();
      var orderId = String(data[i][ACTIVITY_LOG.idx("ORDER_ID")] || "").trim();

      if (ts.getTime() >= startOfToday) {
        if (event === "RECEIVED") result.receivedToday++;
        else if (event === "SHIPPED") result.shippedToday++;
      }

      if (event === "RECEIVED" && orderId) {
        var existing = receivedMap[orderId];
        if (!existing || ts.getTime() < existing) {
          receivedMap[orderId] = ts.getTime();
        }
      }
    }

    // Second pass: scan All Orders for PENDING rows, find oldest by RECEIVED ts
    if (mainSheet) {
      var mainLastRow = mainSheet.getLastRow();
      if (mainLastRow >= Schema.dataStartRow) {
        var mainData = mainSheet.getRange(
          Schema.dataStartRow, 1,
          mainLastRow - Schema.dataStartRow + 1,
          Schema.cols.STATUS
        ).getValues();
        var oldestMs = null;
        for (var j = 0; j < mainData.length; j++) {
          var status = String(mainData[j][Schema.idx("STATUS")] || "").trim().toUpperCase();
          if (status !== Schema.status.PENDING) continue;
          var oid = String(mainData[j][Schema.idx("SALES_ORDER")] || "").trim();
          if (!oid || !receivedMap[oid]) continue;
          if (oldestMs === null || receivedMap[oid] < oldestMs) {
            oldestMs = receivedMap[oid];
          }
        }
        if (oldestMs !== null) {
          result.oldestPendingMinutes = Math.floor((Date.now() - oldestMs) / 60000);
        }
      }
    }
  } catch (err) {
    try { Logger.log("getTodayMetrics error: " + err); } catch (_) {}
  }

  return result;
}


/**
 * One-stop dashboard read. Combines today's metrics + a timeline of today's
 * events + current PENDING count + last-sync freshness into a single
 * round-trip — used by the sidebar's Operations Cockpit panel.
 *
 * Returns:
 *   {
 *     shippedToday:           int,
 *     receivedToday:          int,
 *     oldestPendingMinutes:   int | null,
 *     pendingCount:           int,
 *     lastSyncMinutes:        int | null,
 *     // Queue strip counts (2026-05-16 — sidebar v3.5 Phase 5):
 *     // Cockpit's queue strip uses ebayPending + directPending + prepQueueCount.
 *     // directPending is ALREADY MERGED (DIRECT-table non-terminal + Zoho not-yet-pulled);
 *     // zohoPending is exposed separately for tooltip use.
 *     ebayPending:            int,    // eBay table PENDING + PREPARING
 *     directPending:          int,    // DIRECT non-terminal + Zoho pending (combined)
 *     zohoPending:            int,    // sub-component: SOs waiting to be Pulled
 *     prepQueueCount:         int,
 *     timeline: [
 *       { hourFraction: 0..24, event: str, orderId: str, sku: str, picker: str }
 *       ...
 *     ]  // today's significant events (NOTE excluded — reduces visual noise)
 *   }
 *
 * timeline entries are pre-converted to a 0..24 hour fraction in America/
 * Chicago time so the client can render the tape without doing date math.
 * Array is sorted oldest → newest.
 */
// Classify an order id as Direct vs eBay — mirrors the Floor Board's
// inferChannel(): SO-/INV- prefixes and any non eBay-digit-dash id are Direct.
function _floorOrderIsDirect(orderId) {
  var u = String(orderId || "").trim().toUpperCase();
  if (!u) return false;                                  // unknown → count with eBay so the split sums to the total
  if (u.indexOf("SO-") === 0 || u.indexOf("INV-") === 0) return true;
  if (/^[0-9][0-9\-]+$/.test(u)) return false;           // clean eBay digit-dash id
  return true;
}

function getDashboardSnapshot() {
  var result = {
    shippedToday:         0,
    receivedToday:        0,
    receivedEbay:         0,    // today's intake split by channel (Received-card breakdown)
    receivedDirect:       0,
    oldestPendingMinutes: null,
    pastRedlineCount:     0,    // # of PENDING orders aged past the 3h redline
    pendingCount:         0,
    lastSyncMinutes:      null,
    ebayPending:          0,
    directPending:        0,
    ebayGrab:             0,
    directGrab:           0,
    zohoPending:          0,
    prepQueueCount:       0,
    timeline:             []
  };

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName(ACTIVITY_LOG.sheetName);
    var mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);

    // String-comparison "is this today" — DST-safe (no fixed-offset math).
    var todayStr = Utilities.formatDate(new Date(), "America/Chicago", "yyyy-MM-dd");

    // ---- One pass through the activity log ----
    var receivedMap = {};   // orderId → earliest RECEIVED ms (for oldest-pending)
    var TIMELINE_EVENTS = { RECEIVED: 1, PREPARING: 1, SHIPPED: 1, CANCELED: 1, PRINTED: 1 };

    if (logSheet) {
      var lastLogRow = logSheet.getLastRow();
      if (lastLogRow >= ACTIVITY_LOG.dataStartRow) {
        var data = logSheet.getRange(
          ACTIVITY_LOG.dataStartRow, 1,
          lastLogRow - ACTIVITY_LOG.dataStartRow + 1,
          ACTIVITY_LOG.dataWidth
        ).getValues();

        for (var i = 0; i < data.length; i++) {
          var ts = data[i][ACTIVITY_LOG.idx("TIMESTAMP")];
          if (!(ts instanceof Date)) continue;
          var event   = String(data[i][ACTIVITY_LOG.idx("EVENT")]  || "").toUpperCase();
          var orderId = String(data[i][ACTIVITY_LOG.idx("ORDER_ID")] || "").trim();
          var sku     = String(data[i][ACTIVITY_LOG.idx("SKU")]    || "").trim();
          var picker  = String(data[i][ACTIVITY_LOG.idx("PICKER")] || "").trim();

          var eventDateStr = Utilities.formatDate(ts, "America/Chicago", "yyyy-MM-dd");
          var isToday = (eventDateStr === todayStr);

          if (isToday) {
            if (event === "RECEIVED") {
              result.receivedToday++;
              // channel by order-id shape (mirrors the board's inferChannel);
              // anything not a clean eBay digit-dash id counts as Direct.
              if (_floorOrderIsDirect(orderId)) result.receivedDirect++;
              else                              result.receivedEbay++;
            }
            else if (event === "SHIPPED") result.shippedToday++;

            // Timeline entry — convert to 0..24 hour fraction in Chicago TZ
            if (TIMELINE_EVENTS[event]) {
              var hh = parseInt(Utilities.formatDate(ts, "America/Chicago", "H"), 10);
              var mm = parseInt(Utilities.formatDate(ts, "America/Chicago", "m"), 10);
              result.timeline.push({
                hourFraction: hh + (mm / 60),
                event:        event,
                orderId:      orderId,
                sku:          sku,
                picker:       picker
              });
            }
          }

          // Earliest RECEIVED per order (across all days) — feeds oldest-pending
          if (event === "RECEIVED" && orderId) {
            var prev = receivedMap[orderId];
            if (!prev || ts.getTime() < prev) {
              receivedMap[orderId] = ts.getTime();
            }
          }
        }
      }
    }

    // ---- One pass through the main sheet for PENDING + oldest + queue split ----
    // 2026-05-16: extended to compute ebay/direct non-terminal counts for the
    // cockpit queue strip. Boundary row is detected inline (col A holds
    // Schema.boundaryMarker == "DIRECT") to avoid a second sheet read.
    if (mainSheet) {
      var mainLastRow = mainSheet.getLastRow();
      if (mainLastRow >= Schema.dataStartRow) {
        var mainData = mainSheet.getRange(
          Schema.dataStartRow, 1,
          mainLastRow - Schema.dataStartRow + 1,
          Schema.cols.STATUS
        ).getValues();

        var oldestMs = null;
        var boundaryArrayIdx = -1;   // index in mainData where col A == "DIRECT"

        for (var j = 0; j < mainData.length; j++) {
          var skuCell = String(mainData[j][Schema.idx("SKU")] || "").trim().toUpperCase();

          // Boundary divider row — flip to DIRECT-side from here on, then skip
          if (skuCell === Schema.boundaryMarker) {
            boundaryArrayIdx = j;
            continue;
          }
          // DIRECT header row sits immediately after the boundary — skip it
          if (boundaryArrayIdx !== -1 && j === boundaryArrayIdx + 1) continue;

          var status = String(mainData[j][Schema.idx("STATUS")] || "").trim().toUpperCase();
          var inDirect = (boundaryArrayIdx !== -1 && j > boundaryArrayIdx);

          // Queue-strip counts: PENDING + PREPARING = "workload waiting"
          if (status === Schema.status.PENDING || status === Schema.status.PREPARING) {
            if (inDirect) result.directPending++;
            else          result.ebayPending++;
          }

          // PENDING-only counters: pendingCount + per-channel "to grab".
          // ebayGrab/directGrab are the Floor Board's ORDERS-TO-GRAB hero —
          // strictly not-yet-started, table rows only (excludes PREPARING and
          // the Pending-SO mirror), so the number reflects real pick work.
          if (status !== Schema.status.PENDING) continue;
          result.pendingCount++;
          if (inDirect) result.directGrab++;
          else          result.ebayGrab++;
          var oid = String(mainData[j][Schema.idx("SALES_ORDER")] || "").trim();
          if (oid && receivedMap[oid]) {
            if (oldestMs === null || receivedMap[oid] < oldestMs) {
              oldestMs = receivedMap[oid];
            }
            // how many are aging past the 3h redline (companion to the oldest figure)
            if (Date.now() - receivedMap[oid] > 180 * 60000) result.pastRedlineCount++;
          }
        }
        if (oldestMs !== null) {
          result.oldestPendingMinutes = Math.floor((Date.now() - oldestMs) / 60000);
        }
      }

      // Parse the human-readable last-sync cell (E1) into a freshness number.
      var syncRaw = mainSheet.getRange(Schema.cellSyncTime).getValue();
      var match = String(syncRaw || "").match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
      if (match) {
        var hour = parseInt(match[1], 10);
        var min  = parseInt(match[2], 10);
        var mer  = match[3].toUpperCase();
        if (mer === "PM" && hour < 12) hour += 12;
        if (mer === "AM" && hour === 12) hour = 0;
        // E1 is wall-clock time in the SPREADSHEET timezone (America/Chicago).
        // Compare against "now" in the SAME timezone via formatDate — the old
        // new Date()+setHours() ran in the SCRIPT timezone, and that mismatch
        // produced bogus multi-hour "stale" readings (2026-06-02 fix). Pure
        // minutes-of-day diff, midnight-wrapped.
        var nowParts = Utilities.formatDate(new Date(), "America/Chicago", "H:m").split(":");
        var nowMin   = parseInt(nowParts[0], 10) * 60 + parseInt(nowParts[1], 10);
        var syncMin  = hour * 60 + min;
        var diff = nowMin - syncMin;
        if (diff < 0) diff += 1440;   // sync stamped before midnight, now after
        result.lastSyncMinutes = diff;
      }
    }
    // ---- Zoho pending + Prep Queue (defensive: silent 0 if helpers missing) ----
    // 2026-06-02: directPending is NO LONGER merged with the Pending-SO mirror.
    // Per picker feedback, "Direct" must mean what's actually on the DIRECT
    // table (pickable) — not not-yet-Pulled SOs sitting in the Pending Sales
    // Orders sheet (that mirror holds many old/shipped/closed SOs and inflated
    // the count). zohoPending stays exposed separately so both surfaces show it
    // as a distinct "N to pull" signal WITHOUT inflating Direct. (Reverses the
    // v3.5 merge; affects the sidebar cockpit Direct pill too — intentional.)
    try {
      if (typeof getPendingZohoCount === "function") {
        result.zohoPending = getPendingZohoCount() || 0;
      }
    } catch (e1) { /* swallow — never fail the snapshot for a sub-count */ }

    try {
      if (typeof _getPrepQueueSize === "function") {
        result.prepQueueCount = _getPrepQueueSize() || 0;
      }
    } catch (e2) { /* swallow */ }

  } catch (err) {
    try { Logger.log("getDashboardSnapshot error: " + err); } catch (_) {}
  }

  return result;
}


/**
 * Returns the last N events as objects, newest first. For the sidebar's
 * "Recent Activity" feed.
 */
function getRecentActivity(n) {
  n = parseInt(n) || 5;
  var out = [];
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ACTIVITY_LOG.sheetName);
    if (!sheet) return out;

    var lastRow = sheet.getLastRow();
    if (lastRow < ACTIVITY_LOG.dataStartRow) return out;

    var startRow = Math.max(ACTIVITY_LOG.dataStartRow, lastRow - n + 1);
    var rows = sheet.getRange(
      startRow, 1,
      lastRow - startRow + 1,
      ACTIVITY_LOG.dataWidth
    ).getValues();

    // Reverse so newest is first
    rows.reverse();

    for (var i = 0; i < rows.length; i++) {
      var ts = rows[i][ACTIVITY_LOG.idx("TIMESTAMP")];
      out.push({
        timestamp: (ts instanceof Date) ? ts.getTime() : null,
        event:     String(rows[i][ACTIVITY_LOG.idx("EVENT")] || ""),
        orderId:   String(rows[i][ACTIVITY_LOG.idx("ORDER_ID")] || ""),
        sku:       String(rows[i][ACTIVITY_LOG.idx("SKU")] || ""),
        qty:       rows[i][ACTIVITY_LOG.idx("QTY")] || "",
        source:    String(rows[i][ACTIVITY_LOG.idx("SOURCE")] || ""),
        detail:    String(rows[i][ACTIVITY_LOG.idx("DETAIL")] || ""),
        note:      String(rows[i][ACTIVITY_LOG.idx("NOTE")] || ""),
        picker:    String(rows[i][ACTIVITY_LOG.idx("PICKER")] || "")
      });
    }
  } catch (err) {
    try { Logger.log("getRecentActivity error: " + err); } catch (_) {}
  }
  return out;
}


/**
 * Removes rows older than ACTIVITY_LOG.retentionDays. Run daily by trigger.
 *
 * Reads timestamps in column A, finds the cutoff index, deletes rows in one
 * call. Faster than per-row deletes.
 */
function purgeOldActivityLog() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ACTIVITY_LOG.sheetName);
    if (!sheet) return "Activity Log sheet not found.";

    var lastRow = sheet.getLastRow();
    if (lastRow < ACTIVITY_LOG.dataStartRow) return "Nothing to purge.";

    var cutoff = Date.now() - (ACTIVITY_LOG.retentionDays * 24 * 60 * 60 * 1000);

    var ts = sheet.getRange(
      ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.TIMESTAMP,
      lastRow - ACTIVITY_LOG.dataStartRow + 1, 1
    ).getValues();

    // Rows are appended in chronological order, so the first row that's NEWER
    // than cutoff is where we stop deleting. Everything before it is old.
    var deleteCount = 0;
    for (var i = 0; i < ts.length; i++) {
      var t = ts[i][0];
      if (t instanceof Date && t.getTime() < cutoff) {
        deleteCount++;
      } else {
        break;  // first non-old row → stop
      }
    }

    if (deleteCount > 0) {
      sheet.deleteRows(ACTIVITY_LOG.dataStartRow, deleteCount);
    }

    return "✅ Purged " + deleteCount + " old log row(s).";
  } catch (err) {
    Logger.log("purgeOldActivityLog error: " + err);
    return "❌ Purge error: " + err;
  }
}


// =======================================================================================
// TRIGGER MANAGEMENT
// =======================================================================================

/**
 * Run ONCE from the Apps Script Editor to install the daily 3am purge trigger.
 * Idempotent — removes any prior purge trigger before creating the new one.
 */
function setupActivityLogTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'purgeOldActivityLog') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('purgeOldActivityLog')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  Logger.log("Activity Log purge trigger installed: daily 3am");
  try {
    SpreadsheetApp.getUi().alert(
      "Trigger Installed",
      "Activity Log will purge rows older than " + ACTIVITY_LOG.retentionDays +
      " days every day at 3am.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log("Trigger installed. (No UI context for alert)");
  }
}

/** Removes the daily purge trigger. Manual cleanup helper. */
function removeActivityLogTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'purgeOldActivityLog') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  Logger.log("Removed " + removed + " purgeOldActivityLog trigger(s).");
  return "Removed " + removed + " trigger(s).";
}


// =======================================================================================
// DAILY PICKER RESET — pairs with the Activity Log's picker-capture mechanism.
// =======================================================================================

/**
 * Blanks G2 (Pick ID for Shipping) and I2 (Pick ID for Adjustment) so each
 * shift starts with a fresh selection. Run by a daily 4am trigger — between
 * the Activity Log purge (3am) and warehouse-open hours.
 *
 * Why this matters: the print path refuses to run when G2 is empty. That
 * forcing-function only "forces" if G2 actually starts empty each day. Without
 * a reset, yesterday's value lingers and the picker may forget to update it,
 * causing the day's events to be attributed to the wrong person.
 *
 * Safe to run any time. Only writes empty strings — does NOT touch the
 * dropdown validation rules on those cells (those are preserved).
 */
function resetDailyPickIds() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return "Main sheet not found.";

    sheet.getRange(Schema.cellEmployeeId).setValue("Pick ID for Shipping");
    sheet.getRange(Schema.cellAdjustmentId).setValue("Pick ID for Adjustment");

    Logger.log("Daily Pick IDs reset (" + Schema.cellEmployeeId + ", " +
               Schema.cellAdjustmentId + ") at " + new Date().toISOString());
    return "✅ Pick IDs reset.";
  } catch (err) {
    Logger.log("resetDailyPickIds error: " + err);
    return "❌ Reset error: " + err;
  }
}

/**
 * Run ONCE from the Apps Script Editor to install the daily 4am picker reset.
 * Idempotent — removes any prior reset trigger before creating the new one.
 */
function setupDailyPickerResetTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'resetDailyPickIds') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('resetDailyPickIds')
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();

  Logger.log("Daily picker reset trigger installed: 4am");
  try {
    SpreadsheetApp.getUi().alert(
      "Trigger Installed",
      "Pick IDs (G2 + I2) will be cleared every day at 4am. The picker " +
      "must reselect their ID before printing the picking list.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log("Trigger installed. (No UI context for alert)");
  }
}

/** Removes the daily picker reset trigger. */
function removeDailyPickerResetTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'resetDailyPickIds') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  Logger.log("Removed " + removed + " resetDailyPickIds trigger(s).");
  return "Removed " + removed + " trigger(s).";
}
