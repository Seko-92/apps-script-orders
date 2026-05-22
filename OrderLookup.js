// =======================================================================================
// ORDER_LOOKUP.gs — Customer-service order-history lookup
// =======================================================================================
//
// PURPOSE
//   When a customer calls about an order ("what's the status of #02-14623?"),
//   the warehouse needs ONE consolidated view: current row in All Orders +
//   the full Activity Log timeline for that order. Today that's two Ctrl+F
//   passes across two sheets. This service collapses both into one query.
//
// IT IS NOT A SEARCH BAR
//   We deliberately did NOT build a generic substring-search-as-you-type bar
//   (that would mostly duplicate Ctrl+F). This is an explicit "look this order
//   up" action — type an order ID, see everything we know about it.
//
// PUBLIC API
//   lookupOrder(query)   — returns { found, query, rows, events, summary }
//   jumpToOrderRows(rows) — activate All Orders + multi-select matched rows
//
// MATCHING RULES
//   - Query is trimmed + uppercased + alphanumeric-normalized before comparison
//     (warehouse staff paste IDs with surrounding spaces; an exact-match-only
//     rule would frustrate that)
//   - Matches All Orders col D (SALES_ORDER) and Activity Log col C (ORDER_ID)
//     — these hold the same value, just under different column names per sheet
//   - Substring match on the normalized form (so a partial paste like
//     "14623-46718" still finds "02-14623-46718")
//   - Header rows + the DIRECT divider are skipped automatically (col A
//     contains "DIRECT" or "◈ SKU" etc. — they don't have SALES_ORDER values)
// =======================================================================================


// =======================================================================================
// PUBLIC: lookupOrder(query)
// =======================================================================================

/**
 * Looks up an order across All Orders + Activity Log.
 *
 * @param {string} query — order ID (full or fragment). Case-insensitive,
 *                         non-alphanumeric chars ignored for forgiving paste.
 * @returns {{
 *   found: boolean,
 *   query: string,
 *   normalizedQuery: string,
 *   rows: Array<{row, table, sku, qty, location, status, salesOrder, note, shipping, shipCost}>,
 *   events: Array<{timestamp, event, source, detail, sku, qty, note, picker}>,
 *   summary: {rowCount, totalQty, statuses, pickers, earliestReceived, latestEvent, firstRow}
 * }}
 *
 * Defensive: any error in one half (rows or events) is logged and the other
 * half still returns. found=true iff EITHER half produced data.
 */
function lookupOrder(query) {
  var raw = String(query == null ? "" : query).trim();
  var normalized = _normalizeOrderId(raw);

  var result = {
    found: false,
    query: raw,
    normalizedQuery: normalized,
    rows: [],
    events: [],
    summary: {
      rowCount: 0,
      totalQty: 0,
      statuses: [],
      pickers: [],
      earliestReceived: null,
      latestEvent: null,
      firstRow: null
    }
  };

  if (!normalized) return result;

  // --- Half 1: scan All Orders for matching rows ---
  try {
    result.rows = _findOrderRows(normalized);
  } catch (err) {
    try { console.log("lookupOrder: row-scan failed: " + err); } catch (_) {}
  }

  // --- Half 2: scan Activity Log for matching events ---
  try {
    result.events = _findOrderEvents(normalized);
  } catch (err) {
    try { console.log("lookupOrder: event-scan failed: " + err); } catch (_) {}
  }

  result.found = result.rows.length > 0 || result.events.length > 0;
  result.summary = _summarize(result.rows, result.events);

  return result;
}


/**
 * Sidebar callback: activate All Orders sheet, multi-select the given rows.
 * Mirrors the Alerts card's jumpToAlertRows pattern so the UX feels uniform.
 *
 * @param {Array<number>} rows — 1-based sheet row numbers
 */
function jumpToOrderRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return "ℹ️ No rows to jump to.";
  }

  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet";

  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ All Orders sheet not found";

  ss.setActiveSheet(sheet);

  var ranges = rows.map(function(r) {
    return "A" + r + ":" + _olColLetter(Schema.dataWidth) + r;
  });
  sheet.getRangeList(ranges).activate();

  return "✓ Selected " + rows.length + " row(s) for the order.";
}


// =======================================================================================
// PRIVATE
// =======================================================================================

/**
 * Strips non-alphanumeric chars and uppercases. Lets users paste an order ID
 * with surrounding noise (space, "Order#", etc.) and still get a match.
 * Same normalization applied to sheet values during comparison.
 */
function _normalizeOrderId(s) {
  return String(s == null ? "" : s).toUpperCase().replace(/[^A-Z0-9]/g, "");
}


/**
 * Walks All Orders col A:J, finds rows where SALES_ORDER (col D) contains
 * the normalized query as a substring. Labels each row eBay or DIRECT based
 * on whether it sits above or below the DIRECT boundary row.
 *
 * Skips: rows 1-3 (banner/header), the boundary divider row itself, and the
 * DIRECT table header row immediately following the boundary.
 */
function _findOrderRows(normalizedQuery) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return [];

  var nRows = lastRow - Schema.dataStartRow + 1;
  var data = sheet.getRange(Schema.dataStartRow, 1, nRows, Schema.dataWidth).getValues();

  var boundaryRow = getBoundaryRow();   // sheet-row number of "DIRECT" cell, -1 if missing
  var hits = [];

  var SKU_I        = Schema.idx("SKU");
  var QTY_I        = Schema.idx("QTY");
  var LOC_I        = Schema.idx("LOCATION");
  var SO_I         = Schema.idx("SALES_ORDER");
  var NOTE_I       = Schema.idx("NOTE");
  var STATUS_I     = Schema.idx("STATUS");
  var SHIPPING_I   = Schema.idx("SHIPPING");
  var SHIP_COST_I  = Schema.idx("SHIP_COST");

  for (var i = 0; i < data.length; i++) {
    var sheetRow = Schema.dataStartRow + i;

    // Skip the DIRECT boundary divider row itself
    if (boundaryRow > 0 && sheetRow === boundaryRow) continue;
    // Skip the DIRECT table's column-header row (row immediately after divider)
    if (boundaryRow > 0 && sheetRow === boundaryRow + 1) continue;

    var soNormalized = _normalizeOrderId(data[i][SO_I]);
    if (!soNormalized || soNormalized.indexOf(normalizedQuery) === -1) continue;

    var table = (boundaryRow > 0 && sheetRow > boundaryRow) ? "DIRECT" : "eBay";

    hits.push({
      row:        sheetRow,
      table:      table,
      sku:        String(data[i][SKU_I]        || ""),
      qty:        data[i][QTY_I] || 0,
      location:   String(data[i][LOC_I]        || ""),
      salesOrder: String(data[i][SO_I]         || ""),
      note:       String(data[i][NOTE_I]       || ""),
      status:     String(data[i][STATUS_I]     || ""),
      shipping:   String(data[i][SHIPPING_I]   || ""),
      shipCost:   data[i][SHIP_COST_I] || ""
    });
  }

  return hits;
}


/**
 * Walks Activity Log col A:I, finds events where ORDER_ID (col C) contains
 * the normalized query as a substring. Returns events in chronological order
 * (oldest → newest) so the timeline reads naturally top-to-bottom.
 */
function _findOrderEvents(normalizedQuery) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ACTIVITY_LOG.sheetName);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < ACTIVITY_LOG.dataStartRow) return [];

  var nRows = lastRow - ACTIVITY_LOG.dataStartRow + 1;
  var data = sheet.getRange(ACTIVITY_LOG.dataStartRow, 1, nRows, ACTIVITY_LOG.dataWidth).getValues();

  var TS_I     = ACTIVITY_LOG.idx("TIMESTAMP");
  var EVENT_I  = ACTIVITY_LOG.idx("EVENT");
  var OID_I    = ACTIVITY_LOG.idx("ORDER_ID");
  var SKU_I    = ACTIVITY_LOG.idx("SKU");
  var QTY_I    = ACTIVITY_LOG.idx("QTY");
  var SOURCE_I = ACTIVITY_LOG.idx("SOURCE");
  var DETAIL_I = ACTIVITY_LOG.idx("DETAIL");
  var NOTE_I   = ACTIVITY_LOG.idx("NOTE");
  var PICKER_I = ACTIVITY_LOG.idx("PICKER");

  var hits = [];

  for (var i = 0; i < data.length; i++) {
    var oid = _normalizeOrderId(data[i][OID_I]);
    if (!oid || oid.indexOf(normalizedQuery) === -1) continue;

    var ts = data[i][TS_I];
    hits.push({
      timestamp: (ts instanceof Date) ? ts.getTime() : null,
      event:     String(data[i][EVENT_I]  || ""),
      source:    String(data[i][SOURCE_I] || ""),
      detail:    String(data[i][DETAIL_I] || ""),
      sku:       String(data[i][SKU_I]    || ""),
      qty:       data[i][QTY_I] || "",
      note:      String(data[i][NOTE_I]   || ""),
      picker:    String(data[i][PICKER_I] || "")
    });
  }

  // Sort chronologically — Activity Log is APPEND-ONLY so it's usually already
  // in order, but defensive sort costs nothing on small N.
  hits.sort(function(a, b) {
    return (a.timestamp || 0) - (b.timestamp || 0);
  });

  return hits;
}


/**
 * Builds the summary object — unique statuses, unique pickers, unique SKUs,
 * totals, and earliest/latest event times. Used for the sidebar's header banner.
 *
 * Important: when the order has no current row on All Orders (e.g., shipped
 * weeks ago and cleaned up), the row list is empty. We still want to show
 * status / SKU / picker in the summary so the user can answer a customer's
 * "what happened with this order" question. So we derive those fields from
 * the Activity Log events as a fallback.
 *
 * The status derivation maps each lifecycle event to the status it represents:
 *     RECEIVED → PENDING, PREPARING → PREPARING, SHIPPED → SHIPPED,
 *     CANCELED → CANCELED
 * Then we pick the LATEST such status-bearing event as the order's final state.
 *
 * Precedence: if rows exist, status comes from the sheet (current truth).
 * Only when rows is empty do we derive status from events.
 */
function _summarize(rows, events) {
  var summary = {
    rowCount: rows.length,
    totalQty: 0,
    statuses: [],
    skus: [],
    pickers: [],
    earliestReceived: null,
    latestEvent: null,
    latestEventType: null,
    firstRow: null
  };

  var statusSet = {};
  var skuSet = {};
  var pickerSet = {};

  // --- Pull from rows (sheet state, current truth) ---
  if (rows.length > 0) {
    summary.firstRow = rows[0].row;
    for (var i = 0; i < rows.length; i++) {
      var q = parseInt(rows[i].qty);
      if (!isNaN(q)) summary.totalQty += q;
      var s = String(rows[i].status || "").trim().toUpperCase();
      if (s) statusSet[s] = true;
      var sk = String(rows[i].sku || "").trim();
      if (sk) skuSet[sk] = true;
    }
  }

  // --- Pull from events (timeline) ---
  var EVENT_TO_STATUS = {
    "RECEIVED":  "PENDING",
    "PREPARING": "PREPARING",
    "SHIPPED":   "SHIPPED",
    "CANCELED":  "CANCELED"
  };

  if (events.length > 0) {
    var lastStatusEvent = null;  // most-recent event whose type maps to a status
    for (var j = 0; j < events.length; j++) {
      var p = String(events[j].picker || "").trim();
      if (p) pickerSet[p] = true;
      var esk = String(events[j].sku || "").trim();
      if (esk) skuSet[esk] = true;

      var ts = events[j].timestamp;
      if (ts) {
        if (events[j].event === "RECEIVED") {
          if (summary.earliestReceived == null || ts < summary.earliestReceived) {
            summary.earliestReceived = ts;
          }
        }
        if (summary.latestEvent == null || ts > summary.latestEvent) {
          summary.latestEvent = ts;
          summary.latestEventType = events[j].event;
        }
      }

      // events[] is sorted oldest→newest, so the last EVENT_TO_STATUS hit wins
      if (EVENT_TO_STATUS[events[j].event]) {
        lastStatusEvent = events[j];
      }
    }

    // No current row → derive status from the latest status-bearing event
    if (rows.length === 0 && lastStatusEvent) {
      var derived = EVENT_TO_STATUS[lastStatusEvent.event];
      if (derived) statusSet[derived] = true;
    }
  }

  summary.statuses = Object.keys(statusSet);
  summary.skus     = Object.keys(skuSet);
  summary.pickers  = Object.keys(pickerSet);

  return summary;
}


/** 1 → "A", 26 → "Z", 27 → "AA". Local copy to avoid colliding with Alerts._colLetter. */
function _olColLetter(n) {
  var s = "";
  while (n > 0) {
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
