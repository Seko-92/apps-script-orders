/**
 * Reads the last n8n sync timestamp from the banner cell on the All Orders sheet.
 * The sidebar calls this on load (and on every poll) so the displayed
 * "Last sync" reflects sheet truth, not a localStorage guess.
 *
 * Returns the raw cell text — e.g. "⏱ Last sync · 8:45 PM" — or empty string
 * if the cell is empty.
 */
function getLastSyncFromSheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return "";
    var v = sheet.getRange(Schema.cellSyncTime).getValue();
    return v ? String(v) : "";
  } catch (e) {
    console.error("getLastSyncFromSheet error: " + e);
    return "";
  }
}


/**
 * Fetches the latest API usage from the "API Usage" sheet.
 *
 * Sheet layout (n8n-driven, populated by the API Usage Monitor workflow,
 * which RESETS the sheet on every run — only the latest snapshot is kept):
 *   Run Date | Run Time | API Type | API Name | Endpoint
 *   | Daily Limit | Used | Remaining | Usage % | Reset Time | Status
 *
 * Returns four endpoint-level metrics + worstPct for the header indicator:
 *   {
 *     tradingApi:  { pct, used, limit, label, endpoint },  ← shared quota across
 *                                                            all Name="TradingAPI"
 *                                                            rows; eBay's analytics
 *                                                            endpoint reports the
 *                                                            same total for every
 *                                                            Trading endpoint
 *     fulfillment: { pct, used, limit, label, endpoint },  ← sell.fulfillment row
 *     feed:        { pct, used, limit, label, endpoint },  ← sell.feed row
 *     analytics:   { pct, used, limit, label, endpoint },  ← developer.analytics.*
 *     worstPct:    number  ← max of the four; drives header status indicator
 *     tradingMonitorAvailable: boolean
 *   }
 *
 * Why the Trading-API number is meaningful now (it wasn't on a previous run):
 * eBay's `developer.analytics.app_rate_limit` endpoint reports a SINGLE shared
 * Trading-API quota (5,000 calls/day) across every Trading operation. The n8n
 * workflow exposes that by emitting one row per Trading endpoint, all with
 * identical Used/Limit values — so picking any single TradingAPI row gives us
 * the real number. That's why GetItem shows "3011/5000 60%" and so does
 * AddOrder, RevokeToken, etc. — they all share the pool.
 *
 * Type="Trading" / Name="Trading API" rows (legacy "Check Portal" entries) are
 * deliberately skipped — those are placeholders, not real measurements.
 */
function getLatestApiMetrics() {
  var EMPTY = {
    tradingApi:  { pct: 0, used: 0, limit: 5000,    label: 'Trading API',    endpoint: 'shared quota'        },
    fulfillment: { pct: 0, used: 0, limit: 100000,  label: 'Fulfillment',    endpoint: 'sell.fulfillment'    },
    feed:        { pct: 0, used: 0, limit: 100000,  label: 'Inventory Feed', endpoint: 'sell.feed'           },
    analytics:   { pct: 0, used: 0, limit: 5000,    label: 'Analytics',      endpoint: 'developer.analytics' },
    worstPct: 0,
    tradingMonitorAvailable: false
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("API Usage");
    if (!sheet) return EMPTY;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return EMPTY;

    var headers = data[0].map(function(h) { return String(h).toLowerCase().trim(); });
    var typeIdx     = _findHeader(headers, ['api type', 'apitype', 'type']);
    var nameIdx     = _findHeader(headers, ['api name', 'apiname', 'name']);
    var endpointIdx = _findHeader(headers, ['endpoint']);
    var usageIdx    = _findHeader(headers, ['usage %', 'usage%', 'usagepercent', 'usage_percent']);
    var usedIdx     = _findHeader(headers, ['used']);
    var limitIdx    = _findHeader(headers, ['daily limit', 'dailylimit', 'limit']);

    if (endpointIdx < 0 || usageIdx < 0) return EMPTY;

    var result = {
      tradingApi:  { pct: 0, used: 0, limit: 5000,    label: 'Trading API',    endpoint: 'shared quota'        },
      fulfillment: { pct: 0, used: 0, limit: 100000,  label: 'Fulfillment',    endpoint: 'sell.fulfillment'    },
      feed:        { pct: 0, used: 0, limit: 100000,  label: 'Inventory Feed', endpoint: 'sell.feed'           },
      analytics:   { pct: 0, used: 0, limit: 5000,    label: 'Analytics',      endpoint: 'developer.analytics' },
      worstPct: 0,
      tradingMonitorAvailable: false
    };

    for (var i = 1; i < data.length; i++) {
      var type     = typeIdx     >= 0 ? String(data[i][typeIdx]     || '').trim() : '';
      var name     = nameIdx     >= 0 ? String(data[i][nameIdx]     || '').trim() : '';
      var endpoint =                    String(data[i][endpointIdx] || '').trim();

      // Parse usage %. Sheet may store as string ("60%"), number percent (60),
      // or fraction (0.6) depending on cell format. Handle all three.
      var pct = _parsePercent(data[i][usageIdx]);
      if (isNaN(pct)) continue;

      var used  = usedIdx  >= 0 ? (parseFloat(data[i][usedIdx])  || 0) : 0;
      var limit = limitIdx >= 0 ? (parseFloat(data[i][limitIdx]) || 0) : 0;

      // Skip the legacy Type="Trading" rows — those report "Check Portal" / N/A
      if (type.toLowerCase() === 'trading') continue;

      // Trading API (shared quota across all Name="TradingAPI" rows).
      // All such rows report identical Used/Limit values — eBay reports
      // the SHARED quota total per row. Take the highest-pct row we see;
      // it'll always be the same number anyway, but defensive max is safer.
      if (name === 'TradingAPI') {
        if (pct > result.tradingApi.pct) {
          result.tradingApi.pct   = pct;
          result.tradingApi.used  = used;
          result.tradingApi.limit = limit || 5000;
          result.tradingApi.endpoint = used && limit
            ? 'shared · ' + Math.round(used) + '/' + Math.round(limit)
            : 'shared quota';
          result.tradingMonitorAvailable = true;
        }
        continue;
      }

      // sell.fulfillment — orders workflow's REST endpoint
      if (endpoint === 'sell.fulfillment') {
        if (pct >= result.fulfillment.pct) {
          result.fulfillment.pct   = pct;
          result.fulfillment.used  = used;
          result.fulfillment.limit = limit || 100000;
        }
        continue;
      }

      // sell.feed — LMS inventory feed
      if (endpoint === 'sell.feed') {
        if (pct >= result.feed.pct) {
          result.feed.pct   = pct;
          result.feed.used  = used;
          result.feed.limit = limit || 100000;
        }
        continue;
      }

      // developer.analytics.* (developer.analytics.app_rate_limit etc.)
      if (endpoint.indexOf('developer.analytics') === 0) {
        if (pct > result.analytics.pct) {
          result.analytics.pct   = pct;
          result.analytics.used  = used;
          result.analytics.limit = limit || 5000;
        }
        continue;
      }
    }

    // Worst-case across the four — drives the header status pill color tier.
    result.worstPct = Math.max(
      result.tradingApi.pct,
      result.fulfillment.pct,
      result.feed.pct,
      result.analytics.pct
    );

    return result;
  } catch (e) {
    console.error("Failed to fetch API metrics: " + e.toString());
    return EMPTY;
  }
}

/**
 * Parse a usage-percent value that may arrive in any of three forms:
 *   "60%"  (string with % suffix)         → 60
 *   60     (number, already percent)       → 60
 *   0.6    (number, fraction 0..1)         → 60
 *   "N/A"  / ""                            → NaN (caller skips)
 */
function _parsePercent(raw) {
  if (raw == null) return NaN;
  if (typeof raw === 'number') {
    // 0..1 = fraction, treat as percent
    if (raw > 0 && raw <= 1) return raw * 100;
    return raw;
  }
  var s = String(raw).replace('%', '').trim();
  if (!s || s.toUpperCase() === 'N/A') return NaN;
  var n = parseFloat(s);
  if (isNaN(n)) return NaN;
  // String like "0.6" → fraction; "60" → already percent
  if (n > 0 && n <= 1 && s.indexOf('.') >= 0) return n * 100;
  return n;
}

function _findHeader(headers, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var idx = headers.indexOf(candidates[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}
