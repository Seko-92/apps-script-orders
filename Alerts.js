// =======================================================================================
// ALERTS.gs — Actionable alerts surface for the sidebar///
// =======================================================================================
//
// PURPOSE
//   Turns the sidebar from "panel of buttons" into "panel that tells you
//   what needs your attention right now." Single sheet scan → counts of
//   actionable items the warehouse should look at.
//
//   Each alert is mapped to a clickable row in the sidebar. Click jumps the
//   user's view to those rows in the All Orders sheet (or opens Prep Queue
//   for the queue-size alert).
//
// ALERTS DETECTED (eBay + DIRECT, terminal-state rows excluded)
//   1. paidShipping — non-zero SHIP_COST, not yet SHIPPED. Warehouse must
//                     print labels at the exact service level.
//   2. intl         — SHIPPING contains "[INTL]", not yet SHIPPED. Customs
//                     paperwork required.
//   3. lowStock     — HAND ≤ 20 in PENDING/PREPARING rows. Restock window.
//   4. notFound     — LOCATION = "NOT FOUND" (SKU not in Master Inventory).
//                     Data quality issue — fix before fulfillment errors.
//   5. queueSize    — Items currently in the Prep Queue sheet.
//   6. outOfStock   — SKUs in the Out of Stock sheet (Master Inventory rows
//                     where quantity - quantitySold ≤ 0). Refreshed weekly
//                     by refreshOutOfStock(); the alert reads the snapshot,
//                     not Master Inventory directly — cheap on every poll.
//   7. newFromZoho  — Zoho direct-channel SOs in the Pending Sales Orders
//                     sheet whose PULLED column is blank. Click → opens the
//                     Pending sheet so the picker can decide which to pull.
//
// PERFORMANCE
//   _scanAlerts() reads the All Orders sheet ONCE and partitions rows in a
//   single pass. Called from getActionableAlerts (polled every 30s) and
//   jumpToAlertRows (on user click).
//
// PUBLIC API
//   getActionableAlerts() — { paidShipping:{count,rows}, intl:{...}, lowStock:{...},
//                              notFound:{...}, queueSize:{...}, outOfStock:{...} }
//   jumpToAlertRows(key)  — activates All Orders, multi-selects matching rows
//                            (queueSize/outOfStock open their dedicated sheet instead)
// =======================================================================================

/**
 * Reads All Orders once and partitions rows into the four data-driven alert
 * categories. Skips boundary rows and terminal-state rows (SHIPPED/CANCELED
 * orders aren't actionable — they're done).
 *
 * Returns { paidShipping: [rowNumbers], intl: [...], lowStock: [...], notFound: [...] }
 */
function _scanAlerts() {
  var out = { paidShipping: [], intl: [], lowStock: [], notFound: [] };

  var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return out;

  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return out;

  var data = sheet.getRange(
    Schema.dataStartRow, 1,
    lastRow - Schema.dataStartRow + 1,
    Schema.dataWidth
  ).getValues();

  var boundary = getBoundaryRow();

  for (var i = 0; i < data.length; i++) {
    var rowNum = Schema.dataStartRow + i;

    // PRIMARY skip: boundary divider + DIRECT header row (by row number)
    if (boundary > 0 && (rowNum === boundary || rowNum === boundary + 1)) continue;

    var sku = String(data[i][Schema.idx("SKU")]).trim();
    if (!sku) continue;

    // BELT-AND-SUSPENDERS skip: any header-like row that leaked past the
    // boundary check. The DIRECT header's column J literally contains the
    // text "SHIP COST" — without this, an undetected header row would parse
    // as paid-shipping false positive (which is the bug we're fixing).
    var skuUpper = sku.toUpperCase();
    if (skuUpper === Schema.boundaryMarker) continue;       // "DIRECT" divider value
    if (skuUpper.indexOf('◈') !== -1) continue;              // header glyph from "◈ SKU"
    if (skuUpper === 'SKU' || skuUpper === '# SKU' || skuUpper === '◈ SKU') continue;

    var status = String(data[i][Schema.idx("STATUS")]).trim().toUpperCase();
    if (Schema.isTerminal(status)) continue;  // SHIPPED/CANCELED — done, not actionable

    var location = String(data[i][Schema.idx("LOCATION")]).trim();
    var hand     = data[i][Schema.idx("HAND")];
    var shipping = String(data[i][Schema.idx("SHIPPING")]).trim();
    var shipCost = data[i][Schema.idx("SHIP_COST")];

    // Paid Shipping: parse as a dollar amount, flag only if > 0.
    // This rejects "FREE", "SHIP COST" (header label), "", and zero values
    // without needing a hardcoded literal-string blocklist.
    var shipCostStr = String(shipCost == null ? "" : shipCost).trim();
    var shipCostNum = parseFloat(shipCostStr.replace(/[^0-9.\-]/g, ''));
    if (!isNaN(shipCostNum) && shipCostNum > 0) {
      out.paidShipping.push(rowNum);
    }

    // International: [INTL] flag in shipping service (case-insensitive)
    if (shipping.toUpperCase().indexOf('[INTL]') !== -1) {
      out.intl.push(rowNum);
    }

    // Low Stock: HAND is numeric and ≤ 20
    if (typeof hand === 'number' && hand <= 20) {
      out.lowStock.push(rowNum);
    }

    // NOT FOUND: location lookup failed (case-insensitive match)
    if (location.toUpperCase() === 'NOT FOUND') {
      out.notFound.push(rowNum);
    }
  }

  return out;
}

/**
 * Counts non-empty SKU rows in the Prep Queue sheet.
 * Defensive: returns 0 if sheet doesn't exist yet.
 */
function _getPrepQueueSize() {
  var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < PREP_QUEUE.dataStartRow) return 0;

  var skus = sheet.getRange(
    PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.SKU,
    lastRow - PREP_QUEUE.dataStartRow + 1, 1
  ).getValues();

  var count = 0;
  for (var i = 0; i < skus.length; i++) {
    var v = String(skus[i][0]).trim();
    if (!v) continue;
    // Two-table layout (2026-07-16): the INCOMING divider + its header row
    // hold text in col A but are structure, not items. Both tables' real
    // rows count toward the queue size (total prep workload).
    if (v.toUpperCase() === PREP_QUEUE.boundaryMarker || v.charAt(0) === '◈') continue;
    count++;
  }
  return count;
}

/**
 * Sidebar entry point — returns counts (and matching row numbers, in case
 * the sidebar wants to display them).
 */
function getActionableAlerts() {
  try {
    var alerts = _scanAlerts();
    return {
      paidShipping: { count: alerts.paidShipping.length, rows: alerts.paidShipping },
      intl:         { count: alerts.intl.length,         rows: alerts.intl },
      lowStock:     { count: alerts.lowStock.length,     rows: alerts.lowStock },
      notFound:     { count: alerts.notFound.length,     rows: alerts.notFound },
      queueSize:    { count: _getPrepQueueSize(),        rows: [] },
      outOfStock:   { count: getOutOfStockCount(),       rows: [] },
      newFromZoho:  { count: _safeZohoCount(),           rows: [] }
    };
  } catch (e) {
    console.error("getActionableAlerts error: " + e);
    return {
      paidShipping: { count: 0, rows: [] },
      intl:         { count: 0, rows: [] },
      lowStock:     { count: 0, rows: [] },
      notFound:     { count: 0, rows: [] },
      queueSize:    { count: 0, rows: [] },
      outOfStock:   { count: 0, rows: [] },
      newFromZoho:  { count: 0, rows: [] }
    };
  }
}

/** Defensive wrapper — getPendingZohoCount lives in ZohoSalesOrders.js. If
 *  that file isn't loaded yet or the sheet doesn't exist, we silently return 0
 *  rather than breaking the entire alerts payload.
 */
function _safeZohoCount() {
  try {
    return (typeof getPendingZohoCount === "function") ? getPendingZohoCount() : 0;
  } catch (e) {
    return 0;
  }
}

/**
 * Click handler: activate All Orders sheet (or Prep Queue for queueSize),
 * select all matching rows so the user can see the cluster at once.
 *
 * Uses sheet.getRangeList(...).activate() which highlights non-contiguous
 * ranges with a single visible selection.
 */
function jumpToAlertRows(alertKey) {
  // queueSize / outOfStock / newFromZoho jump to their own sheets (not All Orders)
  if (alertKey === 'queueSize')   return openPrepQueue();
  if (alertKey === 'outOfStock')  return openOutOfStock();
  if (alertKey === 'newFromZoho') return openPendingSalesOrders();

  var alerts = _scanAlerts();
  var rows = alerts[alertKey];
  if (!rows || rows.length === 0) return "ℹ️ Nothing to jump to.";

  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet";

  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  ss.setActiveSheet(sheet);

  // Multi-select all matching data rows (full row width). The first one
  // in the list becomes the visual focus.
  var ranges = rows.map(function(r) {
    return "A" + r + ":" + _colLetter(Schema.dataWidth) + r;
  });
  sheet.getRangeList(ranges).activate();

  return "✓ " + rows.length + " row(s) selected";
}

/** Tiny helper: 1 → "A", 26 → "Z", 27 → "AA". (We only need up to "J" today.) */
function _colLetter(n) {
  var s = "";
  while (n > 0) {
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
