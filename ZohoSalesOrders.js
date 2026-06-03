// =======================================================================================
// ZOHO_SALES_ORDERS.gs — Pending sheet mirror + Invoice webhook + VOID handler
// =======================================================================================
//
// PURPOSE
//   Mirrors Zoho's direct-channel sales orders and invoices into the "Pending
//   Sales Orders" sheet in real time via webhook. The picker brings each SO
//   into the DIRECT table via the explicit Pull modal (see ZohoPull.js).
//
// DATA FLOW
//   Zoho (SO created/edited) → n8n proxy → doPost (action=zohoSalesOrder)
//                                            ↓
//                                     upsertPendingSalesOrder()
//                                            ↓
//                                [Pending Sales Orders sheet] — always current
//                                            ↓ (sidebar Pull → opens modal)
//                                     [ZohoPull.js / ZohoPullModal.html]
//                                            ↓ (picker reviews diff, hits Apply)
//                                     [DIRECT table on All Orders]
//
//   Zoho status: void  →  _handleZohoVoid() — flips DIRECT rows to CANCELED
//                          (the one signal we still propagate automatically;
//                          customer cancellation is unambiguous + time-sensitive,
//                          picker action is identical regardless of source)
//
// RETIRED (2026-05-23, Option C build)
//   The background _propagateToDirectRows feature (line-item add auto-insert,
//   qty-change flag, removal flag, SHIPPED auto-flip) was deleted. Every
//   recurring "phantom row" / "first-fire duplicate" / row-shift class of
//   bug traced back to that implicit background reconciler. Replaced with
//   the picker-driven Pull modal which puts a human in the loop for every
//   Zoho→DIRECT change. Git history is the safety net.
//
//   Also retired: pullSalesOrderToDirect() (the old direct-commit pull) and
//   _deleteDirectRowsForSo() (force re-pull helper). The Pull modal subsumes
//   both via per-line selective apply.
//
// FILTERING (webhook level — skip ingest entirely)
//   - sales_channel must equal "direct_sales" (eBay/Amazon SOs come through
//     the n8n eBay workflow already)
//   - is_test_order must be false (Zoho's own test-order flag)
//
// SCHEMA (Pending Sales Orders sheet)
//   SO# · Customer · Date · Order · Payment · Shipment · Items · Total
//   · Last Updated · Pulled? · Pulled At · Payload (hidden JSON cache)
//   · Invoice · Price Check
// =======================================================================================


var PENDING_SO = {
  sheetName: "Pending Sales Orders",

  cols: {
    SO_NUMBER:    1,   // A
    CUSTOMER:     2,   // B
    DATE:         3,   // C
    ORDER_STATUS: 4,   // D — Confirmed / Closed / Void / Draft (mirrored from Zoho)
    PAYMENT:      5,   // E
    SHIPMENT:     6,   // F
    ITEMS_COUNT:  7,   // G
    TOTAL:        8,   // H
    LAST_UPDATED: 9,   // I
    PULLED:      10,   // J — "PULLED" when pulled, blank otherwise
    PULLED_AT:   11,   // K — Date when pulled
    PAYLOAD:     12,   // L — cached JSON of last Zoho payload (hidden by narrow width)
    INVOICE:     13,   // M — Zoho invoice number (e.g. "INV-022496"), stamped by Invoice webhook
    PRICE_CHECK: 14    // N — Zoho line-item rates vs MI.currentPrice summary
  },

  idx: function(name) { return PENDING_SO.cols[name] - 1; },

  dataWidth:    14,
  headerRow:    1,
  dataStartRow: 2,

  headers: ["SO #", "CUSTOMER", "DATE", "ORDER", "PAYMENT", "SHIPMENT",
            "ITEMS", "TOTAL", "⏱ UPDATED", "PULLED?", "PULLED AT", "_PAYLOAD",
            "INVOICE #", "PRICE CHECK"],

  pulledFlag: "PULLED",

  // Thresholds for the price check — only flag mismatches that are large
  // enough to matter, suppressing rounding noise.
  // A line item is "off" when |zohoPrice - ebayPrice| > max(absThreshold, pctThreshold × ebayPrice).
  // $1.00 OR 2% of price, whichever is greater. Default values are
  // conservative; tune in this object if needed (no other code reads them).
  priceCheck: { absThreshold: 1.00, pctThreshold: 0.02 }
};


// =======================================================================================
// SETUP — idempotent
// =======================================================================================

/**
 * Creates the Pending Sales Orders sheet if missing, applies brand styling.
 * Safe to re-run — preserves existing data, just refreshes formatting.
 */
function setupPendingSalesOrdersSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PENDING_SO.sheetName);
  if (!sheet) sheet = ss.insertSheet(PENDING_SO.sheetName);

  // --- HEADERS ---
  sheet.getRange(PENDING_SO.headerRow, 1, 1, PENDING_SO.dataWidth)
    .setValues([PENDING_SO.headers])
    .setBackground('#1d1d1b')
    .setFontColor('#ffd966')
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(PENDING_SO.headerRow, 1, 1, PENDING_SO.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(PENDING_SO.cols.SO_NUMBER,    100);
  sheet.setColumnWidth(PENDING_SO.cols.CUSTOMER,     200);
  sheet.setColumnWidth(PENDING_SO.cols.DATE,         100);
  sheet.setColumnWidth(PENDING_SO.cols.ORDER_STATUS, 100);
  sheet.setColumnWidth(PENDING_SO.cols.PAYMENT,      110);
  sheet.setColumnWidth(PENDING_SO.cols.SHIPMENT,     110);
  sheet.setColumnWidth(PENDING_SO.cols.ITEMS_COUNT,   60);
  sheet.setColumnWidth(PENDING_SO.cols.TOTAL,        100);
  sheet.setColumnWidth(PENDING_SO.cols.LAST_UPDATED, 140);
  sheet.setColumnWidth(PENDING_SO.cols.PULLED,        80);
  sheet.setColumnWidth(PENDING_SO.cols.PULLED_AT,    140);
  // Payload col: very narrow + hide it (it's a JSON blob, not for humans)
  sheet.setColumnWidth(PENDING_SO.cols.PAYLOAD,       20);
  try { sheet.hideColumns(PENDING_SO.cols.PAYLOAD); } catch (e) { /* already hidden */ }
  sheet.setColumnWidth(PENDING_SO.cols.INVOICE,      120);
  sheet.setColumnWidth(PENDING_SO.cols.PRICE_CHECK,  150);

  // --- DATA AREA FORMATS ---
  var maxDataRow = 2000;
  var dataRows = maxDataRow - PENDING_SO.dataStartRow + 1;

  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.SO_NUMBER, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.CUSTOMER, dataRows, 1)
    .setFontFamily('Roboto').setFontSize(10).setHorizontalAlignment('left');
  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.DATE, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9).setHorizontalAlignment('center');
  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.ORDER_STATUS, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.PAYMENT, dataRows, 1)
    .setFontFamily('Oswald').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.SHIPMENT, dataRows, 1)
    .setFontFamily('Oswald').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.ITEMS_COUNT, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.TOTAL, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.LAST_UPDATED, dataRows, 1)
    .setNumberFormat('M/d/yy h:mm am/pm')
    .setFontFamily('Roboto Mono').setFontSize(9).setFontColor('#5f5f5f').setHorizontalAlignment('center');
  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.PULLED, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.PULLED_AT, dataRows, 1)
    .setNumberFormat('M/d/yy h:mm am/pm')
    .setFontFamily('Roboto Mono').setFontSize(9).setFontColor('#5f5f5f').setHorizontalAlignment('center');
  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.INVOICE, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.PRICE_CHECK, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');

  sheet.getRange(PENDING_SO.dataStartRow, 1, dataRows, PENDING_SO.dataWidth)
    .setVerticalAlignment('middle');

  // --- BANDING (cream alternation) ---
  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  var bandRange = sheet.getRange(1, 1, maxDataRow, PENDING_SO.dataWidth);
  var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b')
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#fff8e7');

  // --- CONDITIONAL FORMATTING ---
  // PULLED column tints green when populated; ORDER status tints by value;
  // PRICE_CHECK tints green (✓ prefix) or amber/orange (⚠ prefix) by direction.
  var existingRules = sheet.getConditionalFormatRules() || [];
  var keep = existingRules.filter(function(r) {
    var ranges = r.getRanges();
    if (!ranges || ranges.length === 0) return true;
    return !ranges.some(function(rg) {
      if (rg.getSheet().getName() !== PENDING_SO.sheetName) return false;
      var c = rg.getColumn();
      return c === PENDING_SO.cols.PULLED
          || c === PENDING_SO.cols.ORDER_STATUS
          || c === PENDING_SO.cols.PRICE_CHECK;
    });
  });

  var pulledRange = sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.PULLED, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(PENDING_SO.pulledFlag)
    .setBackground('#e8f5e9').setFontColor('#1b5e20').setBold(true)
    .setRanges([pulledRange]).build());

  var orderRange = sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.ORDER_STATUS, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('CONFIRMED')
    .setBackground('#fff8e7').setFontColor('#1d1d1b').setBold(true)
    .setRanges([orderRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('CLOSED')
    .setBackground('#e8f5e9').setFontColor('#1b5e20').setBold(true)
    .setRanges([orderRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('VOID')
    .setBackground('#f0f0f0').setFontColor('#5f5f5f').setBold(true)
    .setRanges([orderRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('DRAFT')
    .setBackground('#f5f5f5').setFontColor('#888888').setBold(false)
    .setRanges([orderRange]).build());

  // PRICE_CHECK column — distinguishes three states via prefix-match:
  //   "✓ ..."   = all line items within threshold → soft green
  //   "⚠ Zoho HIGH ..." = Zoho > eBay (customer-complaint case) → amber-yellow
  //   "⚠ Zoho LOW ..."  = Zoho < eBay (under-quoting case)       → orange
  //   "⚠ NOT FOUND ..." = couldn't look up some SKUs in MI       → soft gray
  // Empty cell renders plain (no rule fires).
  var priceRange = sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.PRICE_CHECK, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith('✓')
    .setBackground('#e8f5e9').setFontColor('#1b5e20').setBold(true)
    .setRanges([priceRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Zoho HIGH')
    .setBackground('#fff4b0').setFontColor('#7d5d00').setBold(true)
    .setRanges([priceRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Zoho LOW')
    .setBackground('#ffd699').setFontColor('#7a3d00').setBold(true)
    .setRanges([priceRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('NOT FOUND')
    .setBackground('#f0f0f0').setFontColor('#5f5f5f').setBold(false)
    .setRanges([priceRange]).build());

  sheet.setConditionalFormatRules(keep);

  sheet.setFrozenRows(1);

  return "✅ Pending Sales Orders sheet ready.";
}


/** Sidebar: switch view to Pending Sales Orders sheet. */
function openPendingSalesOrders() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet";
  var sheet = ss.getSheetByName(PENDING_SO.sheetName);
  if (!sheet) {
    setupPendingSalesOrdersSheet();
    sheet = ss.getSheetByName(PENDING_SO.sheetName);
  }
  ss.setActiveSheet(sheet);
  return "✅ Opened " + PENDING_SO.sheetName;
}


// =======================================================================================
// WEBHOOK ENTRY POINT — called from doPost (action=zohoSalesOrder)
// =======================================================================================

/**
 * Upserts the Pending row for this SO from the Zoho payload. If the SO was
 * already pulled to DIRECT, runs the asymmetric propagation rule:
 *   - new line items → auto-insert as new DIRECT rows
 *   - removed items  → flag existing DIRECT rows (strikethrough + NOTE)
 *   - qty changes    → flag existing DIRECT rows
 *   - status: void   → flip DIRECT rows to CANCELED
 *
 * Returns { status, soNumber, actionTaken, propagated }.
 *
 * NEVER throws — webhook errors return as JSON to n8n. A logging failure
 * inside propagation never rolls back the Pending sheet upsert.
 */
function upsertPendingSalesOrder(salesorder) {
  if (!salesorder || typeof salesorder !== 'object') {
    return { status: "skipped", reason: "Empty or invalid payload", soNumber: "" };
  }

  var soNumber = String(salesorder.salesorder_number || "").trim();
  if (!soNumber) {
    return { status: "skipped", reason: "No salesorder_number in payload", soNumber: "" };
  }

  // --- WEBHOOK-LEVEL FILTERS ---
  // Only direct-channel SOs reach the Pending sheet. eBay/Amazon SOs are
  // already handled by the n8n eBay workflow; mirroring them again would
  // double-populate the working sheet.
  var channel = String(salesorder.sales_channel || "").trim().toLowerCase();
  if (channel && channel !== "direct_sales") {
    return { status: "skipped", reason: "Non-direct channel: " + channel, soNumber: soNumber };
  }
  if (salesorder.is_test_order === true) {
    return { status: "skipped", reason: "Zoho test order", soNumber: soNumber };
  }

  // --- ENSURE SHEET EXISTS ---
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PENDING_SO.sheetName);
  if (!sheet) {
    setupPendingSalesOrdersSheet();
    sheet = ss.getSheetByName(PENDING_SO.sheetName);
  }

  // --- LOCATE EXISTING ROW (if any) ---
  var existingRowNum = _findPendingRow(sheet, soNumber);
  var oldPayload = null;
  var wasPulled = false;
  if (existingRowNum > 0) {
    var existingValues = sheet.getRange(existingRowNum, 1, 1, PENDING_SO.dataWidth).getValues()[0];
    wasPulled = String(existingValues[PENDING_SO.idx("PULLED")] || "").trim().toUpperCase()
                === PENDING_SO.pulledFlag;
    var rawJson = String(existingValues[PENDING_SO.idx("PAYLOAD")] || "");
    if (rawJson) {
      try { oldPayload = JSON.parse(rawJson); } catch (e) { oldPayload = null; }
    }
  }

  // --- BUILD ROW VALUES FROM PAYLOAD ---
  var rowValues = _payloadToRow(salesorder);

  // Preserve PULLED + PULLED_AT + INVOICE on existing row (upsert doesn't un-pull
  // or wipe the invoice number; INVOICE is stamped by the separate invoice webhook
  // and must survive subsequent SO-edit fires).
  if (existingRowNum > 0) {
    var preservePulled   = sheet.getRange(existingRowNum, PENDING_SO.cols.PULLED).getValue();
    var preservePulledAt = sheet.getRange(existingRowNum, PENDING_SO.cols.PULLED_AT).getValue();
    var preserveInvoice  = sheet.getRange(existingRowNum, PENDING_SO.cols.INVOICE).getValue();
    rowValues[PENDING_SO.idx("PULLED")]    = preservePulled;
    rowValues[PENDING_SO.idx("PULLED_AT")] = preservePulledAt;
    rowValues[PENDING_SO.idx("INVOICE")]   = preserveInvoice;
  }

  // --- WRITE TO SHEET ---
  var targetRow;
  if (existingRowNum > 0) {
    targetRow = existingRowNum;
    sheet.getRange(existingRowNum, 1, 1, PENDING_SO.dataWidth).setValues([rowValues]);
  } else {
    // Insert at the top (just below header) so newest is visible first
    sheet.insertRowsBefore(PENDING_SO.dataStartRow, 1);
    targetRow = PENDING_SO.dataStartRow;
    sheet.getRange(targetRow, 1, 1, PENDING_SO.dataWidth).setValues([rowValues]);
  }

  var actionTaken = (existingRowNum > 0) ? "updated" : "inserted";

  // --- PRICE CHECK — stamp col N after write (needs MI lookup, not in _payloadToRow) ---
  // Computed on every webhook fire so the flag stays current as Zoho line
  // items change AND as MI's currentPrice gets refreshed by Inventory Lite Sync.
  // Best-effort: a price-check failure doesn't roll back the row write.
  try {
    var priceCheck = _computePriceCheck(salesorder);
    sheet.getRange(targetRow, PENDING_SO.cols.PRICE_CHECK).setValue(priceCheck.summary);
  } catch (priceErr) {
    console.log("upsertPendingSalesOrder: price check stamp failed for " + soNumber + ": " + priceErr);
  }

  // --- AUTO-LINK — RETIRED 2026-05-20 ---
  // The auto-link feature inferred PULLED status when DIRECT rows existed for
  // an SO that hadn't been explicitly Pulled (handled the manual-entry-first
  // workflow). It was the source of multiple subtle bugs:
  //   - First-fire diff with null oldPayload → false duplicate inserts
  //   - Mid-day state inference unpredictable across webhook timing
  //   - Hard to reason about which Pending rows were "really pulled"
  //
  // Retired in favor of an explicit rule: ONLY rows where the picker
  // explicitly clicked Pull receive Zoho-driven propagation. Manual entries
  // in DIRECT are treated as standalone — if the picker wants Zoho sync on
  // those, they Pull explicitly (Force-Pull if needed to dedupe).
  //
  // The Pending sheet still gets the mirrored SO data on every webhook
  // (useful for customer-service lookups), but propagation to DIRECT only
  // fires when PULLED is truly set. Single source of truth, no inference.
  //
  // autoLinked stays in the return shape for back-compat with any client
  // checking it; always false going forward.
  var autoLinked = false;

  // --- AUTO-PROPAGATION: VOID + SHIPPED (the two unambiguous Zoho signals) ---
  // The full line-item propagation (_propagateToDirectRows) was retired
  // 2026-05-23 in favor of the picker-driven Pull modal. The two signals that
  // still propagate automatically are status=void → CANCELED and
  // shipped_status=shipped/fulfilled → SHIPPED — both unambiguous, time-
  // sensitive, and picker-action-identical regardless of who triggered them.
  // PREPARING stays employee-driven. Every other Zoho→DIRECT data movement
  // goes through the Pull modal.
  //
  // 2026-06-02: these run REGARDLESS of wasPulled — so a Direct SO typed into
  // DIRECT by hand (out of habit, e.g. a quick one-item order) still reflects
  // Zoho's shipped/void state. This is SAFE where the old auto-link was not:
  // both are pure status flips on EXISTING rows matched by SO number
  // (updateOrderStatus) — they never insert/diff/delete, so the reconciler bug
  // class doesn't apply. If no DIRECT row carries this SO, updateOrderStatus
  // flips nothing (safe no-op). Each handler self-gates on the status value
  // (early-return otherwise) so updateOrderStatus only runs on actual
  // shipped/void fires. Each is wrapped so one failing can't block the other
  // or the row write. (wasPulled is still computed + returned for back-compat.)
  var propagated = {};
  try {
    propagated.voided = _handleZohoVoid(soNumber, salesorder);
  } catch (propErr) {
    console.log("Zoho void handler error for " + soNumber + ": " + propErr);
    propagated.voided = { error: propErr.toString() };
  }
  try {
    propagated.shipped = _handleZohoShipped(soNumber, salesorder);
  } catch (shipErr) {
    console.log("Zoho shipped handler error for " + soNumber + ": " + shipErr);
    propagated.shipped = { error: shipErr.toString() };
  }

  return {
    status:      "success",
    soNumber:    soNumber,
    actionTaken: actionTaken,
    wasPulled:   wasPulled,
    autoLinked:  autoLinked,
    propagated:  propagated
  };
}


// =======================================================================================
// WEBHOOK ENTRY POINT — invoice (called from doPost when payload.invoice present)
// =======================================================================================

/**
 * Stamps the Zoho invoice number onto the matching Pending row's INVOICE column.
 *
 * Fired by a separate Zoho Workflow Rule on the INVOICE module (Created + Edited),
 * pointing at the same n8n proxy as the SO rule. The proxy hardcodes
 * action=zohoSalesOrder; the doPost dispatcher sniffs payload.invoice to route here.
 *
 * Linkage strategy:
 *   1. PRIMARY:  invoice.salesorder_id → scan PAYLOAD col for cached SO with matching id
 *   2. FALLBACK: invoice.reference_number (Zoho stamps the SO# here) → scan SO_NUMBER col
 *
 * Why two strategies: salesorder_id is the bulletproof Zoho-internal link
 * (always present in invoice payloads). reference_number is a human-set field
 * that Zoho usually populates with the SO# automatically, but power-users can
 * change it. Using both = belt-and-suspenders.
 *
 * Filter: sales_channel === "direct_sales" (same as SO). Invoices on eBay/Amazon
 * SOs would never have a matching Pending row anyway.
 *
 * Returns { status, invoiceNumber, soNumber?, actionTaken, reason? }.
 * NEVER throws — webhook errors return as JSON to n8n.
 */
function upsertInvoiceFromZoho(invoice) {
  if (!invoice || typeof invoice !== 'object') {
    return { status: "skipped", reason: "Empty or invalid invoice payload" };
  }

  var invoiceNumber = String(invoice.invoice_number || "").trim();
  if (!invoiceNumber) {
    return { status: "skipped", reason: "No invoice_number in payload" };
  }

  // --- WEBHOOK-LEVEL FILTER ---
  var channel = String(invoice.sales_channel || "").trim().toLowerCase();
  if (channel && channel !== "direct_sales") {
    return { status: "skipped", invoiceNumber: invoiceNumber,
             reason: "Non-direct channel: " + channel };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PENDING_SO.sheetName);
  if (!sheet) {
    return { status: "skipped", invoiceNumber: invoiceNumber,
             reason: "Pending Sales Orders sheet not set up" };
  }

  // --- FIND MATCHING PENDING ROW ---
  var salesorderId = String(invoice.salesorder_id || "").trim();
  var refSoNumber  = String(invoice.reference_number || "").trim();
  var matchedRow = -1;
  var matchedBy  = "";

  if (salesorderId) {
    matchedRow = _findPendingRowBySalesorderId(sheet, salesorderId);
    if (matchedRow > 0) matchedBy = "salesorder_id";
  }
  if (matchedRow < 1 && refSoNumber) {
    matchedRow = _findPendingRow(sheet, refSoNumber);
    if (matchedRow > 0) matchedBy = "reference_number";
  }

  if (matchedRow < 1) {
    return { status: "skipped", invoiceNumber: invoiceNumber,
             reason: "No matching Pending row (salesorder_id=" + salesorderId
                   + ", reference_number=" + refSoNumber + ")" };
  }

  // --- STAMP INVOICE NUMBER ---
  // Idempotent — re-fires with the same value are a no-op write.
  var existingInvoice = String(sheet.getRange(matchedRow, PENDING_SO.cols.INVOICE).getValue() || "").trim();
  sheet.getRange(matchedRow, PENDING_SO.cols.INVOICE).setValue(invoiceNumber);
  sheet.getRange(matchedRow, PENDING_SO.cols.LAST_UPDATED).setValue(new Date());

  var soNumberOnRow = String(sheet.getRange(matchedRow, PENDING_SO.cols.SO_NUMBER).getValue() || "").trim();
  var actionTaken = (existingInvoice === invoiceNumber) ? "unchanged"
                  : (existingInvoice ? "updated" : "stamped");

  return {
    status:        "success",
    invoiceNumber: invoiceNumber,
    soNumber:      soNumberOnRow,
    matchedBy:     matchedBy,
    actionTaken:   actionTaken
  };
}


// =======================================================================================
// PULL — RETIRED (Option C, 2026-05-23)
// =======================================================================================
//
// `pullSalesOrderToDirect()` and `_deleteDirectRowsForSo()` were the original
// direct-commit Pull path: type SO# in sidebar → server immediately wrote all
// line items to DIRECT → optional force-rewrite via delete+insert.
//
// Replaced by the picker-driven Pull modal (ZohoPull.js + ZohoPullModal.html):
// sidebar opens the modal showing a per-line diff (unchanged/new/qty_changed/
// removed); picker selects what to apply; server commits via
// applyZohoPullSelection with all-or-nothing optimistic locking.
//
// Why this architecture won out: every recurring "phantom row" / "first-fire
// duplicate" / row-shift class of bug was an implicit reconciler running in
// the background. Picker-in-the-loop eliminates that bug class by
// construction. See CLAUDE.md 2026-05-23 narrative for the full story.
//
// Deleted in step 8 of the build. Git history is the safety net.
// =======================================================================================


/**
 * Sidebar preview entry point — returns what would be pulled without writing.
 * Useful for showing the user a confirmation card before they commit.
 *
 * Accepts either SO# (e.g. "SO-22815") or invoice number (e.g. "INV-022496"),
 * resolved via _resolvePendingRow.
 *
 * Returns { status, soNumber, invoiceNumber?, customer, total,
 *           items: [{sku, qty, location, hand, zohoPrice, ebayPrice, priceStatus, priceDelta}],
 *           alreadyPulled, orderStatus, payment, shipment,
 *           priceCheck: { summary, direction, offCount, notFoundCount, totalCount, totalDelta } }
 *
 * `priceCheck` is computed fresh against current MI on every preview call
 * (cell stamp on Pending row may be slightly stale between Inventory Lite
 * Sync runs; preview always uses current MI snapshot).
 */
function previewPendingSalesOrder(soNumber) {
  var query = String(soNumber || "").trim();
  if (!query) return { status: "error", reason: "Empty SO / invoice number" };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PENDING_SO.sheetName);
  if (!sheet) return { status: "error", reason: "Pending sheet not set up" };

  var rowNum = _resolvePendingRow(sheet, query);
  if (rowNum < 1) {
    return { status: "not_found", soNumber: query,
             reason: "Not in Pending. Check Zoho or wait for webhook." };
  }

  var rowValues = sheet.getRange(rowNum, 1, 1, PENDING_SO.dataWidth).getValues()[0];
  var canonicalSo = String(rowValues[PENDING_SO.idx("SO_NUMBER")] || "").trim() || query;
  var invoiceOnRow = String(rowValues[PENDING_SO.idx("INVOICE")] || "").trim();
  var alreadyPulled = String(rowValues[PENDING_SO.idx("PULLED")] || "").trim().toUpperCase()
                      === PENDING_SO.pulledFlag;

  var rawJson = String(rowValues[PENDING_SO.idx("PAYLOAD")] || "");
  if (!rawJson) return { status: "error", reason: "Cached payload missing" };

  var payload;
  try { payload = JSON.parse(rawJson); }
  catch (e) { return { status: "error", reason: "Cached payload corrupted" }; }

  var maps = buildLocationAndInventoryMaps();
  var locationMap = maps.locationMap;
  var inventoryMap = maps.inventoryMap;

  // Fresh price-check against current MI snapshot (the cell stamp may be
  // up to one upsert behind; preview always recomputes for full accuracy).
  // _computePriceCheck returns per-SKU detail we merge into each item below.
  var priceCheck = _computePriceCheck(payload);
  var priceBySku = {};
  for (var pi = 0; pi < priceCheck.items.length; pi++) {
    var pe = priceCheck.items[pi];
    priceBySku[String(pe.sku).toLowerCase()] = pe;
  }

  var items = (payload.line_items || []).map(function(li) {
    var sku = String(li.sku || "").trim().toUpperCase();
    var skuLower = sku.toLowerCase();
    var inv = inventoryMap.get(skuLower);
    var pp = priceBySku[skuLower] || {};
    return {
      sku:         sku,
      qty:         parseInt(li.quantity, 10) || 1,
      name:        String(li.name || ""),
      location:    locationMap.get(skuLower) || "NOT FOUND",
      hand:        inv ? inv.available : 0,
      zohoPrice:   (pp.zohoPrice != null) ? pp.zohoPrice : (parseFloat(li.rate) || 0),
      ebayPrice:   (pp.ebayPrice != null) ? pp.ebayPrice : null,
      priceStatus: pp.status || "ok",
      priceDelta:  pp.delta != null ? pp.delta : 0
    };
  }).filter(function(it) { return it.sku; });

  return {
    status:        "ok",
    soNumber:      canonicalSo,
    invoiceNumber: invoiceOnRow,
    customer:      String(payload.customer_name || ""),
    total:         String(payload.total_formatted || ""),
    alreadyPulled: alreadyPulled,
    orderStatus:   String(rowValues[PENDING_SO.idx("ORDER_STATUS")] || ""),
    payment:       String(rowValues[PENDING_SO.idx("PAYMENT")] || ""),
    shipment:      String(rowValues[PENDING_SO.idx("SHIPMENT")] || ""),
    items:         items,
    priceCheck:    {
      summary:       priceCheck.summary,
      direction:     priceCheck.direction,
      offCount:      priceCheck.offCount,
      notFoundCount: priceCheck.notFoundCount,
      totalCount:    priceCheck.totalCount,
      totalDelta:    priceCheck.totalDelta
    }
  };
}


// =======================================================================================
// ALERTS SUPPORT — count of unpulled Pending SOs
// =======================================================================================

/**
 * Returns the count of Pending sheet rows where PULLED is blank.
 * Used by the sidebar Alerts card ("New from Zoho: N").
 */
function getPendingZohoCount() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(PENDING_SO.sheetName);
    if (!sheet) return 0;
    var lastRow = sheet.getLastRow();
    if (lastRow < PENDING_SO.dataStartRow) return 0;
    var pulledCol = sheet.getRange(
      PENDING_SO.dataStartRow, PENDING_SO.cols.PULLED,
      lastRow - PENDING_SO.dataStartRow + 1, 1
    ).getValues();
    var count = 0;
    for (var i = 0; i < pulledCol.length; i++) {
      var v = String(pulledCol[i][0] || "").trim().toUpperCase();
      if (v !== PENDING_SO.pulledFlag) count++;
    }
    return count;
  } catch (e) {
    return 0;
  }
}


// =======================================================================================
// VOID HANDLER — the one signal still propagated automatically
// =======================================================================================
//
// Replaces the entire `_propagateToDirectRows` family (retired 2026-05-23 as
// part of Option C). Every other Zoho→DIRECT data movement now goes through
// the explicit picker-driven Pull modal (ZohoPull.js).
//
// Why VOID stays automatic: customer cancellation is unambiguous, time-
// sensitive (picker shouldn't pick a canceled order), and the picker's
// action is identical regardless of who triggers it. No room for "review
// before applying" to add value.

/**
 * If the incoming Zoho SO payload has status=void, flip all non-terminal
 * DIRECT rows for this SO to CANCELED. Safe no-op for any other status.
 *
 * Called from upsertPendingSalesOrder on every fire (2026-06-02: no longer
 * gated on wasPulled — see the call site). Pure status flip on existing rows
 * matched by SO number; no-op when no DIRECT row carries this SO.
 *
 * @returns {{ canceledCount: number, error?: string }}
 */
function _handleZohoVoid(soNumber, salesorder) {
  var status = String(salesorder && salesorder.status || "").trim().toLowerCase();
  if (status !== "void") return { canceledCount: 0 };

  try {
    var result = updateOrderStatus(soNumber, Schema.status.CANCELED, {
      source:       "zoho",
      syncTelegram: false
    });
    return { canceledCount: (result && result.count) || 0 };
  } catch (e) {
    console.log("Zoho void handler failed for " + soNumber + ": " + e);
    return { canceledCount: 0, error: e.toString() };
  }
}

/**
 * If the incoming Zoho SO payload reports the order shipped/fulfilled, flip
 * all non-terminal DIRECT rows for this SO to SHIPPED. Safe no-op otherwise.
 *
 * RESTORED 2026-06-02. The shipped auto-flip originally shipped 2026-05-16,
 * then was removed as collateral when the buggy line-item reconciler
 * (_propagateToDirectRows) was deleted in the 2026-05-23 Option C pivot. It
 * is the second of the two UNAMBIGUOUS Zoho signals that auto-propagate
 * (SHIPPED + CANCELED); PREPARING stays employee-driven. This restores ONLY
 * the clean shipped signal — the line-item diff reconciler stays dead.
 *
 *   - 'partially_shipped' is intentionally IGNORED: an SO can ship in several
 *     tracking numbers; the final shipped/fulfilled state catches us up
 *     without needing per-line tracking.
 *   - updateOrderStatus's terminal-state guard means already-SHIPPED/CANCELED
 *     rows are left untouched, so re-fires can't double-flip.
 *
 * Called from upsertPendingSalesOrder on every fire (2026-06-02: no longer
 * gated on wasPulled, so hand-typed Direct SOs flip too — see the call site).
 * Pure status flip on existing rows matched by SO number; no-op when no DIRECT
 * row carries this SO.
 *
 * @returns {{ shippedCount: number, error?: string }}
 */
function _handleZohoShipped(soNumber, salesorder) {
  var shipped = String(salesorder && salesorder.shipped_status || "").trim().toLowerCase();
  if (shipped !== "shipped" && shipped !== "fulfilled") return { shippedCount: 0 };

  try {
    var result = updateOrderStatus(soNumber, Schema.status.SHIPPED, {
      source:       "zoho",
      syncTelegram: false
    });
    return { shippedCount: (result && result.count) || 0 };
  } catch (e) {
    console.log("Zoho shipped handler failed for " + soNumber + ": " + e);
    return { shippedCount: 0, error: e.toString() };
  }
}




/**
 * Insert one or more line items into DIRECT for an already-pulled SO.
 * Returns count of rows inserted.
 */
function _insertAddedItemsToDirect(sheet, soNumber, lineItems, noteOverride, detailOverride) {
  if (!lineItems || lineItems.length === 0) return 0;

  var boundary = getBoundaryRow();
  if (boundary < 1) return 0;

  var maps = buildLocationAndInventoryMaps();
  var locationMap = maps.locationMap;
  var inventoryMap = maps.inventoryMap;
  var zohoMap = buildZohoStockMap();   // DIRECT rows take HAND from Zoho first

  var newRows = [];
  var activityLogBatch = [];

  // Each line item may carry its own _noteOverride / _detailOverride (set by
  // Fix E to give per-row delta context like "was 12, now 15 total"). If not,
  // fall back to the function-level override, then to the default.
  for (var i = 0; i < lineItems.length; i++) {
    var li = lineItems[i] || {};
    var sku = String(li.sku || "").trim().toUpperCase();
    if (!sku) continue;
    var qty = parseInt(li.quantity, 10) || 1;
    var skuLower = sku.toLowerCase();
    var inv = inventoryMap.get(skuLower);
    // HAND for this DIRECT row: Zoho available first (owns direct sales),
    // MI.available as fallback. No committed subtraction (Zoho already nets it).
    var _zo = zohoMap.get(skuLower);
    var directHand = resolveHandValue(inv ? inv.available : null,
                                      _zo ? _zo.available : null, true);

    // Explicit-undefined checks (not `||`) so callers can pass empty string
    // to mean "leave NOTE blank" — the Pull modal's insert action does this
    // to reduce on-sheet note clutter (picker annotates manually after
    // verification, kit expansion adds its own ↳ from KIT- prefix). Legacy
    // callers that pass undefined still get the default "↳ added in Zoho".
    var noteText   = (li._noteOverride   !== undefined) ? li._noteOverride
                   : (noteOverride       !== undefined) ? noteOverride
                   : "↳ added in Zoho";
    var detailText = (li._detailOverride !== undefined) ? li._detailOverride
                   : (detailOverride     !== undefined) ? detailOverride
                   : "Line added in Zoho on existing SO";

    newRows.push([
      sku, qty,
      locationMap.get(skuLower) || "NOT FOUND",
      soNumber,
      noteText,
      Schema.status.PENDING,
      directHand,
      "", "", ""
    ]);
    activityLogBatch.push([
      "RECEIVED", soNumber, sku, qty,
      "zoho-update",
      detailText,
      undefined, noteText
    ]);
  }

  if (newRows.length === 0) return 0;

  var insertRow = boundary + 2;
  sheet.insertRowsBefore(insertRow, newRows.length);
  sheet.getRange(insertRow, 1, newRows.length, Schema.dataWidth).setValues(newRows);

  try {
    var templateRow = insertRow + newRows.length;
    sheet.getRange(templateRow, 1, 1, Schema.dataWidth).copyFormatToRange(
      sheet, 1, Schema.dataWidth,
      insertRow, insertRow + newRows.length - 1
    );
  } catch (e) { /* no-op */ }

  // Reset ROW-SPECIFIC decorations that may have been copied from the
  // template row — newly-inserted rows should be visually neutral, NOT
  // inherit flag/cancel treatments (soft-red bg + strikethrough) from
  // whatever sits below the insert point.
  //
  // Bug regression test (2026-05-23): after the Pull modal apply handler
  // was reordered so flags run before inserts, the row immediately below
  // the insert was always the freshly-flagged row — so copyFormatToRange
  // pulled the flag tint onto every new insert. Banding + duplicate-SO
  // borders are column-level formats and survive the reset.
  try {
    var newRowsRange = sheet.getRange(insertRow, 1, newRows.length, Schema.dataWidth);
    newRowsRange.setBackground(null);
    newRowsRange.setFontLine('none');
  } catch (e) { /* no-op */ }

  try { logActivityBatch(activityLogBatch); }
  catch (e) { console.log("_insertAddedItemsToDirect: log error: " + e); }

  // Refresh duplicate-SO border tabs — see same call in pullSalesOrderToDirect
  // for rationale (programmatic insert doesn't fire onEdit).
  try { setupDuplicateSalesOrderHighlighting(); }
  catch (e) { console.log("_insertAddedItemsToDirect: highlight refresh error: " + e); }

  // Refresh Kit SKU markers — programmatic insert doesn't fire kitSkuOnEdit,
  // so any auto-inserted DIRECT row whose SKU is a kit wouldn't get the ▣
  // glyph until next manual refresh.
  try { refreshKitSkuMarkers(); }
  catch (e) { console.log("_insertAddedItemsToDirect: kit marker refresh error: " + e); }

  // Enrich the inserted SKUs (title note + listing link) from MI.
  try { refreshSkuEnrichment(); }
  catch (e) { console.log("_insertAddedItemsToDirect: SKU enrichment error: " + e); }

  return newRows.length;
}


/**
 * Apply visual flag to a DIRECT row: strikethrough font + soft red tint +
 * prepended NOTE prefix. Picker sees the row immediately on next glance.
 *
 * @param {boolean} strike — true for removed (strikethrough), false for qty change (no strike)
 */
function _flagDirectRow(sheet, rowNum, notePrefix, strike) {
  if (!rowNum || rowNum < 1) return;
  try {
    var rowRange = sheet.getRange(rowNum, 1, 1, Schema.dataWidth);
    rowRange.setBackground('#ffe5e5');     // soft red tint
    if (strike) rowRange.setFontLine('line-through');

    // Prepend the warning to the NOTE column (col E). Preserve any existing
    // note content.
    var noteCell = sheet.getRange(rowNum, Schema.cols.NOTE);
    var existing = String(noteCell.getValue() || "").trim();
    var combined = notePrefix + (existing ? "\n" + existing : "");
    noteCell.setValue(combined);
  } catch (err) {
    console.log("_flagDirectRow error on row " + rowNum + ": " + err);
  }
}


// =======================================================================================
// PRIVATE HELPERS
// =======================================================================================

/**
 * Convert a Zoho salesorder payload into the Pending sheet row array.
 * 12 cols: see PENDING_SO.cols.
 *
 * Col L stores a SLIM payload, not the full Zoho response. Sheets has a
 * 50,000-char-per-cell limit; full Zoho payloads (custom_fields, packages,
 * contact_persons, nested address objects, taxes breakdown, etc.) can blow
 * past it on SOs with many line items, causing setValues to reject the
 * entire write. See _slimSalesOrder for the kept-field set.
 */
function _payloadToRow(salesorder) {
  var status = _zohoOrderStatusLabel(salesorder.status);
  var payment = _zohoPaymentLabel(salesorder.paid_status);
  var shipment = _zohoShipmentLabel(salesorder.shipped_status);
  var lineItems = Array.isArray(salesorder.line_items) ? salesorder.line_items : [];

  return [
    String(salesorder.salesorder_number || ""),       // A: SO #
    String(salesorder.customer_name || ""),           // B: CUSTOMER
    String(salesorder.date || ""),                    // C: DATE
    status,                                           // D: ORDER STATUS
    payment,                                          // E: PAYMENT
    shipment,                                         // F: SHIPMENT
    lineItems.length,                                 // G: ITEMS COUNT
    String(salesorder.total_formatted || ""),         // H: TOTAL
    new Date(),                                       // I: LAST_UPDATED
    "",                                               // J: PULLED (preserved on update; set on pull)
    "",                                               // K: PULLED_AT (preserved on update; set on pull)
    _serializeSlimPayload(salesorder),                // L: PAYLOAD (slim JSON cache)
    "",                                               // M: INVOICE (preserved on update; set by invoice webhook)
    ""                                                // N: PRICE_CHECK (stamped after row write — needs MI lookup)
  ];
}


/**
 * Build the slim payload kept in col L. Only includes fields read by:
 *   - previewPendingSalesOrder:  customer_name, total_formatted, line_items[sku/qty/name/rate]
 *   - pullSalesOrderToDirect:    customer_name, total_formatted, line_items[sku/qty]
 *   - _propagateToDirectRows:    status, shipped_status, line_items[sku/qty]
 *   - _computePriceCheck:        line_items[sku/rate] vs MI.currentPrice
 *   - _findPendingRowBySalesorderId: salesorder_id (invoice→SO linkage)
 *
 * Everything Zoho sends but we never read is dropped: custom_fields, packages,
 * contact_persons, billing_address, shipping_address, taxes, invoices nested,
 * packing_slip_template_id, attachment metadata, etc. Typical reduction is
 * 90-95% of raw size.
 *
 * `rate` per line item carries Zoho's selling price for the price-check
 * feature (shipped 2026-05-22). Parsed as float; defaults to 0 if missing.
 */
function _slimSalesOrder(salesorder) {
  if (!salesorder || typeof salesorder !== 'object') return {};
  var lineItems = Array.isArray(salesorder.line_items) ? salesorder.line_items : [];
  return {
    salesorder_id:     String(salesorder.salesorder_id     || ""),
    salesorder_number: String(salesorder.salesorder_number || ""),
    customer_name:     String(salesorder.customer_name     || ""),
    total_formatted:   String(salesorder.total_formatted   || ""),
    status:            String(salesorder.status            || ""),
    shipped_status:    String(salesorder.shipped_status    || ""),
    line_items: lineItems.map(function(li) {
      return {
        sku:      String((li && li.sku)  || ""),
        quantity: (li && li.quantity != null) ? li.quantity : 1,
        name:     String((li && li.name) || ""),
        rate:     (li && li.rate != null) ? parseFloat(li.rate) || 0 : 0
      };
    })
  };
}


/**
 * Serialize the slim payload with a defensive 49K backstop. If even the
 * slim version exceeds the limit (only realistic with hundreds of line
 * items), drop `name` fields and retry. Logs which path was taken.
 */
function _serializeSlimPayload(salesorder) {
  var LIMIT = 49000;
  var slim = _slimSalesOrder(salesorder);
  var json = JSON.stringify(slim);
  if (json.length <= LIMIT) return json;

  console.log("_serializeSlimPayload: slim JSON " + json.length + " chars exceeds " +
              LIMIT + " for SO " + slim.salesorder_number + " — dropping item names");
  slim.line_items = slim.line_items.map(function(li) {
    return { sku: li.sku, quantity: li.quantity };
  });
  json = JSON.stringify(slim);
  if (json.length <= LIMIT) return json;

  console.log("_serializeSlimPayload: ULTRA-slim JSON still " + json.length + " chars for SO " +
              slim.salesorder_number + " — storing minimal stub; pull will fail until shrunk");
  return JSON.stringify({
    salesorder_id:     slim.salesorder_id,
    salesorder_number: slim.salesorder_number,
    customer_name:     slim.customer_name,
    total_formatted:   slim.total_formatted,
    status:            slim.status,
    shipped_status:    slim.shipped_status,
    line_items:        [],
    _truncated:        true
  });
}


/**
 * Compare Zoho line-item prices (rate) against MI.currentPrice per SKU.
 * Returns:
 *   {
 *     summary: string,             — single-cell label for the PRICE_CHECK column
 *     totalDelta: number,           — sum of (zohoPrice - ebayPrice) across off items
 *     offCount: number,             — items mismatching beyond threshold
 *     notFoundCount: number,        — items with no MI row
 *     totalCount: number,           — total line items considered
 *     direction: "OK" | "HIGH" | "LOW" | "MIXED" | "NOT_FOUND" | "EMPTY",
 *     items: [                      — per-item detail for the sidebar Preview
 *       { sku, name, qty, zohoPrice, ebayPrice, delta, pctDelta, status }
 *     ]
 *   }
 *
 * Status per item:
 *   "ok"        — within threshold
 *   "zoho_high" — Zoho > eBay beyond threshold (customer complaints here)
 *   "zoho_low"  — Zoho < eBay beyond threshold (under-quoting)
 *   "not_found" — SKU not in MI
 *
 * Threshold: |delta| > max(absThreshold, pctThreshold × ebayPrice).
 * Defaults: $1.00 OR 2% of eBay price (whichever is greater).
 */
function _computePriceCheck(salesorder) {
  var EMPTY = {
    summary: "—", totalDelta: 0, offCount: 0, notFoundCount: 0,
    totalCount: 0, direction: "EMPTY", items: []
  };
  if (!salesorder || typeof salesorder !== 'object') return EMPTY;
  var lineItems = Array.isArray(salesorder.line_items) ? salesorder.line_items : [];
  if (lineItems.length === 0) return EMPTY;

  var maps;
  try { maps = buildLocationAndInventoryMaps(); }
  catch (e) {
    console.log("_computePriceCheck: buildLocationAndInventoryMaps failed: " + e);
    return { summary: "— MI unavailable", totalDelta: 0, offCount: 0, notFoundCount: 0,
             totalCount: lineItems.length, direction: "EMPTY", items: [] };
  }
  // currentPrice lives on MI but isn't in the lightweight inventoryMap shape
  // (which is just qty/sold/available). Read MI directly for the price column.
  var priceMap = _buildEbayPriceMap();

  var absT = (PENDING_SO.priceCheck && PENDING_SO.priceCheck.absThreshold) || 1.0;
  var pctT = (PENDING_SO.priceCheck && PENDING_SO.priceCheck.pctThreshold) || 0.02;

  var items = [];
  var offCount = 0;
  var notFoundCount = 0;
  var totalDelta = 0;
  var sawHigh = false;
  var sawLow  = false;

  for (var i = 0; i < lineItems.length; i++) {
    var li = lineItems[i] || {};
    var sku = String(li.sku || "").trim();
    var skuLower = sku.toLowerCase();
    var qty = parseInt(li.quantity, 10) || 1;
    var zohoPrice = parseFloat(li.rate) || 0;
    var ebayPrice = priceMap.get(skuLower);

    var entry = {
      sku:       sku,
      name:      String(li.name || ""),
      qty:       qty,
      zohoPrice: zohoPrice,
      ebayPrice: (ebayPrice != null) ? ebayPrice : null,
      delta:     0,
      pctDelta:  0,
      status:    "ok"
    };

    if (ebayPrice == null || ebayPrice <= 0) {
      entry.status = "not_found";
      notFoundCount++;
      items.push(entry);
      continue;
    }

    var delta = zohoPrice - ebayPrice;
    var threshold = Math.max(absT, pctT * ebayPrice);
    entry.delta    = delta;
    entry.pctDelta = ebayPrice > 0 ? (delta / ebayPrice) : 0;

    if (Math.abs(delta) > threshold) {
      if (delta > 0) { entry.status = "zoho_high"; sawHigh = true; }
      else           { entry.status = "zoho_low";  sawLow  = true; }
      offCount++;
      totalDelta += delta;
    }
    items.push(entry);
  }

  // Build the single-cell summary + direction enum
  var totalCount = lineItems.length;
  var direction, summary;

  if (offCount === 0 && notFoundCount === 0) {
    direction = "OK";
    summary = "✓ OK · " + totalCount + "/" + totalCount;
  } else if (offCount === 0 && notFoundCount > 0) {
    direction = "NOT_FOUND";
    summary = "⚠ NOT FOUND · " + notFoundCount + "/" + totalCount;
  } else {
    if (sawHigh && sawLow)      direction = "MIXED";
    else if (sawHigh)           direction = "HIGH";
    else                        direction = "LOW";

    // Always lead with the dominant direction word so the CF rule matches.
    // MIXED is rare but real (one item high, another low on same SO).
    var dirWord = (direction === "MIXED")
      ? (totalDelta >= 0 ? "Zoho HIGH" : "Zoho LOW")  // pick by sign of net
      : (direction === "HIGH" ? "Zoho HIGH" : "Zoho LOW");

    var sign = totalDelta >= 0 ? "+" : "";
    summary = "⚠ " + dirWord + " · " + offCount + "/" + totalCount + " · " +
              sign + "$" + Math.abs(totalDelta).toFixed(2);
    if (notFoundCount > 0) summary += " · " + notFoundCount + " ??";
  }

  return {
    summary:       summary,
    totalDelta:    totalDelta,
    offCount:      offCount,
    notFoundCount: notFoundCount,
    totalCount:    totalCount,
    direction:     direction,
    items:         items
  };
}


/**
 * Read MI's eBay price columns into a SKU→price map (lowercased SKU key).
 *
 * Source preference: `currentPrice` first, fall back to `startPrice` when
 * currentPrice is null/0. This handles OOS items — eBay returns null
 * `<CurrentPrice>` for out-of-stock listings but `<StartPrice>` (the
 * listing's set price) stays populated. For audit purposes, startPrice IS
 * what the customer will be charged once stock returns, so it's a valid
 * reference for OOS items.
 *
 * Returns an empty Map if MI is missing, has neither price column, or any
 * other read failure — caller treats missing entries as no-reference.
 */
function _buildEbayPriceMap() {
  var map = new Map();
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(DB_SHEET_NAME);
    if (!sheet) return map;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var skuIdx = -1;
    var currentIdx = -1;
    var startIdx = -1;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] || "").trim().toLowerCase();
      if (h === DB_SKU_HEADER.toLowerCase()) skuIdx     = i;
      else if (h === 'currentprice')         currentIdx = i;
      else if (h === 'startprice')           startIdx   = i;
    }
    if (skuIdx < 0 || (currentIdx < 0 && startIdx < 0)) return map;
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (var r = 0; r < data.length; r++) {
      var sku = String(data[r][skuIdx] || "").trim().toLowerCase();
      if (!sku) continue;
      var current = currentIdx >= 0 ? parseFloat(data[r][currentIdx]) : NaN;
      var start   = startIdx   >= 0 ? parseFloat(data[r][startIdx])   : NaN;
      // Prefer currentPrice; fall back to startPrice when current is null/0
      var price = (!isNaN(current) && current > 0) ? current
                : (!isNaN(start)   && start   > 0) ? start
                : NaN;
      if (!isNaN(price)) map.set(sku, price);
    }
  } catch (e) {
    console.log("_buildEbayPriceMap error: " + e);
  }
  return map;
}


/**
 * Resolve a user-typed query to a Pending row. Accepts either SO# (e.g.
 * "SO-22815") or invoice number (e.g. "INV-022496"). Returns row number, -1 if
 * not found. The "INV-" prefix is the discriminator (case-insensitive); anything
 * else falls through to SO# lookup. Both lookups are case/whitespace insensitive.
 */
function _resolvePendingRow(sheet, query) {
  var q = String(query || "").trim();
  if (!q) return -1;
  if (/^INV[-_]/i.test(q)) {
    return _findPendingRowByInvoice(sheet, q);
  }
  return _findPendingRow(sheet, q);
}


/**
 * Find the row number in the Pending sheet matching this SO#. Returns -1 if
 * not found. Case-insensitive trim match.
 */
function _findPendingRow(sheet, soNumber) {
  var lastRow = sheet.getLastRow();
  if (lastRow < PENDING_SO.dataStartRow) return -1;
  var soCol = sheet.getRange(
    PENDING_SO.dataStartRow, PENDING_SO.cols.SO_NUMBER,
    lastRow - PENDING_SO.dataStartRow + 1, 1
  ).getValues();
  var target = String(soNumber).trim().toUpperCase();
  for (var i = 0; i < soCol.length; i++) {
    if (String(soCol[i][0]).trim().toUpperCase() === target) {
      return PENDING_SO.dataStartRow + i;
    }
  }
  return -1;
}


/**
 * Find the Pending row whose cached PAYLOAD JSON has matching salesorder_id.
 * Returns -1 if not found. JSON.parses each cached payload; safe (each parse
 * try/catched) but O(N) — Pending stays small in practice.
 *
 * Used by the invoice handler since invoice payloads carry salesorder_id
 * (Zoho-internal ID) but usually leave salesorder_number empty.
 */
function _findPendingRowBySalesorderId(sheet, salesorderId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < PENDING_SO.dataStartRow) return -1;
  var payloadCol = sheet.getRange(
    PENDING_SO.dataStartRow, PENDING_SO.cols.PAYLOAD,
    lastRow - PENDING_SO.dataStartRow + 1, 1
  ).getValues();
  var target = String(salesorderId).trim();
  if (!target) return -1;
  for (var i = 0; i < payloadCol.length; i++) {
    var raw = String(payloadCol[i][0] || "");
    if (!raw) continue;
    try {
      var p = JSON.parse(raw);
      if (String(p.salesorder_id || "").trim() === target) {
        return PENDING_SO.dataStartRow + i;
      }
    } catch (e) { /* corrupted payload — skip */ }
  }
  return -1;
}


/**
 * Find the Pending row by INVOICE column value (e.g. "INV-022496").
 * Returns -1 if not found. Case-insensitive trim match.
 *
 * Used by the sidebar lookup when input starts with "INV-".
 */
function _findPendingRowByInvoice(sheet, invoiceNumber) {
  var lastRow = sheet.getLastRow();
  if (lastRow < PENDING_SO.dataStartRow) return -1;
  var invCol = sheet.getRange(
    PENDING_SO.dataStartRow, PENDING_SO.cols.INVOICE,
    lastRow - PENDING_SO.dataStartRow + 1, 1
  ).getValues();
  var target = String(invoiceNumber).trim().toUpperCase();
  if (!target) return -1;
  for (var i = 0; i < invCol.length; i++) {
    if (String(invCol[i][0]).trim().toUpperCase() === target) {
      return PENDING_SO.dataStartRow + i;
    }
  }
  return -1;
}


/**
 * Returns { skuUpper: [rowNumbers] } for All Orders rows whose SALES_ORDER
 * matches. One scan of col A-D below the boundary divider.
 */
/**
 * Normalize an SKU value to a comparable string. Handles:
 *   - number → string with no decimal (e.g., 175608 → "175608")
 *   - float with trailing zeros (e.g., "175608.0" → "175608")
 *   - whitespace and case (trim + uppercase)
 * Used everywhere SKUs from different sources (Zoho payload, Sheets cells,
 * Master Inventory) need to be compared. Without consistent normalization,
 * the same SKU stored as `175608` in MI and `"175608"` in Zoho would silently
 * mismatch and trigger false inserts.
 */
function _normalizeSku(rawSku) {
  if (rawSku == null) return "";
  var s;
  if (typeof rawSku === "number") {
    s = String(Math.trunc(rawSku));
  } else {
    s = String(rawSku).trim();
    if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, "");
  }
  return s.trim().toUpperCase();
}


/**
 * Rich per-SKU state of DIRECT rows for a given SO. Returns:
 *   {
 *     skus: {
 *       "<SKU>": {
 *         totalActiveQty: <number>,   — sum of qty across non-CANCELED rows
 *         rows: [
 *           { row: <1-based row#>, qty: <number>, status: <string> }
 *         ]
 *       }
 *     }
 *   }
 *
 * SHIPPED and PENDING/PREPARING rows COUNT toward totalActiveQty (they
 * represent real committed inventory the customer has paid for).
 * CANCELED rows do NOT count (they're nullified — customer didn't get those
 * units, so they shouldn't affect Zoho-comparison math).
 *
 * Used exclusively by Fix E propagation (Zoho SO line-item diff). The
 * separate `_findDirectRowsForSo` helper that returns just row numbers is
 * preserved for other callers (auto-link check, SHIPPED propagation gate).
 */
/**
 * One-shot cleanup helper for the kit-expanded-rows false-flag incident
 * (2026-05-20). Walks DIRECT rows for the given SO# (or all SOs if
 * soNumber is empty), and for each row whose NOTE contains "↳ from KIT-"
 * AND has a stale "⚠️ REMOVED IN ZOHO" warning prefix, strips the warning
 * + clears strikethrough + restores the row background.
 *
 * Run from the Apps Script editor:
 *   clearKitExpansionFalseFlags("SO-22750")   // single SO
 *   clearKitExpansionFalseFlags()              // all SOs (slower)
 *
 * Safe to re-run. Only touches rows that match BOTH conditions (kit-marker
 * present + warning prefix present) — won't disturb legitimately-flagged
 * rows or unflagged kit-expanded rows.
 */
function clearKitExpansionFalseFlags(soNumber) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ All Orders sheet not found";

  var boundary = getBoundaryRow();
  if (boundary < 1) return "❌ DIRECT boundary not found";

  var lastRow = sheet.getLastRow();
  var startRow = boundary + 2;
  var nRows = lastRow - startRow + 1;
  if (nRows < 1) return "ℹ️ No DIRECT rows to scan";

  var data = sheet.getRange(startRow, 1, nRows, Schema.dataWidth).getValues();
  var NOTE_I = Schema.idx("NOTE");
  var SO_I   = Schema.idx("SALES_ORDER");

  var target = soNumber ? String(soNumber).trim().toUpperCase() : "";
  var cleared = 0;
  var WARN_PREFIX_REGEX = /^⚠️ REMOVED IN ZOHO [^\n]*\n?/;

  for (var i = 0; i < data.length; i++) {
    var rowSo = String(data[i][SO_I] || "").trim().toUpperCase();
    if (target && rowSo !== target) continue;

    var note = String(data[i][NOTE_I] || "");
    if (note.indexOf("↳ from KIT-") === -1) continue;          // not kit-expanded
    if (!WARN_PREFIX_REGEX.test(note)) continue;                // not falsely flagged

    var rowNum = startRow + i;

    // Strip the warning prefix from NOTE
    var cleanedNote = note.replace(WARN_PREFIX_REGEX, "");
    sheet.getRange(rowNum, NOTE_I + 1).setValue(cleanedNote);

    // Clear strikethrough + red background across the row
    var rowRange = sheet.getRange(rowNum, 1, 1, Schema.dataWidth);
    rowRange.setFontLine("none");
    rowRange.setBackground(null);  // banding restores natural color

    cleared++;
  }

  return "✅ Cleared false flags on " + cleared + " kit-expanded row(s)"
       + (target ? " for " + soNumber : " (all SOs)");
}


function _readDirectStateForSo(sheet, soNumber) {
  var out = { skus: {} };
  var lastRow  = sheet.getLastRow();
  var boundary = getBoundaryRow();
  if (boundary < 1 || lastRow < boundary + 2) return out;

  var startRow = boundary + 2;
  var nRows    = lastRow - startRow + 1;
  if (nRows < 1) return out;

  var data = sheet.getRange(startRow, 1, nRows, Schema.dataWidth).getValues();
  var target = String(soNumber).trim().toUpperCase();

  var SKU_I    = Schema.idx("SKU");
  var QTY_I    = Schema.idx("QTY");
  var SO_I     = Schema.idx("SALES_ORDER");
  var STATUS_I = Schema.idx("STATUS");
  var NOTE_I   = Schema.idx("NOTE");

  for (var i = 0; i < data.length; i++) {
    var so = String(data[i][SO_I] || "").trim().toUpperCase();
    if (so !== target) continue;

    var sku = _normalizeSku(data[i][SKU_I]);
    if (!sku) continue;

    // SKIP KIT-EXPANDED COMPONENT ROWS (fix shipped 2026-05-20)
    // When the picker expands a kit via Kit Expansion, the kit's component
    // SKUs are inserted as separate DIRECT rows — but they don't exist as
    // separate line items in Zoho (Zoho stores them as TEXT inside the kit
    // line item's description). If we include these rows in the directState,
    // Fix E's diff sees them as "in DIRECT but not in Zoho" and flags every
    // single component as "REMOVED IN ZOHO" on every webhook fire.
    //
    // The kit-expansion NOTE prefix `↳ from KIT-<sku>` is the reliable
    // signal that a row is locally derived, not a real Zoho line item.
    // The kit row ITSELF (which IS a Zoho line item) doesn't carry this
    // prefix and stays in directState normally — so the diff still
    // correctly compares Zoho's view of the kit SKU vs DIRECT's view.
    //
    // We use indexOf (not startsWith) because a previously-flagged row may
    // have "⚠️ REMOVED IN ZOHO 5/20" prepended to the NOTE, with the kit
    // marker still appearing further down — these false-flagged rows still
    // need to be excluded going forward.
    var note = String(data[i][NOTE_I] || "");
    if (note.indexOf("↳ from KIT-") !== -1) continue;

    var qty    = parseInt(data[i][QTY_I], 10) || 0;
    var status = Schema.normalize(String(data[i][STATUS_I] || "").trim().toUpperCase());
    var rowNum = startRow + i;

    if (!out.skus[sku]) {
      out.skus[sku] = { totalActiveQty: 0, rows: [] };
    }
    out.skus[sku].rows.push({ row: rowNum, qty: qty, status: status });

    if (status !== Schema.status.CANCELED) {
      out.skus[sku].totalActiveQty += qty;
    }
  }
  return out;
}


function _findDirectRowsForSo(sheet, soNumber) {
  var out = {};
  var lastRow = sheet.getLastRow();
  var boundary = getBoundaryRow();
  if (boundary < 1 || lastRow < boundary + 2) return out;

  var startRow = boundary + 2;   // first DIRECT data row
  var nRows = lastRow - startRow + 1;
  if (nRows < 1) return out;

  var data = sheet.getRange(startRow, 1, nRows, Schema.cols.SALES_ORDER).getValues();
  var target = String(soNumber).trim().toUpperCase();

  for (var i = 0; i < data.length; i++) {
    var so = String(data[i][Schema.idx("SALES_ORDER")] || "").trim().toUpperCase();
    if (so !== target) continue;
    var sku = String(data[i][Schema.idx("SKU")] || "").trim().toUpperCase();
    if (!sku) continue;
    var rowNum = startRow + i;
    if (!out[sku]) out[sku] = [];
    out[sku].push(rowNum);
  }
  return out;
}


/**
 * Zoho order-status values → user-facing labels for the Pending sheet.
 * draft / open ("confirmed") / closed / void are the documented set.
 */
function _zohoOrderStatusLabel(rawStatus) {
  var s = String(rawStatus || "").trim().toLowerCase();
  if (!s) return "";
  if (s === "draft")  return "DRAFT";
  if (s === "open")   return "CONFIRMED";
  if (s === "closed") return "CLOSED";
  if (s === "void")   return "VOID";
  return s.toUpperCase();
}

function _zohoPaymentLabel(rawPaid) {
  var s = String(rawPaid || "").trim().toLowerCase();
  if (!s) return "";
  if (s === "unpaid")          return "UNPAID";
  if (s === "partially_paid")  return "PARTIAL";
  if (s === "paid")            return "PAID";
  if (s === "overdue")         return "OVERDUE";
  return s.toUpperCase().replace(/_/g, " ");
}

function _zohoShipmentLabel(rawShipped) {
  var s = String(rawShipped || "").trim().toLowerCase();
  if (!s) return "";
  if (s === "shipped")             return "SHIPPED";
  if (s === "partially_shipped")   return "PARTIAL";
  if (s === "fulfilled")           return "FULFILLED";
  if (s === "pending")             return "PENDING";
  return s.toUpperCase().replace(/_/g, " ");
}
