// =======================================================================================pl
// FULFILLMENT_SERVICE.gs - Fulfillment and Printing Functions////
// =======================================================================================

/**
 * Gathers 'PREPARING' rows, separating them into eBay vs DIRECT lists
 * Now includes HAND (Col G) and LEFT (Col H) columns
 */
function preparePrintSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);

  // ---- ACCOUNTABILITY GUARD (added 2026-05-01) ----
  // The picker MUST select their Pick ID for Shipping (cell G2) before
  // printing. This is the single chokepoint for warehouse-side accountability
  // — once set, every status event for the rest of the shift naturally
  // captures the picker name in the Activity Log.
  //
  // The dropdown's placeholder/header text is the literal string "Pick ID for
  // Shipping" — we must reject that as "unset," not accept it as a valid
  // picker. Real values follow the shape "Shipping - [Name] [Id]" (e.g.
  // "Shipping - YAwiss 1"). Pattern-match instead of just checking non-empty.
  var pickIdRaw = sheet.getRange(Schema.cellEmployeeId).getValue();
  var pickIdRawStr = String(pickIdRaw || "").trim();
  if (!pickIdRawStr || !/^Shipping\s*-\s*/i.test(pickIdRawStr)) {
    return "❌ Pick a real Pick ID for Shipping (cell " + Schema.cellEmployeeId +
           ") — the dropdown header doesn't count.";
  }

  var data = sheet.getDataRange().getValues();

  // SO badge exactly as painted on the sheet — read from column D's number
  // formats (the badge lives in the DISPLAY layer; values stay clean).
  // Reading the sheet's own assignment guarantees the printed badge matches
  // what the picker sees on screen. Format shape: '"1️⃣ "@'.
  // The GLYPH is mapped to its plain NUMBER here — the print renders it as
  // a drawn ink circle-digit (.so-badge), because emoji keycaps print as
  // gray mush on B&W printers. Legacy filled-circle glyphs (❶…⓴, from the
  // first badge iteration) are mapped too, in case a print happens before
  // the painter has repainted the sheet.
  var soBadgeFormats = sheet.getRange(1, Schema.cols.SALES_ORDER, data.length, 1)
                            .getNumberFormats();
  var LEGACY_BADGE_GLYPHS = ["❶","❷","❸","❹","❺","❻","❼","❽","❾","❿",
                             "⓫","⓬","⓭","⓮","⓯","⓰","⓱","⓲","⓳","⓴"];
  function _badgeFromFormat(fmt) {
    var bm = /^"(.+) "@$/.exec(String(fmt || ""));
    if (!bm) return "";
    var glyph = bm[1];
    var digits = /^\d+/.exec(glyph);          // keycap "1️⃣" starts with a plain digit char
    if (digits) return digits[0];
    if (glyph === "🔟") return "10";
    var li = LEGACY_BADGE_GLYPHS.indexOf(glyph);
    if (li >= 0) return String(li + 1);
    return "";
  }

  // Pull operational identifiers from the banner.
  // G2 = Pick ID for Shipping (dropdown), I2 = Pick ID for Adjustment (dropdown).
  // _extractPickIdData strips the "Shipping - " / "Adjustments - " prefix and
  // standardizes the trailing ID separator (e.g., "YAwiss · 1").
  var pickIdShipping   = _extractPickIdData(pickIdRaw);
  var pickIdAdjustment = _extractPickIdData(sheet.getRange(Schema.cellAdjustmentId).getValue());
  // Backward-compatible alias used by older code paths
  var employeeId = pickIdShipping;

  var ebayItems = [];
  var directItems = [];
  var isDirectSection = false; // Flag to track which section we are in

  // Iterate through data rows
  for (var i = Schema.dataStartRow - 1; i < data.length; i++) {
    var row = data[i];

    // --- 1. DETECT SECTION SPLIT ---
    // Substring match on the divider value (Schema.boundaryMarker = "DIRECT")
    var rowString = row.join("||").toUpperCase();
    if (rowString.indexOf(Schema.boundaryMarker) > -1 && rowString.length < 200) {
      isDirectSection = true;
      continue; // Skip the divider row itself
    }

    // --- 2. CHECK STATUS ---
    var status = String(row[Schema.idx("STATUS")] || "").trim().toUpperCase();

    if (status === Schema.status.PREPARING) {
      if (String(row[Schema.idx("SKU")]).trim().toUpperCase().indexOf('SKU') !== -1) continue;
      var itemData = [
        row[Schema.idx("SKU")],            // 0: SKU
        row[Schema.idx("QTY")],            // 1: QTY
        row[Schema.idx("LOCATION")],       // 2: LOCATION
        row[Schema.idx("SALES_ORDER")],    // 3: SALES_ORDER
        row[Schema.idx("NOTE")] || "",     // 4: NOTE
        row[Schema.idx("HAND")] || "",     // 5: HAND
        row[Schema.idx("LEFT")] || "",     // 6: LEFT
        row[Schema.idx("SHIPPING")] || "", // 7: SHIPPING
        row[Schema.idx("SHIP_COST")] || "",// 8: SHIP_COST
        _badgeFromFormat(soBadgeFormats[i][0]) // 9: SO badge glyph ("" if none)
      ];

      if (isDirectSection) {
        directItems.push(itemData);
      } else {
        ebayItems.push(itemData);
      }
    }
  }

  // Check if both are empty
  if (ebayItems.length === 0 && directItems.length === 0) {
    throw new Error("No items found with status '" + Schema.status.PREPARING + "' in the STATUS column.");
  }

  // (Print border-color map removed 2026-07-14 — multi-item order grouping
  // is carried by the SO badge glyph read from the sheet's col-D number
  // formats [itemData[9]], which survives B&W printing. Git history has the
  // old orderColorMap if ever wanted back.)

  // Format print timestamp in Houston timezone
  var now = new Date();
  var houstonTZ = "America/Chicago";
  var printDate = Utilities.formatDate(now, houstonTZ, "M/d/yyyy");
  var printTime = Utilities.formatDate(now, houstonTZ, "h:mm a");
  // Document number — stamped on every printed copy like a part number.
  // Format: "FUL · MM/DD · HH:mm" (slash-separated, human-readable, still sortable
  // within a calendar year — date and time use 24-hour, zero-padded fields).
  var docNumber = "FUL  ·  " + Utilities.formatDate(now, houstonTZ, "MM/dd") + "  ·  " + Utilities.formatDate(now, houstonTZ, "HH:mm");

  // Day name (e.g. "Monday") and short month-day format for the title block
  var printDay      = Utilities.formatDate(now, houstonTZ, "EEEE");
  var printDateLong = Utilities.formatDate(now, houstonTZ, "EEEE, MMMM d, yyyy");
  // Estimate total page count for the running header. Tunable; matches the
  // empty-row filling logic in the template (see fillTableToPage there).
  // Values held conservative to avoid phantom "thead-only" overflow pages.
  var ROWS_PER_FIRST_PAGE = 17;
  var ROWS_PER_CONT_PAGE  = 22;
  var estimatedPages = _estimatePageCount(ebayItems.length, directItems.length, ROWS_PER_FIRST_PAGE, ROWS_PER_CONT_PAGE);

  // Compute closing-page batch metrics (KPI cards on the audit page)
  var metrics = _computeBatchMetrics(ebayItems, directItems);

  var htmlTemplate = HtmlService.createTemplateFromFile('PrintFulfillment');
  htmlTemplate.ebayItems = ebayItems;
  htmlTemplate.directItems = directItems;
  htmlTemplate.employeeId = employeeId;            // back-compat
  htmlTemplate.pickIdShipping = pickIdShipping;
  htmlTemplate.pickIdAdjustment = pickIdAdjustment;
  htmlTemplate.printDate = printDate;
  htmlTemplate.printTime = printTime;
  htmlTemplate.printDay = printDay;
  htmlTemplate.printDateLong = printDateLong;
  htmlTemplate.docNumber = docNumber;
  htmlTemplate.estimatedPages = estimatedPages;
  htmlTemplate.metrics = metrics;
  // Soft-delete toggle for paid-shipping signals on the print template.
  // Default false (alerts OFF) per 2026-05-19 business request. Flip back ON
  // by running enablePrintPaidShippingAlerts() (allowlisted to yassinqurabi@gmail.com).
  htmlTemplate.showPaidShippingAlerts = isPrintPaidShippingAlertsEnabled();
  
  var ui = htmlTemplate.evaluate()
      .setTitle('Print Picking List')
      .setWidth(1000)
      .setHeight(800);

  // Activity Log: PRINTED event captures who printed and the batch shape.
  // Best-effort — a logging failure must never block the modal.
  try {
    var totalItems = ebayItems.length + directItems.length;
    logActivity(
      "PRINTED",
      "",                                                    // no single orderId for a batch
      "",                                                    // no SKU
      totalItems,                                            // qty = batch size
      "sidebar",                                             // source (warehouse-side, captures picker)
      "eBay: " + ebayItems.length + " · Direct: " + directItems.length + " · Doc: " + docNumber
      // picker auto-captured from G2 because source is 'sidebar' (warehouse-side)
    );
  } catch (logErr) { /* swallow — print must proceed */ }

  SpreadsheetApp.getUi().showModalDialog(ui, ' ');
  return "✅ Printing list ready — picker " + pickIdShipping + ".";
}

/**
 * Changes the status in Column F to "PREPARING" for all selected rows
 */
/**
 * Changes the status in Column F to "PREPARING" and syncs to Telegram.
 */
/**
 * Sidebar bulk-action: marks the user's currently-selected rows as PREPARING.
 *
 * Now delegates to updateOrderStatus() so terminal-state guards (SHIPPED /
 * CANCELED rows are skipped, not silently overwritten), Telegram sync, stats
 * refresh, and the canonical sequence all match every other status path.
 *
 * Don't take a local lock — updateOrderStatus acquires its own.
 */
function markSelectedPreparing() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range) return "No selection.";

  var startRow = range.getRow();
  var numRows  = range.getNumRows();

  // Clip the selection to the data area
  if (startRow < Schema.dataStartRow) {
    var adj = Schema.dataStartRow - startRow;
    startRow = Schema.dataStartRow;
    numRows -= adj;
  }
  if (numRows <= 0) return "Selection is above the data area.";

  var result = updateOrderStatus(
    { startRow: startRow, numRows: numRows },
    Schema.status.PREPARING,
    { source: "sidebar-bulk", sortAfter: true }
  );

  if (!result.success) {
    return "❌ " + (result.error || "Could not mark selection.");
  }
  var msg = "✅ Marked " + result.count + " row(s) as " + Schema.status.PREPARING + ".";
  if (result.blockedCount > 0) {
    msg += " (" + result.blockedCount + " blocked — already in terminal state.)";
  }
  return msg;
}

// =======================================================================================
// HELPERS — used by preparePrintSheet()
// =======================================================================================

/**
 * Strips the "Shipping - " / "Adjustments - " prefix from a Pick ID dropdown
 * value and inserts " · " before any trailing numeric ID for clean display.
 *   "Shipping - YAwiss 1"      → "YAwiss · 1"
 *   "Adjustments - AShamma 12" → "AShamma · 12"
 *   ""                         → "—"
 */
function _extractPickIdData(raw) {
  var s = String(raw || "").trim();
  if (!s) return "—";
  var match = s.match(/^(?:shipping|adjustments?)\s*[-:·]\s*(.+)$/i);
  var data = match ? match[1].trim() : s;
  return data.replace(/\s+(\d+)$/, " · $1");
}

/**
 * Estimate total printed page count given item counts and per-page capacities.
 * Matches the row-fill logic in PrintFulfillment.html so the page count shown
 * in the running header is accurate.
 */
function _estimatePageCount(ebayCount, directCount, firstCap, contCap) {
  function pagesFor(count) {
    if (count <= 0) return 0;
    if (count <= firstCap) return 1;
    return 1 + Math.ceil((count - firstCap) / contCap);
  }
  var ebayPages   = pagesFor(ebayCount);
  var directPages = pagesFor(directCount);
  var total = ebayPages + directPages;
  return Math.max(1, total);
}

/**
 * Computes batch metrics for the closing summary page (audit/sign-off page).
 * Returns 7 KPIs that get rendered as visual cards in the print template.
 *
 * Item array layout (per FulfillmentService row mapping):
 *   [0] SKU · [1] QTY · [2] LOC · [3] ORDER · [4] NOTE
 *   [5] HAND · [6] LEFT · [7] SHIPPING · [8] SHIP COST
 */
function _computeBatchMetrics(ebayItems, directItems) {
  var allItems = ebayItems.concat(directItems);

  // Sum of all QTY values — distinct from item count when QTY > 1
  var totalQty = allItems.reduce(function(sum, item) {
    var q = parseInt(item[1]);
    return sum + (isNaN(q) ? 0 : q);
  }, 0);

  // Distinct SKUs touched in this batch
  var skuSet = {};
  allItems.forEach(function(item) {
    var sku = String(item[0] || '').trim().toUpperCase();
    if (sku) skuSet[sku] = true;
  });
  var distinctSkus = Object.keys(skuSet).length;

  // Items where HAND ≤ 20 — early restock warning
  var lowStock = allItems.filter(function(item) {
    var hand = parseFloat(item[5]);
    return !isNaN(hand) && hand <= 20;
  }).length;

  // Paid shipping (eBay only — DIRECT items don't carry shipping data)
  var paidShippingItems = ebayItems.filter(function(item) {
    var cost = String(item[8] || '').trim();
    return cost && cost !== 'FREE' && cost !== '0' && cost !== '$0.00';
  });
  var paidShippingCount = paidShippingItems.length;
  var paidShippingTotal = paidShippingItems.reduce(function(sum, item) {
    var raw = String(item[8] || '').replace(/[^0-9.\-]/g, '');
    var num = parseFloat(raw);
    return sum + (isNaN(num) ? 0 : num);
  }, 0);

  return {
    ebayCount:         ebayItems.length,
    directCount:       directItems.length,
    totalItems:        allItems.length,
    totalQty:          totalQty,
    distinctSkus:      distinctSkus,
    lowStock:          lowStock,
    paidShippingCount: paidShippingCount,
    paidShippingTotal: paidShippingTotal.toFixed(2)
  };
}

// =======================================================================================
// PRINT-TEMPLATE SOFT-DELETE TOGGLES (added 2026-05-19)
// =======================================================================================
// State lives in ScriptProperties — server-side, shared across ALL users of the
// project. When one operator flips a toggle, every picker's next print run
// reflects the change immediately. No per-device variation possible.
//
// Asymmetric authorization model:
//   - DISABLE is open (anyone can turn signals OFF — removal is easy)
//   - ENABLE is allowlisted (only listed accounts can turn signals back ON —
//     restoration is governed). Matches the principle "you don't want a
//     curious picker un-doing a deliberate business decision."
//
// Each toggle action logs the actor + timestamp to console (visible in Apps
// Script Executions panel) — audit trail in case the state ever drifts in a
// way nobody remembers.
// =======================================================================================

var PRINT_TOGGLE_ALLOWLIST = ["yassinqurabi@gmail.com"];

/**
 * Authorization gate for enable functions.
 * @returns {string|null} Error message if denied, null if authorized.
 */
function _assertPrintToggleAuthorized() {
  var email = Session.getActiveUser().getEmail();
  if (PRINT_TOGGLE_ALLOWLIST.indexOf(email) === -1) {
    return "❌ Only " + PRINT_TOGGLE_ALLOWLIST.join(", ") +
           " can run this. Current user: " + (email || "unknown");
  }
  return null;
}

/**
 * Re-enable the SHIPPING + SHIP COST signals on the print template.
 * Reverses everything disablePrintPaidShippingAlerts() turned off.
 *
 * When ON:
 *   - "Buyer Paid" column header + body cells visible on eBay table
 *   - Per-row paid-shipping alert callout rows visible (yellow dom + intl)
 *   - Closing-summary "Paid Shipping $" KPI card visible
 *   - Closing-summary "SHIPPING ALERTS · This Batch" block populates
 *
 * Backend data (cols I + J on sheet, Alerts.js paid-shipping count, n8n's
 * deliveryDiscount math) is unaffected by this toggle — it's purely a
 * print-rendering switch. Run anytime to restore the visual signals.
 *
 * ALLOWLISTED — only emails in PRINT_TOGGLE_ALLOWLIST can run this.
 */
function enablePrintPaidShippingAlerts() {
  var denied = _assertPrintToggleAuthorized();
  if (denied) return denied;
  PropertiesService.getScriptProperties().setProperty(
    "print.paidShippingAlertsEnabled", "true"
  );
  console.log("[PrintToggle] Paid-shipping alerts ENABLED by " +
              Session.getActiveUser().getEmail() + " at " + new Date().toISOString());
  return "✅ Print paid-shipping alerts ENABLED. " +
         "Next print run will show the 'Buyer Paid' column + alert callouts + KPI.";
}

/**
 * Soft-delete the SHIPPING + SHIP COST signals from the print template.
 * Disabled state as of 2026-05-19 — business requested removal because "table
 * felt too crowded." See CLAUDE.md timeline entry for historical context
 * (deliveryDiscount math, under-shipping prevention) and full restoration recipe.
 *
 * When OFF:
 *   - "Buyer Paid" column header + body cells hidden on eBay table
 *   - Per-row paid-shipping alert callout rows hidden (both dom + intl)
 *   - Closing-summary "Paid Shipping $" KPI card hidden
 *   - Closing-summary "SHIPPING ALERTS · This Batch" block hidden
 *
 * Open authorization — anyone can run this. Use enablePrintPaidShippingAlerts()
 * to restore (allowlisted).
 */
function disablePrintPaidShippingAlerts() {
  PropertiesService.getScriptProperties().setProperty(
    "print.paidShippingAlertsEnabled", "false"
  );
  console.log("[PrintToggle] Paid-shipping alerts DISABLED by " +
              Session.getActiveUser().getEmail() + " at " + new Date().toISOString());
  return "✅ Print paid-shipping alerts DISABLED. " +
         "Next print run will hide the column + callouts + KPI.";
}

/**
 * Read-only accessor for the current toggle state.
 * Used by preparePrintSheet to pass the flag to the HTML template.
 * Default state when no property has been set yet: false (alerts disabled).
 */
function isPrintPaidShippingAlertsEnabled() {
  var val = PropertiesService.getScriptProperties()
    .getProperty("print.paidShippingAlertsEnabled");
  return val === "true";
}