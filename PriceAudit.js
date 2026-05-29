// =======================================================================================
// PriceAudit.js — bulk eBay-vs-Zoho price-drift audit + Zoho hygiene gap detector
// Shipped 2026-05-22; INACTIVE CANDIDATE detection added 2026-05-25.
// =======================================================================================
//
// PROBLEM
// -------
// Zoho's eBay-sync mirrors qty fine but NOT prices reliably. When admin changes
// an eBay price for competition, Zoho's selling_price stays stale. Direct-sale
// quotes use Zoho's price → customer compares to eBay → calls about the
// discrepancy. The intervention is to keep Zoho prices fixed BEFORE quotes
// happen by surfacing all drifts and fixing them in bulk.
//
// SECONDARY USE — Zoho hygiene gap detection (2026-05-25): the audit also
// surfaces items still ACTIVE in Zoho whose eBay listing has been ended.
// The operational rule: end on eBay → mark inactive in Zoho. When this gets
// missed, the audit catches it for cleanup.
//
// HOW IT WORKS
// ------------
// 1. Read MI's `sku` + `currentPrice`/`startPrice` + `listingStatus` columns.
//    Build TWO structures (single MI read):
//      activeSkus: Set of SKUs where listingStatus === "Active"
//      prices:     Map of {sku → currentPrice||startPrice} for active items with a price
// 2. Read the Zoho Stock sheet (mirrored every 2 min by the scheduled n8n push;
//    see ZohoStock.js). Each row carries {sku, item_name, selling_price, ...}.
//    Repointed 2026-05-28 — was calling the bulk-fetch proxy directly which
//    blocked Apps Script ~60-90s; now reads the cached sheet. Audit freshness
//    becomes "as fresh as last sync" (2-min during 9-5) — fine for drift
//    detection. The SYNCED timestamp is surfaced to the picker so they know.
// 3. For each Zoho item, classify into one of three buckets:
//      a. SKU IS in activeSkus AND has eBay price → price drift comparison
//         → if |delta| > threshold → driftRow (ZOHO HIGH / ZOHO LOW)
//      b. SKU IS in activeSkus but NO eBay price → "OOS / NO REF" row
//         (eBay listing active but returned null Price — usually temp OOS)
//      c. SKU NOT in activeSkus → "INACTIVE CANDIDATE" row
//         (either never on eBay, or ended on eBay — Zoho should be deactivated)
// 4. Write to "Price Audit" sheet in order: drifts (|delta| DESC), INACTIVE
//    CANDIDATEs (SKU ASC), OOS (SKU ASC). CF rules tint by direction.
//
// THRESHOLD: |delta| > max(absThreshold, pctThreshold × ebayPrice).
// Defaults: $1.00 OR 2% of eBay price (whichever is greater).
//
// SCOPE: read-only v1. Picker fixes Zoho manually after reviewing audit.
// v2 (parked): per-row "Push to Zoho" button using Zoho write scope.
// =======================================================================================


var PRICE_AUDIT = {
  sheetName: "Price Audit",

  cols: {
    SKU:          1,   // A
    ITEM_NAME:    2,   // B
    EBAY_LIVE:    3,   // C — from MI.currentPrice
    ZOHO_LIVE:    4,   // D — fresh from Zoho /items
    DELTA:        5,   // E — zoho - ebay  ($)
    PCT_DELTA:    6,   // F — delta / ebay  (%)
    DIRECTION:    7,   // G — ZOHO HIGH / ZOHO LOW
    LAST_CHECKED: 8    // H
  },

  idx: function(name) { return PRICE_AUDIT.cols[name] - 1; },

  dataWidth:    8,
  headerRow:    1,
  dataStartRow: 2,

  headers: ["SKU", "ITEM NAME", "eBay LIVE", "Zoho LIVE",
            "DELTA $", "DELTA %", "DIRECTION", "LAST CHECKED"],

  threshold: { abs: 1.00, pct: 0.02 }
};


// =======================================================================================
// HELPER — read MI once, return BOTH a price map AND the active-SKU set
// =======================================================================================
//
// Used by runPriceAudit to distinguish three audit outcomes per Zoho-active item:
//   1. Zoho SKU in activeSkus + price present → run normal drift comparison
//   2. Zoho SKU in activeSkus + no price       → "OOS / NO REF" (active on eBay, null price)
//   3. Zoho SKU NOT in activeSkus              → "INACTIVE CANDIDATE" — Zoho hygiene gap
//      (either SKU never existed in MI, or its MI row has listingStatus != "Active")
//
// NOT a replacement for _buildEbayPriceMap in ZohoSalesOrders.js — that one is used by
// the per-SO price strip which still wants the "include ended items with startPrice
// fallback" behavior (sale already happened on that SO). Audit's needs are different.
//
// Single MI read for efficiency.
// =======================================================================================

function _buildActiveEbayMaps() {
  var prices = new Map();
  var activeSkus = new Set();
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(DB_SHEET_NAME);
    if (!sheet) return { prices: prices, activeSkus: activeSkus };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { prices: prices, activeSkus: activeSkus };
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    var skuIdx = -1, currentIdx = -1, startIdx = -1, statusIdx = -1;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] || "").trim().toLowerCase();
      if      (h === DB_SKU_HEADER.toLowerCase()) skuIdx     = i;
      else if (h === 'currentprice')              currentIdx = i;
      else if (h === 'startprice')                startIdx   = i;
      else if (h === 'listingstatus')             statusIdx  = i;
    }
    // Required columns: sku + listingStatus. Without listingStatus we can't tell
    // active from ended, which kills the INACTIVE CANDIDATE detection.
    if (skuIdx < 0 || statusIdx < 0) {
      console.log("_buildActiveEbayMaps: missing required column (sku or listingStatus). skuIdx=" + skuIdx + " statusIdx=" + statusIdx);
      return { prices: prices, activeSkus: activeSkus };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (var r = 0; r < data.length; r++) {
      var sku = String(data[r][skuIdx] || "").trim().toLowerCase();
      if (!sku) continue;
      var status = String(data[r][statusIdx] || "").trim().toLowerCase();
      if (status !== 'active') continue;          // <-- the listingStatus filter

      activeSkus.add(sku);

      var current = currentIdx >= 0 ? parseFloat(data[r][currentIdx]) : NaN;
      var start   = startIdx   >= 0 ? parseFloat(data[r][startIdx])   : NaN;
      var price = (!isNaN(current) && current > 0) ? current
                : (!isNaN(start)   && start   > 0) ? start
                : NaN;
      if (!isNaN(price)) prices.set(sku, price);
    }
  } catch (e) {
    console.log("_buildActiveEbayMaps error: " + e);
  }
  return { prices: prices, activeSkus: activeSkus };
}


// =======================================================================================
// SETUP — idempotent
// =======================================================================================

function setupPriceAuditSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PRICE_AUDIT.sheetName);
  if (!sheet) sheet = ss.insertSheet(PRICE_AUDIT.sheetName);

  // --- HEADERS ---
  sheet.getRange(PRICE_AUDIT.headerRow, 1, 1, PRICE_AUDIT.dataWidth)
    .setValues([PRICE_AUDIT.headers])
    .setBackground('#1d1d1b')
    .setFontColor('#ffd966')
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(PRICE_AUDIT.headerRow, 1, 1, PRICE_AUDIT.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(PRICE_AUDIT.cols.SKU,          100);
  sheet.setColumnWidth(PRICE_AUDIT.cols.ITEM_NAME,    320);
  sheet.setColumnWidth(PRICE_AUDIT.cols.EBAY_LIVE,    100);
  sheet.setColumnWidth(PRICE_AUDIT.cols.ZOHO_LIVE,    100);
  sheet.setColumnWidth(PRICE_AUDIT.cols.DELTA,         90);
  sheet.setColumnWidth(PRICE_AUDIT.cols.PCT_DELTA,     80);
  sheet.setColumnWidth(PRICE_AUDIT.cols.DIRECTION,    120);
  sheet.setColumnWidth(PRICE_AUDIT.cols.LAST_CHECKED, 140);

  // --- DATA AREA FORMATS ---
  var maxDataRow = 4000;
  var dataRows = maxDataRow - PRICE_AUDIT.dataStartRow + 1;

  sheet.getRange(PRICE_AUDIT.dataStartRow, PRICE_AUDIT.cols.SKU, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(PRICE_AUDIT.dataStartRow, PRICE_AUDIT.cols.ITEM_NAME, dataRows, 1)
    .setFontFamily('Roboto').setFontSize(10).setHorizontalAlignment('left');
  sheet.getRange(PRICE_AUDIT.dataStartRow, PRICE_AUDIT.cols.EBAY_LIVE, dataRows, 1)
    .setNumberFormat('$#,##0.00').setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(PRICE_AUDIT.dataStartRow, PRICE_AUDIT.cols.ZOHO_LIVE, dataRows, 1)
    .setNumberFormat('$#,##0.00').setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(PRICE_AUDIT.dataStartRow, PRICE_AUDIT.cols.DELTA, dataRows, 1)
    .setNumberFormat('+$#,##0.00;-$#,##0.00').setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(PRICE_AUDIT.dataStartRow, PRICE_AUDIT.cols.PCT_DELTA, dataRows, 1)
    .setNumberFormat('+0.0%;-0.0%').setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(PRICE_AUDIT.dataStartRow, PRICE_AUDIT.cols.DIRECTION, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(PRICE_AUDIT.dataStartRow, PRICE_AUDIT.cols.LAST_CHECKED, dataRows, 1)
    .setNumberFormat('M/d/yy h:mm am/pm')
    .setFontFamily('Roboto Mono').setFontSize(9).setFontColor('#5f5f5f').setHorizontalAlignment('center');

  sheet.getRange(PRICE_AUDIT.dataStartRow, 1, dataRows, PRICE_AUDIT.dataWidth)
    .setVerticalAlignment('middle');

  // --- BANDING ---
  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  var bandRange = sheet.getRange(1, 1, maxDataRow, PRICE_AUDIT.dataWidth);
  var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b')
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#fff8e7');

  // --- CONDITIONAL FORMATTING ---
  // DIRECTION column: ZOHO HIGH = amber-yellow, ZOHO LOW = orange
  // DELTA column: tint by sign (red for positive=customer-complaint, gray for negative=under-quote)
  var existingRules = sheet.getConditionalFormatRules() || [];
  var keep = existingRules.filter(function(r) {
    var ranges = r.getRanges();
    if (!ranges || ranges.length === 0) return true;
    return !ranges.some(function(rg) {
      if (rg.getSheet().getName() !== PRICE_AUDIT.sheetName) return false;
      var c = rg.getColumn();
      return c === PRICE_AUDIT.cols.DIRECTION || c === PRICE_AUDIT.cols.DELTA;
    });
  });

  var dirRange = sheet.getRange(PRICE_AUDIT.dataStartRow, PRICE_AUDIT.cols.DIRECTION, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('ZOHO HIGH')
    .setBackground('#fff4b0').setFontColor('#7d5d00').setBold(true)
    .setRanges([dirRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('ZOHO LOW')
    .setBackground('#ffd699').setFontColor('#7a3d00').setBold(true)
    .setRanges([dirRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('OOS / NO REF')
    .setBackground('#f0f0f0').setFontColor('#5f5f5f').setBold(false)
    .setRanges([dirRange]).build());
  // INACTIVE CANDIDATE: Zoho still active but eBay listing ended (or never existed).
  // Slate-blue tint signals "action needed in Zoho" — distinct from gray OOS (no
  // action) and amber/orange price drifts (different action).
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('INACTIVE CANDIDATE')
    .setBackground('#cfd8dc').setFontColor('#37474f').setBold(true)
    .setRanges([dirRange]).build());

  // DELTA: tint cells by sign for at-a-glance scan
  var deltaRange = sheet.getRange(PRICE_AUDIT.dataStartRow, PRICE_AUDIT.cols.DELTA, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor('#7d5d00')
    .setRanges([deltaRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#7a3d00')
    .setRanges([deltaRange]).build());

  sheet.setConditionalFormatRules(keep);

  sheet.setFrozenRows(1);

  return "✅ Price Audit sheet ready.";
}


/** Sidebar: switch view to Price Audit sheet. */
function openPriceAudit() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet";
  var sheet = ss.getSheetByName(PRICE_AUDIT.sheetName);
  if (!sheet) {
    setupPriceAuditSheet();
    sheet = ss.getSheetByName(PRICE_AUDIT.sheetName);
  }
  ss.setActiveSheet(sheet);
  return "✅ Opened Price Audit";
}


// =======================================================================================
// MAIN — runPriceAudit (sidebar button)
// =======================================================================================

/**
 * Run the full audit. Reads MI's currentPrice (eBay side), calls n8n bulk
 * fetch (Zoho side), joins by SKU, filters to mismatches beyond threshold,
 * writes the Price Audit sheet sorted by abs(delta) DESC.
 *
 * Returns summary stats for the sidebar status bar:
 *   {
 *     ok:           boolean,
 *     message:      string,
 *     totalAudited: number,             — items present in BOTH sides
 *     mismatches:   number,             — actionable drifts written to sheet
 *     zohoHigh:     number,             — Zoho > eBay count
 *     zohoLow:      number,             — Zoho < eBay count
 *     totalDelta:   number,             — net sum across mismatches
 *     noRef:        number,             — Zoho items where MI has no currentPrice (OOS)
 *     onlyInZoho:   number,             — Zoho SKUs absent from MI entirely
 *     onlyInMi:     number,             — MI SKUs absent from Zoho (not audit-relevant; reported for sanity)
 *     durationSec:  number
 *   }
 */
function runPriceAudit() {
  var start = Date.now();
  try {
    // --- ENSURE SHEET ---
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(PRICE_AUDIT.sheetName);
    if (!sheet) {
      setupPriceAuditSheet();
      sheet = ss.getSheetByName(PRICE_AUDIT.sheetName);
    }

    // --- LAYER 1: Zoho prices — read from the cached Zoho Stock sheet
    // (mirrored every 2 min by the scheduled n8n push). Repointed 2026-05-28
    // from the bulk-fetch proxy: avoids the ~60-90s blocking call and lets
    // multiple admins run the audit concurrently without piling Zoho API calls.
    var zohoMap = buildZohoStockMap();
    if (!zohoMap || zohoMap.size === 0) {
      return { ok: false,
               message: "Zoho Stock sheet is empty — run Sync Zoho Stock first",
               durationSec: ((Date.now() - start) / 1000).toFixed(1) };
    }
    var zohoSyncedAt = getZohoStockSyncedAt();

    // --- LAYER 2: eBay live prices + active SKU set (from MI, filtered by listingStatus="Active") ---
    var miMaps = _buildActiveEbayMaps();
    var ebayMap = miMaps.prices;
    var activeSkus = miMaps.activeSkus;
    if (activeSkus.size === 0) {
      return { ok: false, message: "MI active-SKU set is empty — check that MI has a 'listingStatus' column and rows with status='Active'",
               durationSec: ((Date.now() - start) / 1000).toFixed(1) };
    }

    // --- JOIN & DIFF ---
    var absT = PRICE_AUDIT.threshold.abs;
    var pctT = PRICE_AUDIT.threshold.pct;
    var now = new Date();
    var driftRows = [];        // actionable price drifts — sorted to top by |delta| DESC
    var inactiveRows = [];     // INACTIVE CANDIDATE — middle section, Zoho hygiene gap
    var oosRows   = [];        // OOS / NO REF — bottom section, active-on-eBay but null price
    var zohoHigh = 0, zohoLow = 0, totalDelta = 0;
    var noRef = 0, inactiveCandidates = 0, totalAudited = 0;
    var seenSkus = {};  // for the only-in-MI count later

    // Iterate the Zoho Stock map. `forEach` callback returns are the equivalent
    // of `continue` in a classic for-loop — same control flow as the prior
    // bulk-fetch version.
    zohoMap.forEach(function(rec, skuLower) {
      var sku = rec.skuOriginal;          // case-preserved for sheet display
      seenSkus[skuLower] = true;

      var zohoPrice = rec.sellingPrice || 0;
      if (zohoPrice <= 0) return;          // no selling_price — nothing meaningful

      var itemName = rec.itemName || "";

      // STEP A: Is this SKU active on eBay? (in MI AND listingStatus == "Active")
      if (!activeSkus.has(skuLower)) {
        // Not active on eBay — either the SKU isn't in MI at all, OR it's in MI
        // with listingStatus != Active (i.e., Completed/Ended). Either way, Zoho
        // shouldn't still have it as active. The picker should review and
        // deactivate in Zoho. Slate-blue CF tint signals "Zoho action needed".
        inactiveCandidates++;
        inactiveRows.push([
          sku,                          // A: SKU
          itemName,                     // B: ITEM NAME
          "",                           // C: EBAY_LIVE — blank (no active reference)
          zohoPrice,                    // D: ZOHO_LIVE
          "",                           // E: DELTA $ — blank (not comparable)
          "",                           // F: DELTA % — blank
          "INACTIVE CANDIDATE",         // G: DIRECTION
          now                           // H: LAST_CHECKED
        ]);
        return;
      }

      // STEP B: SKU IS active on eBay. Look up its current price.
      var ebayPrice = ebayMap.get(skuLower);

      // OOS / NO REF — Active on eBay but eBay returned null currentPrice AND
      // null startPrice. Typically a temporary OOS state on an active listing.
      // Zoho's stored selling_price is still quotable once stock returns, so
      // we surface for spot-checking — but not as a hygiene gap.
      if (ebayPrice == null || ebayPrice <= 0) {
        noRef++;
        oosRows.push([
          sku,                          // A: SKU
          itemName,                     // B: ITEM NAME
          "",                           // C: EBAY_LIVE — blank
          zohoPrice,                    // D: ZOHO_LIVE
          "",                           // E: DELTA $ — blank
          "",                           // F: DELTA % — blank
          "OOS / NO REF",               // G: DIRECTION
          now                           // H: LAST_CHECKED
        ]);
        return;
      }

      totalAudited++;
      var delta = zohoPrice - ebayPrice;
      var threshold = Math.max(absT, pctT * ebayPrice);
      if (Math.abs(delta) <= threshold) return;   // within tolerance

      var direction = delta > 0 ? "ZOHO HIGH" : "ZOHO LOW";
      if (delta > 0) zohoHigh++;
      else           zohoLow++;
      totalDelta += delta;

      driftRows.push([
        sku,                                          // A: SKU
        itemName,                                     // B: ITEM NAME
        ebayPrice,                                    // C: EBAY_LIVE
        zohoPrice,                                    // D: ZOHO_LIVE
        delta,                                        // E: DELTA $
        ebayPrice > 0 ? (delta / ebayPrice) : 0,      // F: DELTA %
        direction,                                    // G: DIRECTION
        now                                           // H: LAST_CHECKED
      ]);
    });

    // Only-in-MI count (sanity check — items active on eBay but Zoho doesn't have).
    // Reads from activeSkus (post-filter), so it counts Zoho-side gaps cleanly.
    var onlyInMi = 0;
    activeSkus.forEach(function(sku) {
      if (!seenSkus[sku]) onlyInMi++;
    });

    // Sort: drifts (by |delta| DESC) → INACTIVE CANDIDATEs (SKU ASC) → OOS / NO REF (SKU ASC).
    // Picker's eye lands on actionable price drifts first, then the Zoho hygiene
    // queue, then the lower-priority watch list at the bottom.
    driftRows.sort(function(a, b) { return Math.abs(b[4]) - Math.abs(a[4]); });
    inactiveRows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])); });
    oosRows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])); });
    var rows = driftRows.concat(inactiveRows).concat(oosRows);

    // --- WRITE TO SHEET ---
    // Wipe prior audit data (preserve headers + formats)
    var lastRow = sheet.getLastRow();
    if (lastRow >= PRICE_AUDIT.dataStartRow) {
      sheet.getRange(PRICE_AUDIT.dataStartRow, 1,
                     lastRow - PRICE_AUDIT.dataStartRow + 1, PRICE_AUDIT.dataWidth)
           .clearContent();
    }

    if (rows.length > 0) {
      sheet.getRange(PRICE_AUDIT.dataStartRow, 1, rows.length, PRICE_AUDIT.dataWidth)
           .setValues(rows);
    }

    SpreadsheetApp.flush();

    var durationSec = ((Date.now() - start) / 1000).toFixed(1);
    var sign = totalDelta >= 0 ? "+" : "";
    var driftCount    = driftRows.length;
    var inactiveCount = inactiveRows.length;
    var oosCount      = oosRows.length;
    return {
      ok:                   true,
      message:              driftCount + " drift(s) · " + inactiveCount + " INACTIVE · " + oosCount + " OOS · net " + sign + "$" + totalDelta.toFixed(2),
      totalAudited:         totalAudited,
      mismatches:           driftCount,
      inactiveCandidates:   inactiveCount,
      oosRowCount:          oosCount,
      zohoHigh:             zohoHigh,
      zohoLow:              zohoLow,
      totalDelta:           parseFloat(totalDelta.toFixed(2)),
      noRef:                noRef,
      onlyInZoho:           inactiveCount,   // legacy alias — same number now
      onlyInMi:             onlyInMi,
      zohoStockSyncedAt:    zohoSyncedAt ? zohoSyncedAt.getTime() : null,   // ms epoch; sidebar formats "N min ago"
      durationSec:          durationSec
    };
  } catch (err) {
    try { console.log("runPriceAudit error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return {
      ok: false,
      message: "Audit failed: " + (err.message || err),
      durationSec: ((Date.now() - start) / 1000).toFixed(1)
    };
  }
}
