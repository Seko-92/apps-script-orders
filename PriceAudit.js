// =======================================================================================
// PriceAudit.js — bulk eBay-vs-Zoho price-drift audit (shipped 2026-05-22)
// =======================================================================================
//
// PROBLEM
// -------
// Zoho's eBay-sync mirrors qty fine but NOT prices reliably. When admin changes
// an eBay price for competition, Zoho's selling_price stays stale. Direct-sale
// quotes use Zoho's price → customer compares to eBay → calls about the
// discrepancy. By the time the SO is created in our system, the price is
// already locked in. The right intervention is to keep Zoho prices fixed
// BEFORE quotes happen — by surfacing all drifts and fixing them in bulk.
//
// HOW IT WORKS
// ------------
// 1. Read MI's `currentPrice` column → SKU→price map (eBay-side truth)
// 2. Call n8n bulk fetch proxy → paginates Zoho's `GET /items` → SKU→price map
// 3. Join by SKU, compute deltas, filter to mismatches beyond threshold
// 4. Write to "Price Audit" sheet sorted by abs(delta) DESC
//
// NO-EBAY-REFERENCE handling: when MI has no currentPrice for a SKU (typically
// OOS items where eBay returns null Price), the audit classifies as "NO REF"
// in the summary stats and does NOT write to the sheet (no actionable drift
// vs no reference data). Picker decides whether to handle OOS items separately.
//
// THRESHOLD: |delta| > max(absThreshold, pctThreshold × ebayPrice).
// Defaults: $1.00 OR 2% of eBay price (whichever is greater). Conservative —
// suppresses rounding noise without missing real drifts.
//
// SCOPE: read-only v1. Picker fixes Zoho items manually after reviewing audit.
// v2 (parked) would add a per-row "Push to Zoho" button using Zoho write scope.
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

    // --- LAYER 1: Zoho live prices (via n8n bulk fetch) ---
    var fetchResult = triggerZohoBulkItemsFetch();
    if (!fetchResult.ok) {
      return {
        ok: false,
        message: "Zoho bulk fetch failed: " + (fetchResult.message || "unknown"),
        durationSec: ((Date.now() - start) / 1000).toFixed(1)
      };
    }
    var zohoItems = (fetchResult.data && fetchResult.data.items) || [];
    if (zohoItems.length === 0) {
      return { ok: false, message: "Zoho returned 0 items — check the workflow",
               durationSec: ((Date.now() - start) / 1000).toFixed(1) };
    }

    // --- LAYER 2: eBay live prices (from MI) ---
    var ebayMap = _buildEbayPriceMap();   // shared helper in ZohoSalesOrders.js
    if (ebayMap.size === 0) {
      return { ok: false, message: "MI currentPrice map is empty — Lite Sync / MAIN may not be writing prices",
               durationSec: ((Date.now() - start) / 1000).toFixed(1) };
    }

    // --- JOIN & DIFF ---
    var absT = PRICE_AUDIT.threshold.abs;
    var pctT = PRICE_AUDIT.threshold.pct;
    var now = new Date();
    var driftRows = [];    // actionable price drifts — sorted to top
    var oosRows   = [];    // OOS/no-reference items — appended at bottom
    var zohoHigh = 0, zohoLow = 0, totalDelta = 0;
    var noRef = 0, onlyInZoho = 0, totalAudited = 0;
    var seenSkus = {};  // for the only-in-MI count later

    for (var i = 0; i < zohoItems.length; i++) {
      var z = zohoItems[i] || {};
      var sku = String(z.sku || "").trim();
      var skuLower = sku.toLowerCase();
      if (!sku) continue;
      seenSkus[skuLower] = true;

      var zohoPrice = parseFloat(z.selling_price) || 0;
      if (zohoPrice <= 0) {
        // Zoho item with no selling_price — skip; nothing meaningful to surface
        continue;
      }

      var ebayPrice = ebayMap.get(skuLower);

      // NO REFERENCE case — MI has no currentPrice AND no startPrice fallback.
      // Typically OOS items where eBay returns null Price. Still INCLUDE these
      // in the audit (user feedback 2026-05-22): when stock returns tomorrow,
      // Zoho's stored selling_price becomes a quotable price again. Picker
      // needs visibility on the Zoho number so they can spot-check before the
      // next customer agrees to buy. Empty eBay/Delta/Pct cells, direction
      // "OOS / NO REF" so CF tints them gray and they sort to the bottom.
      if (ebayPrice == null || ebayPrice <= 0) {
        if (ebayPrice == null) onlyInZoho++;
        else                   noRef++;
        oosRows.push([
          sku,                          // A: SKU
          String(z.item_name || ""),    // B: ITEM NAME
          "",                           // C: EBAY_LIVE — blank (no reference)
          zohoPrice,                    // D: ZOHO_LIVE
          "",                           // E: DELTA $ — blank (can't compute)
          "",                           // F: DELTA % — blank
          "OOS / NO REF",               // G: DIRECTION
          now                           // H: LAST_CHECKED
        ]);
        continue;
      }

      totalAudited++;
      var delta = zohoPrice - ebayPrice;
      var threshold = Math.max(absT, pctT * ebayPrice);
      if (Math.abs(delta) <= threshold) continue;   // within tolerance

      var direction = delta > 0 ? "ZOHO HIGH" : "ZOHO LOW";
      if (delta > 0) zohoHigh++;
      else           zohoLow++;
      totalDelta += delta;

      driftRows.push([
        sku,                                          // A: SKU
        String(z.item_name || ""),                    // B: ITEM NAME
        ebayPrice,                                    // C: EBAY_LIVE
        zohoPrice,                                    // D: ZOHO_LIVE
        delta,                                        // E: DELTA $
        ebayPrice > 0 ? (delta / ebayPrice) : 0,      // F: DELTA %
        direction,                                    // G: DIRECTION
        now                                           // H: LAST_CHECKED
      ]);
    }

    // Only-in-MI count (sanity check — items eBay has that Zoho doesn't)
    var onlyInMi = 0;
    ebayMap.forEach(function(_price, sku) {
      if (!seenSkus[sku]) onlyInMi++;
    });

    // Sort drifts by |delta| DESC (worst drifts first), then append OOS rows
    // (sorted by SKU ASC for stable scan). Picker's eye lands on the
    // actionable items at the top; OOS section is a separate watch list below.
    driftRows.sort(function(a, b) { return Math.abs(b[4]) - Math.abs(a[4]); });
    oosRows.sort(function(a, b) {
      return String(a[0]).localeCompare(String(b[0]));
    });
    var rows = driftRows.concat(oosRows);

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
    var driftCount = driftRows.length;
    var oosCount   = oosRows.length;
    return {
      ok:           true,
      message:      driftCount + " drift(s) · " + oosCount + " OOS · net " + sign + "$" + totalDelta.toFixed(2),
      totalAudited: totalAudited,
      mismatches:   driftCount,
      oosRowCount:  oosCount,
      zohoHigh:     zohoHigh,
      zohoLow:      zohoLow,
      totalDelta:   parseFloat(totalDelta.toFixed(2)),
      noRef:        noRef,
      onlyInZoho:   onlyInZoho,
      onlyInMi:     onlyInMi,
      durationSec:  durationSec
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
