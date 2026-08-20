// =======================================================================================
// KitHealth.js — unified "Kit Health" cockpit: price drift + buildability, one row per kit
// Shipped 2026-07-25. (Kit Price Audit phase 2, elevated — see roadmap #1/#2.)
// =======================================================================================
//
// WHAT
//   Runs the SHIPPED kit-pricing engine (computeKitPrice) + the SHIPPED OOS
//   buildability math (_oosComputeKitBuild) over EVERY registry kit, and writes
//   one row per kit answering both questions at a glance:
//
//     PRICE   — is the kit's own listed price in line with what its parts are
//               worth?  computed = round( (1−discount) × Σ(qty × unit price) ).
//               Δ = listed − computed. Negative = UNDERPRICED (leaving money on
//               the table — the money-losing case, sorted to the very top).
//     BUILD   — how many complete kits could I assemble from component stock
//               right now?  min over components of floor(available ÷ qty-per-kit).
//
//   This is the cockpit the roadmap flagged as the elevated alternative to a
//   price-only audit: restock a piston → glance here to see which kits it
//   unblocks AND whether they're priced right, in one pass.
//
// THE ENGINES ARE REUSED UNCHANGED
//   - computeKitPrice(components, {maps})           — KitPricing.js
//   - _resolveComponentPrice(sku, maps)             — KitPricing.js (kit's own listed $)
//   - _buildKitComponentPriceMaps()                 — KitPricing.js (ONE MI+Zoho read)
//   - _oosComputeKitBuild(kit, resolveAvail)        — OutOfStock.js
//   - _oosResolveAvailFactory(zohoMap, invMap)      — OutOfStock.js (Zoho-first→MI)
//   - buildKitMap()                                 — KitRegistry.js
//   - buildLocationAndInventoryMaps()               — LiveSync.js (MI avail + location)
//
//   PERF NOTE: the audit builds every map ONCE and calls computeKitPrice on
//   kit.components directly — it does NOT call computeKitPriceBySku/getKitInfo
//   per kit (those each rebuild the whole kit map → O(n²) over the registry).
//
// AUTO-CALIBRATION (2026-07-25 — the fix for the "whole catalog flags underpriced"
//   problem). A FIXED 10% assumption flagged ~60% of the catalog because these
//   kits are really priced ~15% off parts. So the audit no longer assumes a
//   number: pass 1 measures every complete+listed kit's implied discount
//   (1 − listed/partsValue) and takes the CATALOG MEDIAN; pass 2 prices + classifies
//   at that discount. "UNDERPRICED" now means "discounted MORE than your own norm"
//   (the true outliers), "OVERPRICED" means "less than your norm". The calibrated
//   discount is shown in the KPI band. Empty/degenerate catalog → falls back to
//   KIT_PRICING.discount, clamped to [0, 0.6].
//
// IMPLIED-DISCOUNT COLUMN (DISC%)
//   DISC% = 1 − listed/partsValue is still shown per kit — it's the raw signal the
//   calibration is built on, and lets a human eyeball how each kit sits vs the
//   catalog median (shown in the band). Decision-support — no magic threshold.
//
// HONESTY RULE (same as OOS BUILDABLE / the kit-expansion modal / the calculator)
//   A kit with unparsed PD lines (registry ⚠ UNPARSED rows) or any unpriced
//   component has an UNTRUSTWORTHY computed price — computeKitPrice only sees the
//   PARSED components, so its total would be understated. Those kits are marked
//   ⚠ INCOMPLETE, the computed/Δ cells are LEFT BLANK (never a wrong number),
//   and they park at the bottom. A kit with components but no resolvable listed
//   price is NO LISTED $ — its computed IS a "list it at $X" suggestion.
//
// SHARED-COMPONENT CAVEAT (roadmap #2): BUILDABLE is per-kit and INDEPENDENT —
//   it assumes you build only THAT kit. The same piston feeding three kits shows
//   its full count on each row; the numbers are deliberately NOT summed into a
//   false "total kits" (a real total would need contention accounting).
//
// SCOPE: read-only decision-support. No writes, no /exec, no auto-push.
// DEPLOY: editor-bound (sidebar google.script.run + a weekly time trigger).
//   `clasp push` is the whole deploy; the trigger installs via one editor run of
//   setupKitHealthTrigger() (clasp can't install triggers). NO New Version.
// =======================================================================================


var KIT_HEALTH = {
  sheetName: "Kit Health",

  cols: {
    KIT_SKU:      1,   // A
    KIT_NAME:     2,   // B
    LOCATION:     3,   // C — kit's own aisle (NOT FOUND when it has no shelf)
    TYPE:         4,   // D — READY / MANUAL
    KIT_QTY:      5,   // E — kit's OWN on-hand stock (Zoho-first → MI). "How many
                       //     of THIS kit do we have?" Read against BUILDABLE it
                       //     tells the story: "3 on the shelf but 0 buildable = a
                       //     sub-item ran out." Blank when neither source knows it.
    PARTS_VALUE:  6,   // F — Σ(qty × unit price)  (rawSum; blank when incomplete)
    COMPUTED:     7,   // G — round((1−disc) × rawSum)  (the suggested price)
    LISTED:       8,   // H — kit's own eBay-resolved price
    DELTA:        9,   // I — listed − computed  ($)   negative = UNDERPRICED
    PCT_DELTA:    10,  // J — delta / computed  (%)
    DISC_PCT:     11,  // K — implied discount = 1 − listed/partsValue
    PRICE_STATUS: 12,  // L — UNDERPRICED / OVERPRICED / IN LINE / NO LISTED $ / ⚠ INCOMPLETE
    BUILDABLE:    13,  // M — complete kits assemblable now (number, or ⚠)
    LIMITED_BY:   14,  // N — bottleneck component, NAMED ("Piston Kit (167517) · has 12 / needs 6")
    COMPONENTS:   15,  // O — price-completeness ("8 ok" / "⚠ 2 unpriced" / "⚠ 1 PD unreadable")
    STOCK_STATUS: 16,  // P — KIT_QTY × BUILDABLE verdict: STOCK+BUILD / IN STOCK / BUILD-ONLY / OOS / ⚠
    AT_RISK:      17,  // Q — units ADVERTISED that we cannot currently assemble
                       //     (advertised − buildable, MANUAL kits only, blank when covered).
                       //     "advertised" is the HIGHER of eBay and Zoho: either channel
                       //     can take the order, so the exposure is the louder of the two.
    LAST_CHECKED: 18   // R
  },

  idx: function (name) { return KIT_HEALTH.cols[name] - 1; },

  dataWidth:    18,
  // Dashboard layout: row 1 = KPI summary band, row 2 = column headers (both
  // frozen), data from row 3. setupKitHealthSheet migrates the pre-band layout
  // (headers at row 1) forward automatically.
  bannerRow:    1,
  headerRow:    2,
  dataStartRow: 3,

  headers: ["📦 KIT SKU", "KIT NAME", "LOC", "TYPE", "◫ KIT QTY", "PARTS $", "COMPUTED $",
            "LISTED $", "Δ $", "Δ %", "DISC %", "PRICE STATUS",
            "BUILDABLE", "LIMITED BY", "COMPONENTS", "STOCK STATUS", "⚠ AT RISK", "⏱ CHECKED"],

  // Actionable price-drift tolerance. Δ within max($2, 3% of computed) = IN LINE.
  // The engine rounds to the nearest dollar, so a couple dollars of slack keeps
  // rounding + minor component-price granularity from false-flagging.
  threshold: { abs: 2.00, pct: 0.03 },

  // PRICE_STATUS enum (kept here so the badge + CF + sort agree on one spelling)
  status: {
    UNDER:      "UNDERPRICED",
    OVER:       "OVERPRICED",
    INLINE:     "IN LINE",
    NO_LISTED:  "NO LISTED $",
    INCOMPLETE: "⚠ INCOMPLETE"
  },

  // STOCK_STATUS enum — the KIT_QTY × BUILDABLE verdict (CF + KPI count agree on spelling)
  stock: {
    STOCK_BUILD: "STOCK+BUILD",  // have some AND can make more (healthiest)
    IN_STOCK:    "IN STOCK",     // have some but can't replenish (the "blocked" watch case)
    BUILD_ONLY:  "BUILD-ONLY",   // none assembled, but N buildable from parts
    OOS:         "OOS",          // none on shelf AND can't build (urgent)
    UNKNOWN:     "⚠",            // buildability untrustable → can't judge stock

    // ⚠⚠ THE OVERSELL PAIR (2026-08-19). A MANUAL kit has NO assembled box —
    // its listed quantity is a PROMISE backed by component stock, typed once and
    // never walked down as parts sold. So "we are advertising more than we can
    // assemble" is a real oversell exposure, and the old label for exactly that
    // state was "IN STOCK", which read as reassurance. These two name it instead.
    // READY kits keep IN STOCK: a box on a K-* shelf is real inventory, and
    // "can't build more" is a fact about resupply, not a promise we can't keep.
    CANT_BUILD:  "⚠ CAN'T BUILD",  // advertised, buildable 0 — the NEXT sale fails
    OVER_LISTED: "⚠ OVER-LISTED"   // advertised more than we can assemble, partly covered
  }
};


/** STOCK STATUS verdict from the kit's own on-hand qty × its buildable count.
 *  This is the KIT_QTY-vs-BUILDABLE contrast turned into one scannable label.
 *  Pure — Node-testable. */
function _kitStockStatus(kitQty, buildable, kitType, advertised) {
  if (typeof buildable !== 'number') return KIT_HEALTH.stock.UNKNOWN;   // ⚠ buildable

  var adv = (typeof advertised === 'number') ? advertised
          : (typeof kitQty === 'number' ? kitQty : 0);
  var isManual = String(kitType || "MANUAL").toUpperCase() !== "READY";

  // MANUAL kit advertised beyond what its components can assemble → oversell.
  if (isManual && adv > 0 && adv > buildable) {
    return (buildable === 0) ? KIT_HEALTH.stock.CANT_BUILD
                             : KIT_HEALTH.stock.OVER_LISTED;
  }

  var haveStock = (typeof kitQty === 'number' && kitQty > 0);
  var canBuild  = buildable > 0;
  if (haveStock && canBuild)  return KIT_HEALTH.stock.STOCK_BUILD;
  if (haveStock && !canBuild) return KIT_HEALTH.stock.IN_STOCK;
  if (!haveStock && canBuild) return KIT_HEALTH.stock.BUILD_ONLY;
  return KIT_HEALTH.stock.OOS;
}

/** Units advertised that cannot currently be assembled. MANUAL kits only —
 *  a READY kit's quantity is a box on a shelf, not a promise. Returns "" when
 *  covered, so the column stays quiet except where it matters. Pure. */
function _kitAtRisk(advertised, buildable, kitType) {
  if (typeof buildable !== 'number') return "";
  if (String(kitType || "MANUAL").toUpperCase() === "READY") return "";
  var adv = (typeof advertised === 'number') ? advertised : 0;
  var gap = adv - buildable;
  return gap > 0 ? gap : "";
}

/** Median of a numeric array (0 on empty). Used to calibrate the audit to the
 *  catalog's OWN implied discount instead of a fixed assumption. */
function _median(nums) {
  if (!nums || nums.length === 0) return 0;
  var a = nums.slice().sort(function (x, y) { return x - y; });
  var mid = Math.floor(a.length / 2);
  return (a.length % 2) ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}


// =======================================================================================
// SETUP — idempotent styling / CF (mirrors setupPriceAuditSheet)
// =======================================================================================

function setupKitHealthSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(KIT_HEALTH.sheetName);
  if (!sheet) sheet = ss.insertSheet(KIT_HEALTH.sheetName);

  // --- MIGRATE the pre-dashboard layout forward ---
  // The first Kit Health layout put column headers at row 1. The dashboard
  // layout puts a KPI band at row 1 and headers at row 2. Detect the old shape
  // (📦 header glyph at A1 but NOT at A2) and insert a row so the header slides
  // to row 2; the stale data below rides down and is regenerated on the next
  // audit. Idempotent: on the new layout A1 holds the KPI string (no 📦) so this
  // never fires twice.
  var a1v = String(sheet.getRange(1, 1).getValue() || "").trim();
  var a2v = String(sheet.getRange(2, 1).getValue() || "").trim();
  if (a1v.indexOf("📦") === 0 && a2v.indexOf("📦") !== 0) {
    sheet.insertRowBefore(1);
  }

  // --- HEADERS (row 2) ---
  sheet.getRange(KIT_HEALTH.headerRow, 1, 1, KIT_HEALTH.dataWidth)
    .setValues([KIT_HEALTH.headers])
    .setBackground('#1d1d1b')
    .setFontColor('#ffd966')
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(KIT_HEALTH.headerRow, 1, 1, KIT_HEALTH.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(KIT_HEALTH.cols.KIT_SKU,       95);
  sheet.setColumnWidth(KIT_HEALTH.cols.KIT_NAME,      235);
  sheet.setColumnWidth(KIT_HEALTH.cols.LOCATION,      75);
  sheet.setColumnWidth(KIT_HEALTH.cols.TYPE,          72);
  sheet.setColumnWidth(KIT_HEALTH.cols.KIT_QTY,       78);
  sheet.setColumnWidth(KIT_HEALTH.cols.PARTS_VALUE,   90);
  sheet.setColumnWidth(KIT_HEALTH.cols.COMPUTED,      95);
  sheet.setColumnWidth(KIT_HEALTH.cols.LISTED,        95);
  sheet.setColumnWidth(KIT_HEALTH.cols.DELTA,         85);
  sheet.setColumnWidth(KIT_HEALTH.cols.PCT_DELTA,     65);
  sheet.setColumnWidth(KIT_HEALTH.cols.DISC_PCT,      68);
  sheet.setColumnWidth(KIT_HEALTH.cols.PRICE_STATUS,  128);
  sheet.setColumnWidth(KIT_HEALTH.cols.BUILDABLE,     82);
  sheet.setColumnWidth(KIT_HEALTH.cols.LIMITED_BY,    260);
  sheet.setColumnWidth(KIT_HEALTH.cols.COMPONENTS,    150);
  sheet.setColumnWidth(KIT_HEALTH.cols.STOCK_STATUS,  120);
  sheet.setColumnWidth(KIT_HEALTH.cols.AT_RISK,       88);
  sheet.setColumnWidth(KIT_HEALTH.cols.LAST_CHECKED,  140);

  // --- DATA AREA: column-level formats so re-runs inherit ---
  var maxDataRow = 2500;
  var dsr = KIT_HEALTH.dataStartRow;
  var dataRows = maxDataRow - dsr + 1;

  sheet.getRange(dsr, KIT_HEALTH.cols.KIT_SKU, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(dsr, KIT_HEALTH.cols.KIT_NAME, dataRows, 1)
    .setFontFamily('Roboto').setFontSize(10).setHorizontalAlignment('left');
  sheet.getRange(dsr, KIT_HEALTH.cols.LOCATION, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(10).setFontColor('#5f5f5f').setHorizontalAlignment('center');
  sheet.getRange(dsr, KIT_HEALTH.cols.TYPE, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  sheet.getRange(dsr, KIT_HEALTH.cols.KIT_QTY, dataRows, 1)
    .setNumberFormat('0').setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(dsr, KIT_HEALTH.cols.PARTS_VALUE, dataRows, 1)
    .setNumberFormat('$#,##0.00').setFontFamily('Roboto Mono').setFontSize(10).setFontColor('#5f5f5f').setHorizontalAlignment('right');
  sheet.getRange(dsr, KIT_HEALTH.cols.COMPUTED, dataRows, 1)
    .setNumberFormat('$#,##0.00').setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(dsr, KIT_HEALTH.cols.LISTED, dataRows, 1)
    .setNumberFormat('$#,##0.00').setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(dsr, KIT_HEALTH.cols.DELTA, dataRows, 1)
    .setNumberFormat('+$#,##0.00;-$#,##0.00').setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(dsr, KIT_HEALTH.cols.PCT_DELTA, dataRows, 1)
    .setNumberFormat('+0.0%;-0.0%').setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(dsr, KIT_HEALTH.cols.DISC_PCT, dataRows, 1)
    .setNumberFormat('0.0%').setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(dsr, KIT_HEALTH.cols.PRICE_STATUS, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(dsr, KIT_HEALTH.cols.BUILDABLE, dataRows, 1)
    .setNumberFormat('0').setFontFamily('Oswald').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sheet.getRange(dsr, KIT_HEALTH.cols.LIMITED_BY, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9).setFontColor('#434343').setHorizontalAlignment('left');
  sheet.getRange(dsr, KIT_HEALTH.cols.COMPONENTS, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9).setHorizontalAlignment('center');
  sheet.getRange(dsr, KIT_HEALTH.cols.STOCK_STATUS, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  /* ⚠⚠ THE '0' IS LOAD-BEARING, AND ITS ABSENCE COST A MORNING (2026-08-21).
     AT_RISK was added at column 17 on 2026-08-19 — exactly where the old 17-wide
     schema had LAST_CHECKED, a DATE column — and it was the ONLY data column in
     this block with no setNumberFormat. Sheets keeps a cell's number format
     through clearContent, so every gap this audit wrote rendered as a 1900 date
     and came BACK from getValues() as a Date object. getKitOversellSnapshot then
     parseFloat'd a Date, got NaN, skipped it, and reported "30 kits · 0 UNITS" —
     a customer-facing exposure silently understated to zero.
     THIRD instance of this class here: OOS DAYS OUT (2026-07-18) and the Zoho
     Stock SELLING PRICE column (2026-05-28) were the first two.
     RULE: a column that holds code-written NUMBERS must SET its number format.
     Inheriting one is not neutral — it is a silent type change on read. */
  sheet.getRange(dsr, KIT_HEALTH.cols.AT_RISK, dataRows, 1)
    .setNumberFormat('0').setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(dsr, KIT_HEALTH.cols.LAST_CHECKED, dataRows, 1)
    .setNumberFormat('M/d/yy h:mm am/pm')
    .setFontFamily('Roboto Mono').setFontSize(9).setFontColor('#5f5f5f').setHorizontalAlignment('center');

  sheet.getRange(dsr, 1, dataRows, KIT_HEALTH.dataWidth).setVerticalAlignment('middle');

  // --- BANDING (cream alternation) — starts at the HEADER row so row 1 (the KPI
  // band) is left free; banding's header slot would otherwise repaint it (the
  // Prep/OOS title-band lesson). ---
  sheet.getBandings().forEach(function (b) { try { b.remove(); } catch (e) {} });
  var bandRange = sheet.getRange(KIT_HEALTH.headerRow, 1,
                                 maxDataRow - KIT_HEALTH.headerRow + 1, KIT_HEALTH.dataWidth);
  var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b')
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#fff8e7');

  // --- CONDITIONAL FORMATTING ---
  // Strip prior KitHealth-scoped rules (PRICE_STATUS / DELTA / DISC_PCT /
  // BUILDABLE columns) then re-add — idempotent.
  var existing = sheet.getConditionalFormatRules() || [];
  var keep = existing.filter(function (r) {
    var ranges = r.getRanges();
    if (!ranges || ranges.length === 0) return true;
    return !ranges.some(function (rg) {
      if (rg.getSheet().getName() !== KIT_HEALTH.sheetName) return false;
      var c = rg.getColumn();
      return c === KIT_HEALTH.cols.PRICE_STATUS || c === KIT_HEALTH.cols.DELTA
          || c === KIT_HEALTH.cols.DISC_PCT     || c === KIT_HEALTH.cols.BUILDABLE
          || c === KIT_HEALTH.cols.STOCK_STATUS || c === KIT_HEALTH.cols.AT_RISK;
    });
  });

  // PRICE_STATUS — one color per verdict. UNDERPRICED is the loud one (red,
  // money leak); OVERPRICED amber; IN LINE green; NO LISTED $ slate (a listing
  // opportunity); ⚠ INCOMPLETE gray (can't trust the number).
  var statusRange = sheet.getRange(dsr, KIT_HEALTH.cols.PRICE_STATUS, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_HEALTH.status.UNDER)
    .setBackground('#ffcdd2').setFontColor('#b71c1c').setBold(true)
    .setRanges([statusRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_HEALTH.status.OVER)
    .setBackground('#ffd699').setFontColor('#7a3d00').setBold(true)
    .setRanges([statusRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_HEALTH.status.INLINE)
    .setBackground('#e8f5e9').setFontColor('#1b5e20').setBold(true)
    .setRanges([statusRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_HEALTH.status.NO_LISTED)
    .setBackground('#cfd8dc').setFontColor('#37474f').setBold(true)
    .setRanges([statusRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_HEALTH.status.INCOMPLETE)
    .setBackground('#f0f0f0').setFontColor('#5f5f5f').setBold(false)
    .setRanges([statusRange]).build());

  // DELTA — tint by sign. Negative (underpriced/losing money) = red; positive
  // (overpriced/uncompetitive) = amber. Matches the PRICE_STATUS palette.
  var deltaRange = sheet.getRange(dsr, KIT_HEALTH.cols.DELTA, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setFontColor('#b71c1c')
    .setRanges([deltaRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setFontColor('#7a3d00')
    .setRanges([deltaRange]).build());

  // DISC_PCT — a soft "typical kit-discount range" cue (8%–20%) reads green;
  // the audit now calibrates to the catalog's own median (~15%), so this band is
  // just a rough eye-guide — the real verdict lives in PRICE STATUS. Discounts
  // far outside the band (a near-zero or a very deep one) stay default so they
  // catch the eye.
  var discRange = sheet.getRange(dsr, KIT_HEALTH.cols.DISC_PCT, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.08, 0.20).setFontColor('#1b5e20')
    .setRanges([discRange]).build());

  // BUILDABLE — >0 green (can assemble some now). ⚠ amber (untrustable). 0 left
  // plain: for the whole-catalog view, "no spare component stock" is common and
  // NOT alarming, so no red here (unlike the OOS kit table where 0 = blocked).
  var buildRange = sheet.getRange(dsr, KIT_HEALTH.cols.BUILDABLE, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#c8e6c9').setFontColor('#1b5e20').setBold(true)
    .setRanges([buildRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('⚠')
    .setBackground('#fff3b0').setFontColor('#7a5c00').setBold(true)
    .setRanges([buildRange]).build());

  // AT RISK — units advertised we cannot assemble. Blank on a healthy kit, so any
  // ink in this column is a finding. Deepens with the size of the exposure.
  var riskRange = sheet.getRange(dsr, KIT_HEALTH.cols.AT_RISK, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(3)
    .setBackground('#ffcdd2').setFontColor('#b71c1c').setBold(true)
    .setRanges([riskRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#ffe0b2').setFontColor('#7a3d00').setBold(true)
    .setRanges([riskRange]).build());

  // STOCK_STATUS — the KIT_QTY × BUILDABLE verdict. STOCK+BUILD green (healthiest);
  // IN STOCK amber (have some but can't replenish — the watch case); BUILD-ONLY
  // cool slate (info); OOS red (urgent); ⚠ gray (can't judge).
  var stockRange = sheet.getRange(dsr, KIT_HEALTH.cols.STOCK_STATUS, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_HEALTH.stock.STOCK_BUILD)
    .setBackground('#c8e6c9').setFontColor('#1b5e20').setBold(true)
    .setRanges([stockRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_HEALTH.stock.IN_STOCK)
    .setBackground('#ffe0b2').setFontColor('#7a3d00').setBold(true)
    .setRanges([stockRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_HEALTH.stock.BUILD_ONLY)
    .setBackground('#e1f0f5').setFontColor('#2a6478').setBold(true)
    .setRanges([stockRange]).build());
  // ⚠ CAN'T BUILD — advertised and the next sale cannot be fulfilled. RED, and
  // it is the ONE place on this sheet red is spent, because it is the only state
  // here that reaches a customer. OVER-LISTED is amber: exposed, but partly covered.
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_HEALTH.stock.CANT_BUILD)
    .setBackground('#ffcdd2').setFontColor('#b71c1c').setBold(true)
    .setRanges([stockRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_HEALTH.stock.OVER_LISTED)
    .setBackground('#ffe0b2').setFontColor('#7a3d00').setBold(true)
    .setRanges([stockRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_HEALTH.stock.OOS)
    .setBackground('#ffcdd2').setFontColor('#b71c1c').setBold(true)
    .setRanges([stockRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_HEALTH.stock.UNKNOWN)
    .setBackground('#f0f0f0').setFontColor('#5f5f5f').setBold(false)
    .setRanges([stockRange]).build());

  sheet.setConditionalFormatRules(keep);

  // --- KPI SUMMARY BAND (row 1) — styled AFTER banding so it isn't repainted.
  // Dark full-width band; runKitHealthAudit writes the live KPI string into A1,
  // left-aligned so it overflows across the (empty) band cells — no merge, which
  // keeps banding/insert logic simple. ---
  sheet.getRange(KIT_HEALTH.bannerRow, 1, 1, KIT_HEALTH.dataWidth)
    .setBackground('#1d1d1b')
    // OVERFLOW, not wrap — the KPI string lives in A1 and must flow across the
    // empty band cells on ONE line. Without this the migrated row 1 inherits
    // wrap=true from the old header and stacks the text vertically in column A.
    .setWrap(false)
    .setBorder(null, null, true, null, null, null, '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange(KIT_HEALTH.bannerRow, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(11).setFontColor('#ffffff')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  if (!String(sheet.getRange(KIT_HEALTH.bannerRow, 1).getValue() || "").trim()) {
    sheet.getRange(KIT_HEALTH.bannerRow, 1).setValue("🩺 KIT HEALTH — run the audit to populate the summary");
  }
  sheet.setRowHeight(KIT_HEALTH.bannerRow, 40);

  sheet.setFrozenRows(KIT_HEALTH.headerRow);   // KPI band + headers both stay pinned

  return "✅ Kit Health sheet ready.";
}


/** Sidebar: switch view to the Kit Health sheet (create if missing). */
function openKitHealth() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet";
  var sheet = ss.getSheetByName(KIT_HEALTH.sheetName);
  if (!sheet) {
    setupKitHealthSheet();
    sheet = ss.getSheetByName(KIT_HEALTH.sheetName);
  }
  ss.setActiveSheet(sheet);
  return "✅ Opened Kit Health";
}


// =======================================================================================
// CLASSIFICATION — pure helper (Node-testable in isolation, no Sheets calls)
// =======================================================================================
//
// Given one kit's priced result + buildability + its own listed price, decide
// the PRICE_STATUS and assemble the 15-wide sheet row. Kept pure so the
// bucket/sort logic can be unit-checked without a spreadsheet.
//
// priced   : computeKitPrice() output over kit.components
// build    : _oosComputeKitBuild() output ({buildable, limitedBy, limiter, components})
// listed   : kit's own resolved price (number) or null
// kitQty   : kit's own on-hand stock (number) or "" when unknown
// now      : Date for LAST_CHECKED
//
// Returns { row:[17], bucket, status } where bucket ∈
//   "under" | "over" | "inline" | "noListed" | "incomplete"
// =======================================================================================

/** Bottleneck text for LIMITED BY — surfaces the component's NAME so a reviewer
 *  reads "Piston Kit (167517) · has 12 / needs 6" instead of a bare SKU. Falls
 *  back to build.limitedBy for the ⚠ cases (which carry a message, no limiter). */
function _kitHealthLimitedByText(build) {
  if (build && build.limiter) {
    var L = build.limiter;
    var nm = L.name ? String(L.name) : "";
    if (nm.length > 30) nm = nm.substring(0, 29) + "…";
    return (nm ? nm + " " : "") + "(" + L.sku + ") · has " + L.avail + " / needs " + L.qtyPer;
  }
  return (build && build.limitedBy) || "";
}

function _kitHealthClassify(kit, priced, build, listed, kitQty, now, discount, threshold, advertised) {
  var S = KIT_HEALTH.status;
  var hasUnparsed = (kit.unparsedLines || []).length > 0;
  var noComps     = (kit.components   || []).length === 0;
  var hasUnpriced = (priced.unpricedComponents || []).length > 0;
  var complete    = !hasUnparsed && !noComps && !hasUnpriced;

  // COMPONENTS text — price-completeness lens (buildability has its own columns).
  // "N bundled" names parts that ship INSIDE another component and are therefore
  // deliberately out of the price and the build math — stated, never silent, so a
  // reviewer can see WHY the parts value is lower than the line count suggests.
  var nBundled = (priced.excludedComponents || []).length;
  var bundledNote = nBundled ? " · " + nBundled + " bundled" : "";
  var compsText;
  if (hasUnparsed)      compsText = "⚠ " + kit.unparsedLines.length + " PD unreadable";
  else if (noComps)     compsText = "⚠ no components";
  else if (hasUnpriced) compsText = "⚠ " + priced.unpricedComponents.length + " unpriced" + bundledNote;
  else                  compsText = ((kit.components.length - nBundled) + " ok") + bundledNote;

  var loc     = kit._loc || kit.location || "NOT FOUND";
  var type    = kit.type || "MANUAL";
  var kitName = kit.name || "";

  var limitedBy   = _kitHealthLimitedByText(build);
  var adv         = (typeof advertised === 'number') ? advertised
                  : (typeof kitQty === 'number' ? kitQty : 0);
  var stockStatus = _kitStockStatus(kitQty, build.buildable, type, adv);
  var atRisk      = _kitAtRisk(adv, build.buildable, type);

  // Untrustworthy computed → ⚠ INCOMPLETE, blank the price math, show listed if any.
  if (!complete) {
    return {
      status: S.INCOMPLETE, bucket: "incomplete", stockStatus: stockStatus,
      row: [kit.sku, kitName, loc, type, kitQty, "", "", (listed != null ? listed : ""),
            "", "", "", S.INCOMPLETE, build.buildable, limitedBy, compsText, stockStatus, atRisk, now]
    };
  }

  var parts    = priced.rawSum;
  var computed = priced.roundedTotal;

  // No own listed price → the computed number is a "list it at $X" suggestion.
  if (listed == null) {
    return {
      status: S.NO_LISTED, bucket: "noListed", stockStatus: stockStatus,
      row: [kit.sku, kitName, loc, type, kitQty, parts, computed, "",
            "", "", "", S.NO_LISTED, build.buildable, limitedBy, compsText, stockStatus, atRisk, now]
    };
  }

  var delta   = listed - computed;                        // negative = underpriced
  var pct     = computed > 0 ? delta / computed : 0;
  var discPct = parts > 0 ? (1 - listed / parts) : "";
  var tol     = Math.max(threshold.abs, threshold.pct * computed);

  var status, bucket;
  if (Math.abs(delta) <= tol) { status = S.INLINE; bucket = "inline"; }
  else if (delta < 0)         { status = S.UNDER;  bucket = "under";  }
  else                        { status = S.OVER;   bucket = "over";   }

  return {
    status: status, bucket: bucket, stockStatus: stockStatus,
    row: [kit.sku, kitName, loc, type, kitQty, parts, computed, listed,
          delta, pct, discPct, status, build.buildable, limitedBy, compsText, stockStatus, atRisk, now]
  };
}


// =======================================================================================
// MAIN — runKitHealthAudit (sidebar button + weekly trigger)
// =======================================================================================
//
// Returns a summary for the sidebar status bar:
//   { ok, message, totalKits, underpriced, overpriced, inLine, noListed,
//     incomplete, buildableNow, totalUnderBy, durationSec }
// =======================================================================================

function runKitHealthAudit() {
  var start = Date.now();
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(KIT_HEALTH.sheetName);
    if (!sheet) {
      setupKitHealthSheet();
      sheet = ss.getSheetByName(KIT_HEALTH.sheetName);
    } else {
      // Auto-migrate a pre-dashboard sheet (headers still at row 1, no KPI band)
      // so the audit never writes 17-wide rows into an old-shaped sheet. Cheap
      // one-cell check; setupKitHealthSheet's own guard does the row insert.
      var a2 = String(sheet.getRange(2, 1).getValue() || "").trim();
      // Also re-run setup when the sheet is NARROWER than the current schema —
      // the 2026-08-19 AT RISK column widened it 17 → 18, and writing 18-wide rows
      // under a 17-wide header would leave the last column unlabelled and unformatted.
      var lastHdr = String(sheet.getRange(KIT_HEALTH.headerRow, KIT_HEALTH.dataWidth).getValue() || "").trim();
      if (a2.indexOf("📦") !== 0 || lastHdr !== KIT_HEALTH.headers[KIT_HEALTH.dataWidth - 1]) {
        setupKitHealthSheet();
        sheet = ss.getSheetByName(KIT_HEALTH.sheetName);
      }
    }

    var kitMap = buildKitMap();
    if (!kitMap || kitMap.size === 0) {
      return { ok: false,
               message: "Kit Registry is empty — import kits first (Kit Registry → import).",
               durationSec: ((Date.now() - start) / 1000).toFixed(1) };
    }

    // ONE MI+Zoho read for prices, ONE MI read for availability/location.
    var maps    = _buildKitComponentPriceMaps();          // {ebay(price), zoho(stock)}
    var invMaps = buildLocationAndInventoryMaps();          // {locationMap, inventoryMap}
    var resolveAvail = _oosResolveAvailFactory(maps.zoho, invMaps.inventoryMap);

    var threshold = KIT_HEALTH.threshold;
    var now = new Date();

    // ------------------------------------------------------------------------
    // PASS 1 — gather each kit + collect the catalog's ACTUAL implied discounts.
    // A fixed 10% assumption flagged ~60% of the catalog as underpriced because
    // these kits are really priced ~15% off parts. So we CALIBRATE the baseline
    // to the catalog's own median implied discount (impliedDisc = 1 − listed/parts,
    // per complete + listed kit). "Underpriced" then means "discounted MORE than
    // your own norm" — the true outliers — instead of "cheaper than a guessed 10%."
    // rawSum is discount-independent, so pass 1 prices at the default just to read
    // rawSum + price-completeness.
    // ------------------------------------------------------------------------
    var entries = [];
    var impliedDiscs = [];
    var buildableNow = 0;
    kitMap.forEach(function (kit, skuLower) {
      kit._loc = invMaps.locationMap.get(skuLower) || kit.location || "NOT FOUND";

      var priced0 = computeKitPrice(kit.components, { maps: maps });   // rawSum + completeness
      var build   = _oosComputeKitBuild(kit, resolveAvail);
      if (typeof build.buildable === 'number' && build.buildable > 0) buildableNow++;

      var listed = _resolveComponentPrice(kit.sku, maps).price;
      var kitQty = resolveAvail(skuLower);
      if (kitQty === null || kitQty === undefined) kitQty = "";

      // ⚠ ADVERTISED = the HIGHER of eBay and Zoho, deliberately.
      // KIT QTY follows the house rule (Zoho-first → MI). But for an OVERSELL
      // question the number that matters is what a buyer can still order, and the
      // two channels can disagree — kit 215756 sat at eBay 0 / Zoho 1 on 08-19.
      // Either channel can take that order, so the exposure is the louder of the
      // two. Same fail-toward-showing instinct as the rest of this system.
      var zoAvail = maps.zoho.get(skuLower);
      var miRec   = invMaps.inventoryMap.get(skuLower);
      var advertised = Math.max(
        (zoAvail && zoAvail.available != null) ? zoAvail.available : 0,
        (miRec  && miRec.available  != null) ? miRec.available  : 0
      );

      var complete = (kit.unparsedLines || []).length === 0
                  && (kit.components   || []).length > 0
                  && priced0.unpricedComponents.length === 0;
      if (complete && listed != null && priced0.rawSum > 0) {
        impliedDiscs.push(1 - listed / priced0.rawSum);
      }
      entries.push({ kit: kit, build: build, listed: listed, kitQty: kitQty, advertised: advertised });
    });

    // Calibrated baseline = catalog median implied discount (fallback to the
    // KIT_PRICING default when nothing is priceable). Clamp to a sane band so a
    // freak data point can't invert the model.
    var catalogDiscount = impliedDiscs.length ? _median(impliedDiscs) : KIT_PRICING.discount;
    catalogDiscount = Math.max(0, Math.min(0.6, catalogDiscount));

    // ------------------------------------------------------------------------
    // PASS 2 — price + classify every kit at the calibrated discount.
    // ------------------------------------------------------------------------
    var buckets = { under: [], over: [], inline: [], noListed: [], incomplete: [] };
    var counts  = { totalKits: entries.length, buildableNow: buildableNow, totalUnderBy: 0,
                    inStockBlocked: 0, cantBuild: 0, overListed: 0, unitsAtRisk: 0 };

    entries.forEach(function (e) {
      var priced = computeKitPrice(e.kit.components, { maps: maps, discount: catalogDiscount });
      var c = _kitHealthClassify(e.kit, priced, e.build, e.listed, e.kitQty, now, catalogDiscount, threshold, e.advertised);
      buckets[c.bucket].push(c.row);
      if (c.bucket === "under") counts.totalUnderBy += c.row[KIT_HEALTH.idx("DELTA")]; // negative
      if (c.stockStatus === KIT_HEALTH.stock.IN_STOCK) counts.inStockBlocked++;         // have some, can't build
      if (c.stockStatus === KIT_HEALTH.stock.CANT_BUILD)  { counts.cantBuild++;  counts.unitsAtRisk += (c.row[KIT_HEALTH.idx("AT_RISK")] || 0); }
      if (c.stockStatus === KIT_HEALTH.stock.OVER_LISTED) { counts.overListed++; counts.unitsAtRisk += (c.row[KIT_HEALTH.idx("AT_RISK")] || 0); }
    });

    // Sort each bucket, then stack in triage order:
    //   money leaks first (UNDER, biggest loss on top) → competitiveness (OVER,
    //   biggest overprice) → healthy (IN LINE) → listing opportunities
    //   (NO LISTED, biggest suggested price) → broken (INCOMPLETE), at bottom.
    var D = KIT_HEALTH.idx("DELTA");
    var F = KIT_HEALTH.idx("COMPUTED");
    buckets.under.sort(function (a, b) { return a[D] - b[D]; });            // most negative first
    buckets.over.sort(function (a, b) { return b[D] - a[D]; });            // biggest positive first
    buckets.inline.sort(function (a, b) { return String(a[0]).localeCompare(String(b[0])); });
    buckets.noListed.sort(function (a, b) { return (b[F] || 0) - (a[F] || 0); }); // biggest suggested
    buckets.incomplete.sort(function (a, b) { return String(a[0]).localeCompare(String(b[0])); });

    var rows = buckets.under
      .concat(buckets.over)
      .concat(buckets.inline)
      .concat(buckets.noListed)
      .concat(buckets.incomplete);

    // --- WRITE (clear prior data, preserve headers + column formats) ---
    var lastRow = sheet.getLastRow();
    if (lastRow >= KIT_HEALTH.dataStartRow) {
      sheet.getRange(KIT_HEALTH.dataStartRow, 1,
                     lastRow - KIT_HEALTH.dataStartRow + 1, KIT_HEALTH.dataWidth)
           .clearContent();
    }
    if (rows.length > 0) {
      sheet.getRange(KIT_HEALTH.dataStartRow, 1, rows.length, KIT_HEALTH.dataWidth)
           .setValues(rows);
    }

    var durationSec = ((Date.now() - start) / 1000).toFixed(1);
    var underN = buckets.under.length, overN = buckets.over.length;
    var underByStr = "$" + Math.abs(counts.totalUnderBy).toFixed(0);
    var medPct = (catalogDiscount * 100).toFixed(1);

    // --- KPI SUMMARY BAND (row 1) — the headline metrics + the calibrated
    // baseline + a refresh stamp, written into A1 (overflows across the dark
    // band). This is what makes the sheet reviewable at a glance. ---
    var stamp = Utilities.formatDate(now, "America/Chicago", "M/d h:mm a");
    var kpi = "🩺 " + counts.totalKits + " KITS"
            + "   ·   📉 calibrated to " + medPct + "% discount"
            + "   ·   💸 " + underN + " UNDERPRICED (" + underByStr + ")"
            + "   ·   🔧 " + counts.buildableNow + " BUILDABLE NOW"
            + "   ·   ⛔ " + counts.inStockBlocked + " IN-STOCK BLOCKED"
            + "   ·   🚨 " + counts.cantBuild + " CAN'T BUILD (" + counts.unitsAtRisk + " units at risk)"
            + "   ·   ⚠ " + buckets.incomplete.length + " NEED A FIX"
            + "        ⟳ " + stamp;
    sheet.getRange(KIT_HEALTH.bannerRow, 1).setValue(kpi);
    // Keep the band a single overflowing line even on an already-migrated sheet
    // that won't re-run setup (self-heals the inherited-wrap tall-row bug).
    sheet.getRange(KIT_HEALTH.bannerRow, 1, 1, KIT_HEALTH.dataWidth).setWrap(false);
    sheet.setRowHeight(KIT_HEALTH.bannerRow, 40);

    SpreadsheetApp.flush();

    return {
      ok:             true,
      message:        "calibrated " + medPct + "% · " + underN + " underpriced (" + underByStr + ") · "
                      + overN + " overpriced · " + counts.cantBuild + " can't build · "
                      + counts.inStockBlocked + " blocked · " + buckets.incomplete.length + " ⚠",
      totalKits:      counts.totalKits,
      medianDiscount: catalogDiscount,          // fraction; sidebar shows as %
      underpriced:    underN,
      overpriced:     overN,
      inLine:         buckets.inline.length,
      noListed:       buckets.noListed.length,
      incomplete:     buckets.incomplete.length,
      buildableNow:   counts.buildableNow,
      inStockBlocked: counts.inStockBlocked,
      cantBuild:      counts.cantBuild,        // MANUAL, advertised, buildable 0
      overListed:     counts.overListed,       // MANUAL, advertised beyond buildable
      unitsAtRisk:    counts.unitsAtRisk,
      totalUnderBy:   parseFloat(counts.totalUnderBy.toFixed(2)),   // negative sum
      durationSec:    durationSec
    };
  } catch (err) {
    try { console.log("runKitHealthAudit error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return {
      ok: false,
      message: "Kit Health audit failed: " + (err.message || err),
      durationSec: ((Date.now() - start) / 1000).toFixed(1)
    };
  }
}


// =======================================================================================
// SIDEBAR BADGE — cheap actionable count (reads the sheet, does NOT re-audit)
// =======================================================================================
//
// Mirrors getPriceDriftCount(). Counts the actionable price rows (UNDERPRICED +
// OVERPRICED) on the Kit Health sheet — a snapshot of the last audit, safe on
// every 30s alerts poll. NO LISTED $ / ⚠ INCOMPLETE excluded (different concern,
// would inflate the badge). Returns 0 when the sheet is missing/empty.
// =======================================================================================

function getKitPriceDriftCount() {
  try {
    var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(KIT_HEALTH.sheetName);
    if (!sheet) return 0;
    var lastRow = sheet.getLastRow();
    if (lastRow < KIT_HEALTH.dataStartRow) return 0;

    var vals = sheet.getRange(
      KIT_HEALTH.dataStartRow, KIT_HEALTH.cols.PRICE_STATUS,
      lastRow - KIT_HEALTH.dataStartRow + 1, 1
    ).getValues();

    var count = 0;
    for (var i = 0; i < vals.length; i++) {
      var s = String(vals[i][0]).trim().toUpperCase();
      if (s === KIT_HEALTH.status.UNDER || s === KIT_HEALTH.status.OVER) count++;
    }
    return count;
  } catch (e) {
    console.log("getKitPriceDriftCount error: " + e);
    return 0;
  }
}


/**
 * THE OVERSELL EXPOSURE, read from the Kit Health SHEET SNAPSHOT — never a
 * re-audit. Sibling of getKitPriceDriftCount above and it earns its keep the
 * same way: runKitHealthAudit REWRITES and re-sorts the whole sheet and pays for
 * a full kit-map + MI + Zoho pass, so anything sitting on a poll has to read
 * what the last audit already wrote rather than recompute it.
 *
 * ⚠⚠ COUNTS "⚠ CAN'T BUILD" ONLY, NOT "⚠ OVER-LISTED". They are different
 * claims and the resting panel says one of them out loud. OVER-LISTED is
 * advertised-more-than-we-can-assemble but PARTLY covered — the next sale
 * probably still ships. CAN'T BUILD is buildable 0: the next sale FAILS. The
 * panel's row reads "Advertised, can't build", so it gets the number that
 * sentence is actually true of. Widening this to both would inflate the row
 * with kits that are, today, still shippable.
 *
 * ⚠ STOCK_STATUS (P) and AT_RISK (Q) are ADJACENT, so this is ONE 2-wide read.
 * If either column ever moves, this becomes two reads or a wrong one — the
 * range width is derived from nothing, so check it after any schema change.
 *
 * @returns {{kits:number, units:number}} zeros when the sheet is absent, empty,
 *          or unreadable — a count of zero and "I could not look" are
 *          deliberately NOT distinguished here, because the only consumer draws
 *          nothing at zero either way. Do not reuse this for an alarm.
 */
/**
 * ⚠ DIAGNOSTIC (2026-08-21) — why does the oversell snapshot report 0 units?
 *
 * getKitOversellSnapshot() reads 30 rows whose STOCK STATUS is "⚠ CAN'T BUILD"
 * and sums AT RISK to ZERO. Reading the code says that cannot happen: a CAN'T
 * BUILD verdict REQUIRES advertised > 0 and buildable === 0, which forces
 * _kitAtRisk to return advertised. Three separate static readings of the write
 * path all said "impossible", and a fresh audit reproduced it anyway.
 *
 * So stop reading and MEASURE — the standing rule of this project. This dumps
 * what the sheet actually holds, with types, so the next step is decided by data
 * instead of by another guess. Editor-run; output goes to the EXECUTION LOG,
 * because the Run button does not display return values.
 *
 * Delete this once the cause is found.
 */
function diagnoseKitOversellNow() {
  var L = [];
  try {
    var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(KIT_HEALTH.sheetName);
    if (!sheet) { console.log("no Kit Health sheet"); return; }

    var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    L.push("── KIT HEALTH · OVERSELL DIAGNOSTIC ────────────────────────");
    L.push("geometry   lastRow " + lastRow + "  ·  lastCol " + lastCol
           + "   (schema expects dataWidth " + KIT_HEALTH.dataWidth
           + ", data from row " + KIT_HEALTH.dataStartRow + ")");

    // Headers 14..18 — proves whether P/Q/R are where the schema thinks.
    var hdr = sheet.getRange(KIT_HEALTH.headerRow, 14, 1, 5).getValues()[0];
    L.push("headers    N=" + JSON.stringify(hdr[0]) + "  O=" + JSON.stringify(hdr[1])
           + "  P=" + JSON.stringify(hdr[2]) + "  Q=" + JSON.stringify(hdr[3])
           + "  R=" + JSON.stringify(hdr[4]));
    L.push("expected   P=" + JSON.stringify(KIT_HEALTH.headers[15])
           + "  Q=" + JSON.stringify(KIT_HEALTH.headers[16])
           + "  R=" + JSON.stringify(KIT_HEALTH.headers[17]));

    var n = lastRow - KIT_HEALTH.dataStartRow + 1;
    if (n < 1) { console.log(L.join("\n")); return; }

    // Read the WHOLE row so nothing depends on my column arithmetic being right.
    var all = sheet.getRange(KIT_HEALTH.dataStartRow, 1, n, KIT_HEALTH.dataWidth).getValues();

    var iType = KIT_HEALTH.idx("TYPE"), iQty = KIT_HEALTH.idx("KIT_QTY");
    var iBuild = KIT_HEALTH.idx("BUILDABLE"), iStock = KIT_HEALTH.idx("STOCK_STATUS");
    var iRisk = KIT_HEALTH.idx("AT_RISK");

    var cantBuild = 0, overListed = 0, riskNumeric = 0, riskSum = 0, shown = 0;
    for (var i = 0; i < all.length; i++) {
      var st = String(all[i][iStock]).trim().toUpperCase();
      var isCB = st === String(KIT_HEALTH.stock.CANT_BUILD).trim().toUpperCase();
      var isOL = st === String(KIT_HEALTH.stock.OVER_LISTED).trim().toUpperCase();
      if (isCB) cantBuild++;
      if (isOL) overListed++;

      var rv = all[i][iRisk], rn = parseFloat(rv);
      if (!isNaN(rn) && rn > 0) { riskNumeric++; riskSum += rn; }

      // Show the first few of EACH kind, with types — that is the discriminator.
      if ((isCB || isOL) && shown < 8) {
        shown++;
        L.push("row " + (KIT_HEALTH.dataStartRow + i)
             + "  sku " + all[i][0]
             + "  type " + JSON.stringify(all[i][iType])
             + "  kitQty " + JSON.stringify(all[i][iQty])
             + "  buildable " + JSON.stringify(all[i][iBuild]) + " (" + (typeof all[i][iBuild]) + ")"
             + "\n        STOCK " + JSON.stringify(all[i][iStock])
             + "   AT_RISK " + JSON.stringify(rv) + " (" + (typeof rv) + ")");
      }
    }

    L.push("counts     CAN'T BUILD " + cantBuild + "  ·  OVER-LISTED " + overListed
           + "  ·  rows with a numeric AT_RISK > 0: " + riskNumeric
           + "  (sum " + riskSum + ")");
    L.push("verdict    " + (riskNumeric === 0
           ? "column Q is EMPTY across the WHOLE sheet — the writer never populated it"
           : "column Q HAS values — so this is specific to the CAN'T BUILD rows"));
    L.push("────────────────────────────────────────────────────────────");
  } catch (e) {
    L.push("diagnostic failed: " + e);
  }
  try { console.log(L.join("\n")); } catch (_) {}
}


/**
 * AT RISK cell → a number, surviving the stale-DATE-format trap.
 *
 * ⚠ WHY THIS EXISTS RATHER THAN A BARE parseFloat (2026-08-21). If the column
 * carries a date number format — which it did for two days, see the note in
 * setupKitHealthSheet — Sheets hands getValues() a Date OBJECT instead of the
 * number underneath. parseFloat(Date) is NaN, so the row scored ZERO and the
 * oversell total silently read "0 units" against 30 unbuildable kits.
 *
 * The underlying value is still a serial day count, so it is recoverable:
 * Sheets' epoch is 1899-12-30, and building the epoch with the SAME local-time
 * constructor cancels the timezone offset on both sides (Math.round absorbs any
 * historical LMT/DST remainder — these are small integers, not timestamps).
 *
 * Belt-and-braces on top of the format fix: the format is the real repair, this
 * makes a REGRESSION of it visible-but-correct instead of silently zero. Blank
 * stays blank — a healthy kit legitimately has nothing at risk.
 */
function _kitRiskToNumber(v) {
  if (v === "" || v === null || v === undefined) return NaN;
  if (typeof v === "number") return v;
  if (v instanceof Date) {
    var epoch = new Date(1899, 11, 30).getTime();
    return Math.round((v.getTime() - epoch) / 86400000);
  }
  return parseFloat(v);
}


function getKitOversellSnapshot() {
  var out = { kits: 0, units: 0, unreadable: 0 };
  try {
    var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(KIT_HEALTH.sheetName);
    if (!sheet) return out;
    var lastRow = sheet.getLastRow();
    if (lastRow < KIT_HEALTH.dataStartRow) return out;

    var want = String(KIT_HEALTH.stock.CANT_BUILD).trim().toUpperCase();
    var vals = sheet.getRange(
      KIT_HEALTH.dataStartRow, KIT_HEALTH.cols.STOCK_STATUS,
      lastRow - KIT_HEALTH.dataStartRow + 1, 2      // P..Q, one read
    ).getValues();

    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).trim().toUpperCase() !== want) continue;
      out.kits++;
      var n = _kitRiskToNumber(vals[i][1]);
      if (!isNaN(n) && n > 0) out.units += n;
      else out.unreadable++;
    }
    out.units = Math.round(out.units);
    return out;
  } catch (e) {
    console.log("getKitOversellSnapshot error: " + e);
    return { kits: 0, units: 0, unreadable: 0 };
  }
}


// =======================================================================================
// WEEKLY AUTO-AUDIT — Monday ~4am, off-hours by design
// =======================================================================================
//
// Same rationale as the eBay Price Audit's weekly run: runKitHealthAudit REWRITES
// and re-sorts the whole sheet, so it fires off-hours (before the 6am work gate)
// so it can never yank the sheet out from under a mid-review selection. 4am
// (vs the eBay audit's 5am) just staggers them; they write different sheets so
// they wouldn't actually collide. The manual sidebar button is unaffected.
//
// Editor-bound (time trigger runs head code) → deploy = `clasp push` + one run of
// setupKitHealthTrigger(). NO New Version.
// =======================================================================================

function runWeeklyKitHealthAudit() {
  var res = runKitHealthAudit();
  try { console.log("runWeeklyKitHealthAudit: " + (res && res.message ? res.message : "(no message)")); } catch (_) {}
  return res;
}

/** EDITOR-RUN once: install the weekly (Monday ~4am) Kit Health trigger.
 *  Idempotent — clears any existing runWeeklyKitHealthAudit trigger first. */
function setupKitHealthTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runWeeklyKitHealthAudit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runWeeklyKitHealthAudit')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(4).create();
  return "✅ Weekly Kit Health trigger installed (Mondays ~4am).";
}

/** EDITOR-RUN: remove the weekly Kit Health trigger (manual button unaffected). */
function removeKitHealthTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runWeeklyKitHealthAudit') { ScriptApp.deleteTrigger(t); removed++; }
  });
  return "✅ Removed " + removed + " weekly Kit Health trigger(s).";
}


/** EDITOR-RUN test: run the audit and log the summary. */
function testKitHealthAudit() {
  var r = runKitHealthAudit();
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}
