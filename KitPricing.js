// =======================================================================================
// KitPricing.js — component-priced kit valuation engine + calculator modal
// Shipped 2026-07-24.
// =======================================================================================
//
// PURPOSE
//   Automates the manual kit-pricing method the user has been doing in Excel:
//
//       kit price = round( 0.9 × Σ( component_qty × component_unit_price ) )
//
//   i.e. "what the parts would sell for individually, minus a 10% bundle
//   discount, rounded to the nearest dollar." Validated 2026-07-24 against the
//   user's own Excel examples — the engine reproduced their numbers to the cent
//   wherever component prices hadn't drifted, and every delta was explained by
//   either a real price move or a deliberate hand-override.
//
// TWO CONSUMERS (both sit on this one engine):
//   A. Kit Price Calculator (modal, this file) — price a NEW kit before listing,
//      or re-price an existing one. Pick an existing kit (components pulled from
//      the Kit Registry) OR paste a raw BOM (parsed via _parseKitPd). Editable
//      per-line overrides + editable discount, live client-side recompute.
//   B. Kit Price Audit (PARKED — phase 2) — run computeKitPrice over every
//      registry kit, compare computed vs the kit's current listed price, flag
//      drift. Reuses computeKitPrice + _buildKitComponentPriceMaps unchanged.
//
// PRICE SOURCE (agreed 2026-07-24): eBay is the authority.
//   component unit price = MI currentPrice → MI startPrice → Zoho sellingPrice.
//   A component with NO price anywhere is UNPRICED — the engine NEVER folds a
//   guess into the total; it flags loudly and marks the total incomplete (same
//   honesty rule as OOS BUILDABLE's "⚠ N unparsed" and the kit-expansion modal).
//
// DISCOUNT: global default 10% (editable per calculation). ROUNDING: nearest $1.
//
// DEPLOY: editor-bound (modal via showModalDialog + sidebar google.script.run).
//   `clasp push` is the whole deploy — nothing here touches /exec, no New Version.
// =======================================================================================


var KIT_PRICING = {
  discount:  0.10,   // global default bundle discount (10% off Σ component prices)
  roundTo:   1       // round the final kit price to the nearest whole dollar
};


// =======================================================================================
// PRICE MAPS — read MI (eBay) + Zoho Stock ONCE; reused across every component
// =======================================================================================
//
// eBay map: skuLower → currentPrice||startPrice (>0), across ALL MI rows (NOT
// active-filtered — a component's price is valid regardless of its listing
// state; the audit's active filter is a different concern). Zoho map: the
// standard buildZohoStockMap (skuLower → {sellingPrice, itemName, ...}).
//
// For a single-kit calculator call this is one MI read (~1s); the audit builds
// it ONCE and passes it into computeKitPrice per kit via opts.maps.
// =======================================================================================

function _buildKitComponentPriceMaps() {
  var ebay = new Map();
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(DB_SHEET_NAME);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var lastCol  = sheet.getLastColumn();
        var headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        var skuIdx = -1, curIdx = -1, startIdx = -1;
        for (var i = 0; i < headers.length; i++) {
          var h = String(headers[i] || "").trim().toLowerCase();
          if      (h === DB_SKU_HEADER.toLowerCase()) skuIdx   = i;
          else if (h === 'currentprice')              curIdx   = i;
          else if (h === 'startprice')                startIdx = i;
        }
        if (skuIdx >= 0) {
          var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
          for (var r = 0; r < data.length; r++) {
            var sku = String(data[r][skuIdx] || "").trim().toLowerCase();
            if (!sku) continue;
            var cur   = curIdx   >= 0 ? parseFloat(data[r][curIdx])   : NaN;
            var start = startIdx >= 0 ? parseFloat(data[r][startIdx]) : NaN;
            var price = (!isNaN(cur) && cur > 0)     ? cur
                      : (!isNaN(start) && start > 0) ? start
                      : NaN;
            if (!isNaN(price)) ebay.set(sku, price);
          }
        }
      }
    }
  } catch (e) {
    console.log("_buildKitComponentPriceMaps eBay error: " + e);
  }

  var zoho = new Map();
  try { zoho = buildZohoStockMap() || new Map(); }
  catch (e) { console.log("_buildKitComponentPriceMaps zoho error: " + e); }

  return { ebay: ebay, zoho: zoho };
}


/** eBay-first component price resolution. Returns {price, source}.
 *  source ∈ "eBay" | "Zoho" | "unpriced". */
function _resolveComponentPrice(sku, maps) {
  var s = String(sku || "").trim().toLowerCase();
  if (!s) return { price: null, source: "unpriced" };
  var e = maps.ebay.get(s);
  if (e != null && e > 0) return { price: e, source: "eBay" };
  var z = maps.zoho.get(s);
  if (z && z.sellingPrice > 0) return { price: z.sellingPrice, source: "Zoho" };
  return { price: null, source: "unpriced" };
}

function _kpRound2(x) { return Math.round(x * 100) / 100; }


// =======================================================================================
// THE ENGINE — computeKitPrice(components, opts)
// =======================================================================================
//
// components : [{sku, qty, name, bundled?}]  (from buildKitMap or parseBomToComponents)
// opts       : {
//                discount:  fraction (default KIT_PRICING.discount = 0.10),
//                overrides: { skuLower: unitPrice } — hand-set component prices,
//                include:   { skuLower: true|false } — hand-set IN/OUT, wins over
//                           the `bundled` flag in BOTH directions,
//                maps:      pre-built {ebay, zoho} (built here if omitted)
//              }
//
// Returns:
//   {
//     lines: [{ sku, name, qty, unitPrice, source, lineTotal, overridden, unpriced,
//               excluded, bundledInto }],
//     rawSum, discount, discountedTotal, roundedTotal,
//     unpricedComponents: [sku], excludedComponents: [sku], complete: bool
//   }
//
// complete=false when ANY *included* component is unpriced — the caller must treat
// roundedTotal as untrustworthy and surface the warning, never the number alone.
//
// ⚠ EXCLUDED COMPONENTS STILL APPEAR IN `lines`, flagged, never silently removed.
//   A part that vanished from the breakdown would be indistinguishable from one
//   the parser never saw, and this engine's whole contract is that an untrustworthy
//   number is announced rather than quietly produced. The UI dims them instead.
//
// ⚠ A BUNDLED COMPONENT CANNOT MAKE A KIT "INCOMPLETE". It isn't being sold as a
//   separate part, so its missing price says nothing about the kit's price. Before
//   this, an unpriced head gasket that ships inside the gasket set blanked the whole
//   kit's price math for no reason.
// =======================================================================================

function computeKitPrice(components, opts) {
  opts = opts || {};
  var maps      = opts.maps || _buildKitComponentPriceMaps();
  var discount  = (opts.discount != null && !isNaN(opts.discount)) ? opts.discount : KIT_PRICING.discount;
  var overrides = opts.overrides || {};
  var include   = opts.include   || {};

  var lines = [];
  var rawSum = 0;
  var unpriced = [];
  var excluded = [];

  (components || []).forEach(function (c) {
    var sku = String(c.sku || "").trim();
    var qty = parseInt(c.qty, 10) || 0;

    // IN or OUT. Bundled parts ship inside another component, so they are not
    // part of what the kit costs. An explicit include[] entry (the calculator's
    // per-line control) wins in either direction — the human at the bench knows
    // the rare exception the data cannot express.
    var ov  = include[sku.toLowerCase()];
    var isExcluded = (ov === true) ? false : (ov === false ? true : (c.bundled === true));

    var ovp = overrides[sku.toLowerCase()];
    var unit, source, overridden = false;
    if (ovp != null && !isNaN(parseFloat(ovp)) && parseFloat(ovp) > 0) {
      unit = parseFloat(ovp); source = "override"; overridden = true;
    } else {
      var res = _resolveComponentPrice(sku, maps);
      unit = res.price; source = res.source;
    }

    var isUnpriced = (unit == null || isNaN(unit));
    var line = isUnpriced ? null : _kpRound2(unit * qty);
    if (isExcluded)      excluded.push(sku);
    else if (isUnpriced) unpriced.push(sku);
    else                 rawSum += unit * qty;

    lines.push({
      sku: sku, name: c.name || "", qty: qty,
      unitPrice: isUnpriced ? null : _kpRound2(unit),
      source: source, lineTotal: line,
      overridden: overridden, unpriced: isUnpriced,
      excluded: isExcluded, bundledInto: c.bundledInto || ""
    });
  });

  var discounted = rawSum * (1 - discount);
  var rounded = Math.round(discounted / KIT_PRICING.roundTo) * KIT_PRICING.roundTo;

  return {
    lines: lines,
    rawSum: _kpRound2(rawSum),
    discount: discount,
    discountedTotal: _kpRound2(discounted),
    roundedTotal: rounded,
    unpricedComponents: unpriced,
    excludedComponents: excluded,
    complete: unpriced.length === 0
  };
}


/** Existing-kit convenience: resolve components from the Kit Registry, price
 *  them, and also resolve the kit's OWN current listed price (eBay-first) for
 *  the calculator's compare line + the future audit. */
function computeKitPriceBySku(kitSku, opts) {
  opts = opts || {};
  var maps = opts.maps || _buildKitComponentPriceMaps();
  var kit  = getKitInfo(kitSku);
  if (!kit) return { found: false, kitSku: String(kitSku || "").trim() };

  var result = computeKitPrice(kit.components, {
    maps: maps, discount: opts.discount, overrides: opts.overrides, include: opts.include
  });

  var listed = _resolveComponentPrice(kit.sku, maps);
  result.found        = true;
  result.kitSku       = kit.sku;
  result.kitName      = kit.name;
  result.engine       = kit.engine;
  result.kitType      = kit.type;
  result.unparsedLines = kit.unparsedLines || [];
  result.listedPrice  = listed.price;
  result.listedSource = listed.source;
  return result;
}


/** Parse a pasted raw BOM (Zoho Purchase Description text) into components,
 *  reusing the SAME parser the Kit Registry uses (_parseKitPd) so a pasted BOM
 *  behaves identically to a registered one — including the "⚠ could not read"
 *  lines, which we surface so the picker fixes the PD rather than mispricing. */
function parseBomToComponents(text) {
  var parsed = _parseKitPd(text);
  return {
    components: (parsed.components || []).map(function (c) {
      return { sku: c.sku, qty: c.qty, name: c.name };
    }),
    unparsed: parsed.unparsed || [],
    engineModel: parsed.engineModel || ""
  };
}


// =======================================================================================
// MODAL — Kit Price Calculator
// =======================================================================================

/** Sidebar entry point — opens the calculator modal. Passes a lightweight
 *  {sku,name} list of all registry kits for input autocomplete. */
function openKitPriceCalculator() {
  try {
    var kits = [];
    try {
      buildKitMap().forEach(function (k) { kits.push({ sku: k.sku, name: k.name }); });
      kits.sort(function (a, b) { return String(a.sku).localeCompare(String(b.sku)); });
    } catch (e) { console.log("openKitPriceCalculator kit list error: " + e); }

    var template = HtmlService.createTemplateFromFile("KitCalculatorModal");
    // <?!= ?> force-unescaped JSON injection (Gotcha #2) + </-guard.
    template.kitListJson    = JSON.stringify(kits).replace(/<\//g, "<\\/");
    template.defaultDiscount = KIT_PRICING.discount;

    var html = template.evaluate().setWidth(1000).setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(html, "Kit Price Calculator");
    return { ok: true, kits: kits.length };
  } catch (err) {
    try { console.log("openKitPriceCalculator error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}


/** Modal → server: price a kit (existing or pasted BOM). The modal handles
 *  override/discount edits client-side (pure arithmetic); this call is the
 *  price-fetch + authoritative first compute.
 *
 * payload = { mode:'kit'|'bom', kitSku, bomText, discountPct, overrides, include }
 */
function getKitPriceBreakdown(payload) {
  try {
    payload = payload || {};
    var discount = (payload.discountPct != null && !isNaN(parseFloat(payload.discountPct)))
      ? parseFloat(payload.discountPct) / 100
      : KIT_PRICING.discount;
    var overrides = payload.overrides || {};
    var include   = payload.include   || {};
    var maps = _buildKitComponentPriceMaps();

    if (payload.mode === 'bom') {
      var parsed = parseBomToComponents(payload.bomText || "");
      if (parsed.components.length === 0) {
        return { ok: false, reason: "No component lines could be read from that BOM. Each line should end with its 6-digit SKU." };
      }
      var res = computeKitPrice(parsed.components, { maps: maps, discount: discount, overrides: overrides, include: include });
      res.ok = true; res.mode = 'bom';
      res.engine = parsed.engineModel;
      res.unparsedLines = parsed.unparsed;
      res.listedPrice = null;
      return res;
    }

    var kitSku = String(payload.kitSku || "").trim();
    if (!kitSku) return { ok: false, reason: "Enter a kit SKU (or switch to Paste BOM)." };
    var r = computeKitPriceBySku(kitSku, { maps: maps, discount: discount, overrides: overrides, include: include });
    if (!r.found) return { ok: false, reason: "Kit " + kitSku + " isn't in the Kit Registry. Paste its BOM instead, or import it first." };
    r.ok = true; r.mode = 'kit';
    return r;
  } catch (err) {
    try { console.log("getKitPriceBreakdown error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}


/** EDITOR-RUN test: price a known kit and log the full breakdown. */
function testKitPriceCalc() {
  var r = computeKitPriceBySku("217475");
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}
