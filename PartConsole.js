// =======================================================================================
// PartConsole.js — the SKU / Part console: one window that knows everything about a part
// Shipped 2026-07-29.
// =======================================================================================
//
// WHY
//   Every other surface is order-, kit-, or audit-centric. But the atomic unit
//   under all of them is the PART, and a part flows through every system: stock
//   in Master Inventory, price in Zoho, a component in the Kit Registry, health
//   in Kit Health, an OOS row, demand in open orders. To understand one part you
//   used to open five sheets. This console is the JOIN — search once, see it
//   everywhere. Pure Tier 1: no new data, no new dependency, just a lens.
//
// THE KILLER PANEL — the ripple ("used in N kits")
//   Inverts buildKitMap so a component shows the kits that USE it, each with its
//   live BUILDABLE count and whether THIS part is the one BLOCKING it. Payoff
//   line: "restocking this part unblocks kits X, Y." That's the daily piston-
//   restock workflow (the Telegram channel) turned into a single search.
//
// WORKS FOR BOTH
//   - A COMPONENT / single item → stock · price · location · OOS · demand ·
//     the ripple · images.
//   - A KIT (SKU is in the registry) → its composition + buildable + computed-vs-
//     listed price + images. (A sub-assembly that's both shows both.)
//
// IMAGES — Master Inventory already stores pictureUrl1..5 per item; the modal
//   renders them (main + thumbnails). No new fetch.
//
// TIES TOGETHER (all existing): Master Inventory + Zoho Stock (stock/price) ·
//   Kit Registry inverted (which kits) · the Kit Health / OOS engines
//   (_oosComputeKitBuild, computeKitPrice) · All Orders (committed demand) ·
//   the Kit Health sheet (price status join) · MI viewItemURL (listing link).
//
// DEPLOY: editor-bound (sidebar google.script.run + showModalDialog).
//   `clasp push` is the whole deploy; nothing on /exec, no New Version.
// =======================================================================================


// =======================================================================================
// SKU NORMALIZATION — MI stores some SKUs as floats ("161361.0"); normalize both sides.
// =======================================================================================
function _normPartSku(raw) {
  if (raw === null || raw === undefined) return "";
  var s;
  if (typeof raw === 'number') s = String(Math.trunc(raw));
  else {
    s = String(raw).trim();
    if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, "");
  }
  return s.trim().toLowerCase();
}


// =======================================================================================
// MASTER INVENTORY — one part's row (stock, price, location, status, listing, IMAGES)
// =======================================================================================
//
// Two targeted reads: the SKU column to locate the row, then that ONE row across
// its columns. Columns found by HEADER NAME (same robustness as the audit maps),
// so a column re-order can't break it.
// =======================================================================================
function _partMasterInfo(normSku) {
  var out = { found: false, images: [] };
  try {
    if (!normSku) return out;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(DB_SHEET_NAME);
    if (!sheet) return out;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return out;
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    function col(name) {
      var t = String(name).toLowerCase();
      for (var i = 0; i < headers.length; i++) {
        if (String(headers[i] || "").trim().toLowerCase() === t) return i;
      }
      return -1;
    }
    var skuIdx = col(DB_SKU_HEADER);
    if (skuIdx < 0) return out;

    var skus = sheet.getRange(2, skuIdx + 1, lastRow - 1, 1).getValues();
    var rowNum = -1;
    for (var r = 0; r < skus.length; r++) {
      if (_normPartSku(skus[r][0]) === normSku) { rowNum = 2 + r; break; }
    }
    if (rowNum < 0) return out;

    var row = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    function g(name) { var c = col(name); return c >= 0 ? row[c] : ""; }

    var qty  = parseFloat(g(DB_QUANTITY_HEADER)) || 0;
    var sold = parseFloat(g(DB_QUANTITY_SOLD_HEADER)) || 0;
    var cur  = parseFloat(g('currentPrice'));
    var strt = parseFloat(g('startPrice'));
    var images = [];
    for (var p = 1; p <= 5; p++) {
      var u = String(g('pictureUrl' + p) || "").trim();
      if (u) images.push(u);
    }

    out.found        = true;
    out.title        = String(g(DB_TITLE_HEADER) || "");
    out.quantity     = qty;
    out.sold         = sold;
    out.miAvailable  = qty - sold;
    out.currentPrice = (!isNaN(cur)  && cur  > 0) ? cur  : null;
    out.startPrice   = (!isNaN(strt) && strt > 0) ? strt : null;
    out.location     = String(g(DB_LOCATION_HEADER) || "");
    out.listingStatus= String(g(DB_LISTING_STATUS_HEADER) || "");
    out.viewItemURL  = String(g(DB_VIEWURL_HEADER) || "");
    out.images       = images;
  } catch (e) {
    try { console.log("_partMasterInfo error: " + e); } catch (_) {}
  }
  return out;
}


// =======================================================================================
// RIPPLE — invert the kit map: component SKU → the kits that USE it
// =======================================================================================
function _invertKitMap(kitMap) {
  var index = new Map();   // compSkuLower → [{ kitSku, kitName, qtyPer, kitType }]
  kitMap.forEach(function (kit) {
    (kit.components || []).forEach(function (c) {
      var cs = _normPartSku(c.sku);
      if (!cs) return;
      if (!index.has(cs)) index.set(cs, []);
      index.get(cs).push({
        kitSku: kit.sku, kitName: kit.name || "",
        qtyPer: (c.qty > 0 ? c.qty : 1), kitType: kit.type || "MANUAL"
      });
    });
  });
  return index;
}

/** Best-effort join: kit SKU (normalized) → {priceStatus, buildable} from the
 *  last Kit Health run. Empty map if the sheet isn't there — the ripple still
 *  shows live buildability; price status just reads "—". */
function _readKitHealthStatusMap() {
  var map = new Map();
  try {
    if (typeof KIT_HEALTH === 'undefined') return map;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(KIT_HEALTH.sheetName);
    if (!sheet) return map;
    var lastRow = sheet.getLastRow();
    if (lastRow < KIT_HEALTH.dataStartRow) return map;
    var n = lastRow - KIT_HEALTH.dataStartRow + 1;
    var data = sheet.getRange(KIT_HEALTH.dataStartRow, 1, n, KIT_HEALTH.dataWidth).getValues();
    var SKU = KIT_HEALTH.idx("KIT_SKU"), PS = KIT_HEALTH.idx("PRICE_STATUS"), B = KIT_HEALTH.idx("BUILDABLE");
    for (var i = 0; i < data.length; i++) {
      var s = _normPartSku(data[i][SKU]);
      if (!s) continue;
      map.set(s, { priceStatus: String(data[i][PS] || ""), buildable: data[i][B] });
    }
  } catch (e) { /* best-effort */ }
  return map;
}


// =======================================================================================
// THE DOSSIER — build everything for one SKU
// =======================================================================================
function _buildPartDossier(raw) {
  var rawTrim = String(raw == null ? "" : raw).trim();
  var normSku = _normPartSku(rawTrim);

  // --- shared maps (built once) ---
  var mi = _partMasterInfo(normSku);
  var zohoMap; try { zohoMap = buildZohoStockMap() || new Map(); } catch (e) { zohoMap = new Map(); }
  var invMaps; try { invMaps = buildLocationAndInventoryMaps(); } catch (e) { invMaps = { locationMap: new Map(), inventoryMap: new Map() }; }
  var resolveAvail = _oosResolveAvailFactory(zohoMap, invMaps.inventoryMap);
  var kitMap; try { kitMap = buildKitMap(); } catch (e) { kitMap = new Map(); }
  var ripple = _invertKitMap(kitMap);
  var healthMap = _readKitHealthStatusMap();
  var committed; try { committed = getCommittedQuantities(); } catch (e) { committed = new Map(); }

  // --- the part's own stock/price ---
  var zRec = zohoMap.get(normSku) || null;
  var available = resolveAvail(normSku);   // Zoho-first → MI, or null
  var loc = (invMaps.locationMap.get(normSku) || mi.location || "").trim();
  if (!loc || loc.toUpperCase() === "NOT FOUND") loc = "NOT FOUND";

  var part = {
    title:        mi.title || (zRec ? zRec.itemName : "") || "",
    available:    (available != null) ? available : (mi.found ? mi.miAvailable : null),
    miAvailable:  mi.found ? mi.miAvailable : null,
    zohoAvailable: zRec ? zRec.available : null,
    quantity:     mi.found ? mi.quantity : null,
    sold:         mi.found ? mi.sold : null,
    ebayPrice:    mi.currentPrice != null ? mi.currentPrice : mi.startPrice,
    zohoPrice:    zRec ? zRec.sellingPrice : null,
    location:     loc,
    listingStatus: mi.listingStatus || "",
    viewItemURL:  mi.viewItemURL || "",
    images:       mi.images || [],
    committed:    committed.get(normSku) || 0
  };
  part.oos = (part.available != null) && (part.available <= 0);

  // eBay link: the real listing URL, else a search for the SKU (still one click)
  var ebayUrl = part.viewItemURL
    || ("https://www.ebay.com/sch/i.html?_nkw=" + encodeURIComponent(rawTrim));

  // --- is THIS SKU a kit? ---
  var kitObj = null;
  kitMap.forEach(function (k, kitSku) { if (!kitObj && _normPartSku(kitSku) === normSku) kitObj = k; });

  var kitView = null;
  if (kitObj) {
    var build = _oosComputeKitBuild(kitObj, resolveAvail);
    var comps = (kitObj.components || []).map(function (c) {
      var av = resolveAvail(_normPartSku(c.sku));
      return { sku: String(c.sku), qty: (c.qty > 0 ? c.qty : 1), name: c.name || "",
               available: (av != null) ? av : null };
    });
    // price (component-summed vs the kit's own listed price)
    var priced = null;
    try { priced = computeKitPriceBySku(kitObj.sku); } catch (e) { priced = null; }
    var hs = healthMap.get(normSku) || null;
    kitView = {
      name:       kitObj.name || "",
      type:       kitObj.type || "MANUAL",
      engine:     kitObj.engine || "",
      components: comps,
      buildable:  build.buildable,
      limitedBy:  build.limitedBy,
      unparsed:   (kitObj.unparsedLines || []).length,
      partsValue: priced ? priced.rawSum : null,
      computed:   priced ? priced.roundedTotal : null,
      listed:     priced ? priced.listedPrice : null,
      complete:   priced ? priced.complete : false,
      priceStatus: hs ? hs.priceStatus : ""    // from the last Kit Health run
    };
  }

  // --- ripple: kits that USE this SKU ---
  var usedIn = [];
  var unblock = [];
  (ripple.get(normSku) || []).forEach(function (entry) {
    var k = null;
    kitMap.forEach(function (kk, ks) { if (!k && ks === entry.kitSku) k = kk; });
    var b = k ? _oosComputeKitBuild(k, resolveAvail) : { buildable: "?", limiter: null };
    var blocked = (b.buildable === 0) && b.limiter && (_normPartSku(b.limiter.sku) === normSku);
    var hs2 = healthMap.get(_normPartSku(entry.kitSku)) || null;
    usedIn.push({
      kitSku:      entry.kitSku,
      kitName:     entry.kitName,
      qtyPer:      entry.qtyPer,
      buildable:   b.buildable,
      blocked:     !!blocked,
      priceStatus: hs2 ? hs2.priceStatus : ""
    });
    if (blocked) unblock.push(entry.kitSku);
  });
  usedIn.sort(function (a, b) {
    if (a.blocked !== b.blocked) return a.blocked ? -1 : 1;          // blocked first
    return String(a.kitSku).localeCompare(String(b.kitSku));
  });

  var found = mi.found || !!kitObj || usedIn.length > 0 || !!zRec;

  return {
    sku:      rawTrim,
    found:    found,
    part:     part,
    ebayUrl:  found ? ebayUrl : null,
    isKit:    !!kitObj,
    kit:      kitView,
    usedIn:   usedIn,
    unblock:  unblock,
    zohoSyncedAt: (function () { try { var d = getZohoStockSyncedAt(); return d ? d.getTime() : null; } catch (e) { return null; } })()
  };
}


// =======================================================================================
// PUBLIC — modal opener + in-window search
// =======================================================================================

/** In-window search (the console's search bar). Returns {ok, dossier}. */
function getPartData(query) {
  try {
    var raw = String(query == null ? "" : query).trim();
    if (!raw) return { ok: false, reason: "Type a SKU." };
    return { ok: true, dossier: _buildPartDossier(raw) };
  } catch (err) {
    try { console.log("getPartData error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}

/** Open the Part console. Always opens on a non-empty SKU (even not-found → empty
 *  state + a working search bar), same console UX as Order Case. */
function openPartConsole(query) {
  try {
    var raw = String(query == null ? "" : query).trim();
    if (!raw) return { ok: false, reason: "Type a SKU first." };

    var dossier = _buildPartDossier(raw);
    var template = HtmlService.createTemplateFromFile("PartConsoleModal");
    template.dossierJson = JSON.stringify(dossier).replace(/<\//g, "<\\/");   // Gotcha #2

    var html = template.evaluate().setWidth(1080).setHeight(720);
    SpreadsheetApp.getUi().showModalDialog(html, "🧩 Part Console");
    return { ok: true, found: dossier.found, isKit: dossier.isKit, usedIn: dossier.usedIn.length };
  } catch (err) {
    try { console.log("openPartConsole error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}

/** EDITOR-RUN test: dump a part dossier (no modal). */
function testPartConsole(sku) {
  var d = _buildPartDossier(sku || "");
  Logger.log(JSON.stringify(d, null, 2));
  return d;
}
