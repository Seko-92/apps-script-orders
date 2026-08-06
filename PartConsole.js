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

/**
 * LEAN Master-Inventory maps — same shape as buildLocationAndInventoryMaps(),
 * a fraction of the cost.
 *
 * WHY THIS EXISTS (2026-08-06): the shared helper does
 * `getDataRange().getValues()` on Master Inventory — ~174 columns x ~3,500 rows,
 * about 600,000 cells — and the dossier needs exactly FOUR of those columns.
 * On the sidebar that was merely slow; once the Floor Board could tap a row it
 * became the difference between a drawer that paints and one that times out.
 * (Same class of mistake that OOM-killed the n8n container in May by reading
 * all of MI through the Sheets node.)
 *
 * Apps Script cannot fetch non-contiguous columns in one call, so this is four
 * single-column reads — ~14,000 cells instead of ~600,000. Four round trips beat
 * one enormous payload by a wide margin here.
 *
 * @returns {{locationMap: Map, inventoryMap: Map}} keyed by lowercased SKU
 */

/**
 * ONE Master-Inventory pass that serves the whole dossier.
 *
 * MEASURED, 2026-08-06 — this replaced two separate readers that between them
 * cost ~6.6s of a 7.9s lookup:
 *     mi  = 2170ms   (_partMasterInfo scanned the SKU column, then re-read a row)
 *     inv = 4417ms   (_pcLeanInventory read FOUR single columns)
 *
 * THE LESSON: in Apps Script the ROUND TRIP dominates, not the payload. Four
 * single-column reads of 3,500 rows cost ~1.1s EACH. Replacing one giant
 * getDataRange() with four narrow reads barely helped; collapsing them into ONE
 * contiguous read is what actually wins.
 *
 * So: read the header once, work out the SPAN that covers the columns the maps
 * need, and pull that block in a single call. The target SKU's rich fields
 * (title, images, listing URL) come from a second read of ONE row, which is
 * trivially cheap at any width.
 *
 * ⚠ The span guard matters. These columns are clustered in the eBay export
 * today, but MI has ~193 columns and nothing stops one drifting to the far end.
 * If the span ever gets wide the block read would be worse than what it
 * replaced, so it falls back to per-column reads and logs the span.
 *
 * @returns {{locationMap:Map, inventoryMap:Map, priceMap:Map, row:Object|null,
 *            headers:Array, rowIndex:number}}
 */

/**
 * The same shape _partMasterInfo returns, built from the row _pcMasterSnapshot
 * ALREADY read — so MI is not scanned a second time. Kept byte-compatible with
 * the old return so every downstream reader is untouched.
 */
function _partMasterInfoFromRow(snap, normSku) {
  var out = { found: false, title: "", quantity: null, sold: null, miAvailable: null,
              currentPrice: null, startPrice: null, location: "", listingStatus: "",
              viewItemURL: "", images: [] };
  try {
    if (!normSku || !snap || !snap.row || !snap.headers || !snap.headers.length) return out;
    var headers = snap.headers, row = snap.row;
    function col(name) {
      var t = String(name).toLowerCase();
      for (var i = 0; i < headers.length; i++) {
        if (String(headers[i] || "").trim().toLowerCase() === t) return i;
      }
      return -1;
    }
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

    out.found         = true;
    out.title         = String(g(DB_TITLE_HEADER) || "");
    out.quantity      = qty;
    out.sold          = sold;
    out.miAvailable   = qty - sold;
    out.currentPrice  = (!isNaN(cur)  && cur  > 0) ? cur  : null;
    out.startPrice    = (!isNaN(strt) && strt > 0) ? strt : null;
    out.location      = String(g(DB_LOCATION_HEADER) || "");
    out.listingStatus = String(g(DB_LISTING_STATUS_HEADER) || "");
    out.viewItemURL   = String(g(DB_VIEWURL_HEADER) || "");
    out.images        = images;
  } catch (e) {
    try { console.log("_partMasterInfoFromRow: " + e); } catch (_) {}
  }
  return out;
}

function _pcMasterSnapshot(normSku) {
  var out = { locationMap: new Map(), inventoryMap: new Map(), priceMap: new Map(),
              row: null, headers: [], rowIndex: -1 };
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DB_SHEET_NAME);
    if (!sheet) return out;
    var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 2) return out;

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    out.headers = headers;
    function col(name) {
      var t = String(name).toLowerCase();
      for (var i = 0; i < headers.length; i++) {
        if (String(headers[i] || "").trim().toLowerCase() === t) return i + 1;
      }
      return -1;
    }
    var want = {
      sku:   col(DB_SKU_HEADER),
      qty:   col(DB_QUANTITY_HEADER),
      sold:  col(DB_QUANTITY_SOLD_HEADER),
      loc:   col(DB_LOCATION_HEADER),
      cur:   col('currentPrice'),
      strt:  col('startPrice')
    };
    if (want.sku < 1) return out;

    var present = [];
    for (var k in want) if (want[k] > 0) present.push(want[k]);
    var lo = Math.min.apply(null, present), hi = Math.max.apply(null, present);
    var span = hi - lo + 1;
    var n = lastRow - 1;

    var block = null, base = lo;
    if (span <= 60) {
      block = sheet.getRange(2, lo, n, span).getValues();       // ONE read
    } else {
      // Columns drifted apart — fall back rather than pull a huge block.
      try { console.log('_pcMasterSnapshot: span ' + span + ' too wide, per-column fallback'); } catch (e) {}
      block = [];
      var cols = {}, order = [];
      for (var kk in want) if (want[kk] > 0) { order.push(want[kk]); }
      var byCol = {};
      for (var oi = 0; oi < order.length; oi++) {
        byCol[order[oi]] = sheet.getRange(2, order[oi], n, 1).getValues();
      }
      for (var r = 0; r < n; r++) {
        var rowArr = [];
        for (var ci = lo; ci <= hi; ci++) rowArr.push(byCol[ci] ? byCol[ci][r][0] : "");
        block.push(rowArr);
      }
      base = lo;
    }
    function at(rowArr, c) { return (c > 0) ? rowArr[c - base] : ""; }

    for (var i = 0; i < block.length; i++) {
      var raw = String(at(block[i], want.sku) || "").trim();
      if (!raw) continue;
      var key = raw.toLowerCase();
      if (out.rowIndex < 0 && _normPartSku(raw) === normSku) out.rowIndex = 2 + i;

      if (want.loc > 0) out.locationMap.set(key, String(at(block[i], want.loc) || "").trim() || "NOT FOUND");
      if (want.qty > 0 && want.sold > 0) {
        var q = parseInt(at(block[i], want.qty), 10) || 0;
        var sd = parseInt(at(block[i], want.sold), 10) || 0;
        out.inventoryMap.set(key, { quantity: q, sold: sd, available: q - sd, status: "" });
      }
      // price map in the SAME pass — this is what lets the kit-price call skip
      // rebuilding its own maps (another full MI read plus another Zoho read).
      var pc = parseFloat(at(block[i], want.cur));
      var ps = parseFloat(at(block[i], want.strt));
      var price = (!isNaN(pc) && pc > 0) ? pc : ((!isNaN(ps) && ps > 0) ? ps : null);
      if (price != null) out.priceMap.set(key, price);
    }

    // the target row's rich fields — one row, cheap at any width
    if (out.rowIndex > 0) {
      out.row = sheet.getRange(out.rowIndex, 1, 1, lastCol).getValues()[0];
    }
  } catch (e) {
    try { console.log("_pcMasterSnapshot: " + e); } catch (_) {}
  }
  return out;
}

function _pcLeanInventory() {
  var out = { locationMap: new Map(), inventoryMap: new Map() };
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DB_SHEET_NAME);
    if (!sheet) return out;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return out;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    function col(name) {
      var t = String(name).toLowerCase();
      for (var i = 0; i < headers.length; i++) {
        if (String(headers[i] || "").trim().toLowerCase() === t) return i + 1;
      }
      return -1;
    }
    var cSku = col(DB_SKU_HEADER), cQty = col(DB_QUANTITY_HEADER),
        cSold = col(DB_QUANTITY_SOLD_HEADER), cLoc = col(DB_LOCATION_HEADER);
    if (cSku < 1) return out;

    var n = lastRow - 1;
    function grab(c) { return c > 0 ? sheet.getRange(2, c, n, 1).getValues() : null; }
    var skus = grab(cSku), qtys = grab(cQty), solds = grab(cSold), locs = grab(cLoc);

    for (var r = 0; r < n; r++) {
      var sku = String(skus[r][0] || "").trim().toLowerCase();
      if (!sku) continue;
      if (locs) out.locationMap.set(sku, String(locs[r][0] || "").trim() || "NOT FOUND");
      if (qtys && solds) {
        var q = parseInt(qtys[r][0], 10) || 0, sd = parseInt(solds[r][0], 10) || 0;
        out.inventoryMap.set(sku, { quantity: q, sold: sd, available: q - sd, status: "" });
      }
    }
  } catch (e) {
    try { console.log("_pcLeanInventory: " + e); } catch (_) {}
  }
  return out;
}

function _buildPartDossier(raw) {
  var rawTrim = String(raw == null ? "" : raw).trim();
  var normSku = _normPartSku(rawTrim);

  // --- shared maps (built once) ---
  // TIMED. Every stage here is a separate sheet read, and which one dominates
  // is not guessable — measure, then optimise the real one. The summary lands
  // in the Apps Script execution log as "partDossier <sku>: mi=… zoho=… …".
  var _t0 = Date.now(), _tPrev = _t0, _lap = {};
  function _t(name) { var now = Date.now(); _lap[name] = now - _tPrev; _tPrev = now; }

  // ONE Master Inventory pass now serves BOTH the per-SKU detail and the
  // all-SKU maps. Previously _partMasterInfo and _pcLeanInventory each read MI
  // separately — measured at 2170ms + 4417ms of a 7946ms lookup.
  var snap = _pcMasterSnapshot(normSku);                                                 _t('snap');
  var mi = _partMasterInfoFromRow(snap, normSku);                                        _t('mi');
  var zohoMap; try { zohoMap = buildZohoStockMap() || new Map(); } catch (e) { zohoMap = new Map(); }   _t('zoho');
  var invMaps = { locationMap: snap.locationMap, inventoryMap: snap.inventoryMap };
  var resolveAvail = _oosResolveAvailFactory(zohoMap, invMaps.inventoryMap);
  var kitMap; try { kitMap = buildKitMap(); } catch (e) { kitMap = new Map(); }           _t('kitMap');
  var ripple = _invertKitMap(kitMap);                                                     _t('invert');
  var healthMap = _readKitHealthStatusMap();                                              _t('health');
  var committed; try { committed = getCommittedQuantities(); } catch (e) { committed = new Map(); }     _t('committed');

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
    // Measured at 4985ms because it rebuilt the component price maps (another
    // full MI read + another Zoho read) that we already hold. Hand them over.
    // _buildKitComponentPriceMaps returns { ebay, zoho } — same shape as this.
    var priced = null;
    try {
      priced = computeKitPriceBySku(kitObj.sku, {
        maps: { ebay: snap.priceMap, zoho: zohoMap }
      });
    } catch (e) { priced = null; }
    _t('kitPrice');
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
  // Index once. This used to scan the WHOLE kit map for every ripple entry —
  // a part used in 11 kits meant ~2,500 iterations. Same key semantics as the
  // old `ks === entry.kitSku` comparison, just O(1).
  var kitBySku = new Map();
  kitMap.forEach(function (kk, ks) { kitBySku.set(String(ks), kk); });

  (ripple.get(normSku) || []).forEach(function (entry) {
    var k = kitBySku.get(String(entry.kitSku)) || null;
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

  _t('ripple');
  try {
    var _parts = [];
    for (var _k in _lap) if (_lap[_k] > 0) _parts.push(_k + '=' + _lap[_k]);
    console.log('partDossier ' + rawTrim + ': ' + _parts.join(' ') +
                '  TOTAL=' + (Date.now() - _t0) + 'ms');
  } catch (e) {}

  return {
    sku:      rawTrim,
    found:    found,
    _ms:      (Date.now() - _t0),
    _lap:     _lap,
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

    // ⚠ DELIBERATELY NOT CACHED. This is a DECISION surface — a picker reads
    // "on hand 3" and walks to a shelf — and the underlying numbers move
    // constantly: the Zoho mirror rewrites every 2 minutes, committed qty
    // changes with every arrival and every status flip. A cache here can only
    // ever hide a change that just landed, and the cost of that is a wasted
    // walk or a short pick.
    //
    // This is the SAME rule the hq-kits proxy already follows and for the same
    // reason. The board's 15s tick cache is not a counter-example: that exists
    // because every screen polls every 20s and would otherwise exhaust the
    // ~90 min/day Apps Script quota. On-demand taps have no such argument, so
    // there is nothing to trade the staleness against.
    //
    // Speed is a READ problem, not a caching problem — fix the reads.
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


// =======================================================================================
// FAST PATH — the identity half of the dossier
// =======================================================================================

/**
 * Everything needed to answer "is this the right part, and how many are there?"
 * — and NOTHING else.
 *
 * WHY THIS EXISTS (2026-08-06). The full dossier is a steady ~4s, but on the
 * board it arrived in 10-45s and varied wildly: Apps Script cold start, plus
 * execution serialisation with the board's own poll, plus the /exec redirect.
 * The fix is not to remove the expensive parts — the ripple is the best thing
 * in that drawer — it is to STOP MAKING THE PICKER WAIT FOR ALL OF IT AT ONCE.
 *
 * So the drawer loads in three stages: what the browser already knows (instant),
 * this (about a second), then the full dossier in the background (~3s). Nothing
 * is lost; it just stops arriving as one lump.
 *
 * Deliberately NOT cached — same reason as getPartData: this is what a picker
 * reads before walking to a shelf.
 *
 * Costs TWO reads: one MI row, one Zoho row. No kit map, no ripple, no
 * committed scan, no all-SKU availability map.
 *
 * @returns {{ok:boolean, basics:Object}|{ok:false, reason:string}}
 */
function getPartBasics(query) {
  try {
    var raw = String(query == null ? "" : query).trim();
    if (!raw) return { ok: false, reason: "Type a SKU." };
    var normSku = _normPartSku(raw);

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DB_SHEET_NAME);
    if (!sheet) return { ok: false, reason: "Master Inventory not found." };

    var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { ok: false, reason: "Master Inventory is empty." };

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    function col(name) {
      var t = String(name).toLowerCase();
      for (var i = 0; i < headers.length; i++) {
        if (String(headers[i] || "").trim().toLowerCase() === t) return i;
      }
      return -1;
    }
    var skuIdx = col(DB_SKU_HEADER);
    if (skuIdx < 0) return { ok: false, reason: "No SKU column in Master Inventory." };

    // One narrow scan to locate the row, then one row. Nothing else.
    var skus = sheet.getRange(2, skuIdx + 1, lastRow - 1, 1).getValues();
    var rowNum = -1;
    for (var r = 0; r < skus.length; r++) {
      if (_normPartSku(skus[r][0]) === normSku) { rowNum = 2 + r; break; }
    }

    var b = {
      sku: raw, found: false, title: "", location: "NOT FOUND", images: [],
      miAvailable: null, zohoAvailable: null, available: null,
      ebayPrice: null, zohoPrice: null, listingStatus: "", ebayUrl: ""
    };

    if (rowNum > 0) {
      var row = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
      function g(name) { var c = col(name); return c >= 0 ? row[c] : ""; }
      var qty  = parseFloat(g(DB_QUANTITY_HEADER)) || 0;
      var sold = parseFloat(g(DB_QUANTITY_SOLD_HEADER)) || 0;
      var cur  = parseFloat(g('currentPrice'));
      var strt = parseFloat(g('startPrice'));
      for (var p = 1; p <= 5; p++) {
        var u = String(g('pictureUrl' + p) || "").trim();
        if (u) b.images.push(u);
      }
      b.found         = true;
      b.title         = String(g(DB_TITLE_HEADER) || "");
      b.miAvailable   = qty - sold;
      b.ebayPrice     = (!isNaN(cur) && cur > 0) ? cur : ((!isNaN(strt) && strt > 0) ? strt : null);
      b.listingStatus = String(g(DB_LISTING_STATUS_HEADER) || "");
      b.location      = String(g(DB_LOCATION_HEADER) || "").trim() || "NOT FOUND";
      b.ebayUrl       = String(g(DB_VIEWURL_HEADER) || "").trim();
    }

    // Zoho is the stock master, so it wins when it knows the SKU.
    var z = null;
    try { z = getSingleZohoStock(normSku); } catch (e) { z = null; }
    if (z) {
      b.zohoAvailable = z.available;
      b.zohoPrice     = z.sellingPrice;
      if (!b.title && z.itemName) b.title = z.itemName;
      b.found = true;
    }
    b.available = (b.zohoAvailable != null) ? b.zohoAvailable : b.miAvailable;
    if (!b.ebayUrl) b.ebayUrl = "https://www.ebay.com/sch/i.html?_nkw=" + encodeURIComponent(raw);

    return { ok: true, basics: b };

  } catch (err) {
    try { console.log("getPartBasics: " + err); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}
