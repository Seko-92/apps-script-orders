// =======================================================================================
// ENGINE CATALOGUE — live stock for the /catalogue page (2026-10-03)
// =======================================================================================
// The catalogue itself (manufacturer parts lists + drawings) is static and lives on the
// VPS (catalogue/ in the repo builds it). This file answers ONE question when someone opens
// an engine: "which of these part numbers are on our shelves right now?"
//
// It reuses the Parts Finder's own pieces — _mpnLoadMi (one bounded MI read + Zoho-first
// stock, the house rule) and _mpnBuildIndex / _mpnKey (the match key: letters+digits,
// leading zeros stripped) — so the catalogue and the Parts Finder can never disagree.
//
// ⚠ READ-ONLY and NOT LOGGED: browsing an engine is not a customer asking for a part, so
// it must not flood the MPN Search Log (whose misses are the "what to list next" record).
// Lock-free in doPost (DOPOST_LOCK_FREE) for the same reason as boardFind.
// =======================================================================================

var CATALOGUE = {
  maxPns: 800,          // one engine is ~300-450; a request past this is not a catalogue page
  maxHitsPerPn: 8       // an overhaul-kit number can appear on dozens of listings
};

/**
 * @param {string[]} pns   manufacturer part numbers as printed (dashes and all)
 * @return {{ok:boolean, stock:Object<string,Array>, ms:number, reason?:string}}
 *         stock[pn] = active listings carrying that number, the part before kits
 */
function catalogueStock(pns) {
  var t0 = Date.now();
  try {
    if (!Array.isArray(pns) || !pns.length) return { ok: false, reason: "No part numbers sent." };
    pns = pns.slice(0, CATALOGUE.maxPns).map(function (p) { return String(p == null ? "" : p).slice(0, 40); });

    var mi = _mpnLoadMi();
    var index = _mpnBuildIndex(mi.r.rows, mi.mpnCols);
    var stock = {};
    var cache = {};                                  // row → described, so a row is built once
    pns.forEach(function (pn) {
      var key = _mpnKey(pn);
      if (!key || key.length < 3) { stock[pn] = []; return; }
      var hits = (index.get(key) || []).map(function (h) {
        if (!cache[h.row]) cache[h.row] = mi.describe(h.row);
        return cache[h.row];
      }).filter(function (m) { return m.active; });
      stock[pn] = _catalogueRank(hits).slice(0, CATALOGUE.maxHitsPerPn).map(function (m) {
        return { sku: m.sku, title: String(m.title).slice(0, 110), loc: m.location === "NOT FOUND" ? "" : m.location,
                 avail: m.available, price: m.price, kit: !!m.isKit, url: m.url };
      });
    });
    return { ok: true, stock: stock, ms: Date.now() - t0 };
  } catch (err) {
    try { console.log("catalogueStock: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}

/** The part itself before kits that contain it, then in-stock first, then SKU. Pure. */
function _catalogueRank(hits) {
  return hits.slice().sort(function (a, b) {
    return (a.isKit ? 1 : 0) - (b.isKit ? 1 : 0)
        || ((b.available > 0) ? 1 : 0) - ((a.available > 0) ? 1 : 0)
        || String(a.sku).localeCompare(String(b.sku));
  });
}
