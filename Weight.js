// =======================================================================================
// Weight.js — SHIPPING WEIGHT FOR A SET OF ITEMS (2026-10-02)
// =======================================================================================
//
// To price a label (eBay order bought off-eBay, or a Direct order) the shipping person needs
// the weight of everything the customer is buying. Until now the owner looked each SKU up,
// multiplied by qty, added it all and handed the number over. This does that sum.
//
// THREE DOORS, ONE CALCULATOR:
//   • Telegram  /weight SO-26018  ·  /weight 24-15008-33107  ·  /weight 166500 x2, 173817
//   • Parts Finder quote basket — a weight line + "Copy for shipping"
//   • Order Case window — the order's total weight
//
// WHERE A WEIGHT COMES FROM, best first — and every line says which:
//   1. By Part Type's MEASURED weight, via the Product Data Health sheet (one fast read).
//      ⚠ Not MI first: eBay is LIGHTER than the measured weight on 214 SKUs (2026-09-12),
//      and an understated weight is exactly what gets a label re-billed by the carrier.
//   2. The eBay listing weight (Product Data Health's eBay column, else Master Inventory).
//   3. Nothing → the line is LISTED as having no weight and left out of the total, which then
//      says so. Never counted as 0 — a total that quietly drops a part is worse than none.
//
// PARTS ONLY. No packing allowance — the box is the shipping person's call (owner, 2026-10-02).
// =======================================================================================

var WEIGHT = {
  maxItems: 80,          // a typed list longer than this is almost certainly a paste gone wrong
  maxQty:   999
};

/** A SKU as a match key: case-blind, leading zeros dropped from all-digit SKUs. Pure. */
function _wtKey(s) {
  var k = String(s == null ? "" : s).trim().toUpperCase().replace(/\.0+$/, "");
  return /^\d+$/.test(k) ? k.replace(/^0+(?=\d)/, "") : k;
}

/** A positive number of ounces, or null. Pure. */
function _wtOz(v) {
  var n = parseFloat(v);
  return (isFinite(n) && n > 0) ? n : null;
}

/** "12 x 8 x 6" / "12x8x6" → [12,8,6] or null. Pure. */
function _wtDims(v) {
  if (Array.isArray(v)) v = v.join("x");
  var m = String(v == null ? "" : v).match(/(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)/i);
  if (!m) return null;
  var d = [parseFloat(m[1]), parseFloat(m[2]), parseFloat(m[3])];
  return (d[0] > 0 && d[1] > 0 && d[2] > 0) ? d : null;
}

/** 151 oz → "9 lb 7 oz"; 6 → "6 oz". Pure. */
function _wtFmt(oz) {
  if (oz == null) return "—";
  var t = Math.round(oz * 10) / 10;
  var lb = Math.floor(t / 16), rest = Math.round((t - lb * 16) * 10) / 10;
  if (rest === 16) { lb++; rest = 0; }
  if (!lb) return rest + " oz";
  return lb + " lb" + (rest ? " " + rest + " oz" : "");
}

/** Does this read like an order id rather than a list of SKUs? Pure. */
function _wtLooksLikeOrder(t) {
  t = String(t || "").trim();
  return /^(SO|INV)[-_\s]?\d+$/i.test(t) || /^\d{2,3}-\d{4,6}-\d{4,6}$/.test(t) || /^AMZ-\S+$/i.test(t);
}

/**
 * "166500 x2, 173817, 3x171432, 166527 4" → [{sku, qty}], plus anything unreadable. Pure.
 * Quantity forms: "x2" / "×2" / "*2" after the SKU (glued or spaced), "2x" before it, or a
 * bare small number right after a SKU. A SKU is 4+ characters with a digit (SUP-101 too);
 * a repeated SKU adds up.
 */
function _wtParseList(text) {
  var src = String(text == null ? "" : text)
    .replace(/(^|[\s,;])(\d{1,3})\s*[x×*]\s+(?=\S)/gi, "$1$2x ")   // "2 x 166500" → "2x 166500" (a QTY, never a SKU)
    .replace(/(\S)\s+[x×*]\s*(\d{1,3})\b/gi, "$1 x$2");  // "166500 x 2" → "166500 x2"
  var toks = src.split(/[\s,;]+/).filter(Boolean);
  var out = [], bad = [], pending = null, last = null, lastQtySet = false;
  var isSku = function (t) { return t.length >= 4 && /\d/.test(t) && /^[A-Za-z0-9][A-Za-z0-9._\/-]*$/.test(t); };
  toks.forEach(function (raw) {
    var t = raw.trim(), m;
    if ((m = t.match(/^(\d{1,3})[x×*]$/i)))      { pending = +m[1]; return; }
    if ((m = t.match(/^[x×*](\d{1,3})$/i)))      { if (last) { last.qty = +m[1]; lastQtySet = true; } return; }
    if ((m = t.match(/^(\d{1,3})[x×*](.+)$/i)) && isSku(m[2])) { pending = +m[1]; t = m[2]; }
    else if ((m = t.match(/^(.+?)[x×*](\d{1,3})$/i)) && isSku(m[1])) {
      last = { sku: m[1], qty: +m[2] }; out.push(last); lastQtySet = true; pending = null; return;
    }
    if (/^\d{1,3}$/.test(t) && !isSku(t)) {
      if (last && !lastQtySet) { last.qty = +t; lastQtySet = true; } else pending = +t;
      return;
    }
    if (isSku(t)) {
      last = { sku: t, qty: pending || 1 }; out.push(last);
      lastQtySet = !!pending; pending = null;
      return;
    }
    bad.push(t);
  });
  // fold repeats
  var byKey = {}, items = [];
  out.forEach(function (it) {
    var k = _wtKey(it.sku), q = Math.min(WEIGHT.maxQty, Math.max(1, it.qty || 1));
    if (byKey[k]) byKey[k].qty += q;
    else { byKey[k] = { sku: String(it.sku).trim(), qty: q }; items.push(byKey[k]); }
  });
  return { items: items, bad: bad };
}

/**
 * Weights for the given SKUs: {key: {oz, src, dims, title}}. Product Data Health first
 * (measured, then its eBay column), Master Inventory only for what is still missing.
 */
function _wtWeightMap(skus) {
  var want = {}, map = {};
  (skus || []).forEach(function (s) { want[_wtKey(s)] = true; });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  try {
    var ph = ss.getSheetByName(PRODUCT_HEALTH.sheetName);
    var last = ph ? ph.getLastRow() : 0;
    if (last >= PRODUCT_HEALTH.dataStartRow) {
      var C = PRODUCT_HEALTH.cols;
      var data = ph.getRange(PRODUCT_HEALTH.dataStartRow, 1, last - PRODUCT_HEALTH.dataStartRow + 1, C.EBAY_DIMS).getValues();
      data.forEach(function (r) {
        var k = _wtKey(r[C.SKU - 1]);
        if (!want[k] || map[k]) return;
        var truth = _wtOz(r[C.TRUTH_WT - 1]), ebay = _wtOz(r[C.EBAY_WT - 1]);
        var tD = _wtDims(r[C.TRUTH_DIMS - 1]), eD = _wtDims(r[C.EBAY_DIMS - 1]);
        if (truth === null && ebay === null) return;      // let MI have a try
        map[k] = { oz: truth !== null ? truth : ebay, src: truth !== null ? "measured" : "eBay",
                   dims: (truth !== null ? tD : eD) || tD || eD, title: String(r[C.TITLE - 1] || "") };
      });
    }
  } catch (e) { try { console.log("_wtWeightMap: Product Data Health read failed: " + e); } catch (_) {} }

  var missing = Object.keys(want).filter(function (k) { return !map[k]; });
  if (missing.length) {
    try {
      var mi = ss.getSheetByName(DB_SHEET_NAME);
      var OPT = ["packageWeightOz", "packageLengthIn", "packageWidthIn", "packageDepthIn", DB_TITLE_HEADER];
      var r = MiSchema.readColumns(mi, [DB_SKU_HEADER], { optional: OPT });
      var I = r.idx, need = {};
      missing.forEach(function (k) { need[k] = true; });
      r.rows.forEach(function (row) {
        var k = _wtKey(row[I[DB_SKU_HEADER]]);
        if (!need[k]) return;
        var oz = I.packageWeightOz >= 0 ? _wtOz(row[I.packageWeightOz]) : null;
        var t  = I[DB_TITLE_HEADER] >= 0 ? String(row[I[DB_TITLE_HEADER]] || "") : "";
        var prev = map[k];
        if (oz === null) { if (!prev) map[k] = { oz: null, src: "", dims: null, title: t }; return; }
        if (prev && prev.oz !== null) return;           // an earlier listing row already gave one
        var d = (I.packageLengthIn >= 0 && I.packageWidthIn >= 0 && I.packageDepthIn >= 0)
          ? _wtDims([row[I.packageLengthIn], row[I.packageWidthIn], row[I.packageDepthIn]]) : null;
        map[k] = { oz: oz, src: "eBay", dims: d, title: t };
      });
    } catch (e) { try { console.log("_wtWeightMap: Master Inventory read failed: " + e); } catch (_) {} }
  }
  return map;
}

/** items [{sku, qty, title?}] + a weight map → the result. Pure. */
function _wtCompute(items, wmap) {
  var res = { lines: [], totalOz: 0, pieces: 0, missing: [], largest: null };
  (items || []).forEach(function (it) {
    var w = wmap[_wtKey(it.sku)] || null, qty = Math.max(1, +it.qty || 1);
    var oz = w ? w.oz : null;
    var line = { sku: String(it.sku), qty: qty, title: (w && w.title) || it.title || "",
                 ozEach: oz, ozTotal: oz === null ? null : Math.round(oz * qty * 10) / 10,
                 src: w && oz !== null ? w.src : "", dims: w ? w.dims : null };
    res.pieces += qty;
    if (oz === null) res.missing.push(line.sku); else res.totalOz += oz * qty;
    if (line.dims) {
      var vol = line.dims[0] * line.dims[1] * line.dims[2];
      if (!res.largest || vol > res.largest.vol) res.largest = { sku: line.sku, dims: line.dims, vol: vol };
    }
    res.lines.push(line);
  });
  res.totalOz = Math.round(res.totalOz * 10) / 10;
  return res;
}

/** The message — for Telegram and for "Copy for shipping". Pure. */
function _wtText(res, head) {
  var L = ["⚖ " + (head || "Shipping weight") + " · " + res.pieces + " item" + (res.pieces === 1 ? "" : "s"), ""];
  res.lines.forEach(function (ln) {
    var name = ln.title ? "  " + (ln.title.length > 32 ? ln.title.slice(0, 31) + "…" : ln.title) : "";
    L.push(ln.qty + " × " + ln.sku + name);
    L.push("    " + (ln.ozEach === null ? "⚠ no weight on file"
      : (ln.qty > 1 ? _wtFmt(ln.ozEach) + " each · " + _wtFmt(ln.ozTotal) : _wtFmt(ln.ozEach)) +
        (ln.src === "eBay" ? "  (eBay listing weight)" : "")));
  });
  L.push("");
  if (res.lines.length === res.missing.length) {
    L.push("⚠ No weight on file for any of these.");
  } else {
    L.push("TOTAL ≈ " + _wtFmt(res.totalOz) + "  (" + (Math.round(res.totalOz / 16 * 100) / 100) + " lb) · parts only");
    if (res.missing.length) L.push("⚠ NOT in the total — no weight on file: " + res.missing.join(", "));
  }
  if (res.largest) L.push("Largest item: " + res.largest.dims.join(" × ") + " in (" + res.largest.sku + ")");
  return L.join("\n");
}

/**
 * The lines of an order, for weighing. All Orders first — exact order id, CANCELED lines out,
 * and an EXPANDED kit counted once (its components, not the parent box as well). Not on the
 * sheet yet → a Zoho SO's own lines from the Pending mirror.
 * → { ok, head, items[], reason }
 */
function _wtOrderItems(query) {
  var q = String(query || "").trim();
  var norm = function (s) { return String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, ""); };
  var rows = [];
  try { rows = (_findOrderRows(norm(q)) || []).filter(function (r) { return norm(r.salesOrder) === norm(q); }); } catch (e) {}
  if (!rows.length && /^INV/i.test(q)) {
    // An invoice number: find its SO in the Pending mirror, then look again.
    try { var dz = computeZohoSoDiff(q); if (dz && dz.ok && dz.soNumber) { q = dz.soNumber;
      rows = (_findOrderRows(norm(q)) || []).filter(function (r) { return norm(r.salesOrder) === norm(q); }); } } catch (e) {}
  }
  if (rows.length) {
    var expanded = {};
    rows.forEach(function (r) { var k = kitComponentTag(r.note); if (k) expanded[_wtKey(k)] = true; });
    var items = [];
    rows.forEach(function (r) {
      if (String(r.status).toUpperCase() === Schema.status.CANCELED) return;
      if (!kitComponentTag(r.note) && expanded[_wtKey(r.sku)]) return;   // the kit box, already counted as parts
      if (String(r.sku).trim()) items.push({ sku: String(r.sku).trim(), qty: +r.qty || 1 });
    });
    return { ok: true, head: rows[0].salesOrder, items: items };
  }
  try {
    var d = computeZohoSoDiff(q);
    if (d && d.ok && d.lines && d.lines.length) {
      var zi = d.lines.filter(function (ln) { return ln.status !== "removed" && +ln.zohoQty > 0; })
                      .map(function (ln) { return { sku: ln.sku, qty: +ln.zohoQty, title: ln.name || "" }; });
      return { ok: true, head: d.soNumber + " (from Zoho — not on the sheet yet)", items: zi };
    }
  } catch (e) {}
  return { ok: false, reason: "Order " + q + " is not on the sheet or in Pending Sales Orders." };
}

/** Weigh [{sku, qty}] — the quote basket's door. → { ok, result, text } */
function weighItems(items, head) {
  var list = (items || []).filter(function (x) { return x && String(x.sku || "").trim(); }).slice(0, WEIGHT.maxItems);
  if (!list.length) return { ok: false, reason: "Nothing to weigh." };
  var res = _wtCompute(list, _wtWeightMap(list.map(function (x) { return x.sku; })));
  return { ok: true, result: res, text: _wtText(res, head) };
}

/** Weigh an order by id — the Order Case window's door. → { ok, result, text } | { ok:false } */
function weighOrder(query) {
  var o = _wtOrderItems(query);
  if (!o.ok) return o;
  if (!o.items.length) return { ok: false, reason: "No open lines on " + o.head + "." };
  return weighItems(o.items, o.head);
}

/** Telegram /weight: an order id, or a typed list. → message text */
function weighQueryText(text) {
  var t = String(text || "").trim();
  if (!t) return "Usage: /weight <order>  or  /weight <SKU> x<qty>, …\nExamples:\n  /weight SO-26018\n  /weight 166500 x2, 173817, 171432 x3";
  if (_wtLooksLikeOrder(t)) {
    var r = weighOrder(t);
    return r.ok ? r.text : "⚠ " + r.reason;
  }
  var p = _wtParseList(t);
  if (!p.items.length) return "⚠ No SKUs found in that." + (p.bad.length ? " Could not read: " + p.bad.join(", ") : "");
  if (p.items.length > WEIGHT.maxItems) return "⚠ That is " + p.items.length + " SKUs — " + WEIGHT.maxItems + " at most.";
  var w = weighItems(p.items, "Shipping weight");
  return w.text + (p.bad.length ? "\n\n⚠ Skipped, could not read: " + p.bad.join(", ") : "");
}

/** EDITOR-RUN check. */
function testWeighNow() {
  console.log(weighQueryText("166500 x2, 173817"));
}
