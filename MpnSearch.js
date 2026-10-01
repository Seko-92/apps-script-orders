// =======================================================================================
// MpnSearch.js — "a customer sent part numbers; which ones do we carry?"
// =======================================================================================
//
// THE DAILY PROCESS THIS REPLACES (2026-10-01). A customer sends an engine serial number;
// someone looks it up in the SERPIC app, copies out the MPNs, then searches eBay for them
// ONE AT A TIME — and misses half of what we carry, because eBay's search only matches the
// FIRST number in a listing's MPN field. The extra numbers are stored on the listing; eBay
// just never searches them.
//
// Measured on the 2026-09-29 export: 1,543 of 3,600 active listings carry several numbers
// in `C:MPN` (comma-separated), and ~19 one-off columns (`C:Interchange Part Number`,
// `C:Additional MPN`, `C:MPN 2`, …) carry more on ~40 listings. 5,098 distinct numbers in
// all. Every one of them is searchable here.
//
// SOURCE = MASTER INVENTORY ONLY — the user's call, and the right one: MI is written by the
// eBay sync, so it is what the listings actually say. By Part Type is hand-kept and
// measured 6% of its numbers absent from MI (some of them junk like "821").
//
// ⚠ NO CACHE, NO INDEX JOB. Same rule as getPartData: this is a decision surface (someone
// tells a customer "we have it, shelf E-17, 6 on hand") and the numbers move. One bounded
// MI read is ~2s — the round trip dominates, not the width.
//
// ⚠⚠ THE MATCH KEY IS THE WHOLE CORRECTNESS STORY. Deutz numbers are 8 digits with a
// leading zero (`02102238`), and on hundreds of MI cells the zero is GONE because the
// value was stored as a number (`2102238`). SERPIC also prints them with a space
// (`0415 7075`). So both sides are reduced to: uppercase, letters+digits only, leading
// zeros stripped. Same lesson as the `000000` placeholder SKU (2026-09-29).
//
// Every search is LOGGED (MPN Search Log) — found or not. The misses are the valuable half:
// "asked for 04157738 four times, we don't list it" is what to list next.
// =======================================================================================

var MPN_SEARCH = {
  // A C: column holds part numbers if its NAME says so. Matched by name, never position
  // (MiSchema's whole point) — and a new oddly-named column typed on a listing is picked
  // up automatically.
  colPattern: /^C:.*(mpn|part ?num|interchang|replac|reference)/i,
  mainHeader: "C:MPN",

  // A pasted SERPIC block also carries position numbers, quantities and descriptions.
  // Keys shorter than this are treated as noise ("1", "12", "2x").
  minKeyLen: 5,
  maxNumbers: 60,          // one paste; a longer one is almost certainly not a parts list

  logSheet: "MPN Search Log",
  logHeaders: ["TIMESTAMP", "MPN", "FOUND", "SKUS", "SOURCE", "WHO"]
};


// ---------------------------------------------------------------------------------------
// PURE HELPERS — no Apps Script calls, so Node can test them
// ---------------------------------------------------------------------------------------

/** The comparison key: uppercase, letters+digits only, leading zeros stripped. */
function _mpnKey(s) {
  var k = String(s == null ? "" : s).toUpperCase().replace(/[^A-Z0-9]/g, "");
  var stripped = k.replace(/^0+/, "");
  return stripped || k;
}

/** SKU comparison key — numbers read back as 157554 or "157554.0" must agree. */
function _mpnSkuKey(s) { return _mpnCellText(s).toLowerCase(); }

/** A cell value as a string — a number stored as 4201560 or 4201560.0 reads "4201560". */
function _mpnCellText(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number") return isFinite(v) ? String(Math.round(v)) : "";
  return String(v).replace(/^(\d+)\.0$/, "$1").trim();
}

/** Join Deutz-style "0415 7075" into one number before anything splits on spaces. */
function _mpnJoinSplitNumbers(s) {
  return String(s).replace(/(^|[^0-9A-Za-z])(\d{4})[ \t]+(\d{4})(?![0-9A-Za-z])/g, "$1$2$3");
}

/**
 * Raw numbers stored in one MI cell, in order (index 0 = the one eBay searches).
 * Separators: comma, semicolon, slash, newline, and spaces between whole numbers.
 * When a chunk had spaces we ALSO keep it joined — `1A033- 03043` must still match.
 */
function _mpnSplitCell(v) {
  var s = _mpnCellText(v);
  if (!s) return [];
  var out = [];
  _mpnJoinSplitNumbers(s).split(/[,;\/\n]+/).forEach(function (chunk) {
    chunk = chunk.trim();
    if (!chunk) return;
    var parts = chunk.split(/\s+/);
    parts.forEach(function (p) { if (/\d/.test(p)) out.push(p); });
    if (parts.length > 1 && /\d/.test(chunk)) out.push(chunk);
  });
  return out;
}

/**
 * Turn whatever was pasted into a deduped list of numbers to look up.
 * Accepts one number, a list, or a block copied from SERPIC with positions,
 * descriptions and quantities mixed in.
 * @return {Array<{raw:string,key:string}>}
 */
function _mpnParseQuery(text) {
  var src = _mpnJoinSplitNumbers(String(text == null ? "" : text));
  var seen = {}, out = [], loose = [];
  src.split(/[\s,;\/|]+/).forEach(function (tok) {
    var raw = tok.replace(/^[^0-9A-Za-z]+|[^0-9A-Za-z]+$/g, "");
    if (!raw || !/\d/.test(raw)) return;
    var key = _mpnKey(raw);
    if (!key || seen[key]) return;
    seen[key] = true;
    if (key.length >= MPN_SEARCH.minKeyLen) out.push({ raw: raw, key: key });
    else loose.push({ raw: raw, key: key });
  });
  // A single short number typed on its own is a real query, not noise.
  if (!out.length && loose.length === 1) return loose;
  return out.slice(0, MPN_SEARCH.maxNumbers);
}

/**
 * Index every part number in MI rows.
 * @param {Array<Array>} rows      MI rows (any slice)
 * @param {Array<{name,off}>} mpnCols   part-number columns, offsets into each row
 * @return {Map<string, Array<{row:number, via:string}>>}
 */
function _mpnBuildIndex(rows, mpnCols) {
  var index = new Map();
  for (var r = 0; r < rows.length; r++) {
    var perRow = {};
    for (var c = 0; c < mpnCols.length; c++) {
      var col = mpnCols[c];
      var nums = _mpnSplitCell(rows[r][col.off]);
      for (var n = 0; n < nums.length; n++) {
        var key = _mpnKey(nums[n]);
        if (key.length < 3) continue;   // "1", "12" stored in a messy cell
        var isMain = (col.name.toLowerCase() === MPN_SEARCH.mainHeader.toLowerCase() && n === 0);
        var via = isMain ? "main" : (col.name.toLowerCase() === MPN_SEARCH.mainHeader.toLowerCase()
          ? "extra MPN" : col.name.replace(/^C:/, ""));
        // one entry per row per key — the MAIN match wins if it is both
        if (perRow[key] && perRow[key] !== "main" && isMain) perRow[key] = "main";
        else if (!perRow[key]) perRow[key] = via;
      }
    }
    Object.keys(perRow).forEach(function (key) {
      if (!index.has(key)) index.set(key, []);
      index.get(key).push({ row: r, via: perRow[key] });
    });
  }
  return index;
}

/**
 * Answer the query against an index. Pure.
 * @param {Array<{raw,key}>} query
 * @param {Map} index
 * @param {Function} describe   row → match object (stock, shelf, price…)
 */
function _mpnAnswer(query, index, describe) {
  return query.map(function (q) {
    var hits = (index.get(q.key) || []).map(function (h) {
      var m = describe(h.row);
      m.via = h.via;
      return m;
    });
    // Active listings first, then the PART before the kits that contain it (someone
    // asking for a piston wants the piston, not three overhaul kits — 2026-10-01 floor
    // feedback), then whatever has stock, then by SKU — deterministic.
    hits.sort(function (a, b) {
      var aa = a.active ? 0 : 1, bb = b.active ? 0 : 1;
      if (aa !== bb) return aa - bb;
      var ak = a.isKit ? 1 : 0, bk = b.isKit ? 1 : 0;
      if (ak !== bk) return ak - bk;
      var as = (a.available || 0) > 0 ? 0 : 1, bs = (b.available || 0) > 0 ? 0 : 1;
      if (as !== bs) return as - bs;
      return String(a.sku).localeCompare(String(b.sku));
    });
    return { query: q.raw, key: q.key, matches: hits };
  });
}


// ---------------------------------------------------------------------------------------
// THE SEARCH
// ---------------------------------------------------------------------------------------

// ---------------------------------------------------------------------------------------
// ONE MI READ, SHARED BY BOTH SEARCHES
// ---------------------------------------------------------------------------------------

/**
 * Read Master Inventory once (one bounded span via MiSchema) plus Zoho stock and the
 * kit registry, and hand back a `describe(i)` that turns row i into a result row.
 * @param {string[]=} extra   further optional MI columns the caller needs
 */
function _mpnLoadMi(extra) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DB_SHEET_NAME);
  if (!sheet) throw new Error("Master Inventory not found.");

  var headers = MiSchema.headers(sheet);
  var mpnNames = [];
  headers.forEach(function (h) {
    var s = String(h == null ? "" : h).trim();
    if (MPN_SEARCH.colPattern.test(s) && mpnNames.indexOf(s) === -1) mpnNames.push(s);
  });

  var fields = [DB_TITLE_HEADER, DB_LOCATION_HEADER, DB_QUANTITY_HEADER, DB_QUANTITY_SOLD_HEADER,
                "currentPrice", "startPrice", DB_LISTING_STATUS_HEADER, DB_VIEWURL_HEADER];
  var optional = fields.concat(extra || []).concat(
    mpnNames.filter(function (n) { return n !== MPN_SEARCH.mainHeader; }));
  var r = MiSchema.readColumns(sheet, [DB_SKU_HEADER, MPN_SEARCH.mainHeader], { optional: optional });
  var mpnCols = mpnNames.map(function (n) { return { name: n, off: r.idx[n] }; })
                        .filter(function (c) { return c.off >= 0; });

  var zoho = null;
  try { zoho = buildZohoStockMap(); } catch (e) { zoho = null; }

  // Which SKUs are kits — the registry is the source, same as everywhere else.
  // Best effort: if it can't be read, nothing is tagged rather than the search failing.
  var kits = {};
  try { buildKitMap().forEach(function (v, k) { kits[_mpnSkuKey(k)] = true; }); } catch (e) {}

  function get(row, name) { var o = r.idx[name]; return (o != null && o >= 0) ? row[o] : ""; }
  function describe(i) {
    var row = r.rows[i];
    var sku = _mpnCellText(get(row, DB_SKU_HEADER));
    var qty = parseFloat(get(row, DB_QUANTITY_HEADER)) || 0;
    var sold = parseFloat(get(row, DB_QUANTITY_SOLD_HEADER)) || 0;
    var z = zoho ? zoho.get(sku.toLowerCase()) : null;
    var cur = parseFloat(get(row, "currentPrice")), st = parseFloat(get(row, "startPrice"));
    var status = String(get(row, DB_LISTING_STATUS_HEADER) || "").trim();
    return {
      sku: sku,
      title: String(get(row, DB_TITLE_HEADER) || ""),
      location: String(get(row, DB_LOCATION_HEADER) || "").trim() || "NOT FOUND",
      // Zoho is the stock master; MI is the fallback — the house rule.
      available: z ? z.available : (qty - sold),
      price: (!isNaN(cur) && cur > 0) ? cur : ((!isNaN(st) && st > 0) ? st : null),
      status: status,
      active: !status || status === "Active",
      isKit: !!kits[_mpnSkuKey(sku)],
      url: String(get(row, DB_VIEWURL_HEADER) || "").trim()
    };
  }
  return { r: r, get: get, mpnCols: mpnCols, describe: describe };
}


/**
 * @param {string} text     one number, a list, or a pasted SERPIC block
 * @param {Object=} opts    { source: "console"|"telegram"|..., who: string, noLog: bool }
 * @return {{ok:boolean, results:Array, found:number, missing:number, ms:number, reason?:string}}
 */
function searchMpns(text, opts) {
  opts = opts || {};
  var t0 = Date.now();
  try {
    var query = _mpnParseQuery(text);
    if (!query.length) return { ok: false, reason: "No part numbers found in what was pasted." };

    var mi = _mpnLoadMi();
    var index = _mpnBuildIndex(mi.r.rows, mi.mpnCols);
    var results = _mpnAnswer(query, index, mi.describe);
    var found = results.filter(function (x) { return x.matches.length; }).length;

    if (!opts.noLog) _mpnLog(results, opts.source || "", opts.who || "");

    return { ok: true, mode: "mpn", results: results, found: found, missing: results.length - found,
             columns: mi.mpnCols.length, ms: Date.now() - t0 };
  } catch (err) {
    try { console.log("searchMpns: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}


// =======================================================================================
// KEYWORD SEARCH — "deutz 912 head gasket", "V2203 piston" (2026-10-01)
// =======================================================================================
//
// For engines SERPIC does not cover (Kubota, Perkins, Yanmar, newer Deutz) the way in is
// the engine model and the part's name — which the listings already carry in the title
// and in `C:Model` / `C:Part Type` / the compatibility fields.
//
// THE RULE IS LITERAL ON PURPOSE: every word typed must appear somewhere on the listing.
// eBay's search guesses and ranks by its own relevance; this one never shows something
// that doesn't contain what was asked for. Three tolerances make literal usable:
//   • plurals:   "gaskets" = "gasket", "valves" = "valve"
//   • prefixes:  "bf4m" finds "BF4M1011", "gask" finds "gasket"
//   • model codes split by a space or glued to a prefix: "bf6m1013" finds "BF6 M1013",
//     "912" finds "F3L912"
//
// ⚠ C:Brand is deliberately NOT searched — it is "HQ" on every listing, so a query
//   containing "hq" would match everything. C:Model Year is the SHELF, not a year.

var KW_SEARCH = {
  fields: ["C:Model", "C:Part Type", "C:Compatible Equipment Make",
           "C:Compatible Equipment Type", "C:Engine Type"],
  maxResults: 100,
  minWord: 2
};

/** Lowercase, letters+digits, simple plural fold. Pure. */
function _kwNorm(w) {
  var t = String(w == null ? "" : w).toLowerCase().replace(/[^a-z0-9]/g, "");
  if (t.length > 3 && /[a-z]s$/.test(t) && !/ss$/.test(t)) t = t.slice(0, -1);
  return t;
}

/** The words of a query, deduped. Pure. */
function _kwParseQuery(text) {
  var seen = {}, out = [];
  String(text == null ? "" : text).split(/[\s,;\/|]+/).forEach(function (w) {
    var k = _kwNorm(w);
    if (k.length < KW_SEARCH.minWord || seen[k]) return;
    seen[k] = true; out.push(k);
  });
  return out;
}

/**
 * Every searchable token of a listing: each word, plus each pair of neighbours glued
 * together (so "BF6 M1013" also yields "bf6m1013"). Pure.
 */
function _kwTokens(texts) {
  var words = [], glued = [];
  texts.forEach(function (t) {
    var w = String(t == null ? "" : t).split(/[^A-Za-z0-9]+/).map(_kwNorm).filter(Boolean);
    for (var i = 0; i < w.length; i++) {
      words.push(w[i]);
      if (i + 1 < w.length) glued.push(w[i] + w[i + 1]);
    }
  });
  return { words: words, glued: glued };
}

/**
 * Does one query word appear? Pure.
 * Prefix match on words and glued pairs; the "912 → F3L912" end-of-code match on real
 * words ONLY — on a glued pair it would let "013" hit "part 4013".
 */
function _kwHit(word, toks) {
  var i;
  for (i = 0; i < toks.words.length; i++) if (toks.words[i].indexOf(word) === 0) return true;
  for (i = 0; i < toks.glued.length; i++) if (toks.glued[i].indexOf(word) === 0) return true;
  if (/^\d{3,}$/.test(word)) {
    for (i = 0; i < toks.words.length; i++) {
      var t = toks.words[i];
      if (t.length > word.length && /[a-z]/.test(t) && t.slice(-word.length) === word) return true;
    }
  }
  return false;
}

/**
 * Rank rows for a query. Pure.
 * @param {string[]} words
 * @param {Array<{title:string, other:string[]}>} docs
 * @return {Array<{row:number, score:number}>}  matches only, best first
 */
function _kwMatch(words, docs) {
  var out = [];
  for (var r = 0; r < docs.length; r++) {
    var tTitle = _kwTokens([docs[r].title]), tOther = null, score = 0, ok = true;
    for (var w = 0; w < words.length; w++) {
      if (_kwHit(words[w], tTitle)) { score += 2; continue; }      // a title match counts double
      if (!tOther) tOther = _kwTokens(docs[r].other);
      if (_kwHit(words[w], tOther)) { score += 1; continue; }
      ok = false; break;
    }
    if (ok) out.push({ row: r, score: score });
  }
  return out;
}

/**
 * @param {string} text   words, e.g. "deutz 912 head gasket"
 * @param {Object=} opts  { source, who, noLog }
 */
function searchKeywords(text, opts) {
  opts = opts || {};
  var t0 = Date.now();
  try {
    var words = _kwParseQuery(text);
    if (!words.length) return { ok: false, reason: "Type a few words — e.g. deutz 912 head gasket" };

    var mi = _mpnLoadMi(KW_SEARCH.fields);
    var fieldsPresent = KW_SEARCH.fields.filter(function (f) { return mi.r.idx[f] >= 0; });
    var docs = mi.r.rows.map(function (row) {
      var other = fieldsPresent.map(function (f) { return mi.get(row, f); });
      mi.mpnCols.forEach(function (c) { other.push(row[c.off]); });   // a part number typed among the words
      return { title: mi.get(row, DB_TITLE_HEADER), other: other };
    });

    var hits = _kwMatch(words, docs).map(function (h) {
      var m = mi.describe(h.row); m.score = h.score; return m;
    }).filter(function (m) { return m.sku; });

    // Best match first (title matches beat spec matches), then the same order as the
    // MPN finder: active, the part before kits, in stock, SKU.
    hits.sort(function (a, b) {
      if (a.score !== b.score) return b.score - a.score;
      var aa = a.active ? 0 : 1, bb = b.active ? 0 : 1;
      if (aa !== bb) return aa - bb;
      var ak = a.isKit ? 1 : 0, bk = b.isKit ? 1 : 0;
      if (ak !== bk) return ak - bk;
      var as = (a.available || 0) > 0 ? 0 : 1, bs = (b.available || 0) > 0 ? 0 : 1;
      if (as !== bs) return as - bs;
      return String(a.sku).localeCompare(String(b.sku));
    });

    if (!opts.noLog) {
      _mpnLog([{ query: words.join(" "), matches: hits.slice(0, 10) }],
              (opts.source || "") + " · keywords", opts.who || "");
    }

    return { ok: true, mode: "keywords", query: words.join(" "), total: hits.length,
             matches: hits.slice(0, KW_SEARCH.maxResults), ms: Date.now() - t0 };
  } catch (err) {
    try { console.log("searchKeywords: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}

/** In-window keyword search for the modal. */
function getKeywordSearch(text) {
  return searchKeywords(text, { source: "console", who: _mpnWho() });
}


// ---------------------------------------------------------------------------------------
// THE LOG — best effort; a logging failure must never cost the person their answer
// ---------------------------------------------------------------------------------------

function _mpnLog(results, source, who) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(MPN_SEARCH.logSheet) || _mpnSetupLog(ss);
    var now = new Date();
    var rows = results.map(function (x) {
      return [now, x.query, x.matches.length ? "YES" : "NO",
              x.matches.map(function (m) { return m.sku; }).join(", "), source, who];
    });
    if (!rows.length) return;
    // MPN as plain text, or Sheets turns 02102238 into 2102238 — the bug this file exists for.
    var range = sh.getRange(sh.getLastRow() + 1, 1, rows.length, MPN_SEARCH.logHeaders.length);
    range.offset(0, 1, rows.length, 1).setNumberFormat("@");
    range.setValues(rows);
  } catch (e) {
    try { console.log("_mpnLog: " + e); } catch (_) {}
  }
}

function _mpnSetupLog(ss) {
  var sh = ss.insertSheet(MPN_SEARCH.logSheet);
  sh.getRange(1, 1, 1, MPN_SEARCH.logHeaders.length).setValues([MPN_SEARCH.logHeaders])
    .setBackground("#1d1d1b").setFontColor("#ffd400").setFontFamily("Oswald").setFontWeight("bold");
  sh.setFrozenRows(1);
  sh.getRange("A:A").setNumberFormat("m/d/yy h:mm AM/PM");
  sh.getRange("B:B").setNumberFormat("@").setFontFamily("Roboto Mono");
  sh.setColumnWidths(1, 6, 130);
  sh.setColumnWidth(4, 220);
  return sh;
}

/** Open the log (sidebar button). */
function openMpnSearchLog() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(MPN_SEARCH.logSheet) || _mpnSetupLog(ss);
  ss.setActiveSheet(sh);
  return { ok: true };
}


// ---------------------------------------------------------------------------------------
// SURFACES
// ---------------------------------------------------------------------------------------

/** In-window search for the modal. */
function getMpnSearch(text) {
  return searchMpns(text, { source: "console", who: _mpnWho() });
}

/** Open the finder modal, pre-run on whatever was typed in the sidebar.
 *  @param {string} text
 *  @param {string=} mode  "mpn" (default) or "keywords" */
function openMpnSearch(text, mode) {
  try {
    mode = (mode === "keywords") ? "keywords" : "mpn";
    var has = String(text || "").trim();
    var res = has ? (mode === "keywords" ? getKeywordSearch(text) : getMpnSearch(text)) : null;
    var t = HtmlService.createTemplateFromFile("MpnSearchModal");
    t.initJson = JSON.stringify({ text: String(text || ""), mode: mode, res: res }).replace(/<\//g, "<\\/");
    SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(1080).setHeight(720), "Parts Finder");
    var summary = !res || !res.ok ? {} : (mode === "keywords"
      ? { found: res.total, missing: 0 } : { found: res.found, missing: res.missing });
    return { ok: true, mode: mode, found: summary.found || 0, missing: summary.missing || 0 };
  } catch (err) {
    try { console.log("openMpnSearch: " + err); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}

function _mpnWho() {
  try { var p = _currentPicker(); if (p) return String(p); } catch (_) {}
  return "";
}

/** Telegram /find. */
function _tgFormatFind(argStr, who) {
  var res = searchMpns(argStr, { source: "telegram", who: who || "" });
  if (!res.ok) return "⚠ " + res.reason;

  var L = ["🔎 " + res.results.length + " number" + (res.results.length === 1 ? "" : "s") +
           " · " + res.found + " found · " + res.missing + " not carried", ""];
  res.results.forEach(function (x) {
    if (!x.matches.length) { L.push("✗ " + x.query + " — not on any listing"); return; }
    x.matches.slice(0, 4).forEach(function (m, i) {
      L.push((i === 0 ? "✓ " + x.query : "   ") + " → " + m.sku + " · " + m.location +
             " · on hand " + _tgNum(m.available) + " · " + _tgMoney(m.price) +
             (m.isKit ? " · KIT" : "") + (m.via === "main" ? "" : " (" + m.via + ")") + (m.active ? "" : " · " + (m.status || "ended")));
      if (m.title) L.push("     " + _tgClip(m.title, 52));
    });
    if (x.matches.length > 4) L.push("     … +" + (x.matches.length - 4) + " more listings");
  });
  return L.join("\n");
}

/** Telegram /search. */
function _tgFormatSearch(argStr, who) {
  var res = searchKeywords(argStr, { source: "telegram", who: who || "" });
  if (!res.ok) return "⚠ " + res.reason;
  if (!res.total) return "🔎 " + res.query + "\n\nNothing on any listing contains all of those words.";
  var show = res.matches.slice(0, 12);
  var L = ["🔎 " + res.query + " · " + res.total + " listing" + (res.total === 1 ? "" : "s"), ""];
  show.forEach(function (m) {
    L.push(m.sku + " · " + m.location + " · on hand " + _tgNum(m.available) + " · " + _tgMoney(m.price) +
           (m.isKit ? " · KIT" : "") + (m.active ? "" : " · " + (m.status || "ended")));
    L.push("   " + _tgClip(m.title, 56));
  });
  if (res.total > show.length) L.push("\n… +" + (res.total - show.length) + " more — add a word to narrow it");
  return L.join("\n");
}
