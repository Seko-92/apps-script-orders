// =======================================================================================
// ProductHealth.js — "Product Data Health": one sheet, three jobs
// =======================================================================================
//
// WHAT THIS IS. A computed verdict sheet that reconciles the THREE places product facts
// live, and reports every row where they disagree or where a fact is missing:
//
//   By Part Type  (a SEPARATE Google Sheet)  → SOURCE OF TRUTH for weight + dimensions,
//                                              fitment, part numbers, per-type specs
//   Master Inventory (this file)             → what eBay currently believes
//   Zoho Stock (this file)                   → stock (read via the existing mirror)
//
// ⭐ ONE SHEET, THREE JOBS — filter it three ways and it is three different worklists:
//     BAND = ⚠ UNDER-CHARGING / ⚠ MISMATCH  → the eBay correction list
//     BAND = ○ INCOMPLETE                    → the data-completion worklist
//     BAND = ✓ READY, sorted by SOLD         → the Amazon launch shortlist
//
// ⭐ INCOMPLETE ROWS ARE RANKED BY UNITS SOLD, deliberately. A flat list of 1,700 gaps is
//   a wall nobody starts. Ranked by what actually sells, the top of the list is the part
//   of the backlog worth doing first — the same reasoning as "photos ranked by velocity".
//
// ⚠⚠ IT IS MACHINE-OWNED AND WRITES NOWHERE ELSE. It never edits By Part Type, MI or Zoho.
//   It reports differences; the humans fix them at the source. A row leaves this sheet when
//   it is fixed, so the sheet empties itself — same contract as Out of Stock / Photo Queue.
//
// ⚠⚠ THE `Photo` / `Zoho` / `Website` COLUMNS IN BY PART TYPE ARE DELIBERATELY NOT READ.
//   Measured 2026-09-09: the hand-kept `Photo` flag disagrees with MI's ACTUAL image count
//   on 36.2% of rows (1,277 of 3,527) — and mostly in the direction "flag says 0, eBay has
//   4 images", i.e. the photos were taken and nobody ticked the box. A hand-maintained
//   mirror of a fact the system can already see is worse than no column at all, because it
//   sends someone to re-shoot work that is already done. PHOTOS ARE COUNTED FROM MI HERE.
//   Those three columns should be DELETED from By Part Type once this sheet is live.
//
// ⚠ THE FIRST CROSS-SPREADSHEET READ IN THE PROJECT. Everything else opens SPREADSHEET_ID.
//   The script must be authorised for BOTH files, and a missing grant surfaces first from a
//   TRIGGER context, silently (LocationUpdate.js documents that exact failure). Run
//   probeByPartTypeRead() once from the editor before wiring this to any trigger.
// =======================================================================================

var PRODUCT_HEALTH = {
  sheetName: "Product Data Health",

  cols: {
    SKU:          1,   // A
    BRAND:        2,   // B — the ENGINE make from By Part Type (Deutz/Kubota/…)
    PART_TYPE:    3,   // C
    TITLE:        4,   // D
    TRUTH_WT:     5,   // E — oz, from By Part Type (as-shipped)
    EBAY_WT:      6,   // F — oz, from MI packageWeightOz
    WT_DELTA:     7,   // G — eBay − truth
    TRUTH_DIMS:   8,   // H — "LxWxD"
    EBAY_DIMS:    9,   // I — "LxWxD"
    PHOTOS:      10,   // J — counted from MI, never from a flag
    SOLD:        11,   // K
    ON_HAND:     12,   // L
    GAPS:        13,   // M — what is missing, in words
    EBAY_FIX:    14,   // N — the correction verdict
    AMAZON:      15,   // O — READY, or why not
    LAST_CHECKED:16    // P
  },

  idx: function (n) { return PRODUCT_HEALTH.cols[n] - 1; },

  dataWidth: 16,
  headerRow: 1,
  dataStartRow: 2,

  headers: ["SKU", "BRAND", "PART TYPE", "TITLE",
            "◫ TRUTH WT", "◉ EBAY WT", "WT Δ",
            "◫ TRUTH DIMS", "◉ EBAY DIMS",
            "PHOTOS", "SOLD", "ON HAND",
            "GAPS", "EBAY FIX", "AMAZON", "LAST CHECKED"],

  // Amazon readiness gates — measured 2026-09-09, see product-data-foundation.md
  amazon: { minPhotos: 2, minOnHand: 3, minSold: 1 },

  // A weight difference under this is rounding, not a discrepancy.
  wtToleranceOz: 0.5,
  dimToleranceIn: 0.5,

  bands: {
    UNDER:      "⚠ UNDER-CHARGING",
    MISMATCH:   "⚠ MISMATCH",
    INCOMPLETE: "○ INCOMPLETE",
    READY:      "✓ READY",
    NOT_LISTED: "— NOT ON EBAY"
  },

  // Sort rank: money first, then disagreement, then the worklist, then the shortlist.
  bandRank: { "⚠ UNDER-CHARGING": 0, "⚠ MISMATCH": 1, "○ INCOMPLETE": 2, "✓ READY": 3, "— NOT ON EBAY": 4 }
};

// =======================================================================================
// PURE HELPERS — no Sheets calls, so the logic is Node-testable and cannot drift
// =======================================================================================

/**
 * MI and By Part Type both store numeric SKUs as floats in some rows ("161361.0").
 * Normalise to string-of-integer or the join silently misses. Same rule as
 * KitRegistry._buildMasterInventoryMap.
 */
function _phNormSku(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return String(Math.trunc(v));
  var s = String(v).trim();
  return /^\d+\.0+$/.test(s) ? s.replace(/\.0+$/, "") : s;
}

function _phNum(v) {
  if (v === null || v === undefined || v === "") return null;
  var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? null : n;
}

/**
 * By Part Type stores weight as EITHER pounds OR ounces, not both.
 * ⚠ Ounces is the exact figure (eBay stores two integers); pounds is derived.
 * Returns ounces, or null when the row carries no weight at all.
 */
function _phTruthOz(lb, oz) {
  var L = _phNum(lb), O = _phNum(oz);
  if (L) return L * 16;
  if (O) return O;
  return null;
}

/**
 * "7x5x1" / "7 x 5 x 1" / "7X5X1" → [1,5,7] (SORTED).
 * ⚠ Sorted on purpose: a box is the same box whichever order someone typed the sides in,
 *   so comparing unsorted would manufacture hundreds of false mismatches.
 * Returns null unless exactly three numbers are present.
 */
function _phParseDims(s) {
  if (s === null || s === undefined) return null;
  var parts = String(s).split(/[xX×*]/)
    .map(function (p) { return _phNum(p); })
    .filter(function (n) { return n !== null && n > 0; });
  return parts.length === 3 ? parts.sort(function (a, b) { return a - b; }) : null;
}

function _phFmtDims(arr) {
  return (arr && arr.length === 3) ? arr.join("x") : "";
}

/**
 * THE VERDICT. Pure: takes one SKU's truth record and its MI record, returns the row's
 * classification. Either side may be null.
 *
 * @param {Object|null} t  { pt, title, lb, oz, dim }        from By Part Type
 * @param {Object|null} m  { title, oz, L, W, D, photos, sold, onHand, active } from MI
 * @return {Object} { band, gaps, ebayFix, amazon, truthOz, ebayOz, wtDelta,
 *                    truthDims, ebayDims }
 */
function _phClassify(t, m) {
  var A = PRODUCT_HEALTH.amazon;
  var truthOz  = t ? _phTruthOz(t.lb, t.oz) : null;
  var truthDim = t ? _phParseDims(t.dim) : null;
  var ebayOz   = m ? _phNum(m.oz) : null;
  var ebayDim  = (m && _phNum(m.L) && _phNum(m.W) && _phNum(m.D))
    ? [_phNum(m.L), _phNum(m.W), _phNum(m.D)].sort(function (a, b) { return a - b; })
    : null;

  var out = {
    truthOz: truthOz, ebayOz: ebayOz,
    truthDims: _phFmtDims(truthDim), ebayDims: _phFmtDims(ebayDim),
    wtDelta: null, gaps: [], ebayFix: "", amazon: "", band: ""
  };

  // ---- not on eBay at all: an expansion candidate, not a fault -----------------------
  if (!m) {
    out.band = PRODUCT_HEALTH.bands.NOT_LISTED;
    out.amazon = "not listed on eBay";
    if (truthOz === null) out.gaps.push("weight");
    if (truthDim === null) out.gaps.push("dims");
    return out;
  }

  // ---- eBay vs truth ------------------------------------------------------------------
  if (truthOz !== null && ebayOz !== null) {
    var d = ebayOz - truthOz;
    out.wtDelta = d;
    if (Math.abs(d) >= PRODUCT_HEALTH.wtToleranceOz) {
      // ⚠ eBay LIGHTER than the true shipped weight = under-charging shipping TODAY.
      //   That is the only population with a live cost, so it gets its own band.
      out.ebayFix = (d < 0) ? "eBay LIGHT by " + Math.abs(d).toFixed(0) + "oz"
                            : "eBay heavy by " + d.toFixed(0) + "oz";
    }
  }
  var dimsDisagree = false;
  if (truthDim && ebayDim) {
    for (var i = 0; i < 3; i++) {
      if (Math.abs(truthDim[i] - ebayDim[i]) >= PRODUCT_HEALTH.dimToleranceIn) dimsDisagree = true;
    }
    if (dimsDisagree) out.ebayFix = out.ebayFix ? (out.ebayFix + " · dims differ") : "dims differ";
  }

  // ---- what is missing ----------------------------------------------------------------
  if (!t) out.gaps.push("no truth record");
  if (truthOz === null)  out.gaps.push("weight");
  if (truthDim === null) out.gaps.push("dims");
  if ((m.photos || 0) < A.minPhotos) out.gaps.push("photos");

  // ---- Amazon readiness ---------------------------------------------------------------
  if (!m.active)                                out.amazon = "not Active on eBay";
  else if (truthOz === null || truthDim === null) out.amazon = "needs weight/dims";
  else if ((m.photos || 0) < A.minPhotos)        out.amazon = "needs photos";
  else if ((m.sold || 0) < A.minSold)            out.amazon = "no sales history";
  else if ((m.onHand || 0) < A.minOnHand)        out.amazon = "thin stock";
  else                                           out.amazon = "READY";

  // ---- band ---------------------------------------------------------------------------
  if (out.wtDelta !== null && out.wtDelta <= -PRODUCT_HEALTH.wtToleranceOz) {
    out.band = PRODUCT_HEALTH.bands.UNDER;          // losing money now — always first
  } else if (out.ebayFix) {
    out.band = PRODUCT_HEALTH.bands.MISMATCH;
  } else if (out.amazon === "READY") {
    out.band = PRODUCT_HEALTH.bands.READY;
  } else {
    out.band = PRODUCT_HEALTH.bands.INCOMPLETE;
  }
  return out;
}

/**
 * Sort comparator: band first, then units sold DESCENDING inside the band.
 * ⚠ The sold-desc rule is what makes the INCOMPLETE band a worklist instead of a wall —
 *   the gaps worth closing first are the ones on parts that actually sell.
 */
function _phCompareRows(a, b) {
  var ra = PRODUCT_HEALTH.bandRank[a.band], rb = PRODUCT_HEALTH.bandRank[b.band];
  if (ra !== rb) return ra - rb;
  var sa = a.sold || 0, sb = b.sold || 0;
  if (sa !== sb) return sb - sa;
  return String(a.sku).localeCompare(String(b.sku));
}

// =======================================================================================
// READING THE TRUTH FILE — the one genuinely new thing here
// =======================================================================================
//
// ⚠⚠ WEIGHT AND DIMENSIONS EXIST ONLY IN THE 78 PER-TYPE SHEETS. Checked 2026-09-09:
//   neither `All` (41 cols) nor `Temp` (42 cols) carries them, so there is no single flat
//   tab to read. Every data sheet shares the same 7-column spine —
//   SKU · Title · Brand · Part Type · Weight (lb) · Weight (Oz) · Dimension (in) — and then
//   diverges into type-specific attributes we deliberately ignore (MI has no equivalent
//   field to compare a piston bore against).
//
// ⚠ SO THIS IS ~78 CROSS-FILE ROUND TRIPS. In Apps Script the round trip dominates, not the
//   payload (measured on this project 2026-08-19: 198 cols vs 39 cols = 1,905ms vs 1,872ms).
//   That means reading only 7 columns does NOT make it cheap — the sheet COUNT is the cost.
//   ⏭ Run probeByPartTypeRead() ONCE before scheduling this. If it is slow, the fix is a
//     machine-owned mirror tab refreshed daily (the Zoho Stock pattern), NOT a narrower read.

/** Sheets in By Part Type that are utility, not part-type data. */
var _PH_SKIP_SHEETS = ["Attrs list", "SKUS-PHOTOS", "Website", "Temp",
                       "Delete", "Part Number Update", "All"];

/**
 * Reads the 7-column spine from every part-type sheet.
 * @return {Object} { bySku: {sku: {pt,title,lb,oz,dim}}, sheets: n, rows: n, ms: n }
 */
function _phReadTruth() {
  var t0 = Date.now();
  var ss = SpreadsheetApp.openById(BY_PART_TYPE_ID);
  var sheets = ss.getSheets();
  var bySku = {}, used = 0, rows = 0;

  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var name = sh.getName();
    if (_PH_SKIP_SHEETS.indexOf(name) !== -1) continue;

    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;

    // Resolve the spine BY HEADER NAME — the sheets are hand-made and the trailing
    // columns differ per type, so a positional read is exactly the bug MiSchema exists
    // to prevent. Only the first ~10 columns are ever needed.
    var width = Math.min(lastCol, 12);
    var data = sh.getRange(1, 1, lastRow, width).getValues();
    var hdr = data[0], H = {};
    for (var c = 0; c < hdr.length; c++) {
      var h = String(hdr[c] === null ? "" : hdr[c]).trim().toLowerCase();
      if (h && !H.hasOwnProperty(h)) H[h] = c;
    }
    if (!H.hasOwnProperty("sku")) continue;   // not a data sheet

    // ⚠ "Wright(Oz)" is a live typo in the file. Match on "oz" appearing anywhere
    //   rather than an exact header, or every ounce value is silently dropped.
    var iSku = H["sku"], iPt = H["part type"], iTitle = H["title"], iBrand = H["brand"];
    var iLb = H["weight (lb)"], iOz = null, iDim = null;
    for (var k in H) {
      if (iOz === null && k.indexOf("oz") !== -1) iOz = H[k];
      if (iDim === null && k.indexOf("dimension") === 0) iDim = H[k];
    }
    used++;

    for (var r = 1; r < data.length; r++) {
      var sku = _phNormSku(data[r][iSku]);
      if (!sku || sku.toLowerCase() === "sku") continue;
      rows++;
      // Last sheet wins on the rare cross-sheet duplicate (23 measured). Harmless:
      // the spine is the same in both, only the type-specific tail differs.
      bySku[sku] = {
        pt:    iPt    !== undefined ? String(data[r][iPt]    || "").trim() : "",
        brand: iBrand !== undefined ? String(data[r][iBrand] || "").trim() : "",
        title: iTitle !== undefined ? String(data[r][iTitle] || "").trim() : "",
        lb:    iLb    !== undefined ? data[r][iLb]  : "",
        oz:    iOz    !== null      ? data[r][iOz]  : "",
        dim:   iDim   !== null      ? String(data[r][iDim] || "").trim() : ""
      };
    }
  }
  return { bySku: bySku, sheets: used, rows: rows, ms: Date.now() - t0 };
}

/**
 * ⏭ RUN THIS ONCE FROM THE EDITOR BEFORE SCHEDULING ANYTHING.
 * Answers the only open question in this module: how expensive is the cross-file read?
 * Also proves the script is authorised for the second spreadsheet — a missing grant fails
 * here loudly instead of failing silently inside a trigger later.
 */
function probeByPartTypeRead() {
  var out;
  try {
    out = _phReadTruth();
  } catch (e) {
    var msg = "❌ COULD NOT READ By Part Type: " + e.message +
      "\n\nIf this is an authorisation error, open the spreadsheet once as this account, " +
      "or re-run from the editor and accept the prompt. This is the FIRST cross-spreadsheet " +
      "read in the project, so the grant may never have been given.";
    console.log(msg);
    return msg;
  }
  var wt = 0, dm = 0;
  for (var k in out.bySku) {
    if (_phTruthOz(out.bySku[k].lb, out.bySku[k].oz) !== null) wt++;
    if (_phParseDims(out.bySku[k].dim)) dm++;
  }
  var n = Object.keys(out.bySku).length;
  var lines = [
    "✅ READ By Part Type OK",
    "   data sheets read : " + out.sheets,
    "   rows scanned     : " + out.rows,
    "   unique SKUs      : " + n,
    "   with weight      : " + wt + "  (" + (n ? (wt / n * 100).toFixed(1) : 0) + "%)",
    "   with dimensions  : " + dm + "  (" + (n ? (dm / n * 100).toFixed(1) : 0) + "%)",
    "   ELAPSED          : " + (out.ms / 1000).toFixed(1) + "s",
    "",
    (out.ms > 120000
      ? "⚠ OVER 2 MINUTES — do NOT put this on a trigger as-is. Mirror it to a hidden tab"
      : "⭐ Fast enough to run directly; no mirror tab needed yet.")
  ];
  var msg = lines.join("\n");
  console.log(msg);
  return msg;
}

// =======================================================================================
// SHEET SETUP — brand-styled, idempotent (same shape as setupPriceAuditSheet)
// =======================================================================================

function setupProductDataHealthSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PRODUCT_HEALTH.sheetName);
  if (!sheet) sheet = ss.insertSheet(PRODUCT_HEALTH.sheetName);
  var C = PRODUCT_HEALTH.cols;

  sheet.getRange(PRODUCT_HEALTH.headerRow, 1, 1, PRODUCT_HEALTH.dataWidth)
    .setValues([PRODUCT_HEALTH.headers])
    .setBackground('#1d1d1b').setFontColor('#ffd966')
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(PRODUCT_HEALTH.headerRow, 1, 1, PRODUCT_HEALTH.dataWidth)
    .setBorder(null, null, true, null, null, null, '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.setFrozenRows(1);

  var W = {};
  W[C.SKU]=95; W[C.BRAND]=110; W[C.PART_TYPE]=150; W[C.TITLE]=290;
  W[C.TRUTH_WT]=95; W[C.EBAY_WT]=95; W[C.WT_DELTA]=80;
  W[C.TRUTH_DIMS]=110; W[C.EBAY_DIMS]=110;
  W[C.PHOTOS]=70; W[C.SOLD]=70; W[C.ON_HAND]=80;
  W[C.GAPS]=170; W[C.EBAY_FIX]=190; W[C.AMAZON]=140; W[C.LAST_CHECKED]=130;
  for (var c in W) sheet.setColumnWidth(parseInt(c, 10), W[c]);

  var maxRow = 4200, n = maxRow - PRODUCT_HEALTH.dataStartRow + 1;
  var R = function (col) { return sheet.getRange(PRODUCT_HEALTH.dataStartRow, col, n, 1); };

  R(C.SKU).setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  R(C.PART_TYPE).setFontFamily('Roboto').setFontSize(10);
  R(C.TITLE).setFontFamily('Roboto').setFontSize(10);
  // ⚠ Numeric columns get an EXPLICIT number format — Gotcha #16. A column that inherits a
  //   DATE format renders integers as 1900-era dates AND getValues() returns Date objects.
  [C.TRUTH_WT, C.EBAY_WT, C.PHOTOS, C.SOLD, C.ON_HAND].forEach(function (col) {
    R(col).setNumberFormat('0').setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('center');
  });
  R(C.WT_DELTA).setNumberFormat('+0;-0;0').setFontFamily('Roboto Mono').setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center');
  [C.TRUTH_DIMS, C.EBAY_DIMS].forEach(function (col) {
    R(col).setNumberFormat('@').setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('center');
  });
  R(C.GAPS).setFontFamily('Roboto').setFontSize(9);
  R(C.EBAY_FIX).setFontFamily('Roboto').setFontSize(9);
  R(C.AMAZON).setFontFamily('Oswald').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  R(C.LAST_CHECKED).setNumberFormat('M/d/yy h:mm am/pm')
    .setFontFamily('Roboto Mono').setFontSize(9).setFontColor('#5f5f5f').setHorizontalAlignment('center');
  sheet.getRange(PRODUCT_HEALTH.dataStartRow, 1, n, PRODUCT_HEALTH.dataWidth).setVerticalAlignment('middle');

  _phApplyConditionalFormatting(sheet);
  SpreadsheetApp.flush();
  return "✅ '" + PRODUCT_HEALTH.sheetName + "' ready.";
}

/**
 * ⚠ Rules are applied FIRST-MATCH-WINS by Sheets, so ORDER IS LOAD-BEARING:
 *   the money band must precede the softer ones or it gets painted over.
 */
function _phApplyConditionalFormatting(sheet) {
  // ⚠⚠ DERIVE the column letters from the schema — NEVER hardcode them. These used to
  //    read $M2 / $N2, which was correct only while EBAY_FIX sat at 13 and AMAZON at 14.
  //    Adding the BRAND column on 2026-09-10 shifted both right by one, and a hardcoded
  //    letter would have kept matching — silently painting the wrong column, with no
  //    error and no wrong-looking number. Same class as the KitRegistry positional read
  //    that MiSchema.js exists to prevent.
  var FIXC = _colLetter(PRODUCT_HEALTH.cols.EBAY_FIX);
  var AMZC = _colLetter(PRODUCT_HEALTH.cols.AMAZON);
  var C = PRODUCT_HEALTH.cols, B = PRODUCT_HEALTH.bands;
  var maxRow = 4200;
  var whole = sheet.getRange(PRODUCT_HEALTH.dataStartRow, 1, maxRow - 1, PRODUCT_HEALTH.dataWidth);
  var amazonCol = sheet.getRange(PRODUCT_HEALTH.dataStartRow, C.AMAZON, maxRow - 1, 1);
  var fixCol = sheet.getRange(PRODUCT_HEALTH.dataStartRow, C.EBAY_FIX, maxRow - 1, 1);
  var rules = [];

  // 1 · UNDER-CHARGING — the only band costing money today. Loudest treatment.
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH(TO_TEXT($' + FIXC + '2),"eBay LIGHT")')
    .setBackground('#ffcdd2').setFontColor('#b71c1c').setBold(true)
    .setRanges([whole]).build());

  // 2 · any other eBay disagreement — amber, act on it but nothing is bleeding
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + FIXC + '2<>"",NOT(REGEXMATCH(TO_TEXT($' + FIXC + '2),"eBay LIGHT")))')
    .setBackground('#fff4b0').setRanges([whole]).build());

  // 3 · READY — quiet green, so the Amazon shortlist reads at a glance
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('READY')
    .setBackground('#c8e6c9').setFontColor('#1b5e20').setBold(true)
    .setRanges([amazonCol]).build());

  // 4 · every non-READY reason — muted, it is a worklist not an alarm
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + AMZC + '2<>"",$' + AMZC + '2<>"READY")')
    .setFontColor('#8a7434').setRanges([amazonCol]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('LIGHT').setFontColor('#b71c1c').setBold(true)
    .setRanges([fixCol]).build());

  sheet.setConditionalFormatRules(rules);
}

function openProductDataHealth() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PRODUCT_HEALTH.sheetName);
  if (!sheet) return "❌ Sheet not found — run Setup first.";
  SpreadsheetApp.setActiveSpreadsheet(ss);
  ss.setActiveSheet(sheet);
  return "✅ Opened.";
}

// =======================================================================================
// THE REFRESH — read both sides, classify, sort, write. Writes nowhere else.
// =======================================================================================

/**
 * Reads MI once, resolving every column BY HEADER NAME via MiSchema (never by position).
 * @return {Object} { sku: {title,oz,L,W,D,photos,sold,onHand,active,pt} }
 */
function _phReadMi() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var mi = ss.getSheetByName(DB_SHEET_NAME);
  if (!mi) return {};

  var NEED = [DB_SKU_HEADER, DB_TITLE_HEADER, DB_QUANTITY_HEADER,
              DB_QUANTITY_SOLD_HEADER, DB_LISTING_STATUS_HEADER];
  var OPT = ['packageWeightOz', 'packageLengthIn', 'packageWidthIn', 'packageDepthIn',
             'pictureUrl1', 'pictureUrl2', 'pictureUrl3', 'pictureUrl4', 'pictureUrl5',
             'C:Part Type'];
  var r = MiSchema.readColumns(mi, NEED, { optional: OPT });
  var I = r.idx, out = {};

  for (var i = 0; i < r.rows.length; i++) {
    var row = r.rows[i];
    var sku = _phNormSku(row[I[DB_SKU_HEADER]]);
    if (!sku) continue;
    var pics = 0;
    for (var p = 1; p <= 5; p++) {
      var ix = I['pictureUrl' + p];
      if (ix >= 0 && String(row[ix] || "").trim()) pics++;
    }
    var qty = _phNum(row[I[DB_QUANTITY_HEADER]]) || 0;
    var sold = _phNum(row[I[DB_QUANTITY_SOLD_HEADER]]) || 0;
    out[sku] = {
      title:  String(row[I[DB_TITLE_HEADER]] || "").trim(),
      pt:     I['C:Part Type'] >= 0 ? String(row[I['C:Part Type']] || "").trim() : "",
      oz:     I['packageWeightOz'] >= 0 ? row[I['packageWeightOz']] : "",
      L:      I['packageLengthIn'] >= 0 ? row[I['packageLengthIn']] : "",
      W:      I['packageWidthIn']  >= 0 ? row[I['packageWidthIn']]  : "",
      D:      I['packageDepthIn']  >= 0 ? row[I['packageDepthIn']]  : "",
      photos: pics,
      sold:   sold,
      onHand: qty - sold,
      active: String(row[I[DB_LISTING_STATUS_HEADER]] || "").trim() === "Active"
    };
  }
  return out;
}

/**
 * PURE: joins the two sides into finished sheet rows. Split out from the I/O so the whole
 * join + classify + sort can be driven from Node against fixtures.
 * @return {Object} { rows: [[...]], counts: {...} }
 */
function _phBuildRows(truthBySku, miBySku, now) {
  var C = PRODUCT_HEALTH.cols, B = PRODUCT_HEALTH.bands;
  var seen = {}, recs = [];

  function push(sku) {
    if (seen[sku]) return;
    seen[sku] = true;
    var t = truthBySku[sku] || null, m = miBySku[sku] || null;
    // ⚠ An inactive listing with no truth record is an ended item, not a gap. Skipping
    //   them keeps the sheet about live decisions; MI keeps them for history.
    if (m && !m.active && !t) return;
    var v = _phClassify(t, m);
    recs.push({
      sku: sku, band: v.band, sold: (m && m.sold) || 0,
      cells: (function () {
        var a = [];
        a[C.SKU - 1] = sku;
        // ⚠ BRAND comes ONLY from By Part Type — it is the ENGINE MAKE the part fits.
        // Do NOT fall back to MI's `C:Brand`: that is "HQ" on 3,624 of 3,633 rows (it is
        // OUR manufacturer name, a different question wearing the same word). Measured
        // 2026-09-10 — the two agree on 0.1% of rows, so a fallback would fill this
        // column with noise. MI's real counterpart is `C:Compatible Equipment Make`
        // (93.5% agreement), and it is a MACHINE make vs an ENGINE make, so it is not a
        // mismatch worth reporting either — hence display only, no verdict.
        a[C.BRAND - 1] = (t && t.brand) || "";
        a[C.PART_TYPE - 1] = (t && t.pt) || (m && m.pt) || "";
        a[C.TITLE - 1] = (m && m.title) || (t && t.title) || "";
        a[C.TRUTH_WT - 1] = v.truthOz === null ? "" : v.truthOz;
        a[C.EBAY_WT - 1] = v.ebayOz === null ? "" : v.ebayOz;
        a[C.WT_DELTA - 1] = v.wtDelta === null ? "" : v.wtDelta;
        a[C.TRUTH_DIMS - 1] = v.truthDims;
        a[C.EBAY_DIMS - 1] = v.ebayDims;
        a[C.PHOTOS - 1] = m ? m.photos : "";
        a[C.SOLD - 1] = m ? m.sold : "";
        a[C.ON_HAND - 1] = m ? m.onHand : "";
        a[C.GAPS - 1] = v.gaps.join(" · ");
        a[C.EBAY_FIX - 1] = v.ebayFix;
        a[C.AMAZON - 1] = v.amazon;
        a[C.LAST_CHECKED - 1] = now;
        for (var i = 0; i < PRODUCT_HEALTH.dataWidth; i++) if (a[i] === undefined) a[i] = "";
        return a;
      })()
    });
  }

  for (var s in miBySku) push(s);
  for (var s2 in truthBySku) push(s2);   // the 365 not-on-eBay SKUs come in here

  recs.sort(_phCompareRows);

  var counts = { total: recs.length, under: 0, mismatch: 0, incomplete: 0, ready: 0, notListed: 0 };
  for (var i = 0; i < recs.length; i++) {
    if (recs[i].band === B.UNDER) counts.under++;
    else if (recs[i].band === B.MISMATCH) counts.mismatch++;
    else if (recs[i].band === B.INCOMPLETE) counts.incomplete++;
    else if (recs[i].band === B.READY) counts.ready++;
    else counts.notListed++;
  }
  return { rows: recs.map(function (r) { return r.cells; }), counts: counts };
}

/**
 * The one entry point. Reads both sources, rebuilds the sheet wholesale.
 * ⚠ Wholesale rebuild is deliberate: it is what makes the sheet SELF-CLEANING — a row that
 *   has been fixed simply does not come back. Nothing here is hand-edited, so there is no
 *   state to preserve (unlike OOS's FIRST SEEN or the Photo Queue's ✔ DONE).
 */
function refreshProductDataHealth() {
  var t0 = Date.now();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PRODUCT_HEALTH.sheetName);
  if (!sheet) { setupProductDataHealthSheet(); sheet = ss.getSheetByName(PRODUCT_HEALTH.sheetName); }

  // ⚠⚠ AUTO-MIGRATE A STALE LAYOUT. Adding a column (BRAND, 2026-09-10) makes every
  //    EXISTING sheet the wrong shape, and this function would happily write 16-wide rows
  //    under a 15-wide header — every label off by one from column B rightwards, and the
  //    CF still painting the pre-shift columns. Nothing would throw; the sheet would just
  //    quietly lie. A ONE-CELL probe on the header is enough to detect it, and setup is
  //    idempotent, so re-running costs a few style writes and no data (the rows below are
  //    cleared and rewritten by this function anyway).
  //    Same pattern as runKitHealthAudit's A2 probe and the OOS title-band migrator —
  //    it means the operator never has to remember to click "Re-style Sheet" first.
  var probe = String(sheet.getRange(PRODUCT_HEALTH.headerRow, PRODUCT_HEALTH.cols.BRAND)
                          .getValue() || "").trim();
  if (probe !== PRODUCT_HEALTH.headers[PRODUCT_HEALTH.cols.BRAND - 1]) {
    setupProductDataHealthSheet();
    sheet = ss.getSheetByName(PRODUCT_HEALTH.sheetName);
  }

  var truth;
  try {
    truth = _phReadTruth();
  } catch (e) {
    // ⚠ Fail LOUD and change nothing. A half-built health sheet is worse than a stale one,
    //   because someone would work from it believing the gaps had been recomputed.
    return "❌ Could not read By Part Type (" + e.message + "). Sheet left untouched. " +
           "Run probeByPartTypeRead() to diagnose.";
  }
  var mi = _phReadMi();
  var built = _phBuildRows(truth.bySku, mi, new Date());

  var last = sheet.getLastRow();
  if (last >= PRODUCT_HEALTH.dataStartRow) {
    sheet.getRange(PRODUCT_HEALTH.dataStartRow, 1,
                   last - PRODUCT_HEALTH.dataStartRow + 1, PRODUCT_HEALTH.dataWidth).clearContent();
  }
  if (built.rows.length) {
    var need = PRODUCT_HEALTH.dataStartRow + built.rows.length - 1;
    if (sheet.getMaxRows() < need) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows());
    sheet.getRange(PRODUCT_HEALTH.dataStartRow, 1, built.rows.length, PRODUCT_HEALTH.dataWidth)
      .setValues(built.rows);
  }
  SpreadsheetApp.flush();

  var c = built.counts;
  var msg = "✅ Product Data Health · " + c.total + " rows in " +
    ((Date.now() - t0) / 1000).toFixed(1) + "s\n" +
    "   ⚠ UNDER-CHARGING : " + c.under + "   (fix these first — costing money now)\n" +
    "   ⚠ mismatch       : " + c.mismatch + "\n" +
    "   ○ incomplete     : " + c.incomplete + "   (worklist, ranked by units sold)\n" +
    "   ✓ AMAZON READY   : " + c.ready + "\n" +
    "   — not on eBay    : " + c.notListed + "   (expansion candidates)\n" +
    "   truth file read  : " + truth.sheets + " sheets, " +
    (truth.ms / 1000).toFixed(1) + "s";
  console.log(msg);
  return msg;
}

/** Cheap snapshot read for the sidebar badge — counts the sheet, never re-audits. */
function getProductHealthCounts() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(PRODUCT_HEALTH.sheetName);
    if (!sheet) return { under: 0, incomplete: 0, ready: 0 };
    var last = sheet.getLastRow();
    if (last < PRODUCT_HEALTH.dataStartRow) return { under: 0, incomplete: 0, ready: 0 };
    var n = last - PRODUCT_HEALTH.dataStartRow + 1;
    var fix = sheet.getRange(PRODUCT_HEALTH.dataStartRow, PRODUCT_HEALTH.cols.EBAY_FIX, n, 1).getValues();
    var amz = sheet.getRange(PRODUCT_HEALTH.dataStartRow, PRODUCT_HEALTH.cols.AMAZON, n, 1).getValues();
    var under = 0, ready = 0, incomplete = 0;
    for (var i = 0; i < n; i++) {
      if (String(fix[i][0] || "").indexOf("LIGHT") !== -1) under++;
      var a = String(amz[i][0] || "");
      if (a === "READY") ready++;
      else if (a && a !== "not listed on eBay") incomplete++;
    }
    return { under: under, incomplete: incomplete, ready: ready };
  } catch (e) {
    return { under: 0, incomplete: 0, ready: 0 };
  }
}
