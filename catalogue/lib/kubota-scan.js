// Scanned Kubota "Spare Parts List" — OCR reader (2026-10-03).
//
// Same book layout as kubota.js, but the PDF has no text layer. Each page is rendered at
// 400 dpi and read with tesseract (TSV = every word with its box). Rows are found by POSITION,
// not by text layout — table rules turn into "|" noise and break a plain-text read:
//   * a part number is a word in the left part-number column that fits the Kubota shape
//     (after fixing OCR's usual letter/digit swaps in the digit positions);
//   * the REF is the word just left of it; name / qty / remark are the words on the same row,
//     sorted into columns by the header words (DESIGNATION · STUECK · REMARKS).
// Output has the same shape as parseKubota(), plus per-part `ocr: {conf}` so the report and
// the page can tell a read number from a printed one.

"use strict";

const { PN_SHAPE } = require("./kubota");

/** Parse tesseract TSV into words with boxes. */
function tsvWords(tsv) {
  return String(tsv).split("\n").slice(1).map(l => l.split("\t")).filter(c => c.length >= 12 && c[11].trim())
    .map(c => ({ block: +c[2], par: +c[3], line: +c[4], x: +c[6], y: +c[7], w: +c[8], h: +c[9], conf: +c[10],
                 t: c[11].trim() }));
}

// Digit positions of a Kubota number, after the 5-char prefix: OCR swaps these letters in.
const TO_DIGIT = { O: "0", Q: "0", D: "0", U: "0", I: "1", L: "1", J: "1", T: "1", Z: "2", S: "5", B: "8", G: "6" };

/**
 * Turn an OCR'd token into a Kubota part number, or null.
 * Accepts "15471-3501-Z", "|1A021-3515-0", "154713501 2", "1547135012".
 */
function normalisePn(raw) {
  // a number OCR glued to the next word ("17331-2105-0_[ASSY"): keep the leading number-shaped part
  const lead = String(raw).toUpperCase().match(/^[|\[\](]*([0-9A-Z]{5}[-~—–_]?[0-9A-Z]{4}[-~—–_]?[0-9A-Z])(?=$|[^0-9A-Z])/);
  if (lead && lead[1] !== String(raw).toUpperCase()) { const n = normalisePn(lead[1]); if (n) return n; }
  let s = String(raw).toUpperCase().replace(/[|\[\]{}()!'"`.,;:]/g, "").replace(/[—–_~]/g, "-");
  const flat = s.replace(/-/g, "");
  if (!/^[0-9A-Z]{10}$/.test(flat)) return null;
  const head = flat.slice(0, 5);
  let tail = flat.slice(5).split("").map(ch => (/\d/.test(ch) ? ch : TO_DIGIT[ch] || ch)).join("");
  if (!/^\d+$/.test(tail)) return null;
  if (!/\d/.test(head)) return null;                       // every Kubota prefix holds a digit
  // 10 characters fit two shapes: 5-4-1 (1G780-2111-2) and the 5-5 hardware form (07715-00401).
  // The dashes OCR kept decide; without them, Kubota hardware numbers start with "0".
  let fiveFive;
  if (/^[0-9A-Z]{5}-[0-9A-Z]{5}$/.test(s)) fiveFive = true;
  else if (/^[0-9A-Z]{5}-?[0-9A-Z]{4}-[0-9A-Z]$/.test(s)) fiveFive = false;
  else fiveFive = head[0] === "0";
  if (flat.length !== 10) return null;
  const pn = fiveFive ? head + "-" + tail : head + "-" + tail.slice(0, 4) + "-" + tail.slice(4);
  return PN_SHAPE.test(pn) ? pn : null;
}

/**
 * Read one OCR'd page.
 * @param {Array} words   tsvWords() of the page
 * @param {number} W, H   image size in px
 * @return {{code,name,model,rows,band,isIndex}|null}  null = no parts table on this page
 */
// A section code as OCR reads it → the clean code, or "" if the word isn't one.
//   "0102" (older books) · "E03." · "E02-1." (newer, sub-part) — measured 2026-10-04 on the
//   newer bilingual books: "EO2-1." (O for 0), "ES1." (S for 5), "£11-1." all occur, and the
//   old exact-digit test rejected every one, so whole books came back as 1–7 sections.
// ⚠ At least one REAL digit is required, or top-left words like "BOS" would pass.
function sectionCodeOf(t) {
  const s = String(t || "").trim();
  if (/^\d{4}$/.test(s)) return s;
  const m = s.match(/^([A-Z£€])([0-9OSIlBZ]{2,3})(?:-([0-9Il]{1,2}))?\.?$/);
  if (!m || !/\d/.test(m[2])) return "";
  const fix = x => x.replace(/O/g, "0").replace(/S/g, "5").replace(/[Il]/g, "1").replace(/B/g, "8").replace(/Z/g, "2");
  return (/[£€]/.test(m[1]) ? "E" : m[1]) + fix(m[2]) + (m[3] ? "-" + fix(m[3]) : "");
}

// The section code + English title at the top-left of a page — on a parts TABLE page, or on
// the DRAWING page before it. ⚠ Newer books print the big "E07." only on the drawing page, and
// the table that follows often doesn't repeat it (measured on D1302, 2026-10-04) — so this runs
// on every page and the book loop carries a drawing page's code onto the next table.
function pageSection(words, W, H) {
  // section code at the top-left: "0102" (older books) or "E03." (newer). The English title is
  // the TOPMOST line beside it: above the code line in the first style, on it in the second.
  // (OCR reads the large "E" of "E03." as £ or €, and newer books add a sub-part: "E02-1.")
  const code = words.find(w => sectionCodeOf(w.t) && w.y < 0.12 * H && w.x < 0.2 * W);
  if (code) code.t = sectionCodeOf(code.t);
  let name = "";
  if (code) {
    const beside = words.filter(w => w.x > code.x + code.w && w.y > code.y - 2.4 * code.h && w.y < code.y + 0.6 * code.h && w.y < 0.15 * H);
    if (beside.length) {
      const topY = Math.min(...beside.map(w => w.y));
      name = beside.filter(w => w.y < topY + 0.6 * Math.max(...beside.map(x => x.h))).sort((a, b) => a.x - b.x).map(w => w.t).join(" ").trim();
    }
  }
  return { code: code ? code.t.replace(/\.$/, "") : null, name, codeWord: code };
}

function readScanPage(words, W, H) {
  const txt = words.map(w => w.t).join(" ");
  const hdr = words.filter(w => /^(REFERENCE|BESELL|BESTELL|DESIGNAT|BEZE|STUECK|REMARK|REMARQ)/.test(w.t));
  // ⚠ the CONTENTS page also says "NUMERICAL INDEX" — only a page with no parts header that is
  // mostly part numbers is the index (the caller also requires parts to have been read first)
  const pnCount = words.filter(w => normalisePn(w.t)).length;
  // ⚠⚠ The index has its OWN column headings — four "PART No. / REFERENCE" columns side by side —
  //   so "fewer than 2 headings" missed it in most books and the whole index (200–380 lines) was
  //   swallowed into the last section as fake parts (measured 2026-10-04). Now: the index TITLE
  //   in the top of the page, or 3+ REFERENCE/BESTELL headings spread across the width.
  const titleIdx = words.some(w => /^NUMER(ICAL|IQUE|ISCHEN|QUE)?$/.test(w.t.replace(/[^A-Z]/g, "")) && w.y < 0.15 * H);
  const refCols = words.filter(w => /^(REFERENCE|BESTELL)/.test(w.t))
    .map(w => Math.round(w.x / (0.1 * W))).filter((v, i, a) => a.indexOf(v) === i).length;
  const isIndex = pnCount > 20 && (titleIdx || refCols >= 3 ||
                  (/NUMERICAL|NUMERIQUE|NUMERISCHEN/.test(txt) && hdr.length < 2));
  const sec = pageSection(words, W, H);
  if (!hdr.length || isIndex) return { isIndex, rows: [], code: isIndex ? null : sec.code, name: sec.name };
  const code = sec.codeWord, name = sec.name;
  const hw = re => words.find(x => re.test(x.t) && hdr.some(h => Math.abs(h.y - x.y) < 140));
  const centre = w => w ? w.x + w.w / 2 : null;
  const nameW = words.find(x => x.t === "NAME" && hdr.some(h => Math.abs(h.y - x.y) < 140));
  const desW = hw(/^DESIGNAT/), reW = hw(/^REMARK|^REMARQ/), bezW = hw(/^BEZE/);
  // the qty column's heading: "STUECK" in most books, "UNIT / UNITE / ANZAHL" or "Q'TY" in others
  // (Z400, 2026-10-04 — without it the column was guessed at 0.69 W and every qty was missed).
  // Columns are left-aligned under the heading, so take the LEFTMOST of the heading words.
  // ⚠ tolerant of OCR misreads of a tiny heading: "STUEGK/S." (D782, 2026-10-04)
  const stCands = words.filter(x => /^(STU?E?[CGK]{1,2}K?\b|ST[UÜ]E?[CG]K|ANZAHL|UNITE?$|Q'?T[YE])/.test(x.t) && hdr.some(h => Math.abs(h.y - x.y) < 140));
  const stW = stCands.length ? stCands.reduce((a, c) => c.x < a.x ? c : a) : null;
  // and it ends where the interchangeability / serial-number columns start, if the book has them
  const intW = hw(/^INTERCHANG|^REVISION/);
  // columns are left-aligned under centred headings: a boundary sits halfway between headings
  const nameEnd = (nameW && desW) ? (centre(nameW) + centre(desW)) / 2 : 0.36 * W;
  const stueckX = stW ? stW.x - 40 : 0.69 * W;
  const remarksX = reW ? reW.x - 120 : 0.83 * W;
  const headerTop = Math.min(...hdr.map(w => w.y));
  const headerBottom = Math.max(...hdr.map(w => w.y + w.h)) + 60;
  const footer = words.find(w => /^Interchang/i.test(w.t) && w.y > 0.5 * H);
  const tableBottom = footer ? footer.y - 10 : 0.95 * H;
  // ⚠ capped at 4.5% of the width: the column is narrow and the serial-range arrows / brackets
  //   of the next column otherwise ride along and spoil the read (Z400, 2026-10-04)
  const qtyEnd = Math.min(remarksX - 20, stueckX + 0.045 * W, intW && intW.x > stueckX + 60 ? intW.x - 20 : Infinity);
  // the qty heading's CENTRE: in books with model columns it's centred over A…D, so its left edge
  // sits in B — the variant reader lays the sub-columns out around this instead (D1703, 2026-10-04)
  const stCenter = stCands.length ? (Math.min(...stCands.map(w => w.x)) + Math.max(...stCands.map(w => w.x + w.w))) / 2 : null;
  const layout = { pnX0: 0.06 * W, pnX1: nameW ? nameW.x - 120 : 0.25 * W, nameEnd, stueckX, qtyEnd, stCenter, remarksX, headerBottom, tableBottom, W, H };

  const modelW = words.find(w => /^[A-Z]:[A-Z0-9][A-Z0-9-]{3,}/.test(w.t) && w.y < headerTop + 10);

  // rows: a part number in the left column
  const pnWords = words.filter(w => w.y > headerBottom - 40 && w.y < tableBottom && w.x < 0.32 * W && w.x > 0.06 * W)
    .map(w => ({ w, pn: normalisePn(w.t) })).filter(x => x.pn);
  const rows = pnWords.map(({ w, pn }) => rowAt(words, w, pn, layout));

  // drawing band: below the title block, above the model/qty lines over the table (px)
  const bandTop = code ? code.y + 2.4 * code.h : 0.08 * H;
  const bandBottom = (modelW ? modelW.y : headerTop) - 30;
  const band = bandBottom - bandTop > 0.12 * H ? { y: bandTop, h: bandBottom - bandTop } : null;

  return { isIndex, code: sec.code, name, model: modelW ? modelW.t.slice(2) : "", rows, band, layout };
}

// The French column often spills into the English name on a scan ("KEY,FEATHER CLAVETTE"):
// cut at the first word that only ever starts a French designation.
const FRENCH = new Set(["ENS", "ENS.", "JOINT", "CLAVETTE", "ECROU", "FCROU", "PIGNON", "COUSSINET", "RONDELLE", "BAGUE",
  "VIS", "MIS", "AXE", "CIRCLIP", "CIRGLIP", "BIELLE", "SEGMENT", "TUYAU", "BOUCHON", "COLLIER", "CARTER", "POMPE",
  "VILEBREQUIN", "DEFLECTEUR", "DEFLEGTEUR", "GOUJON", "RESSORT", "SOUPAPE", "CULBUTEUR", "ARBRE", "PLAQUE", "LEVIER",
  "CONTACT", "BRIDE", "DEMARREUR", "ALTERNATEUR", "VOLANT", "COUVERCLE", "CACHE", "TIGE", "GOUPILLE", "ANNEAU",
  "COUDE", "SUPPORT", "ROULEMENT", "COLLECTEUR", "FILTRE", "CARTOUCHE", "INJECTEUR", "BOUGIE", "POULIE", "COURROIE",
  "VENTILATEUR", "TUBE", "RACCORD", "ETIQUETTE", "MANUEL", "CALE", "CLIP", "ATTACHE", "AGRAFE"]);

const NOISE = /^[|\[\]_~.,'`"\-=*]+$/;
const clean = t => t.replace(/^[|\[\]_~]+|[|\[\]_~]+$/g, "").trim();

/** Build a row around a part-number word: REF to its left, name up to the name column's end. */
function rowAt(words, w, pn, L) {
  const mid = w.y + w.h / 2, tol = 0.6 * Math.max(w.h, 40);
  const line = words.filter(x => Math.abs((x.y + x.h / 2) - mid) < tol).sort((a, b) => a.x - b.x);
  const refW = line.filter(x => x.x < w.x).map(x => x.t.replace(/\D/g, "")).find(t => /^\d{3}$/.test(t));
  let nameWords = line.filter(x => x.x > w.x + w.w - 5 && x.x < L.nameEnd).map(x => clean(x.t)).filter(t => t && !NOISE.test(t));
  const cut = nameWords.findIndex((t, i) => i > 0 && FRENCH.has(t.toUpperCase().replace(/[^A-Z.]/g, "")));
  if (cut > 0) nameWords = nameWords.slice(0, cut);
  return { ref: refW || "", pn, name: nameWords.join(" ").replace(/\s+,/g, ",").replace(/,\s+/g, ","),
           remark: "", qty: null, conf: Math.round(w.conf), y: mid, h: Math.max(w.h, 40) };
}

/** Merge column re-reads into rows by vertical position (y in the same 400-dpi page px). */
function attachColumn(rows, colWords, field) {
  // Each row owns the band from halfway-to-the-previous row to halfway-to-the-next one: in
  // these books the qty sits at the top of its cell and the size at the bottom.
  const sorted = rows.slice().sort((a, b) => a.y - b.y);
  const band = i => ({
    lo: i ? (sorted[i - 1].y + sorted[i].y) / 2 : sorted[i].y - sorted[i].h,
    hi: i < sorted.length - 1 ? (sorted[i].y + sorted[i + 1].y) / 2 : sorted[i].y + sorted[i].h * 1.2 });
  const lines = {};
  colWords.forEach(c => { const k = c.block + ":" + c.par + ":" + c.line; (lines[k] = lines[k] || []).push(c); });
  Object.values(lines).forEach(ws => {
    const y = ws.reduce((a, c) => a + c.y + c.h / 2, 0) / ws.length;
    const i = sorted.findIndex((r, j) => { const bd = band(j); return y >= bd.lo && y < bd.hi; });
    if (i < 0) return;
    const row = sorted[i];
    const t = ws.sort((a, b) => a.x - b.x).map(c => clean(c.t)).filter(x => x && !NOISE.test(x)).join(" ");
    if (field === "qty") {
      // the first small number; serial-number ranges ("489911") print under the qty in newer books
      // read from the START of a word: an arrow glued on reads as "1-" (Z400, 2026-10-04)
      const m = t.split(/\s+/).map(x => x.match(/^(\d{1,3})(?!\d)/)).find(Boolean);
      if (m && row.qty == null) row.qty = [Number(m[1])];
    } else if (field === "name") {
      const nm = cleanName(t);
      if (nm && (!row.name || row.rescued || nm.length >= row.name.length - 2)) row.name = nm;
    } else if (!row.remark) row.remark = fixRemark(t);
  });
}

/** A name from the name-column crop: OCR noise out, the French spill cut. */
function cleanName(t) {
  let w = t.split(/\s+/).map(clean).filter(x => x && !NOISE.test(x));
  const cut = w.findIndex((x, i) => i > 0 && FRENCH.has(x.toUpperCase().replace(/[^A-Z.]/g, "")));
  if (cut > 0) w = w.slice(0, cut);
  return w.join(" ").replace(/\s+,/g, ",").replace(/,\s+/g, ",").trim();
}

/** OCR of the tiny remarks: "+0.25mm", "-0.20mm SET", "STD", "STD SET". */
function fixRemark(t) {
  let s = t.toUpperCase().replace(/\s+/g, " ").trim();
  s = s.replace(/^(5TD|ST0|SID|STO|S7D)/, "STD");
  s = s.replace(/^([+-]?)\s*O(?=[.,]?\d)/, "$10").replace(/(\d)\s*[.,]\s*(\d)/g, "$1.$2");
  s = s.replace(/(\d)\s*MM(?=\s|SET|$)/g, "$1mm").replace(/\bMM\b/g, "mm").replace(/(\S)SET$/, "$1 SET");
  // "0.50mm" printed with its sign: a size with no sign is an OCR drop — keep it, flagged by the reader
  return s;
}

/**
 * Second, independent read of the part-number column (REF and PN come out fused:
 * "10007715-00401"). It does two jobs:
 *   * rescue rows the page read missed entirely;
 *   * mark each row `agree` when both reads give the same number — a far stronger signal
 *     than tesseract's own confidence.
 */
// A part number taken BY ITS SHAPE out of a read that ran into the next column
// ("02771-50120NUTFL", "1C011-5510-4HOS"): the dashed Kubota shapes first, then loose.
function pnByShape(t) {
  const u = String(t || "").toUpperCase();
  const m = u.match(/(\d{5}-\d{4}-\d|[0-9][A-Z0-9]\d{3}-\d{4}-\d|\d{5}-\d{5})/) ||
            u.match(/(\d{5}-?\d{4}-?\d|[0-9][A-Z0-9]\d{3}-?\d{4}-?\d|\d{5}-?\d{5})/);
  return m ? normalisePn(m[1]) : null;
}

function rescueRows(rows, words, pnColWords, L) {
  const reads = pnColWords.map(c => {
    const m = String(c.t).toUpperCase().match(/^[|\[(]*(\d{3})?(.+)$/);
    let pn = m ? normalisePn(m[2]) : null, ref = m && m[1];
    if (!pn) { pn = normalisePn(c.t); ref = ""; }
    // ⚠⚠ the part-number column read runs into the NAME column, so most of its words were
    //   "02771-50120NUTFL": no shape → no pn → the two reads NEVER agreed (0 on every page tested,
    //   2026-10-04) and every OCR line showed "check". Take the number by its shape.
    if (!pn) { pn = m ? pnByShape(m[2]) : null; ref = m && m[1] || ""; }
    if (!pn) { pn = pnByShape(c.t); ref = ""; }
    return pn ? { c, pn, ref: ref || "", mid: c.y + c.h / 2 } : null;
  }).filter(Boolean);
  rows.forEach(r => {
    const twin = reads.find(x => Math.abs(x.mid - r.y) < r.h * 0.6);
    // ⚠ OR, never overwrite: the gap filler calls this a second time with only its few cell reads,
    //   and a plain assignment wiped every agreement the column read had found (2026-10-04)
    r.agree = !!r.agree || !!(twin && twin.pn === r.pn);
    if (twin && !r.ref && twin.ref) r.ref = twin.ref;
  });
  const added = [];
  reads.forEach(x => {
    if (rows.some(r => Math.abs(r.y - x.mid) < r.h * 0.6)) return;
    const row = rowAt(words, Object.assign({}, x.c, { x: x.c.x + x.c.w * 0.25 }), x.pn, L);
    if (!row.ref && x.ref) row.ref = x.ref;
    row.agree = false; row.rescued = true;
    added.push(row);
  });
  return rows.concat(added).sort((a, b) => a.y - b.y);
}

module.exports = { pnByShape, sectionCodeOf, tsvWords, normalisePn, readScanPage, attachColumn, rescueRows, fixRemark };
