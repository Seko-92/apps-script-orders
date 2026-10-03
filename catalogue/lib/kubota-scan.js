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
function readScanPage(words, W, H) {
  const txt = words.map(w => w.t).join(" ");
  const isIndex = /NUMERICAL|NUMERIQUE|NUMERISCHEN/.test(txt) && !/STUECK/.test(txt);
  const hdr = words.filter(w => /^(REFERENCE|BESELL|DESIGNAT|BEZE|STUECK|REMARK|REMARQ)/.test(w.t));
  if (!hdr.length || isIndex) return { isIndex, rows: [] , code: null };
  const hw = re => words.find(x => re.test(x.t) && hdr.some(h => Math.abs(h.y - x.y) < 140));
  const centre = w => w ? w.x + w.w / 2 : null;
  const nameW = words.find(x => x.t === "NAME" && hdr.some(h => Math.abs(h.y - x.y) < 140));
  const desW = hw(/^DESIGNAT/), stW = hw(/^STUECK/), reW = hw(/^REMARK|^REMARQ/), bezW = hw(/^BEZE/);
  // columns are left-aligned under centred headings: a boundary sits halfway between headings
  const nameEnd = (nameW && desW) ? (centre(nameW) + centre(desW)) / 2 : 0.36 * W;
  const stueckX = stW ? stW.x - 40 : 0.69 * W;
  const remarksX = reW ? reW.x - 120 : 0.83 * W;
  const headerTop = Math.min(...hdr.map(w => w.y));
  const headerBottom = Math.max(...hdr.map(w => w.y + w.h)) + 60;
  const footer = words.find(w => /^Interchang/i.test(w.t) && w.y > 0.5 * H);
  const tableBottom = footer ? footer.y - 10 : 0.95 * H;
  const layout = { pnX0: 0.06 * W, pnX1: nameW ? nameW.x - 120 : 0.25 * W, nameEnd, stueckX, remarksX, headerBottom, tableBottom, W, H };

  // section code (4 digits, top-left) and the English title on the line above it
  const code = words.find(w => /^\d{4}$/.test(w.t) && w.y < 0.12 * H && w.x < 0.2 * W);
  let name = "";
  if (code) {
    name = words.filter(w => w.x > code.x + code.w && w.y < code.y - 0.25 * code.h && w.y > code.y - 2.2 * code.h)
      .sort((a, b) => a.x - b.x).map(w => w.t).join(" ").trim();
  }
  const modelW = words.find(w => /^[A-Z]:[A-Z0-9][A-Z0-9-]{3,}/.test(w.t) && w.y < headerTop + 10);

  // rows: a part number in the left column
  const pnWords = words.filter(w => w.y > headerBottom - 40 && w.y < tableBottom && w.x < 0.32 * W && w.x > 0.06 * W)
    .map(w => ({ w, pn: normalisePn(w.t) })).filter(x => x.pn);
  const rows = pnWords.map(({ w, pn }) => rowAt(words, w, pn, layout));

  // drawing band: below the title block, above the model/qty lines over the table (px)
  const bandTop = code ? code.y + 2.4 * code.h : 0.08 * H;
  const bandBottom = (modelW ? modelW.y : headerTop) - 30;
  const band = bandBottom - bandTop > 0.12 * H ? { y: bandTop, h: bandBottom - bandTop } : null;

  return { isIndex, code: code ? code.t : null, name, model: modelW ? modelW.t.slice(2) : "", rows, band, layout };
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
    if (field === "qty") { const n = t.match(/\d+/); if (n && row.qty == null) row.qty = [Number(n[0])]; }
    else if (!row.remark) row.remark = fixRemark(t);
  });
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
function rescueRows(rows, words, pnColWords, L) {
  const reads = pnColWords.map(c => {
    const m = String(c.t).toUpperCase().match(/^[|\[(]*(\d{3})?(.+)$/);
    let pn = m ? normalisePn(m[2]) : null, ref = m && m[1];
    if (!pn) { pn = normalisePn(c.t); ref = ""; }
    return pn ? { c, pn, ref: ref || "", mid: c.y + c.h / 2 } : null;
  }).filter(Boolean);
  rows.forEach(r => {
    const twin = reads.find(x => Math.abs(x.mid - r.y) < r.h * 0.6);
    r.agree = !!(twin && twin.pn === r.pn);
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

module.exports = { tsvWords, normalisePn, readScanPage, attachColumn, rescueRows, fixRemark };
