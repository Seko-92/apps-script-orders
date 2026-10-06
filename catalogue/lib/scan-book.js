// OCR a whole scanned Kubota book → the same shape parseKubota() returns.
// One page at a time: pdftoppm (400 dpi, grey) → tesseract TSV → readScanPage().
"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");
const { tsvWords, readScanPage, attachColumn, rescueRows, sectionCodeOf } = require("./kubota-scan");

const DPI = 400;

// ⭐ Page rasters are rendered ONCE per (page, dpi) and every crop is sliced from them in Node.
//   A `pdftoppm -x -y -W -H` crop decodes the WHOLE scanned page every time (~0.6 s), and a page
//   needs up to ~290 crops (per-row qty cells × variants) — measured 2026-10-05: pdftoppm was
//   65–75% of the run (D1703 p30: 185 s of 255 s). The slice is pixel-identical to the crop
//   (verified: 0 px differ), so the reads don't change. Only the current page is kept.
let rasterKey = "", rasters = {};
function pageRaster(pdf, page, dpi, tmp) {
  const key = pdf + "|" + page;
  if (key !== rasterKey) { rasterKey = key; rasters = {}; }
  if (rasters[dpi]) return rasters[dpi];
  const base = path.join(tmp, "r" + page + "-" + dpi);
  execFileSync("pdftoppm", ["-f", String(page), "-l", String(page), "-r", String(dpi), "-gray", "-singlefile", pdf, base],
    { stdio: "ignore" });
  const b = fs.readFileSync(base + ".pgm"); fs.unlinkSync(base + ".pgm");
  const m = b.toString("latin1", 0, 64).match(/^P5\s+(\d+)\s+(\d+)\s+(\d+)\s/);
  return (rasters[dpi] = { W: +m[1], H: +m[2], d: b.subarray(m[0].length) });
}
function dropRasters() { rasterKey = ""; rasters = {}; }

/** Write the crop (x, y, w, h in px AT `dpi`, as pdftoppm's -x -y -W -H) to `base`.pgm. */
function cropPgm(pdf, page, dpi, x, y, w, h, base, tmp) {
  x = Math.round(x); y = Math.round(y); w = Math.round(w); h = Math.round(h);
  if (w < 1 || h < 1) {   // pdftoppm reads -W 0 as "whole width" — keep that exact behaviour
    execFileSync("pdftoppm", ["-f", String(page), "-l", String(page), "-r", String(dpi), "-gray",
      "-x", String(x), "-y", String(y), "-W", String(w), "-H", String(h), "-singlefile", pdf, base], { stdio: "ignore" });
    return;
  }
  const R = pageRaster(pdf, page, dpi, tmp);
  const x0 = Math.max(0, Math.min(x, R.W - 1)), y0 = Math.max(0, Math.min(y, R.H - 1));
  const cw = Math.max(1, Math.min(x + w, R.W) - x0), ch = Math.max(1, Math.min(y + h, R.H) - y0);
  const body = Buffer.alloc(cw * ch, 255);
  for (let r = 0; r < ch; r++) R.d.copy(body, r * cw, (y0 + r) * R.W + x0, (y0 + r) * R.W + x0 + cw);
  fs.writeFileSync(base + ".pgm", Buffer.concat([Buffer.from("P5\n" + cw + " " + ch + "\n255\n", "latin1"), body]));
}

function ocrPage(pdf, page, tmp) {
  const R = pageRaster(pdf, page, DPI, tmp), W = R.W, H = R.H;
  const img = path.join(tmp, "p" + page + ".pgm");
  fs.writeFileSync(img, Buffer.concat([Buffer.from("P5\n" + W + " " + H + "\n255\n", "latin1"), R.d]));
  const tsv = execFileSync("tesseract", [img, "-", "--psm", "6", "tsv"],
    { maxBuffer: 32 << 20, stdio: ["ignore", "pipe", "ignore"], env: Object.assign({}, process.env, { OMP_THREAD_LIMIT: "1" }) }).toString();
  fs.unlinkSync(img);
  return { words: tsvWords(tsv), W, H };
}

/**
 * OCR one column of the page (box in 400-dpi page px) at `dpi`, returning words with
 * coordinates mapped back to 400-dpi page px so they line up with the page read.
 */
// The big section code ("E07.", "E14-1.") is printed so large that the full-page read at DPI
// breaks it into fragments ("EO" + "7"). A small LOW-resolution crop of just that corner, one
// text line, restricted to the characters a code can hold, reads it whole (D1302, 2026-10-04).
// Only called when the page read found no code. Returns the clean code or "".
function ocrSectionCode(pdf, page, W, H, tmp) {
  const dpi = 150, k = dpi / DPI, base = path.join(tmp, "s" + page);
  try {
    // starts at 2% of the width: the page border / binding line at the very edge reads as "E" or "|"
    cropPgm(pdf, page, dpi, 0.02 * W * k, 0, 0.15 * W * k, 0.085 * H * k, base, tmp);
    const img = base + ".pgm";
    // tesseract is mode-sensitive on one big word: measured on D1302, psm 6 reads "E13." where
    // psm 7 returns nothing — so try 6, 11, 8 in turn and keep the first real code
    let found = "";
    for (const psm of ["6", "11", "8"]) {
      const txt = execFileSync("tesseract", [img, "-", "--psm", psm, "-c", "tessedit_char_whitelist=EO0123456789-."],
        { stdio: ["ignore", "pipe", "ignore"], env: Object.assign({}, process.env, { OMP_THREAD_LIMIT: "1" }) }).toString();
      // word by word, last first — a stray "E" from a rule must not glue onto the real code
      for (const t of txt.split(/\s+/).filter(Boolean).reverse()) {
        const c = sectionCodeOf(t.replace(/^EE/, "E"));
        if (c && /^E/.test(c)) { found = c; break; }
      }
      if (found) break;
    }
    fs.unlinkSync(img);
    return found;
  } catch (e) { return ""; }
}

// Vintage books (D850, D1402) head a section with just a big "2." — the page read fragments it
// ("I-", "§", "1" for 3). Re-read only the corner where it sits, digits only. Returns "V02" or "".
// ⚠ The caller enables this only while the book has shown no E-code / 4-digit code, so an older
//   book's "0207" can never be mistaken for a vintage "02" (and "020"/"0207" fail the 1–2 digit test).
function ocrVintageCode(pdf, page, W, H, tmp) {
  const dpi = 150, k = dpi / DPI, base = path.join(tmp, "v" + page);
  try {
    cropPgm(pdf, page, dpi, 0.012 * W * k, 0, 0.055 * W * k, 0.09 * H * k, base, tmp);
    const img = base + ".pgm";
    let found = "";
    for (const psm of ["6", "8", "10"]) {
      const txt = execFileSync("tesseract", [img, "-", "--psm", psm, "-c", "tessedit_char_whitelist=0123456789."],
        { stdio: ["ignore", "pipe", "ignore"], env: Object.assign({}, process.env, { OMP_THREAD_LIMIT: "1" }) }).toString().trim();
      const m = txt.match(/^(\d{1,2})\.?$/);
      if (m && Number(m[1]) > 0) { found = "V" + m[1].padStart(2, "0"); break; }
    }
    fs.unlinkSync(img);
    return found;
  } catch (e) { return ""; }
}

// The English title beside a code we read from the corner crop: the topmost line of words just
// right of the code, near the top of the page.
function titleBeside(words, W, H) {
  const near = words.filter(w => w.x > 0.13 * W && w.x < 0.55 * W && w.y < 0.07 * H && /[A-Z]{2}/.test(w.t));
  if (!near.length) return "";
  const topY = Math.min(...near.map(w => w.y)), h = Math.max(...near.map(w => w.h));
  return near.filter(w => w.y < topY + 0.6 * h).sort((a, b) => a.x - b.x).map(w => w.t).join(" ").trim();
}

// The page's model columns from the header line "A:D1703-BB-EC-1, B:…,…, C:…" → [{col, name}].
// A letter's name runs until the next "X:" word (B lists two models here, comma-separated).
function pageModels(words) {
  const hits = words.filter(w => /^[A-H]:[A-Z0-9]/.test(w.t)).sort((a, b) => a.y - b.y || a.x - b.x);
  if (!hits.length) return [];
  const line = words.filter(w => Math.abs(w.y - hits[0].y) < hits[0].h).sort((a, b) => a.x - b.x).map(w => w.t).join(" ");
  const out = [], re = /\b([A-H]):(\S[^]*?)(?=\s+[A-H]:|$)/g; let m;
  while ((m = re.exec(line))) {
    const name = m[2].replace(/[,\s]+$/, "").replace(/,(?=\S)/g, ", ").trim();
    if (!out.some(x => x.col === m[1])) out.push({ col: m[1], name });
  }
  return out;
}

// Model sub-columns A…n: equal-width gaps between vertical rules, CENTRED under the qty heading.
// One wide column read (from left of the heading to the remarks) supplies the rules. Returns
// [[x0,x1], …] for A…n, or null. Missing dividers are filled in from the found width.
function variantColumns(pdf, page, qtyBox, L, n, W, tmp) {
  const x0 = qtyBox.x - 0.08 * W, x1 = Math.min(L.remarksX - 20, qtyBox.x + 0.25 * W);
  const col = ocrColumn(pdf, page, { x: x0, y: qtyBox.y, w: x1 - x0, h: qtyBox.h }, 600, tmp, ["-c", "tessedit_char_whitelist=0123456789-"], "qty");
  // rules from the OCR crop PLUS a direct scan of the page image: a printed rule is a pixel column dark
  // down most of the table. The crop read found only 2 of D1703 p29's 5 rules (its rows are broken
  // lines), which left the layout to the heading centre → one column off.
  const rules = (col.rules || []).concat(pageVRules(pdf, page, x0, Math.max(x1, Math.min(L.remarksX || x1, qtyBox.x + 0.3 * W)), qtyBox.y, qtyBox.y + qtyBox.h, tmp))
    .sort((a, b) => a - b).filter((x, i, a) => !i || x - a[i - 1] > 0.006 * W);
  const gaps = [];
  for (let i = 0; i < rules.length - 1; i++) { const d = rules[i + 1] - rules[i]; if (d > 0.012 * W && d < 0.06 * W) gaps.push([rules[i], rules[i + 1]]); }
  if (!gaps.length) return null;
  const centre = L.stCenter || qtyBox.x + 0.02 * W;
  // candidate layouts: each found gap as column A, B, C or D, the rest laid out at that width.
  // ⭐ Keep the one whose n+1 EDGES land on the most printed rules; the heading centre only breaks a
  //   tie. Centring alone put D1703 p23 one column LEFT: the rules 2125·2273·2429·2575·2725 were all
  //   found, but "Q'TE STUECK" heads A–B only, so its centre (2334) pulled the layout left — every
  //   variant read its left neighbour's quantity (16423-2111-0 A3 B3 C3 D– came out –,3,3,3).
  const onRule = x => rules.some(r => Math.abs(r - x) < 0.006 * W);
  let best = null;
  gaps.forEach(g => {
    const w = g[1] - g[0];
    [0, 1, 2, 3].slice(0, n).forEach(k => {           // g could be column A, B, C or D
      const a = g[0] - k * w, span = [a, a + n * w];
      const off = Math.abs((span[0] + span[1]) / 2 - centre);
      let hits = 0; for (let e = 0; e <= n; e++) if (onRule(a + e * w)) hits++;
      if (!best || hits > best.hits || (hits === best.hits && off < best.off)) best = { off, a, w, hits };
    });
  });
  if (!best || (best.off > 1.5 * best.w && best.hits < n + 1)) return null;
  const out = []; for (let v = 0; v < n; v++) out.push([best.a + v * best.w, best.a + (v + 1) * best.w]);
  if (process.env.OCR_DEBUG_QTY) console.error("rules p" + page + ": " + rules.map(Math.round).join(" "));
  if (process.env.OCR_DEBUG_QTY) console.error("variants p" + page + ": " + out.map(g => Math.round(g[0]) + "-" + Math.round(g[1])).join(" ") + " (heading centre " + Math.round(centre) + ")");
  return out;
}

/** x positions (400-dpi page px) in [x0, x1] where a vertical rule runs down ≥ 60% of rows y0…y1. */
function pageVRules(pdf, page, x0, x1, y0, y1, tmp) {
  const R = pageRaster(pdf, page, DPI, tmp);
  x0 = Math.max(0, Math.round(x0)); x1 = Math.min(R.W - 1, Math.round(x1));
  y0 = Math.max(0, Math.round(y0)); y1 = Math.min(R.H - 1, Math.round(y1));
  const rows = y1 - y0 + 1; if (rows < 50 || x1 <= x0) return [];
  const frac = [];
  for (let x = x0; x <= x1; x++) {
    let n = 0;
    for (let y = y0; y <= y1; y++) if (R.d[y * R.W + x] < 128) n++;
    frac.push(n / rows);
  }
  const out = [];
  for (let i = 0; i < frac.length; i++) {
    if (frac[i] < 0.6) continue;
    let j = i; while (j + 1 < frac.length && frac[j + 1] >= 0.6) j++;   // a rule is a few px wide
    out.push(x0 + (i + j) / 2); i = j;
  }
  return out;
}

// Variant quantities B…n: the same three cell reads as column A, STRICT vote — a digit must be
// confirmed by two reads (or be the light read's), else "–" → null ("not used on it"), matching the
// text manuals. One quick read turned rule fragments into "13", "21" and dashes into "1"/"7".
function readVariantQty(pdf, page, rows, boxes, tmp) {
  const bands = rowBands(rows), cells = [];
  bands.forEach(({ row, lo, hi }) => { for (let v = 1; v < boxes.length; v++) cells.push({ row, v, lo, hi }); });
  const job = (c, mode, psm) => prepColumn(pdf, page, { x: boxes[c.v].x, y: c.lo, w: boxes[c.v].w, h: c.hi - c.lo }, 600, tmp,
    ["--psm", psm, "-c", "tessedit_char_whitelist=0123456789"], mode);
  const digit = ws => { const d = ws.map(w => String(w.t).match(/^(\d{1,3})(?!\d)/)).find(Boolean); return d ? Number(d[1]) : null; };
  // the 3rd read can only change the vote when the first two disagree or one is empty (see needThird)
  const two = runColumns(cells.flatMap(c => [job(c, "qtycellw", "7"), job(c, "qtycell", "7")]), tmp).map(digit);
  const need = cells.map((c, i) => needThird(two[2 * i], two[2 * i + 1]));
  const third = runColumns(cells.filter((c, i) => need[i]).map(c => job(c, "qtycell", "8")), tmp).map(digit);
  let t = 0;
  const q = new Map(bands.map(({ row }) => [row, [row.qty ? row.qty[0] : null]]));
  cells.forEach((c, i) => q.get(c.row).push(voteStrict(two[2 * i], two[2 * i + 1], need[i] ? third[t++] : null)));
  bands.forEach(({ row }) => { row.qty = q.get(row); });
}

/** Does the 3rd (psm 8) cell read matter? Only when the first two disagree, or exactly one is empty.
 *  When they agree (both a digit) both votes return it; when both are empty, voteStrict ignores a lone
 *  3rd read and voteQty falls back to the column read before it — so the result is IDENTICAL and the
 *  OCR call is saved (tesseract was ~95% of the run once rasters were cached, 2026-10-05). */
function needThird(w, r0) { return !((w != null && w === r0) || (w == null && r0 == null)); }

/** Variant / dash-aware vote: two reads agree → that; else the light read; else null. */
function voteStrict(wiped, raw, word) {
  const all = [wiped, raw, word].filter(v => v != null);
  for (const v of all) if (all.filter(x => x === v).length >= 2) return v;
  return raw != null && word == null && wiped == null ? raw : null;
}

// Each row's band: halfway to its neighbours (shared by the per-row cell reads).
function rowBands(rows) {
  const sorted = rows.slice().sort((a, b) => a.y - b.y), ys = sorted.map(r => r.y);
  const pitch = ys.length > 1 ? Math.min(...ys.slice(1).map((y, i) => y - ys[i]).filter(d => d > 10)) : 120;
  return sorted.map((row, i) => ({ row,
    lo: i ? (ys[i - 1] + ys[i]) / 2 : ys[i] - pitch / 2,
    hi: i < ys.length - 1 ? (ys[i] + ys[i + 1]) / 2 : ys[i] + pitch / 2 }));
}

// One OCR per row, of just that row's qty cell — read TWO ways, then a vote with the column read.
//   "qtycellw": with the row-line wipe — right when dashed row lines sit near the digit (D905)
//   "qtycell":  without it — right when the digit sits ON the row line and the wipe eats its
//               base (V1505: "2" → "7")
// No single cleaning suited every layout (2026-10-04), so: two of three agree → that value;
// otherwise the order in voteQty (light read, column, wiped read).
function readQtyCells(pdf, page, rows, box, tmp, strict) {
  const sorted = rows.slice().sort((a, b) => a.y - b.y);
  const ys = sorted.map(r => r.y);
  const pitch = ys.length > 1 ? Math.min(...ys.slice(1).map((y, i) => y - ys[i]).filter(d => d > 10)) : 120;
  const cellJob = (lo, hi, mode, psm) => prepColumn(pdf, page, { x: box.x, y: lo, w: box.w, h: hi - lo }, 600, tmp,
    ["--psm", psm || "7", "-c", "tessedit_char_whitelist=0123456789"], mode);
  const firstDigit = ws => { for (const w of ws) { const m = String(w.t).match(/^(\d{1,3})(?!\d)/); if (m) return { v: Number(m[1]), conf: w.conf }; } return null; };
  const bands = sorted.map((row, i) => ({ row,
    lo: i ? (ys[i - 1] + ys[i]) / 2 : ys[i] - pitch / 2,
    hi: i < ys.length - 1 ? (ys[i] + ys[i + 1]) / 2 : ys[i] + pitch / 2 }));
  // the two reads every row gets, all in one pass; then the 3rd only where it can matter
  const two = runColumns(bands.flatMap(b => [cellJob(b.lo, b.hi, "qtycellw"), cellJob(b.lo, b.hi, "qtycell")]), tmp).map(firstDigit);
  const need = bands.map((b, i) => needThird(two[2 * i] ? two[2 * i].v : null, two[2 * i + 1] ? two[2 * i + 1].v : null));
  const third = runColumns(bands.filter((b, i) => need[i]).map(b => cellJob(b.lo, b.hi, "qtycell", "8")), tmp).map(firstDigit);
  let t = 0;
  bands.forEach(({ row }, i) => {
    const w = two[2 * i], r0 = two[2 * i + 1], r8 = need[i] ? third[t++] : null;
    const reads = [w, r0, r8];
    const col = row.qty ? row.qty[0] : null;
    const pick = strict ? voteStrict(w ? w.v : null, r0 ? r0.v : null, r8 ? r8.v : null)
                        : voteQty(w ? w.v : null, r0 ? r0.v : null, col, r8 ? r8.v : null);
    if (process.env.OCR_DEBUG_QTY) console.error("qty p" + page + " " + row.pn + "  wipe=" + JSON.stringify(reads[0]) + " raw=" + JSON.stringify(reads[1]) + " word=" + JSON.stringify(reads[2]) + " col=" + col + " → " + pick);
    if (pick != null || strict) row.qty = [pick];
  });
}

/** Two of three agree → that value. Otherwise, in this order: the light-cleaned cell read, the
 *  column read, the line-wiped cell read. ⚠ NOT tesseract's confidence — measured useless here:
 *  correct "2"s scored ~2%, wrong "7"s ~80% (V1505). What IS reliable: when line clutter confuses
 *  the light read it returns NOTHING rather than a wrong digit (D905). */
function voteQty(wiped, raw, col, word) {
  // a strict majority first (≥ 2 of up to 4 reads, and more votes than any rival)
  const all = [wiped, raw, col, word].filter(v => v != null);
  const counts = {}; all.forEach(v => { counts[v] = (counts[v] || 0) + 1; });
  const best = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);
  if (best.length && counts[best[0]] >= 2 && (best.length === 1 || counts[best[0]] > counts[best[1]])) return Number(best[0]);
  for (const v of all) if (all.filter(x => x === v).length >= 2) return v;
  if (raw != null) return raw;
  if (col != null) return col;
  return wiped;
}

function cellBox(box, colWords, W) {
  const xs = [box.x].concat((colWords.rules || []).slice().sort((a, b) => a - b)).concat([box.x + box.w]);
  const gaps = [];
  for (let i = 0; i < xs.length - 1; i++) if (xs[i + 1] - xs[i] > 0.012 * W) gaps.push([xs[i], xs[i + 1]]);
  if (!gaps.length) return box;
  const digitX = colWords.filter(w => /\d/.test(w.t)).map(w => w.x + w.w / 2);
  const score = g => digitX.filter(x => x > g[0] && x < g[1]).length;
  const g = gaps.reduce((a, c) => score(c) > score(a) ? c : a, gaps[0]);
  const m = 8;   // stay clear of the rule's wobble
  // ⚠ and at most 3% of the width: the qty is left-aligned in sub-column "A", and when the A|B
  //   divider is DASHED it isn't found as a rule — the cell then held B's "–", read as "7"
  //   ("1 –" → "177", D782, 2026-10-04)
  return { x: g[0] + m, y: box.y, w: Math.max(20, Math.min(g[1] - g[0] - 2 * m, 0.03 * W)), h: box.h };
}


// Vintage title: the first text line right of the number, top of the page ("OIL PAN GROUP").
function vintageTitle(words, W, H) {
  const near = words.filter(w => w.x > 0.06 * W && w.x < 0.6 * W && w.y < 0.035 * H && /[A-Z]{2}/.test(w.t));
  return near.sort((a, b) => a.x - b.x).map(w => w.t.replace(/^_/, "")).join(" ").trim();
}

// All title words of a vintage heading (EN/FR/DE lines, 3+ letters) — for matching pages.
function vintageHeadSet(words, W, H) {
  return new Set(words.filter(w => w.x > 0.06 * W && w.x < 0.6 * W && w.y < 0.08 * H)
    .flatMap(w => w.t.toUpperCase().split(/[^A-Z]+/)).filter(x => x.length > 2 && !/^(GROUP|CODE|NO)$/.test(x)));
}

// y (page px at DPI) of the horizontal row lines in the part-number column: a cheap 100-dpi scan,
// pixel rows that are mostly ink across the column (dashed lines included), merged into lines.
function rowLines(pdf, page, x0, x1, L, tmp) {
  const dpi = 100, k = dpi / DPI, base = path.join(tmp, "l" + page);
  try {
    const y0 = L.headerBottom - 40, h = L.tableBottom - y0;
    cropPgm(pdf, page, dpi, x0 * k, y0 * k, (x1 - x0) * k, h * k, base, tmp);
    const b = fs.readFileSync(base + ".pgm"); fs.unlinkSync(base + ".pgm");
    const m = b.toString("latin1", 0, 64).match(/^P5\s+(\d+)\s+(\d+)\s+(\d+)\s/); if (!m) return [];
    const W = +m[1], H = +m[2], off = m[0].length, out = [];
    let prev = -10;
    for (let y = 0; y < H; y++) {
      let n = 0; for (let x = 0; x < W; x++) if (b[off + y * W + x] < 160) n++;
      if (n > 0.45 * W) { if (y - prev > 3) out.push(y0 + y / k); else out[out.length - 1] = (out[out.length - 1] + y0 + y / k) / 2; prev = y; }
    }
    return out;
  } catch (e) { return []; }
}

// Fill missing rows from spacing gaps (see the call). Returns the rows, with any rescued ones added.
function fillRowGaps(pdf, page, rows, words, L, x0, x1, tmp) {
  if (rows.length < 2) return rows;
  const ys = rows.map(r => r.y).sort((a, b) => a - b);
  const diffs = ys.slice(1).map((y, i) => y - ys[i]).filter(d => d > 10).sort((a, b) => a - b);
  // the pitch: the SMALLEST plausible gap when few rows were found (their gaps are multiples of it),
  // the median otherwise
  let pitch = rows.length < 5 ? diffs[0] : diffs[Math.floor(diffs.length / 2)];
  if (rows.length < 5) { const rh = Math.max(...rows.map(r => r.h || 40)) * 1.4; if (pitch > 2.5 * rh) pitch = pitch / Math.round(pitch / (1.6 * rh)); }
  if (!pitch) return rows;
  const want = [];
  // ⭐ Row LINES first (most layouts print them, solid or dashed): they give the rows directly —
  //   guessing a pitch from 2–3 found rows got it 2× wrong on D1703 (cells two rows tall).
  // ⚠ only when FEW rows were found (< 8): with many, their own spacing is the better guide, and
  //   faint dashed lines are detected patchily and point to the wrong places (D782 p34)
  const lines = rows.length < 8 ? rowLines(pdf, page, x0, x1, L, tmp) : [];
  if (process.env.OCR_DEBUG_ROWS) console.error("lines p" + page + ": " + lines.map(Math.round).join(" ") + "  rows: " + rows.map(r => Math.round(r.y)).join(" "));
  const sp = lines.slice(1).map((y, i) => y - lines[i]).sort((a, b) => a - b);
  const med = sp[Math.floor(sp.length / 2)];
  // and only if they're REGULAR: most spacings a whole number of rows
  const regular = sp.length >= 4 && sp.filter(d => Math.abs(d / med - Math.round(d / med)) < 0.2).length >= 0.75 * sp.length;
  if (regular) {
    for (let i = 0; i < lines.length - 1; i++) {
      const a = lines[i], b = lines[i + 1], k = Math.round((b - a) / med);
      if (b - a < 0.6 * med || k > 6) continue;
      // a span of k rows (a line OCR-scan missed between them — D905 p32 row 020): one cell each
      for (let j = 0; j < k; j++) {
        const ca = a + j * (b - a) / k, cb = a + (j + 1) * (b - a) / k;
        if (!rows.some(r => r.y > ca && r.y < cb)) want.push((ca + cb) / 2);
      }
    }
    pitch = med;
  } else {
  for (let i = 0; i < ys.length - 1; i++) {
    const k = Math.round((ys[i + 1] - ys[i]) / pitch);
    for (let j = 1; j < k && j < 20; j++) want.push(ys[i] + j * (ys[i + 1] - ys[i]) / k);
  }
  // above the first found row, up to the table header (D1703 p30: rows 010–040 came before the
  // first one the page read found)
  for (let y = ys[0] - pitch, n = 0; y > L.headerBottom - 10 && n < 12; y -= pitch, n++) want.push(y);
  // and below the last found row, while still inside the table (trailing rows lost)
  for (let y = ys[ys.length - 1] + pitch, n = 0; y < L.tableBottom - pitch / 2 && n < 30; y += pitch, n++) want.push(y);
  }
  if (!want.length) return rows;
  const cellWords = [];
  const gapReads = runColumns(want.map(y => prepColumn(pdf, page, { x: x0 - 0.04 * L.W, y: y - pitch / 2, w: x1 - x0 + 0.04 * L.W, h: pitch }, 600, tmp,
    ["--psm", "7", "-c", "tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-"], "raw")), tmp);
  want.forEach((y, gi) => {
    const ws = gapReads[gi];
    // one line → join its words so "16241 6401-2" style splits still parse
    if (process.env.OCR_DEBUG_ROWS) console.error("  gap y=" + Math.round(y) + " read: " + ws.map(w => w.t).join(" | "));
    // the read runs into the name column ("…-6411-0P", "…-6402-0FR"): take the REF (3 digits at the
    // start) and the part number BY ITS SHAPE, not the whole string
    const t = ws.map(w => w.t).join("");
    // the DASHED shape first: with dashes optional, "240115881-9104-0" (REF 240 + a bracket read as
    // "1") split as "24011-5881-9" (D782, 2026-10-04)
    const m = t.match(/(\d{5}-\d{4}-\d|[0-9][A-Z0-9]\d{3}-\d{4}-\d|\d{5}-\d{5})/) ||
              t.match(/(\d{5}-?\d{4}-?\d|[0-9][A-Z0-9]\d{3}-?\d{4}-?\d|\d{5}-?\d{5})/);
    if (m) {
      const ref = (t.match(/^(\d{3})/) || [])[1] || "";
      cellWords.push(Object.assign({}, ws[0], { t: ref + m[1], y: y - 10, h: 20 }));
    }
  });
  const before = rows.length;
  const out = rescueRows(rows, words, cellWords, L);
  if (process.env.OCR_DEBUG_ROWS) console.error("gaps p" + page + ": tried " + want.length + ", rescued " + (out.length - before));
  return out;
}

function ocrColumn(pdf, page, box, dpi, tmp, extra, mode) {
  return runColumns([prepColumn(pdf, page, box, dpi, tmp, extra, mode)], tmp)[0];
}

// A column/cell read in two halves, so many small crops can share ONE tesseract process:
// prepColumn crops + wipes (synchronous, cheap); runColumns OCRs a list of them, one tesseract run
// per distinct settings (`-c` options / psm) using tesseract's file-list input. ⭐ Process start-up
// is ~95% of a one-cell read (40 cells: 9.4 s as separate runs vs 0.4 s batched, 2026-10-05).
// OCR_NO_BATCH=1 restores one run per crop, to compare.
let colSeq = 0;
function prepColumn(pdf, page, box, dpi, tmp, extra, mode) {
  const k = dpi / DPI;
  const base = path.join(tmp, "c" + page + "-" + Math.round(box.x) + "-" + (++colSeq));
  cropPgm(pdf, page, dpi, box.x * k, box.y * k, box.w * k, box.h * k, base, tmp);
  const img = base + ".pgm";
  const pad = wipeRules(img, mode) || 0;
  // OCR_DEBUG_DIR=<dir> keeps every column crop (after the rule wipe) to look at by eye
  if (process.env.OCR_DEBUG_DIR) fs.copyFileSync(img, path.join(process.env.OCR_DEBUG_DIR, "p" + page + "-" + (mode || "col") + "-" + Math.round(box.x) + "-" + Math.round(box.y) + ".pgm"));
  const psm = (extra || []).includes("--psm") ? [] : ["--psm", "6"];   // a caller may pick its own mode
  return { img, pad, box, k, args: psm.concat(extra || []), rules: lastRuleXs.slice() };
}

function runColumns(jobs, tmp) {
  const groups = {};
  jobs.forEach((j, i) => { const key = process.env.OCR_NO_BATCH ? "solo" + i : JSON.stringify(j.args); (groups[key] = groups[key] || []).push(i); });
  const tsvs = new Array(jobs.length);
  for (const key in groups) {
    const idx = groups[key], args = jobs[idx[0]].args, opt = { maxBuffer: 64 << 20, stdio: ["ignore", "pipe", "ignore"],
      env: Object.assign({}, process.env, { OMP_THREAD_LIMIT: "1" }) };
    if (idx.length === 1) { tsvs[idx[0]] = execFileSync("tesseract", [jobs[idx[0]].img, "-"].concat(args).concat(["tsv"]), opt).toString(); continue; }
    const list = path.join(tmp, "list-" + (++colSeq) + ".txt");
    fs.writeFileSync(list, idx.map(i => jobs[i].img).join("\n") + "\n");
    const lines = execFileSync("tesseract", [list, "-"].concat(args).concat(["tsv"]), opt).toString().split("\n");
    fs.unlinkSync(list);
    // page_num (column 2) is the 1-based position in the list; keep the header for tsvWords
    const per = idx.map(() => [lines[0]]);
    lines.slice(1).forEach(l => { const c = l.split("\t"); const n = +c[1]; if (n >= 1 && n <= idx.length) per[n - 1].push(l); });
    idx.forEach((i, n) => { tsvs[i] = per[n].join("\n"); });
  }
  return jobs.map((j, i) => {
    try { fs.unlinkSync(j.img); } catch (e) {}
    const out = tsvWords(tsvs[i]).map(w => Object.assign(w, { x: j.box.x + (w.x - j.pad) / j.k, y: j.box.y + (w.y - j.pad) / j.k, w: w.w / j.k, h: w.h / j.k }));
    out.rules = j.rules.map(x => j.box.x + x / j.k);
    return out;
  });
}

/**
 * Erase table rules from a greyscale PGM in place: a pixel row that is mostly ink is a
 * horizontal rule (solid or dashed); a pixel column that is mostly ink is a vertical rule.
 * Text never fills that much of a row or column, so it survives.
 */
function padOnly(file) {
  const b = fs.readFileSync(file);
  const m = b.toString("latin1", 0, 64).match(/^P5\s+(\d+)\s+(\d+)\s+(\d+)\s/);
  if (!m) return 0;
  const W = +m[1], H = +m[2], off = m[0].length, P = 30, W2 = W + 2 * P, H2 = H + 2 * P;
  const head = Buffer.from("P5\n" + W2 + " " + H2 + "\n255\n", "latin1");
  const body = Buffer.alloc(W2 * H2, 255);
  for (let y = 0; y < H; y++) b.copy(body, (y + P) * W2 + P, off + y * W, off + (y + 1) * W);
  fs.writeFileSync(file, Buffer.concat([head, body]));
  return P;
}

let lastRuleXs = [];   // x (crop px, before padding) of the vertical rules the last wipe erased
function wipeRules(file, mode) {
  lastRuleXs = [];
  // "raw": no cleaning, just the white margin — a single row's cell has no rules worth erasing, and
  // the row-line wipe SLICES a line of part numbers (a pixel row through the 5/8/6/4 bars and the
  // hyphens is > 32% ink) → tesseract read nothing (D782 gap cells, 2026-10-04)
  if (mode === "raw") return padOnly(file);
  const b = fs.readFileSync(file);
  const m = b.toString("latin1", 0, 64).match(/^P5\s+(\d+)\s+(\d+)\s+(\d+)\s/);
  if (!m) return 0;
  const W = +m[1], H = +m[2], off = m[0].length, dark = 128;
  // ⚠ ORDER MATTERS: vertical rules first — the horizontal wipe whites out whole pixel rows,
  //   which cuts a vertical rule into row-height pieces too short for the run test below.
  // Vertical rules: a pixel column that is mostly ink. ⚠ NOT for the part-number column —
  // every number there starts with a digit at the same x, and that stack of "1"s reads as a
  // rule and gets erased (measured 2026-10-03). That crop starts after the rule instead.
  if (mode !== "pn") {
    for (let x = 0; x < W; x++) {
      let n = 0; for (let y = 0; y < H; y++) if (b[off + y * W + x] < dark) n++;
      if (n > 0.45 * H) for (let y = 0; y < H; y++) b[off + y * W + x] = 255;
    }
    // ⚠ A scanned rule is rarely straight and often dashed: it drifts a few px over the page,
    // so no pixel column reaches the 45% above and it survives (D905, 2026-10-04: the qty
    // column's rules read as "1" → "31", "13"). Coverage can't separate a dashed rule (31–62%)
    // from a stack of digits (22%+). LENGTH can: erase any vertical stroke that runs unbroken
    // (in a narrow band, tolerating a 2 px wobble) for ≥ RUN px — no digit is that tall.
    const R = 2, RUN = process.env.OCR_NO_RUN ? 1e9 : 100;
    const inkAt = (x, y) => { for (let d = -R; d <= R; d++) { const xx = x + d; if (xx >= 0 && xx < W && b[off + y * W + xx] < dark) return true; } return false; };
    const kill = [];
    for (let x = 0; x < W; x += 2) {
      let start = -1;
      for (let y = 0; y <= H; y++) {
        const on = y < H && inkAt(x, y);
        if (on && start < 0) start = y;
        if (!on && start >= 0) { if (y - start >= RUN) kill.push([x, start, y]); start = -1; }
      }
    }
    lastRuleXs = [...new Set(kill.map(k => k[0]))];
    kill.forEach(([x, y0, y1]) => { for (let y = y0; y < y1; y++) for (let d = -R - 1; d <= R + 1; d++) { const xx = x + d; if (xx >= 0 && xx < W) b[off + y * W + xx] = 255; } });
  }
  // Horizontal rules (solid or dashed): a pixel row that is mostly ink.
  // ⚠ NOT in a qty CELL: some books print the digit ON the row line, and wiping the line took
  //   the base of a "2" with it ("7", V1505, 2026-10-04). The cell is cropped clear of the
  //   column's own rules, so a row line left in it is harmless to the one-word read.
  // ⚠ NOR in the PART-NUMBER column: a pixel row through a line of numbers (5/8/6/4 bars, hyphens)
  //   is > 32% ink, so the wipe sliced the text — measured: confirmed numbers D782 p34 0 → 19 of 28,
  //   V3600 p25 7 → 10 of 20 without it (2026-10-04). OCR_PN_HWIPE=1 restores the old behaviour.
  if (mode !== "qtycell" && !(mode === "pn" && !process.env.OCR_PN_HWIPE)) for (let y = 0; y < H; y++) {
    let n = 0; for (let x = 0; x < W; x++) if (b[off + y * W + x] < dark) n++;
    if (n > 0.32 * W) for (let x = 0; x < W; x++) b[off + y * W + x] = 255;
  }
  // Dashed / broken horizontal rules: under 32% of a row, so the pass above misses them, and
  // the dash clutter makes tesseract drop the digits between them (D905 qty, 2026-10-04).
  // A dash is THIN (≤ THIN px of vertical ink) while digit and letter strokes are thicker:
  // erase thin ink, but only on rows where a lot of it lines up (a rule, not a stray serif).
  if (mode !== "pn" && mode !== "qtycell" && !process.env.OCR_NO_THIN) {
    const THIN = 4, vr = new Uint16Array(W * H);
    for (let x = 0; x < W; x++) {
      let y = 0;
      while (y < H) {
        if (b[off + y * W + x] >= dark) { y++; continue; }
        let e = y; while (e < H && b[off + e * W + x] < dark) e++;
        for (let k = y; k < e; k++) vr[k * W + x] = e - y;
        y = e;
      }
    }
    const thinRow = new Uint32Array(H);
    for (let y = 0; y < H; y++) { let n = 0; for (let x = 0; x < W; x++) { const v = vr[y * W + x]; if (v && v <= THIN) n++; } thinRow[y] = n; }
    for (let y = 0; y < H; y++) {
      let n = 0; for (let d = -2; d <= 2; d++) if (y + d >= 0 && y + d < H) n += thinRow[y + d];
      if (n > 0.25 * W) for (let x = 0; x < W; x++) { const v = vr[y * W + x]; if (v && v <= THIN) b[off + y * W + x] = 255; }
    }
  }
  // QTY column only: drop every separate ink blob shorter than MINH px. A quantity digit is
  // ~45 px tall at 600 dpi; leftover dash pieces, specks and serial-range ticks are not
  // (D905, 2026-10-04: dashes 6–8 px thick survived every rule test and read as "7").
  // ⚠ Not for names/remarks — commas, hyphens and the "-" of "-0.25mm" are short blobs too.
  if (mode === "qty") {
    // binarise with OUR threshold first: light-grey dash edges are "not ink" to the filters
    // here but tesseract's own binarisation still saw them — so make the two agree
    // ⚠ at a LENIENT threshold (BIN), not `dark`: tiny qty digits draw their thin strokes in
    //   faint grey, and cutting at 128 turned a "2" into a "4" (V1505, 2026-10-04)
    const BIN = Number(process.env.QTY_BIN || 190);
    for (let i = 0; i < W * H; i++) b[off + i] = b[off + i] < BIN ? 0 : 255;
    const MINH = process.env.OCR_NO_BLOB ? 0 : 12, seen = new Uint8Array(W * H), stack = [];
    for (let i = 0; i < W * H; i++) {
      if (seen[i] || b[off + i] >= dark) continue;
      const blob = []; let y0 = H, y1 = -1, x0 = W, x1 = -1; stack.push(i); seen[i] = 1;
      while (stack.length) {
        const j = stack.pop(); blob.push(j); const y = (j / W) | 0, x = j - y * W;
        if (y < y0) y0 = y; if (y > y1) y1 = y; if (x < x0) x0 = x; if (x > x1) x1 = x;
        for (const [dx, dy] of [[1,0],[-1,0],[0,1],[0,-1]]) {
          const nx = x + dx, ny = y + dy; if (nx < 0 || ny < 0 || nx >= W || ny >= H) continue;
          const k = ny * W + nx; if (!seen[k] && b[off + k] < dark) { seen[k] = 1; stack.push(k); }
        }
      }
      // short AND dash-shaped (≥ 3× wider than tall) or a speck (small both ways) — NOT any short blob: the faint
      // top curve of a small "2" breaks off as its own short blob, and losing it reads "4"
      const bh = y1 - y0 + 1, bw = x1 - x0 + 1;
      if (bh < MINH && (bw >= 3 * bh || bw < MINH)) blob.forEach(j => { b[off + j] = 255; });
    }
  }
  // QTY CELL, light read: remove SEPARATE flat blobs (≥ 3× wider than tall, short) — a dashed
  // or solid row line lying near the digit makes the one-line read see two lines (D905). When
  // the digit sits ON the line (V1505) they are one tall blob and survive, base and all.
  if (mode === "qtycell") {
    const seen = new Uint8Array(W * H), stack = [];
    for (let i = 0; i < W * H; i++) {
      if (seen[i] || b[off + i] >= dark) continue;
      const blob = []; let y0 = H, y1 = -1, x0 = W, x1 = -1; stack.push(i); seen[i] = 1;
      while (stack.length) {
        const j = stack.pop(); blob.push(j); const y = (j / W) | 0, x = j - y * W;
        if (y < y0) y0 = y; if (y > y1) y1 = y; if (x < x0) x0 = x; if (x > x1) x1 = x;
        for (const [dx, dy] of [[1,0],[-1,0],[0,1],[0,-1]]) {
          const nx = x + dx, ny = y + dy; if (nx < 0 || ny < 0 || nx >= W || ny >= H) continue;
          const k = ny * W + nx; if (!seen[k] && b[off + k] < dark) { seen[k] = 1; stack.push(k); }
        }
      }
      const bh = y1 - y0 + 1, bw = x1 - x0 + 1;
      if (bh <= 0.12 * H && bw >= 3 * bh) {
        // whiten the blob AND its faint grey halo, or tesseract still sees a ghost line
        blob.forEach(j => { const y = (j / W) | 0, x = j - y * W;
          for (let dy = -2; dy <= 2; dy++) for (let dx = -2; dx <= 2; dx++) {
            const nx = x + dx, ny = y + dy; if (nx < 0 || ny < 0 || nx >= W || ny >= H) continue;
            const v = b[off + ny * W + nx]; if (v >= dark && v < 230) b[off + ny * W + nx] = 255; } });   // grey halo only — never another blob's ink
        blob.forEach(j => { b[off + j] = 255; });
      }
    }
  }
  // Tesseract misses glyphs that touch the image edge (a leading "1" beside a wiped rule):
  // write the column back with a white margin, and report the shift so boxes map back.
  const P = 30, W2 = W + 2 * P, H2 = H + 2 * P;
  const head = Buffer.from("P5\n" + W2 + " " + H2 + "\n255\n", "latin1");
  const body = Buffer.alloc(W2 * H2, 255);
  for (let y = 0; y < H; y++) b.copy(body, (y + P) * W2 + P, off + y * W, off + (y + 1) * W);
  fs.writeFileSync(file, Buffer.concat([head, body]));
  return P;
}

function pageCount(pdf) {
  const info = execFileSync("pdfinfo", [pdf], { stdio: ["ignore", "pipe", "ignore"] }).toString();
  return Number((info.match(/^Pages:\s+(\d+)/m) || [])[1]) || 0;
}

/**
 * @param {string} pdf
 * @param {{pages?:number[], log?:Function}} opts
 */
function ocrBook(pdf, opts) {
  opts = opts || {};
  const n = pageCount(pdf);
  const pages = opts.pages || Array.from({ length: n }, (_, i) => i + 1);
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "hqocr-"));
  const out = { brand: "Kubota", source: "ocr", file: pdf, model: "", codeNo: "", validity: "", models: [],
                sections: [], flags: [], contents: [], imagePages: [], pageCount: n, bands: {} };
  let current = null, coverText = "";
  let pending = null;   // section code seen on a drawing page, waiting for its table
  let vintageOk = true; // the bare-number section read, until the book shows a real code
  let vintageSeq = 0, lastVintage = null;   // vintage sections are numbered in page order (their printed numbers misread)
  try {
    for (const p of pages) {
      const { words, W, H } = ocrPage(pdf, p, tmp);
      if (p <= 4) coverText += " " + words.map(w => w.t).join(" ");
      const r = readScanPage(words, W, H);
      if (!r.code && !r.isIndex) {
        const c = ocrSectionCode(pdf, p, W, H, tmp);
        if (c) { r.code = c; r.name = r.name || titleBeside(words, W, H); }
        else if (vintageOk) {
          const v = ocrVintageCode(pdf, p, W, H, tmp);
          if (v) { r.code = v; r.name = r.name || vintageTitle(words, W, H); }
        }
      }
      if (r.code && !/^V\d/.test(r.code)) vintageOk = false;   // a real E-code / 4-digit code → not a vintage book
      // Vintage numbers misread too often to trust ("V14" for 2) — and a misread repeat would MERGE
      // two sections. So only "a heading is here" counts: number the sections in page order, and
      // a page whose title shares most words with the current section continues it.
      if (r.code && /^V\d/.test(r.code)) {
        // compare with the LAST HEADING SEEN (a drawing page and its table both carry it), using ALL
        // THREE title lines (EN/FR/DE): OCR often catches only some lines, in different languages on
        // different pages (p18 FR+DE, p19 EN — D850, 2026-10-04)
        const a = vintageHeadSet(words, W, H), b = lastVintage ? lastVintage.set : new Set();
        const shared = [...a].filter(x => b.has(x)).length;
        if (lastVintage && a.size && b.size && shared >= 2) r.code = lastVintage.code;
        else { vintageSeq++; r.code = "V" + String(vintageSeq).padStart(2, "0"); }
        const same = lastVintage && lastVintage.code === r.code;
        lastVintage = { code: r.code, set: same ? new Set([...b, ...a]) : a,
                        name: same && /\bGROUP\b/.test(lastVintage.name || "") ? lastVintage.name : r.name };
        r.name = lastVintage.name;
        if (process.env.OCR_DEBUG_SEC) console.error("vintage p" + p + " → " + r.code + " | " + r.name);
      }
      if (r.rows.length) {
        const L = r.layout;
        // ⚠ start ABOVE the first row, not at a fixed offset under the headers: the first
        //   row's qty digit sat across the old crop edge and read as "2" (D905, 2026-10-04)
        const top = Math.min(L.headerBottom - 20, Math.min(...r.rows.map(x => x.y - x.h))), hgt = L.tableBottom - top;
        // 1. part-number column alone, restricted alphabet → rows the page read skipped
        const pnEdge = Math.min(0.26 * W, Math.max(...words.filter(w => r.rows.some(x => x.pn && Math.abs(x.y - (w.y + w.h / 2)) < 30 && w.x < 0.32 * W && /\d{4}/.test(w.t))).map(w => w.x + w.w)) + 20);
        // start just left of the part numbers, NOT at the REF column: with the rule between
        // them wiped, "010" and "16423-…" run together and a digit is lost
        const pnLeft = Math.min(...words.filter(w => r.rows.some(x => Math.abs(x.y - (w.y + w.h / 2)) < 30) && w.x < 0.32 * W && /\d{4}/.test(w.t)).map(w => w.x)) - 12;
        const x0 = isFinite(pnLeft) ? pnLeft : L.pnX0;
        const pnCol = ocrColumn(pdf, p, { x: x0, y: top, w: (isFinite(pnEdge) ? pnEdge : 0.22 * W) - x0, h: hgt }, DPI, tmp,
          ["-c", "tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-"], "pn");
        r.rows = rescueRows(r.rows, words, pnCol, L);
        // 1b. rows NEITHER read found: rows are evenly spaced, so a gap of ~2× the pitch is a missing
        //     row. Read just that row's part-number cell, one line, and rescue it if it's a valid
        //     number (D782 p34: 8 of 28 parts were lost, e.g. "1-64T1-0" for 15881-6411-0, 2026-10-04)
        if (!process.env.OCR_NO_GAPS) r.rows = fillRowGaps(pdf, p, r.rows, words, L, x0, (isFinite(pnEdge) ? pnEdge : 0.22 * W), tmp);
        // 1c. REF column alone, digits only. Rows rescued from the part-number column carry no REF —
        //     on D1703 p23 the page read caught almost none of the table, so 206 of 509 rows had none
        //     and couldn't be linked to the drawing. Only fills a row that has no REF yet.
        if (isFinite(pnLeft) && !process.env.OCR_NO_REFCOL) {
          const rx = Math.max(0, pnLeft - 0.07 * W);
          attachColumn(r.rows, ocrColumn(pdf, p, { x: rx, y: top, w: pnLeft + 8 - rx, h: hgt }, 600, tmp,
            ["-c", "tessedit_char_whitelist=0123456789"]), "ref");
        }
        // 2. names: their own crop (the page read loses whole rows in denser books)
        const nameLeft = isFinite(pnEdge) ? pnEdge + 4 : 0.22 * W;
        attachColumn(r.rows, ocrColumn(pdf, p, { x: nameLeft, y: top, w: L.nameEnd - nameLeft, h: hgt }, 600, tmp), "name");
        // 3. quantity and 4. remarks (sizes) — printed small, so read at 600 dpi
        const qtyBox = { x: L.stueckX, y: top, w: (L.qtyEnd || L.remarksX - 20) - L.stueckX, h: hgt };
        const qtyCol = ocrColumn(pdf, p, qtyBox, 600, tmp, ["-c", "tessedit_char_whitelist=0123456789-"], "qty");
        attachColumn(r.rows, qtyCol, "qty");
        // 3b. then each row's qty CELL on its own, one word, and that read wins. Reading the whole
        //     column let neighbouring marks (rule pieces, dashes, serial arrows) merge with the
        //     digit, and every cleaning rule that fixed one book's layout broke another's
        //     (2026-10-04, D905 / Z400 / V1505). A single cell has none of those neighbours.
        // 3c. books with MODEL COLUMNS (A:… B:… C:… D:… in the header): one qty per variant, "–" =
        //     not used on it. The sub-columns are found first, so column A is read in the RIGHT place.
        const letters = pageModels(words).map(m => m.col);
        const subCols = letters.length > 1 && !process.env.OCR_NO_VARIANTS ? variantColumns(pdf, p, qtyBox, L, letters.length, W, tmp) : null;
        const mid = g => ({ x: (g[0] + g[1]) / 2 - 0.35 * (g[1] - g[0]), y: qtyBox.y, w: 0.7 * (g[1] - g[0]), h: qtyBox.h });
        const aBox = subCols ? mid(subCols[0]) : cellBox(qtyBox, qtyCol, W);
        // ⚠ in a variant book the column read above looked in the wrong place (its left edge is in B):
        //   don't let it vote for A
        if (subCols) r.rows.forEach(x => { x.qty = null; });
        if (!process.env.OCR_NO_CELLS) readQtyCells(pdf, p, r.rows, aBox, tmp, !!subCols);
        if (subCols) readVariantQty(pdf, p, r.rows, subCols.map(mid), tmp);
        attachColumn(r.rows, ocrColumn(pdf, p, { x: L.remarksX, y: top, w: 0.98 * W - L.remarksX, h: hgt }, 600, tmp,
          ["-c", "tessedit_char_whitelist=0123456789+-.STDEmMSOVRIZEABCHGKLNPU/"]), "remark");
      }
      if (opts.log) opts.log(p, r);
      if (r.isIndex && out.sections.length) break;     // the index at the END, never the contents page
      // a DRAWING page (no table) that carries the section code: remember it for the table that
      // follows — newer books print "E07." only there (the FIRST section's drawing comes before
      // any parts, so this can't wait for parts to start). A table page's own code still wins.
      if (!r.rows.length) { if (r.code) pending = { code: r.code, name: r.name }; continue; }
      if (!r.code && pending) { r.code = pending.code; r.name = r.name || pending.name; }
      pending = null;
      pageModels(words).forEach(m => { if (!out.models.some(x => x.col === m.col)) out.models.push(m); });
      if (r.model && !out.models.length) out.models.push({ col: "A", name: r.model });
      out.models.sort((a, b) => a.col.localeCompare(b.col));
      if (r.code && (!current || current.code !== r.code)) {
        current = out.sections.find(s => s.code === r.code);
        if (!current) {
          current = { code: r.code, name: r.name || "(section " + r.code + ")", page: p, parts: [] };
          out.sections.push(current);
          if (r.band) out.bands[r.code] = { x: 0, y: r.band.y * 72 / DPI, w: W * 72 / DPI, h: r.band.h * 72 / DPI };
        }
      } else if (!current) {
        current = { code: "????", name: "(untitled)", page: p, parts: [] }; out.sections.push(current);
        out.flags.push({ page: p, kind: "no-section", line: "" });
      }
      r.rows.forEach(row => {
        const part = { ref: row.ref, pn: row.pn, name: row.name, remark: row.remark, qty: row.qty, page: p,
                        ocr: { conf: row.conf, agree: !!row.agree, rescued: !!row.rescued } };
        current.parts.push(part);
        if (!row.ref) out.flags.push({ page: p, kind: "ocr-no-ref", line: row.pn + " " + row.name });
        if (!row.name) out.flags.push({ page: p, kind: "no-name", line: row.pn });
      });
    }
  } finally { dropRasters(); fs.rmSync(tmp, { recursive: true, force: true }); }

  const mc = coverText.match(/\b([A-Z]\d{3,4}[A-Z0-9-]*-E\dB[A-Z0-9-]*|[A-Z]\d{3,4}-[A-Z0-9-]{3,})\s+([0-9A-Z]{5}-\d{5})\b/);
  if (mc) { out.model = mc[1]; out.codeNo = mc[2]; }
  if (!out.model && out.models.length) out.model = out.models[0].name;
  const mv = coverText.match(/\(\s*\d{2}\.\d{2}\.\d{4}\s*-->\s*[\d.]*\s*\)/);
  if (mv) out.validity = mv[0].replace(/\s+/g, " ");
  return out;
}

module.exports = { ocrBook, ocrPage, ocrSectionCode, voteQty, voteStrict, needThird, pageModels };
