// Look at ONE character of a part number on the page image, instead of trusting tesseract's guess.
//
// Why: in the newer Kubota books tesseract reads the C of "1C010-2395-0" as a 6 (or 0), the J of
// "1J530" as a 4 and the G of "1G520" as a 6 — and the result is a valid-looking number, so no shape
// check can see it. User test 2026-10-06 found ~30–40 per V3300/V3600/V3800 book still wrong after
// lib/pn-fix.js, which only corrects on outside evidence.
//
// The test is TOPOLOGY, not a trained classifier: 0, 4 and 6 enclose a loop; C, G and J don't.
// Measured on V3600-T-E3B-EU-ZH4 pp 15/20/21 at 400 dpi: every 0/4/6/9 had 80–350 px of enclosed
// background, every C/G/J had 0. Then, for an open glyph, which letter:
//   C — almost no ink in the inner-right "bar" zone (12–15 px)
//   J — bar-zone ink 33–36 px AND an empty top-left (the J has no left stroke up top)
//   G — bar-zone ink ~46 px (the inward bar), left stroke present
// Only the DIGIT→LETTER direction is ever changed (the misread runs one way: lib/pn-fix.js notes).
"use strict";

const fs = require("fs");

function loadPgm(file) {
  const b = fs.readFileSync(file);
  const m = b.toString("latin1", 0, 64).match(/^P5\s+(\d+)\s+(\d+)\s+(\d+)\s/);
  return { W: +m[1], H: +m[2], d: b.subarray(m[0].length) };
}

/** Otsu threshold of a grey crop — the cut that best splits it into two tones (ink, paper). */
function otsu(R, x, y, w, h) {
  const hist = new Array(256).fill(0);
  for (let r = 0; r < h; r++) for (let c = 0; c < w; c++) hist[R.d[(y + r) * R.W + x + c]]++;
  const n = w * h;
  let sum = 0; for (let i = 0; i < 256; i++) sum += i * hist[i];
  let sumB = 0, wB = 0, best = 0, th = 150;
  for (let t = 0; t < 256; t++) {
    wB += hist[t]; if (!wB) continue;
    const wF = n - wB; if (!wF) break;
    sumB += t * hist[t];
    const mB = sumB / wB, mF = (sum - sumB) / wF, between = wB * wF * (mB - mF) * (mB - mF);
    if (between > best) { best = between; th = t + 1; }
  }
  return th;
}

/** Ink components inside (x, y, w, h) of raster R = {W, H, d}, sorted left → right. */
function components(R, x, y, w, h, th) {
  x = Math.max(0, Math.round(x)); y = Math.max(0, Math.round(y));
  w = Math.max(1, Math.min(Math.round(w), R.W - x)); h = Math.max(1, Math.min(Math.round(h), R.H - y));
  // ink/paper cut-off from THIS crop (Otsu): a fixed 150 lost the light half of every stroke on
  // faintly printed words (D1503 p14/p15/p34), so whole glyphs came out short and "open" — a G read
  // as C, a 6 with its loop gone read as G
  if (!th) th = otsu(R, x, y, w, h);
  const ink = new Uint8Array(w * h);
  for (let r = 0; r < h; r++) for (let c = 0; c < w; c++) ink[r * w + c] = R.d[(y + r) * R.W + x + c] < th ? 1 : 0;
  const lab = new Int32Array(w * h).fill(-1), list = [];
  for (let i = 0; i < w * h; i++) {
    if (!ink[i] || lab[i] >= 0) continue;
    const id = list.length, st = [i]; lab[i] = id;
    let n = 0, x0 = 1e9, x1 = -1, y0 = 1e9, y1 = -1;
    while (st.length) {
      const p = st.pop(); n++;
      const px = p % w, py = (p / w) | 0;
      if (px < x0) x0 = px; if (px > x1) x1 = px; if (py < y0) y0 = py; if (py > y1) y1 = py;
      for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
        const qx = px + dx, qy = py + dy;
        if (qx < 0 || qy < 0 || qx >= w || qy >= h) continue;
        const q = qy * w + qx;
        if (ink[q] && lab[q] < 0) { lab[q] = id; st.push(q); }
      }
    }
    list.push({ id, n, x0, x1, y0, y1 });
  }
  list.sort((a, b) => a.x0 - b.x0);
  return { w, h, lab, list };
}

/** Background pixels inside the glyph's box that the outside can't reach = its loop(s). */
function holeArea(C, g) {
  const bw = g.x1 - g.x0 + 3, bh = g.y1 - g.y0 + 3, m = new Uint8Array(bw * bh);
  for (let r = 0; r < bh; r++) for (let k = 0; k < bw; k++) {
    const X = g.x0 - 1 + k, Y = g.y0 - 1 + r;
    m[r * bw + k] = X >= 0 && Y >= 0 && X < C.w && Y < C.h && C.lab[Y * C.w + X] === g.id ? 1 : 0;
  }
  const seen = new Uint8Array(bw * bh), st = [0]; seen[0] = 1;
  while (st.length) {
    const p = st.pop(), px = p % bw, py = (p / bw) | 0;
    for (const [dx, dy] of [[1, 0], [-1, 0], [0, 1], [0, -1]]) {
      const qx = px + dx, qy = py + dy;
      if (qx < 0 || qy < 0 || qx >= bw || qy >= bh) continue;
      const q = qy * bw + qx;
      if (!m[q] && !seen[q]) { seen[q] = 1; st.push(q); }
    }
  }
  let a = 0;
  for (let i = 0; i < bw * bh; i++) if (!m[i] && !seen[i]) a++;
  return a;
}

/** Ink of glyph g inside a zone given as fractions of its box. */
function zoneInk(C, g, fx0, fx1, fy0, fy1) {
  let n = 0;
  const gw = Math.max(1, g.x1 - g.x0), gh = Math.max(1, g.y1 - g.y0);
  for (let r = g.y0; r <= g.y1; r++) for (let k = g.x0; k <= g.x1; k++) {
    if (C.lab[r * C.w + k] !== g.id) continue;
    const fx = (k - g.x0) / gw, fy = (r - g.y0) / gh;
    if (fx > fx0 && fx < fx1 && fy > fy0 && fy < fy1) n++;
  }
  return n;
}

/**
 * What the glyph g is, by shape: "loop" (a digit with a hole — 0/4/6/8/9), "C"/"G"/"J", "5", or null.
 * Thresholds are relative to the glyph's own ink so a heavier or lighter scan reads the same.
 */
function classify(C, g) {
  const hole = holeArea(C, g), ink = g.n;
  if (hole >= Math.max(25, ink * 0.08)) return { kind: "loop", hole };
  if (hole > 0) return { kind: null, hole };                     // a speck of a loop: undecided
  const bar = zoneInk(C, g, 0.45, 0.95, 0.45, 0.75) / ink;       // C ≈ .04 · J ≈ .11 · G ≈ .115
  const topLeft = zoneInk(C, g, -0.01, 0.35, -0.01, 0.55) / ink; // C/G: the left stroke · J: almost empty
  // a 5 is open too ("15261" read "16261", V3600-T p43): its bowl opens to the LEFT in the lower
  // middle, where C and G keep a solid stroke — measured 5 ≤ .065 · G .09–.14 · C .13–.16
  const leftLow = zoneInk(C, g, -0.01, 0.25, 0.5, 0.8) / ink;
  // an empty top-left is a J — or the OPEN-TOPPED 4 of the old typewriter face (Z400, D1402: no loop
  // at all). The 4's diagonal reaches the left side halfway down; the J's left is empty until the hook.
  // Measured: J .017–.063 · serif 4 .122–.145. An open 4 is a real digit → reported as "loop".
  if (topLeft < 0.12) {
    const leftMid = zoneInk(C, g, -0.01, 0.35, 0.3, 0.65) / ink;
    return leftMid < 0.09 ? { kind: "J", hole, bar, topLeft, leftMid } : { kind: "loop", hole, openFour: true, leftMid };
  }
  if (leftLow < 0.085) return { kind: "5", hole, bar, topLeft, leftLow };
  if (bar < 0.07) return { kind: "C", hole, bar, topLeft, leftLow };
  if (bar >= 0.09) return { kind: "G", hole, bar, topLeft, leftLow };
  return { kind: null, hole, bar, topLeft, leftLow };
}

/** Widest row in the bottom 8% of glyph g ÷ its width at 60% height (a serif "1" has a foot). */
function footRatio(C, g) {
  const rowW = r => { let a = 1e9, b = -1; for (let k = g.x0; k <= g.x1; k++) if (C.lab[r * C.w + k] === g.id) { if (k < a) a = k; b = k; } return b < 0 ? 0 : b - a + 1; };
  const gh = g.y1 - g.y0;
  let foot = 0;
  for (let r = g.y1 - Math.round(gh * 0.08); r <= g.y1; r++) foot = Math.max(foot, rowW(r));
  return foot / Math.max(1, rowW(g.y0 + Math.round(gh * 0.6)));
}

// Median height of the leading "1" over every part-number-shaped word on the page (cached per page).
const PAGE_H = new WeakMap();
function pageGlyphHeight(R, words) {
  if (PAGE_H.has(words)) return PAGE_H.get(words);
  const hs = [];
  for (const w of words) {
    if (!/^[|\[(]*1[0-9A-Z]{4}-?\d{4}-?\d/.test(String(w.t))) continue;
    const pad = Math.max(3, Math.round(w.h * 0.35));
    const C = components(R, w.x - 3, w.y - pad, Math.min(w.w + 6, Math.round(w.h * 1.2)), w.h + 2 * pad);
    const tall = C.list.filter(g => g.y1 - g.y0 > 8);
    if (tall.length) hs.push(Math.max(...tall.map(g => g.y1 - g.y0)));
  }
  hs.sort((a, b) => a - b);
  const h = hs.length ? hs[hs.length >> 1] : 0;
  PAGE_H.set(words, h);
  return h;
}

// characters tesseract prints for the tall glyphs of a word (dashes/specks/underlines drop out)
const SMALL = /[-_~°.,:;'"`*\s]/g;

/**
 * Find a part number's second character on the page and classify it.
 * @param R      page raster at the dpi the words were read at
 * @param words  tsvWords() of the page
 * @param pn     the number as OCR'd, e.g. "16010-2395-0"
 * @param opts   { pageH } — the page's usual part-number glyph height (else measured from `words`)
 * @return {kind,...}|null   null = not found on the page or the glyphs don't line up
 */
function secondCharShape(R, words, pn, opts) {
  const refH = opts && opts.pageH != null ? opts.pageH : pageGlyphHeight(R, words);   // tests pass the page's value
  const key = pn.replace(/-/g, "").slice(2);          // the 8 chars after the doubtful one
  const hits = [];
  for (const w of words) {
    const big = String(w.t).toUpperCase().replace(SMALL, "");
    const at = big.indexOf(key);
    if (at < 2) continue;
    // tesseract's box can be SHORTER than the printed glyphs (faint lines, D1503 p14): a ±3 px crop cut
    // the tops and bottoms off, leaving 26-px fragments that read "open" — pad by a third of the height
    const pad = Math.max(3, Math.round(w.h * 0.35));
    const C = components(R, w.x - 3, w.y - pad, w.w + 6, w.h + 2 * pad);
    if (!C.list.length) continue;
    const H = Math.max(...C.list.map(g => g.y1 - g.y0));
    const tall = C.list.filter(g => g.y1 - g.y0 > H * 0.5);
    // the text and the glyphs must line up: same count, or the number starts the word (any extra
    // glyphs — a glued name — come after it and can't shift the index)
    // ⚠ never FEWER glyphs than characters: a "1" outside the word box shifts every index by one
    //   (V3600-T p29: the test measured the 0 after the C and called it a loop)
    if (tall.length < big.length || (tall.length !== big.length && at !== 2)) continue;
    const g = tall[at - 1], one = tall[at - 2];
    if (!g || !one) continue;
    // ⚠ smaller than the page's usual part-number glyph → refuse (D1503 p14/15/34: 26–35 px against
    //   ~40, and at that size a G read as C and a 6 with a broken loop read as G; the old fix had them right)
    if ((one.y1 - one.y0) < refH * 0.85) continue;
    // the glyph must BE a character: a scrap of a table rule caught in the word box (3 px wide,
    // V3600-T p12/p21) is "open" and was read as a 5 / C — so size it against its neighbours
    const hs = tall.slice(Math.max(0, at - 2), at + 8).map(t => t.y1 - t.y0).sort((a, b) => a - b);
    const medH = hs[hs.length >> 1], gh = g.y1 - g.y0, gw = g.x1 - g.x0;
    if (gh < medH * 0.8 || gw < gh * 0.35) continue;
    if ((one.x1 - one.x0) > (one.y1 - one.y0) * 0.45) continue;   // the leading "1" must be a narrow stroke
    // the old TYPEWRITER face (Z400, D1402…): its 6 prints with a faint loop that often doesn't close,
    // so "open" proves nothing there (16851 / 16241 were called G). Its "1" has a flat base foot —
    // measured foot/middle width 2.2–4 vs exactly 1 in the newer books → on that face, digits only.
    const v = classify(C, g);
    if (v.kind && v.kind !== "loop" && footRatio(C, one) > 1.6) { hits.push({ kind: null, typewriter: true }); continue; }
    hits.push(v);
  }
  if (!hits.length) return null;
  const k = hits[0].kind;
  return hits.every(h => h.kind === k) ? hits[0] : { kind: null, conflict: hits.length };
}

module.exports = { loadPgm, components, holeArea, classify, footRatio, pageGlyphHeight, secondCharShape };
