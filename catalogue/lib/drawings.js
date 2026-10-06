// Exploded-view drawings ("blueprints") for a Kubota parts book.
//
// Each section's first page carries the drawing on top and the parts table below
// (measured on V2203-M-E2B, 2026-10-03). The drawing is a scanned image, and its callouts
// are the table's REF numbers (010, 020 …), so:
//   1. find the drawing band on the page: below the section title, above "REF.No."
//      (word boxes from `pdftotext -bbox-layout`);
//   2. render that band at 300 dpi and OCR it (tesseract, digits only) → callout boxes;
//   3. render the same band at 150 dpi, 1-bit, as the image the page shows.
// Callouts are stored as fractions of the image (0..1), so any display size works.
// A callout OCR misses simply has no highlight — the drawing still shows.

"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");

const OCR_PASSES_DPI = [300, 400], OCR_PSMS = ["11", "6"], WEB_DPI = 300;   // 1-bit at 150 broke thin lines (2026-10-03) — shapes matter

function wordBoxes(pdf, page) {
  const xml = execFileSync("pdftotext", ["-bbox-layout", "-f", String(page), "-l", String(page), pdf, "-"],
    { maxBuffer: 32 << 20, stdio: ["ignore", "pipe", "ignore"] }).toString();
  const pg = xml.match(/<page width="([\d.]+)" height="([\d.]+)"/);
  const words = [...xml.matchAll(/<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)<\/word>/g)]
    .map(m => ({ x0: +m[1], y0: +m[2], x1: +m[3], y1: +m[4], t: m[5] }));
  return { w: pg ? +pg[1] : 595, h: pg ? +pg[2] : 842, words };
}

/** The drawing band in PDF points, or null when the page has no room above the table. */
function drawingBand(box, sectionCode) {
  const ref = box.words.find(w => /^REF\.No\./.test(w.t));
  if (!ref) return null;
  const code = box.words.find(w => w.t === sectionCode && w.y0 < ref.y0);
  // the 3-line title block ends ~15 pt below the code line
  const top = code ? code.y1 + 16 : 70;
  // the model/qty header lines sit ~25 pt above REF.No.
  const bottom = ref.y0 - 28;
  if (bottom - top < 80) return null;           // table-only page
  return { x: 0, y: top, w: box.w, h: bottom - top };
}

/**
 * KPAD web printouts (D902): each section's page prints its title line TWICE — at the top, then the
 * drawing, then again just above the parts table ("D902-E4B-AVN-1 -> ENGINE -> 000300 CYLINDER HEAD").
 * The drawing is the band between the first title block (+ its "Update Date" line) and the second.
 */
function kpadDrawingBand(box, sectionCode) {
  const codes = box.words.filter(w => w.t === sectionCode).sort((a, b) => a.y0 - b.y0);
  if (codes.length < 2) return null;
  const upd = box.words.find(w => /^Update/.test(w.t) && w.y0 > codes[0].y0 && w.y0 < codes[1].y0);
  const top = (upd ? upd.y1 : codes[0].y1) + 6, bottom = codes[1].y0 - 8;
  if (bottom - top < 80) return null;
  return { x: 0, y: top, w: box.w, h: bottom - top };
}

function render(pdf, page, band, dpi, mono, outBase) {
  const s = dpi / 72;
  const args = ["-f", String(page), "-l", String(page), "-r", String(dpi),
    "-x", String(Math.round(band.x * s)), "-y", String(Math.round(band.y * s)),
    "-W", String(Math.round(band.w * s)), "-H", String(Math.round(band.h * s)),
    "-png", "-singlefile"];
  if (mono) args.push("-mono");
  execFileSync("pdftoppm", args.concat([pdf, outBase]), { stdio: "ignore" });
  return outBase + ".png";
}

function ocrCallouts(png, refs, imgW, imgH, psm) {
  let tsv = "";
  try {
    tsv = execFileSync("tesseract", [png, "-", "--psm", psm || "11", "-c", "tessedit_char_whitelist=0123456789", "tsv"],
      // ⚠ ONE thread: tesseract's default 4 OpenMP threads spin-wait, and with 4 workers on 4 cores
      //   (import.js runs engines in parallel) each call went from ~3 s to 5–6 MINUTES (2026-10-05)
      { maxBuffer: 16 << 20, stdio: ["ignore", "pipe", "ignore"], env: Object.assign({}, process.env, { OMP_THREAD_LIMIT: "1" }) }).toString();
  } catch (e) { return []; }
  const want = new Set(refs);
  const out = [];
  tsv.split("\n").slice(1).forEach(l => {
    const c = l.split("\t");
    if (c.length < 12) return;
    const t = c[11].trim(), x = +c[6], y = +c[7], w = +c[8], h = +c[9], conf = +c[10];
    // a 3-digit callout at 300 dpi is ~55-60 px wide; wider = two numbers merged
    if (!want.has(t) || conf < 40 || w > 2.8 * h) return;
    out.push({ ref: t, x: +(x / imgW).toFixed(4), y: +(y / imgH).toFixed(4), w: +(w / imgW).toFixed(4), h: +(h / imgH).toFixed(4) });
  });
  return out;
}

/** Union of two callout lists; the same ref at (nearly) the same spot counts once. */
function mergeCallouts(a, b) {
  const out = a.slice();
  b.forEach(c => {
    if (!out.some(o => o.ref === c.ref && Math.abs(o.x - c.x) < 0.02 && Math.abs(o.y - c.y) < 0.03)) out.push(c);
  });
  return out;
}

function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}

/**
 * Extract the drawing for every section of a parsed book.
 * @param {string} pdf     source file
 * @param {object} parsed  parseKubota() output
 * @param {string} outDir  where <id>/<code>.png is written
 * @param {string} id      engine id
 * @return {object} code → {img, w, h, callouts, found, refs}
 */
function extractDrawings(pdf, parsed, outDir, id) {
  const dir = path.join(outDir, "drawings", id);
  fs.mkdirSync(dir, { recursive: true });
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "hqcat-"));
  const result = {};
  for (const s of parsed.sections) {
    try {
      // a scanned book has no text layer: its OCR reader already measured the band
      const band = (parsed.bands && parsed.bands[s.code]) ||
        (parsed.source === "kpad" ? kpadDrawingBand(wordBoxes(pdf, s.page), s.code) : drawingBand(wordBoxes(pdf, s.page), s.code));
      if (!band) continue;
      const refs = [...new Set(s.parts.map(p => p.ref))];
      // Several OCR passes see different callouts (leader lines touch the digits); merge them.
      let callouts = [];
      for (const dpi of OCR_PASSES_DPI) {
        const hi = render(pdf, s.page, band, dpi, false, path.join(tmp, s.code + "-" + dpi));
        const hiSize = pngSize(hi);
        for (const psm of OCR_PSMS) callouts = mergeCallouts(callouts, ocrCallouts(hi, refs, hiSize.w, hiSize.h, psm));
      }
      const web = render(pdf, s.page, band, WEB_DPI, true, path.join(dir, s.code));
      const size = pngSize(web);
      result[s.code] = { img: path.relative(outDir, web), w: size.w, h: size.h, callouts,
                         found: new Set(callouts.map(c => c.ref)).size, refs: refs.length };
    } catch (e) {
      result[s.code] = { error: String(e.message).slice(0, 80) };
    }
  }
  fs.rmSync(tmp, { recursive: true, force: true });
  return result;
}

// Worker mode, so import.js can run several engines at once (a scanned engine takes ~6 min, nearly
// all of it tesseract on whole drawings — 4 workers on 4 cores is the win, 2026-10-05):
//   node lib/drawings.js --worker <task.json>   task = {pdf, parsed, outDir, id, result}
if (require.main === module && process.argv[2] === "--worker") {
  const t = JSON.parse(fs.readFileSync(process.argv[3], "utf8"));
  fs.writeFileSync(t.result, JSON.stringify(extractDrawings(t.pdf, t.parsed, t.outDir, t.id)));
}

module.exports = { extractDrawings, drawingBand, kpadDrawingBand, mergeCallouts };
