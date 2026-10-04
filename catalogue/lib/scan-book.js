// OCR a whole scanned Kubota book → the same shape parseKubota() returns.
// One page at a time: pdftoppm (400 dpi, grey) → tesseract TSV → readScanPage().
"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");
const { tsvWords, readScanPage, attachColumn, rescueRows } = require("./kubota-scan");

const DPI = 400;

function ocrPage(pdf, page, tmp) {
  const base = path.join(tmp, "p" + page);
  execFileSync("pdftoppm", ["-f", String(page), "-l", String(page), "-r", String(DPI), "-gray", "-png", "-singlefile", pdf, base],
    { stdio: "ignore" });
  const png = base + ".png";
  const b = fs.readFileSync(png);
  const W = b.readUInt32BE(16), H = b.readUInt32BE(20);
  const tsv = execFileSync("tesseract", [png, "-", "--psm", "6", "tsv"],
    { maxBuffer: 32 << 20, stdio: ["ignore", "pipe", "ignore"], env: Object.assign({}, process.env, { OMP_THREAD_LIMIT: "1" }) }).toString();
  fs.unlinkSync(png);
  return { words: tsvWords(tsv), W, H };
}

/**
 * OCR one column of the page (box in 400-dpi page px) at `dpi`, returning words with
 * coordinates mapped back to 400-dpi page px so they line up with the page read.
 */
function ocrColumn(pdf, page, box, dpi, tmp, extra, mode) {
  const k = dpi / DPI;
  const base = path.join(tmp, "c" + page + "-" + Math.round(box.x));
  execFileSync("pdftoppm", ["-f", String(page), "-l", String(page), "-r", String(dpi), "-gray",
    "-x", String(Math.round(box.x * k)), "-y", String(Math.round(box.y * k)),
    "-W", String(Math.round(box.w * k)), "-H", String(Math.round(box.h * k)), "-singlefile", pdf, base], { stdio: "ignore" });
  const img = base + ".pgm";
  const pad = wipeRules(img, mode) || 0;
  const tsv = execFileSync("tesseract", [img, "-", "--psm", "6"].concat(extra || []).concat(["tsv"]),
    { maxBuffer: 16 << 20, stdio: ["ignore", "pipe", "ignore"], env: Object.assign({}, process.env, { OMP_THREAD_LIMIT: "1" }) }).toString();
  fs.unlinkSync(img);
  return tsvWords(tsv).map(w => Object.assign(w, { x: box.x + (w.x - pad) / k, y: box.y + (w.y - pad) / k, w: w.w / k, h: w.h / k }));
}

/**
 * Erase table rules from a greyscale PGM in place: a pixel row that is mostly ink is a
 * horizontal rule (solid or dashed); a pixel column that is mostly ink is a vertical rule.
 * Text never fills that much of a row or column, so it survives.
 */
function wipeRules(file, mode) {
  const b = fs.readFileSync(file);
  const m = b.toString("latin1", 0, 64).match(/^P5\s+(\d+)\s+(\d+)\s+(\d+)\s/);
  if (!m) return 0;
  const W = +m[1], H = +m[2], off = m[0].length, dark = 128;
  // Horizontal rules (solid or dashed): a pixel row that is mostly ink.
  for (let y = 0; y < H; y++) {
    let n = 0; for (let x = 0; x < W; x++) if (b[off + y * W + x] < dark) n++;
    if (n > 0.32 * W) for (let x = 0; x < W; x++) b[off + y * W + x] = 255;
  }
  // Vertical rules: a pixel column that is mostly ink. ⚠ NOT for the part-number column —
  // every number there starts with a digit at the same x, and that stack of "1"s reads as a
  // rule and gets erased (measured 2026-10-03). That crop starts after the rule instead.
  if (mode !== "pn") {
    for (let x = 0; x < W; x++) {
      let n = 0; for (let y = 0; y < H; y++) if (b[off + y * W + x] < dark) n++;
      if (n > 0.45 * H) for (let y = 0; y < H; y++) b[off + y * W + x] = 255;
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
  try {
    for (const p of pages) {
      const { words, W, H } = ocrPage(pdf, p, tmp);
      if (p <= 4) coverText += " " + words.map(w => w.t).join(" ");
      const r = readScanPage(words, W, H);
      if (r.rows.length) {
        const L = r.layout, top = L.headerBottom - 20, hgt = L.tableBottom - top;
        // 1. part-number column alone, restricted alphabet → rows the page read skipped
        const pnEdge = Math.min(0.26 * W, Math.max(...words.filter(w => r.rows.some(x => x.pn && Math.abs(x.y - (w.y + w.h / 2)) < 30 && w.x < 0.32 * W && /\d{4}/.test(w.t))).map(w => w.x + w.w)) + 20);
        // start just left of the part numbers, NOT at the REF column: with the rule between
        // them wiped, "010" and "16423-…" run together and a digit is lost
        const pnLeft = Math.min(...words.filter(w => r.rows.some(x => Math.abs(x.y - (w.y + w.h / 2)) < 30) && w.x < 0.32 * W && /\d{4}/.test(w.t)).map(w => w.x)) - 12;
        const x0 = isFinite(pnLeft) ? pnLeft : L.pnX0;
        const pnCol = ocrColumn(pdf, p, { x: x0, y: top, w: (isFinite(pnEdge) ? pnEdge : 0.22 * W) - x0, h: hgt }, DPI, tmp,
          ["-c", "tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-"], "pn");
        r.rows = rescueRows(r.rows, words, pnCol, L);
        // 2. names: their own crop (the page read loses whole rows in denser books)
        const nameLeft = isFinite(pnEdge) ? pnEdge + 4 : 0.22 * W;
        attachColumn(r.rows, ocrColumn(pdf, p, { x: nameLeft, y: top, w: L.nameEnd - nameLeft, h: hgt }, 600, tmp), "name");
        // 3. quantity and 4. remarks (sizes) — printed small, so read at 600 dpi
        attachColumn(r.rows, ocrColumn(pdf, p, { x: L.stueckX, y: top, w: L.remarksX - L.stueckX - 20, h: hgt }, 600, tmp,
          ["-c", "tessedit_char_whitelist=0123456789-"]), "qty");
        attachColumn(r.rows, ocrColumn(pdf, p, { x: L.remarksX, y: top, w: 0.98 * W - L.remarksX, h: hgt }, 600, tmp,
          ["-c", "tessedit_char_whitelist=0123456789+-.STDEmMSOVRIZEABCHGKLNPU/"]), "remark");
      }
      if (opts.log) opts.log(p, r);
      if (r.isIndex && out.sections.length) break;     // the index at the END, never the contents page
      if (!r.rows.length) continue;
      if (r.model && !out.models.length) out.models.push({ col: "A", name: r.model });
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
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }

  const mc = coverText.match(/\b([A-Z]\d{3,4}[A-Z0-9-]*-E\dB[A-Z0-9-]*|[A-Z]\d{3,4}-[A-Z0-9-]{3,})\s+([0-9A-Z]{5}-\d{5})\b/);
  if (mc) { out.model = mc[1]; out.codeNo = mc[2]; }
  if (!out.model && out.models.length) out.model = out.models[0].name;
  const mv = coverText.match(/\(\s*\d{2}\.\d{2}\.\d{4}\s*-->\s*[\d.]*\s*\)/);
  if (mv) out.validity = mv[0].replace(/\s+/g, " ");
  return out;
}

module.exports = { ocrBook };
