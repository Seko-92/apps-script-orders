#!/usr/bin/env node
// Second-character check for OCR'd part numbers (see lib/glyph.js for why and how).
//
//   node catalogue/glyph-pass.js <out/ocr/BOOK.pdf.json> [...more]
//
// For every number whose 2nd character OCR read as 0, 4 or 6 ("16010-2395-0"), find it on its page
// and look at that glyph: a loop → the digit stands; open → it is the letter (C / G / J) — or a 5.
// Writes <BOOK>.pdf.glyph.json beside the cache: { "<page>|<pn as read>": "C" | "G" | "J" | "loop" | null }.
// import.js applies it BEFORE lib/pn-fix.js. The OCR cache itself is never modified.
// One book per process; run several with xargs -P (tesseract is single-threaded here).
"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");
const { tsvWords } = require("./lib/kubota-scan");
const { secondCharShape } = require("./lib/glyph");

const DPI = 400;
const CANDIDATE = /^1[046]\d{3}-\d{4}-\d$/;
const LEFT = 0.55;            // the part-number column sits in the left half on every layout seen

function pageRead(pdf, page, tmp) {
  const base = path.join(tmp, "g" + page);
  execFileSync("pdftoppm", ["-f", String(page), "-l", String(page), "-r", String(DPI), "-gray", "-singlefile", pdf, base], { stdio: "ignore" });
  const b = fs.readFileSync(base + ".pgm");
  const m = b.toString("latin1", 0, 64).match(/^P5\s+(\d+)\s+(\d+)\s+(\d+)\s/);
  const R = { W: +m[1], H: +m[2], d: b.subarray(m[0].length) };
  // tesseract on the left part only (same origin → word boxes are page coordinates)
  const cw = Math.round(R.W * LEFT), body = Buffer.alloc(cw * R.H);
  for (let r = 0; r < R.H; r++) R.d.copy(body, r * cw, r * R.W, r * R.W + cw);
  const crop = base + "-l.pgm";
  fs.writeFileSync(crop, Buffer.concat([Buffer.from("P5\n" + cw + " " + R.H + "\n255\n", "latin1"), body]));
  const tsv = execFileSync("tesseract", [crop, "-", "--psm", "6", "tsv"],
    { maxBuffer: 32 << 20, stdio: ["ignore", "pipe", "ignore"], env: Object.assign({}, process.env, { OMP_THREAD_LIMIT: "1" }) }).toString();
  fs.unlinkSync(base + ".pgm"); fs.unlinkSync(crop);
  return { R, words: tsvWords(tsv) };
}

function runBook(cacheFile) {
  const book = JSON.parse(fs.readFileSync(cacheFile, "utf8"));
  const byPage = new Map();
  book.sections.forEach(s => s.parts.forEach(p => {
    if (!CANDIDATE.test(p.pn || "")) return;
    if (!byPage.has(p.page)) byPage.set(p.page, new Set());
    byPage.get(p.page).add(p.pn);
  }));
  const out = {}, tally = { C: 0, G: 0, J: 0, "5": 0, loop: 0, unsure: 0, notFound: 0 };
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "hqglyph-"));
  const t0 = Date.now();
  try {
    for (const [page, pns] of [...byPage.entries()].sort((a, b) => a[0] - b[0])) {
      let pr;
      try { pr = pageRead(book.file, page, tmp); } catch (e) { pns.forEach(pn => { out[page + "|" + pn] = null; tally.notFound++; }); continue; }
      for (const pn of pns) {
        const v = secondCharShape(pr.R, pr.words, pn);
        const k = v ? v.kind : null;
        out[page + "|" + pn] = k;
        if (!v) tally.notFound++; else if (!k) tally.unsure++; else tally[k]++;
      }
    }
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
  const dest = cacheFile.replace(/\.json$/, ".glyph.json");
  fs.writeFileSync(dest, JSON.stringify(out, null, 0));
  console.log(`${path.basename(cacheFile)}: ${byPage.size} pages · ${Math.round((Date.now() - t0) / 1000)} s · ` +
    `C ${tally.C} · G ${tally.G} · J ${tally.J} · 5 ${tally["5"]} · digit ${tally.loop} · unsure ${tally.unsure} · not found ${tally.notFound}`);
}

process.argv.slice(2).forEach(runBook);
