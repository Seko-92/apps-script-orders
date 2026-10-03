// OCR every scanned Kubota parts book, 4 at a time, caching each result so a run can be
// stopped and resumed:   node catalogue/ocr-batch.js <folder> [--jobs 4]
// Writes catalogue/out/ocr/<pdf name>.json (the parsed book) and a .skip file for scans
// that turn out not to be parts lists. import.js picks the cache up.
"use strict";
const fs = require("fs"), path = require("path"), { execFileSync, spawn } = require("child_process");

const cacheDir = path.join(__dirname, "out", "ocr");
fs.mkdirSync(cacheDir, { recursive: true });

if (process.argv[2] === "--one") {                      // worker: one PDF
  const { ocrBook } = require("./lib/scan-book");
  const { tsvWords, readScanPage } = require("./lib/kubota-scan");
  const pdf = process.argv[3], out = path.join(cacheDir, path.basename(pdf) + ".json");
  const t0 = Date.now();
  const r = ocrBook(pdf, {});
  r.ms = Date.now() - t0;
  const lines = r.sections.reduce((a, s) => a + s.parts.length, 0);
  if (lines < 10) fs.writeFileSync(out.replace(/\.json$/, ".skip"), "not a Kubota parts book (" + lines + " lines)");
  else fs.writeFileSync(out, JSON.stringify(r));
  console.log(`${path.basename(pdf)} · ${r.sections.length} sections · ${lines} lines · ${(r.ms / 1000).toFixed(0)} s`);
  process.exit(0);
}

const root = process.argv[2];
const jobs = process.argv.includes("--jobs") ? Number(process.argv[process.argv.indexOf("--jobs") + 1]) : 4;
function list(p) { return fs.statSync(p).isDirectory() ? fs.readdirSync(p).flatMap(n => list(path.join(p, n))) : /\.pdf$/i.test(p) ? [p] : []; }
function hasText(f) {
  try { return execFileSync("pdftotext", ["-layout", f, "-"], { maxBuffer: 256 << 20, stdio: ["ignore", "pipe", "ignore"] }).toString().replace(/\s/g, "").length > 3000; }
  catch (e) { return false; }
}
function pages(f) { try { return Number((execFileSync("pdfinfo", [f], { stdio: ["ignore", "pipe", "ignore"] }).toString().match(/^Pages:\s+(\d+)/m) || [])[1]) || 0; } catch (e) { return 0; } }

// scans only: no text layer, parts-book sized, not a workshop manual by its name
const todo = list(root).filter(f => !/workshop|repair|service|shop manual|operation/i.test(path.basename(f)))
  .filter(f => { const n = pages(f); return n > 20 && n <= 150; })
  .filter(f => !hasText(f))
  .filter(f => !fs.existsSync(path.join(cacheDir, path.basename(f) + ".json")) && !fs.existsSync(path.join(cacheDir, path.basename(f) + ".skip")));
console.log(`${todo.length} scanned book(s) to OCR, ${jobs} at a time`);
let next = 0, running = 0, done = 0;
function launch() {
  while (running < jobs && next < todo.length) {
    const f = todo[next++]; running++;
    const p = spawn(process.execPath, [__filename, "--one", f], { stdio: ["ignore", "pipe", "pipe"] });
    p.stdout.on("data", d => process.stdout.write(`[${++done}/${todo.length}] ` + d));
    p.stderr.on("data", d => process.stdout.write(`  ! ${path.basename(f)}: ${String(d).split("\n")[0]}\n`));
    p.on("exit", () => { running--; launch(); if (!running && next >= todo.length) console.log("OCR batch finished"); });
  }
}
launch();
