// Build the files the VPS serves:  node catalogue/bundle.js [--pdf-dir <folder>]
//   catalogue/dist/catalogue.html            → /opt/hq-app/catalogue.html   (served at /catalogue)
//   catalogue/dist/catalogue-data/…          → /opt/hq-app/catalogue-data/
//       index.json · <ENGINE>.json · drawings/<ENGINE>/<code>.png · pdf/<ENGINE>.pdf
// ⚠ The data folder is NOT named "catalogue/": Caddy's try_files {path} {path}.html would then
//   meet a directory at /catalogue before it reaches catalogue.html.
// No stock is baked in — the live page asks Apps Script (catalogueStock) when an engine opens.
"use strict";
const fs = require("fs"), path = require("path");

const out = path.join(__dirname, "out");
const dist = path.join(__dirname, "dist");
const a = process.argv.slice(2);
const pdfDir = a.includes("--pdf-dir") ? a[a.indexOf("--pdf-dir") + 1] : null;

fs.rmSync(dist, { recursive: true, force: true });
const data = path.join(dist, "catalogue-data");
fs.mkdirSync(path.join(data, "pdf"), { recursive: true });

fs.copyFileSync(path.join(__dirname, "page", "catalogue.html"), path.join(dist, "catalogue.html"));
fs.copyFileSync(path.join(out, "index.json"), path.join(data, "index.json"));

function findPdf(dir, name) {
  for (const n of fs.readdirSync(dir)) {
    const p = path.join(dir, n);
    if (fs.statSync(p).isDirectory()) { const f = findPdf(p, name); if (f) return f; }
    else if (n === name) return p;
  }
  return null;
}

let bytes = 0, pdfs = 0, drawings = 0;
for (const e of JSON.parse(fs.readFileSync(path.join(out, "index.json"), "utf8"))) {
  const src = path.join(out, e.id + ".json");
  fs.copyFileSync(src, path.join(data, e.id + ".json"));
  const d = JSON.parse(fs.readFileSync(src, "utf8"));
  Object.values(d.drawings || {}).forEach(v => {
    if (!v || !v.img) return;
    const to = path.join(data, v.img);
    fs.mkdirSync(path.dirname(to), { recursive: true });
    fs.copyFileSync(path.join(out, v.img), to); drawings++;
  });
  if (pdfDir) {
    const f = findPdf(pdfDir, d.source.file);
    if (f) { fs.copyFileSync(f, path.join(data, "pdf", e.id + ".pdf")); pdfs++; }
    else console.log("  ⚠ source PDF not found for " + e.id + ": " + d.source.file);
  }
}
(function size(p) { for (const n of fs.readdirSync(p)) { const q = path.join(p, n); const st = fs.statSync(q); st.isDirectory() ? size(q) : (bytes += st.size); } })(dist);
console.log(`dist ready: ${drawings} drawings · ${pdfs} PDFs · ${(bytes / 1e6).toFixed(1)} MB → ${dist}`);
