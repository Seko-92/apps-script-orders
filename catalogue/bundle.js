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
const pnIndex = {};                      // part-number key → [engine ids]  (for the Parts Finder)
for (const e of JSON.parse(fs.readFileSync(path.join(out, "index.json"), "utf8"))) {
  const src = path.join(out, e.id + ".json");
  fs.copyFileSync(src, path.join(data, e.id + ".json"));
  const d = JSON.parse(fs.readFileSync(src, "utf8"));
  addToPnIndex(pnIndex, e.id, d);
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
fs.writeFileSync(path.join(out, "pn-index.json"), JSON.stringify(pnIndex));
console.log(`pn index: ${Object.keys(pnIndex).length} part numbers → out/pn-index.json (send it with: node catalogue/push-index.js)`);
(function size(p) { for (const n of fs.readdirSync(p)) { const q = path.join(p, n); const st = fs.statSync(q); st.isDirectory() ? size(q) : (bytes += st.size); } })(dist);
console.log(`dist ready: ${drawings} drawings · ${pdfs} PDFs · ${(bytes / 1e6).toFixed(1)} MB → ${dist}`);

// ---------------------------------------------------------------------------------------
// PART-NUMBER INDEX (2026-10-09) — which engine manuals list a number. The Parts Finder
// shows "in N engine catalogues" from it. The KEY must equal MpnSearch.js _mpnKey():
// letters+digits only, uppercase, leading zeros stripped — or the two will never meet.
// Cross-references (p.xref, a superseding number) count too: the page's own search
// finds a part by them. Keys under 5 characters are noise (REF numbers, stray digits).
// ---------------------------------------------------------------------------------------
function pnKey(s) {
  const k = String(s == null ? "" : s).toUpperCase().replace(/[^A-Z0-9]/g, "");
  return k.replace(/^0+/, "") || k;
}
function addToPnIndex(idx, engineId, d) {
  (d.sections || []).forEach(s => (s.parts || []).forEach(p => [p.pn, p.xref].forEach(n => {
    const k = pnKey(n);
    if (k.length < 5 || !/\d/.test(k)) return;
    const list = idx[k] || (idx[k] = []);
    if (list.indexOf(engineId) < 0) list.push(engineId);
  })));
}
