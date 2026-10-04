// Build the review preview: the real page with the imported engines + a stock SNAPSHOT
// from an MI export baked in (the live page asks Apps Script instead).
//   node catalogue/build-preview.js --mi <mi.json> --out <file.html>
"use strict";
const fs = require("fs"), path = require("path"), vm = require("vm");

const a = process.argv.slice(2);
const mi = a[a.indexOf("--mi") + 1], outFile = a[a.indexOf("--out") + 1];
const outDir = path.join(__dirname, "out");

const ctx = { console }; vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "MpnSearch.js"), "utf8"), ctx);
const M = JSON.parse(fs.readFileSync(mi, "utf8"));
const H = Object.fromEntries(M.headers.map((h, i) => [h, i]));
const cols = M.headers.map((h, i) => ({ name: h, off: i })).filter(c => ctx.MPN_SEARCH.colPattern.test(c.name));
const index = ctx._mpnBuildIndex(M.rows, cols);

const num = v => (v == null || v === "" ? null : Number(v));
function stockFor(pn) {
  const hits = index.get(ctx._mpnKey(pn)) || [];
  return hits.map(h => M.rows[h.row])
    .filter(r => String(r[H.listingStatus] || "Active") === "Active")
    .map(r => ({
      sku: ctx._mpnCellText(r[H.sku]),
      title: String(r[H.title] || "").slice(0, 90),
      loc: String(r[H["C:Model Year"]] || ""),
      avail: (num(r[H.quantity]) || 0) - (num(r[H.quantitySold]) || 0),
      price: num(r[H.currentPrice]) || num(r[H.startPrice])
    }));
}

const engines = JSON.parse(fs.readFileSync(path.join(outDir, "index.json"), "utf8")).map(e => {
  const d = JSON.parse(fs.readFileSync(path.join(outDir, e.id + ".json"), "utf8"));
  const stock = {};
  d.sections.forEach(s => s.parts.forEach(p => { if (!(p.pn in stock)) stock[p.pn] = stockFor(p.pn); }));
  const drawings = {};
  Object.entries(d.drawings || {}).forEach(([code, v]) => {
    if (v.error || !v.img) return;
    const src = "data:image/png;base64," + fs.readFileSync(path.join(outDir, v.img)).toString("base64");
    drawings[code] = { src, w: v.w, h: v.h, callouts: v.callouts };
  });
  return { id: d.id, model: d.model, models: d.models || [], codeNo: d.codeNo, validity: d.validity, source: d.source,
           sections: d.sections, flags: d.flags, stock, drawings };
});

const tpl = fs.readFileSync(path.join(__dirname, "page", "catalogue.html"), "utf8");
const data = JSON.stringify({ snapshot: path.basename(mi), engines }).replace(/</g, "\\u003c");
fs.writeFileSync(outFile, tpl.replace("/*__CATALOGUE_DATA__*/null", data));
console.log("wrote", outFile, (fs.statSync(outFile).size / 1024).toFixed(0) + " KB");
