// Engine Catalogue importer — runs LOCALLY (pdftotext lives here, the VPS has none).
//
//   node catalogue/import.js <pdf-or-folder>... [--mi <mi.json>] [--out catalogue/out]
//
// For every PDF: pick the parser by what the text looks like, extract, validate, then
// write one JSON per engine plus report.md. Nothing is published from here — the report
// is for a human to approve first (plan: ~/.claude/plans/engine-catalogue.md).
//
// --mi takes a {headers, rows} export of Master Inventory. It only feeds the REPORT
// ("we carry 48 of 318"); the page itself looks stock up live.

"use strict";

const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { execFileSync } = require("child_process");
const { parseKubota } = require("./lib/kubota");
const { parseKpad } = require("./lib/kpad");

// ---- args -------------------------------------------------------------------------------
const args = process.argv.slice(2);
const opt = { mi: null, out: path.join(__dirname, "out"), inputs: [] };
for (let i = 0; i < args.length; i++) {
  if (args[i] === "--mi") opt.mi = args[++i];
  else if (args[i] === "--out") opt.out = args[++i];
  else opt.inputs.push(args[i]);
}
if (!opt.inputs.length) { console.error("usage: node catalogue/import.js <pdf|folder>... [--mi mi.json]"); process.exit(1); }

function listPdfs(p) {
  const st = fs.statSync(p);
  if (st.isFile()) return /\.pdf$/i.test(p) ? [p] : [];
  return fs.readdirSync(p).flatMap(n => listPdfs(path.join(p, n)));
}
const files = opt.inputs.flatMap(listPdfs).sort();

// ---- the Parts Finder's own matching code, so the two can never disagree ------------------
const mpn = (() => {
  const src = fs.readFileSync(path.join(__dirname, "..", "MpnSearch.js"), "utf8");
  const ctx = { console };
  vm.createContext(ctx);
  vm.runInContext(src, ctx);
  return ctx;
})();

let miIndex = null, miRows = null, miIdx = null;
if (opt.mi) {
  const mi = JSON.parse(fs.readFileSync(opt.mi, "utf8"));
  miRows = mi.rows;
  miIdx = Object.fromEntries(mi.headers.map((h, i) => [h, i]));
  const mpnCols = mi.headers.map((h, i) => ({ name: h, off: i })).filter(c => mpn.MPN_SEARCH.colPattern.test(c.name));
  miIndex = mpn._mpnBuildIndex(miRows, mpnCols);
}
function weCarry(pn) {
  if (!miIndex) return null;
  const hits = miIndex.get(mpn._mpnKey(pn)) || [];
  const rows = hits.map(h => miRows[h.row]);
  const active = rows.filter(r => String(r[miIdx.listingStatus] || "Active") === "Active");
  return { active: active.map(r => mpn._mpnCellText(r[miIdx.sku])), ended: rows.length - active.length };
}

// ---- classify + parse ---------------------------------------------------------------------
function pdfText(f) {
  return execFileSync("pdftotext", ["-layout", f, "-"], { maxBuffer: 256 << 20, stdio: ["ignore", "pipe", "ignore"] }).toString();
}
function pdfPages(f) {
  try { return Number((execFileSync("pdfinfo", [f], { stdio: ["ignore", "pipe", "ignore"] }).toString().match(/^Pages:\s+(\d+)/m) || [])[1]) || 0; }
  catch (e) { return 0; }
}

const MACHINE_PAGES = 150;     // engine parts lists run ~50-90 pages; whole-machine books 200+
const MIN_LINES = 50;          // fewer readable part lines than this = a scan → OCR batch

const results = [];            // one per file
for (const f of files) {
  const rec = { file: f, name: path.basename(f), pages: pdfPages(f) };
  let txt = "";
  try { txt = pdfText(f); } catch (e) { rec.status = "unreadable"; rec.why = String(e.message).slice(0, 80); results.push(rec); continue; }

  const isKpad = /kpadweb\.kubota/.test(txt);
  const isPartsBook = isKpad || /\bREF\.No\.|\bPos\.Nr\./.test(txt);
  if (!isPartsBook && /\bItem\s+Part No\.\s+Qty\.?\s+Description/.test(txt)) {
    rec.status = "oem-parts-list"; results.push(rec); continue;
  }
  if (!isPartsBook) {
    rec.status = /workshop|repair|service|shop manual|operation/i.test(rec.name + " " + txt.slice(0, 3000)) ? "not-a-parts-list" : "no-text";
    results.push(rec); continue;
  }
  if (rec.pages > MACHINE_PAGES && !isKpad) { rec.status = "machine-manual"; results.push(rec); continue; }

  const r = isKpad ? parseKpad(txt, f) : parseKubota(txt, f);
  const lines = r.sections.reduce((a, s) => a + s.parts.length, 0);
  rec.parsed = r;
  rec.lines = lines;
  rec.status = lines < MIN_LINES ? "mostly-scanned" : "ok";
  results.push(rec);
}

// ---- dedupe: same model (+ code number) from two files → keep the fuller one ----------------
const byModel = new Map();
for (const r of results.filter(r => r.status === "ok")) {
  const key = (r.parsed.model + "|" + r.parsed.codeNo).toUpperCase();
  const prev = byModel.get(key);
  if (!prev) { byModel.set(key, r); continue; }
  const [keep, drop] = r.lines > prev.lines ? [r, prev] : [prev, r];
  drop.status = "duplicate"; drop.dupOf = keep.name;
  byModel.set(key, keep);
}

// ---- write ------------------------------------------------------------------------------
fs.mkdirSync(opt.out, { recursive: true });
const slug = s => s.toUpperCase().replace(/[^A-Z0-9]+/g, "-").replace(/^-|-$/g, "");
const engines = [];
for (const r of byModel.values()) {
  const p = r.parsed;
  const id = slug(p.model);
  const doc = {
    id, brand: p.brand, model: p.model, codeNo: p.codeNo, validity: p.validity, models: p.models,
    source: { file: r.name, pages: r.pages, parser: p.source || "kubota-book" },
    sections: p.sections.map(s => ({ code: s.code, name: s.name, page: s.page, parts: s.parts })),
    flags: p.flags,
    importedAt: new Date().toISOString()
  };
  fs.writeFileSync(path.join(opt.out, id + ".json"), JSON.stringify(doc));

  // stats for the report
  const pns = new Map();
  p.sections.forEach(s => s.parts.forEach(x => { if (!pns.has(x.pn)) pns.set(x.pn, { part: x, section: s }); }));
  let carried = 0;
  const perSection = p.sections.map(s => {
    const set = new Set(s.parts.map(x => x.pn));
    let c = 0;
    set.forEach(pn => { const w = weCarry(pn); if (w && w.active.length) c++; });
    return { code: s.code, name: s.name, total: set.size, carried: c };
  });
  pns.forEach((v, pn) => { const w = weCarry(pn); if (w && w.active.length) carried++; });
  engines.push({ id, model: p.model, file: r.name, pages: r.pages, sections: p.sections.length,
                 lines: r.lines, distinct: pns.size, carried, perSection, flags: p.flags });
}
engines.sort((a, b) => a.model.localeCompare(b.model));
fs.writeFileSync(path.join(opt.out, "index.json"), JSON.stringify(engines.map(e => ({
  id: e.id, model: e.model, sections: e.sections, parts: e.distinct
}))));

// ---- report -------------------------------------------------------------------------------
const L = [];
L.push(`# Catalogue import report — ${new Date().toISOString().slice(0, 16).replace("T", " ")}`, "");
L.push(`${files.length} PDF(s) read · ${engines.length} engine(s) imported` + (opt.mi ? ` · stock from ${path.basename(opt.mi)}` : ""), "");
L.push("## Imported", "", "| Engine | Sections | Parts | We carry | Flags | File |", "|---|---|---|---|---|---|");
engines.forEach(e => L.push(`| ${e.model} | ${e.sections} | ${e.distinct} | ${opt.mi ? e.carried : "—"} | ${e.flags.length} | ${e.file} |`));
L.push("");
engines.forEach(e => {
  L.push(`### ${e.model}`, "");
  if (opt.mi) L.push("Per section (distinct part numbers · we carry):", "", e.perSection.map(s => `- ${s.code} ${s.name} — ${s.total} · **${s.carried}**`).join("\n"), "");
  if (e.flags.length) {
    L.push("To check:", "");
    const label = { "section-image-only": "image page — needs OCR", "section-missing": "in contents, not found",
                    "pn-shape": "odd part number", "no-name": "no part name", "pn-split": "part number split across lines",
                    "no-section": "lines with no section" };
    e.flags.forEach(f => L.push(`- ${label[f.kind] || f.kind}${f.page ? " · page " + f.page : ""} · \`${f.line}\``));
    L.push("");
  }
});
const groups = { "machine-manual": "Whole-machine manuals (engine is one section) — skipped for now",
                 "mostly-scanned": "Mostly scanned — OCR batch", "no-text": "No text layer (scans) — OCR batch",
                 "oem-parts-list": "Machine-maker parts lists (OEM numbers, e.g. Bobcat) — own parser later",
                 "not-a-parts-list": "Not a parts list (workshop/operation) — skipped",
                 "duplicate": "Duplicates (kept the fuller copy)", "unreadable": "Could not be read" };
Object.keys(groups).forEach(k => {
  const g = results.filter(r => r.status === k);
  if (!g.length) return;
  L.push(`## ${groups[k]} (${g.length})`, "");
  g.forEach(r => L.push(`- ${r.name}${r.pages ? " · " + r.pages + " pp" : ""}${r.lines != null ? " · " + r.lines + " lines read" : ""}${r.dupOf ? " · duplicate of " + r.dupOf : ""}`));
  L.push("");
});
fs.writeFileSync(path.join(opt.out, "report.md"), L.join("\n"));

console.log(`imported ${engines.length} engine(s) → ${opt.out}`);
engines.forEach(e => console.log(`  ${e.model.padEnd(24)} ${String(e.distinct).padStart(4)} parts · carry ${opt.mi ? e.carried : "—"} · ${e.flags.length} flag(s)`));
const counts = {};
results.forEach(r => { counts[r.status] = (counts[r.status] || 0) + 1; });
console.log("  by status:", JSON.stringify(counts));
