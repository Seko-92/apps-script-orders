// Engine Catalogue importer — runs LOCALLY (pdftotext lives here, the VPS has none).
//
//   node catalogue/import.js <pdf-or-folder>... [--mi <mi.json>] [--out catalogue/out] [--drawings | --keep-drawings]
//
// For every PDF: pick the parser by what the text looks like, extract, validate, then
// write one JSON per engine plus report.md. Nothing is published from here — the report
// is for a human to approve first (plan: ~/.claude/plans/engine-catalogue.md).
//
// MI EVIDENCE ("is this part number on one of our listings?") — proof for OCR fixes and the
// report's "we carry" counts. LIVE by default (2026-10-07): one doPost catalogueMiKeys call
// through /api/board, saved to out/mi-keys.json; offline → that file, with its age printed.
//   --mi <file>   override with a {headers, rows} MI export (the old way)
//   --mi-offline  skip the live fetch and use out/mi-keys.json
// The page itself always looks stock up live.

"use strict";

const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { execFileSync } = require("child_process");
const { parseKubota } = require("./lib/kubota");
const { parseKpad } = require("./lib/kpad");
const { fixMisreadPrefixes } = require("./lib/pn-fix");
const { dropSerialRangeModels } = require("./lib/models");

// ---- args -------------------------------------------------------------------------------
const args = process.argv.slice(2);
const opt = { mi: null, out: path.join(__dirname, "out"), inputs: [], drawings: false };
for (let i = 0; i < args.length; i++) {
  if (args[i] === "--mi") opt.mi = args[++i];
  else if (args[i] === "--mi-offline") opt.miOffline = true;
  else if (args[i] === "--out") opt.out = args[++i];
  else if (args[i] === "--drawings") opt.drawings = true;   // slow: OCR, a few minutes per engine
  // reuse each engine's drawings from the previous import (out/<ID>.json): they come from the page
  // images + REF numbers only, so a part-number fix doesn't change them — saves the ~45 min OCR
  else if (args[i] === "--keep-drawings") opt.keepDrawings = true;
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

// miKeys: Map key → 1 (on an active listing) | 0 (ended listings only). null = no evidence.
let miKeys = null, miSource = "";
const MI_KEYS_FILE = path.join(opt.out, "mi-keys.json");
const MI_API = "https://hq.yassinqurabi.com/api/board";
if (opt.mi) {
  const mi = JSON.parse(fs.readFileSync(opt.mi, "utf8"));
  const idx = Object.fromEntries(mi.headers.map((h, i) => [h, i]));
  const mpnCols = mi.headers.map((h, i) => ({ name: h, off: i })).filter(c => mpn.MPN_SEARCH.colPattern.test(c.name));
  miKeys = new Map();
  mpn._mpnBuildIndex(mi.rows, mpnCols).forEach((hits, key) => miKeys.set(key,
    hits.some(h => String(mi.rows[h.row][idx.listingStatus] || "Active") === "Active") ? 1 : 0));
  miSource = `${path.basename(opt.mi)} (export, override)`;
} else {
  let live = null;
  if (!opt.miOffline) {
    try {
      const out = execFileSync("curl", ["-s", "--max-time", "120", "-X", "POST", MI_API,
        "-H", "Content-Type: application/json", "--data", JSON.stringify({ action: "catalogueMiKeys" })],
        { maxBuffer: 64 << 20 }).toString();
      const j = JSON.parse(out);
      if (j && j.ok && j.keys && Object.keys(j.keys).length > 1000) live = j;
      else console.warn("⚠ live MI evidence refused: " + ((j && j.reason) || "too few keys"));
    } catch (e) { console.warn("⚠ live MI evidence unreachable: " + String(e.message || e).split("\n")[0]); }
  }
  if (live) {
    fs.mkdirSync(opt.out, { recursive: true });
    fs.writeFileSync(MI_KEYS_FILE, JSON.stringify({ at: live.at, rows: live.rows, keys: live.keys }));
    miKeys = new Map(Object.entries(live.keys));
    miSource = `live MI ${live.at.slice(0, 16).replace("T", " ")} UTC (${live.rows} rows, ${miKeys.size} numbers)`;
  } else if (fs.existsSync(MI_KEYS_FILE)) {
    const f = JSON.parse(fs.readFileSync(MI_KEYS_FILE, "utf8"));
    miKeys = new Map(Object.entries(f.keys));
    const days = ((Date.now() - Date.parse(f.at)) / 864e5).toFixed(1);
    miSource = `saved MI keys from ${f.at.slice(0, 10)} (${days} days old)`;
    console.warn(`⚠ using ${miSource} — numbers listed since then can't count as proof`);
  } else console.warn("⚠ no MI evidence at all — OCR fixes rely on the clean manuals only");
}
if (miSource) console.log("MI evidence: " + miSource);
const miKnown = pn => !!(miKeys && miKeys.has(mpn._mpnKey(pn)));
const weCarry = pn => !!(miKeys && miKeys.get(mpn._mpnKey(pn)) === 1);

// ---- classify + parse ---------------------------------------------------------------------
function pdfText(f) {
  return execFileSync("pdftotext", ["-layout", f, "-"], { maxBuffer: 256 << 20, stdio: ["ignore", "pipe", "ignore"] }).toString();
}
function pdfPages(f) {
  try { return Number((execFileSync("pdfinfo", [f], { stdio: ["ignore", "pipe", "ignore"] }).toString().match(/^Pages:\s+(\d+)/m) || [])[1]) || 0; }
  catch (e) { return 0; }
}

const MACHINE_PAGES = 150;     // engine parts lists run ~50-90 pages; whole-machine books 200+
/** "KUBOTA ENGINE D1302-BBS-1 PARTS MANUAL.pdf" → "D1302-BBS-1"; "V2203-M-E2B(1).pdf" → "V2203-M-E2B". */
function modelFromFileName(name) {
  const m = String(name).replace(/\.pdf$/i, "").replace(/\(\d+\)$/, "").toUpperCase()
    .match(/\b([A-Z]{1,2}\d{3,4}(?:[A-Z]{0,3})?(?:[-. ][A-Z0-9]+)*?)(?=\s+(?:PARTS?|SPARE|DIESEL)\b|\s*$)/);
  return m ? m[1].replace(/ /g, "-").replace(/\./g, "-") : "";
}
const OCR_BAR = { shape: 0.95, qty: 0.85, name: 0.85 };          // see the OCR block below
// A published book that slipped under the bar for a GOOD reason stays live — named here, with why.
const OCR_BAR_KEEP = {
  // 2026-10-06: the variant-column fix removed WRONG C quantities (the old layout read one column
  // left, inventing values); qty share 85.0% → 83.9% with the data now MORE correct. The cells it
  // misses print a serial range under the qty ("1" over "~489915") — an empty cell, not a wrong one.
  "D1105.pdf": "variant-column fix removed wrong quantities"
};
const OCR_PN_SHAPE = /^(\d[A-Z0-9]\d{3}-\d{4}-\d|\d{5}-\d{5})$/; // 1C010-5675-0 · 15221-1443-0 · 04814-10070
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

// The second character LOOKED AT on the page image (catalogue/glyph-pass.js → <book>.glyph.json):
// a loop means the digit is real; an open glyph is C / G / J (or a 5). A direct observation, so it
// runs before lib/pn-fix.js, and a "loop" verdict stops pn-fix from turning a real "16…" into "1C…".
// Only the direction the misread runs is changed: digit → letter (and 6 → 5).
const GLYPH_FROM = { C: "06", G: "06", J: "4", "5": "6" };
function applyGlyphVerdicts(book, file, tally, isKnown) {
  if (!fs.existsSync(file)) return 0;
  const v = JSON.parse(fs.readFileSync(file, "utf8"));
  book.sections.forEach(s => s.parts.forEach(p => {
    const k = v[p.page + "|" + p.pn];
    if (!k) return;
    if (k === "loop") { p.ocr = Object.assign({}, p.ocr, { glyph: "loop" }); tally.digit++; return; }
    if (!(GLYPH_FROM[k] || "").includes(p.pn[1])) return;
    // the shape is sure about letter-vs-digit; C vs G on a small or faint print is its weak spot
    // (D1503 p15: 1G841 measured as C). If only the OTHER of the two is a known number, take that.
    const as = L => p.pn[0] + L + p.pn.slice(2), alt = { C: "G", G: "C" }[k];
    let L = k;
    if (alt && isKnown && !isKnown(as(k)) && isKnown(as(alt))) { L = alt; tally.swapped++; }
    // the two reads that agreed agreed on the wrong character → not confirmed by "agree" any more
    p.ocr = Object.assign({}, p.ocr, { glyph: L, glyphFixedFrom: p.pn, agree: false });
    p.pn = as(L);
    tally.fixed++;
  }));
  return 1;
}

// ---- OCR'd scans (catalogue/ocr-batch.js) ---------------------------------------------------
// A scan's cached parse replaces its "no text" verdict. Part numbers OCR read are CONFIRMED
// when the two independent reads agreed, or the number is known from a clean manual or from
// one of our listings; the rest are flagged for a human check (never silently trusted).
const ocrDir = path.join(opt.out, "ocr");
const textBooks = results.filter(r => r.status === "ok");
const knownName = new Map();                     // pn → name, from the clean manuals
textBooks.forEach(r => r.parsed.sections.forEach(s => s.parts.forEach(p => { if (p.name && !knownName.has(p.pn)) knownName.set(p.pn, p.name); })));
if (fs.existsSync(ocrDir)) {
  // load every OCR book first: the prefix fix (lib/pn-fix.js) needs evidence from ALL of them
  const isKnown = pn => knownName.has(pn) || miKnown(pn);
  const ocrBooks = [], glyphTally = { books: 0, fixed: 0, digit: 0, swapped: 0 };
  results.forEach(rec => {
    if (rec.status === "ok") return;
    const cache = path.join(ocrDir, rec.name + ".json");
    if (fs.existsSync(cache)) {
      rec.ocrParsed = JSON.parse(fs.readFileSync(cache, "utf8"));
      glyphTally.books += applyGlyphVerdicts(rec.ocrParsed, cache.replace(/\.json$/, ".glyph.json"), glyphTally, isKnown);
      ocrBooks.push({ parts: rec.ocrParsed.sections.flatMap(s => s.parts) });
    }
  });
  console.log(`glyph check: ${glyphTally.fixed} second characters corrected from the page image (${glyphTally.books} books), ${glyphTally.digit} confirmed as digits, ${glyphTally.swapped} C/G settled by a known number`);
  const pf = fixMisreadPrefixes(ocrBooks, isKnown);
  console.log(`prefix fix: ${pf.fixed} OCR part numbers corrected (known ${pf.byTier.known} · read ${pf.byTier.read} · stem ${pf.byTier.stem}), ${pf.left} "10…" left for a check`);
  results.forEach(rec => {
    if (rec.status === "ok") return;
    const cache = path.join(ocrDir, rec.name + ".json");
    if (!fs.existsSync(cache)) { if (fs.existsSync(cache.replace(/\.json$/, ".skip"))) rec.status = "scan-not-parts"; return; }
    const parsed = rec.ocrParsed; delete rec.ocrParsed;
    // the model comes from the FILE NAME for a scan: the cover read returns serial ranges ("<=15000"),
    // a French word ("ECHANGE") or another model column — and dedupe keys on the model, so junk
    // names merged 11 different books as "duplicates" (2026-10-05)
    parsed.model = modelFromFileName(rec.name) || parsed.model;
    dropSerialRangeModels(parsed);   // "<=15000" is a serial range, not a variant (lib/models.js)
    let check = 0, named = 0;
    parsed.sections.forEach(s => s.parts.forEach(p => {
      const known = isKnown(p.pn);
      p.ocr.confirmed = !!(p.ocr.agree || known);
      if (knownName.has(p.pn) && (!p.name || p.ocr.rescued || p.name.length < 3 || /[a-z]/.test(p.name))) { p.name = knownName.get(p.pn); named++; }
      if (!p.ocr.confirmed) check++;
    }));
    if (check) parsed.flags.push({ page: null, kind: "ocr-check", line: check + " part numbers read by OCR could not be confirmed — check them against the manual page" });
    rec.parsed = parsed; rec.ocr = { check, named };
    rec.lines = parsed.sections.reduce((a, s) => a + s.parts.length, 0);
    rec.status = rec.lines >= MIN_LINES ? "ok" : "mostly-scanned";
    // the publish bar for OCR'd books (agreed 2026-10-05): enough lines must have a Kubota-shaped
    // number, a quantity and a name — a book below it is reported, not published
    const all = parsed.sections.flatMap(s => s.parts), n = all.length || 1;
    const share = f => all.filter(f).length / n;
    rec.ocr.score = { shape: share(p => OCR_PN_SHAPE.test(p.pn || "")), qty: share(p => Array.isArray(p.qty) && p.qty.some(v => v != null)),
                      name: share(p => p.name && p.name.length > 2) };
    if (rec.status === "ok" && !(rec.ocr.score.shape >= OCR_BAR.shape && rec.ocr.score.qty >= OCR_BAR.qty && rec.ocr.score.name >= OCR_BAR.name)
        && !OCR_BAR_KEEP[rec.name])
      rec.status = "ocr-below-bar";
  });
}

// ---- dedupe: same model (+ code number) from two files → keep the fuller one ----------------
// a scan that is a copy of a clean text book ("V2203-M-E2B(1).pdf" beside "V2203-M-E2B.pdf"): the text
// book wins — its numbers are exact, the scan's are OCR. Its model came from the file name, so the
// model key below would not catch it.
const baseName = n => String(n).replace(/\.pdf$/i, "").replace(/\s*\(\d+\)$/, "").toUpperCase();
const textByBase = new Map(results.filter(r => r.status === "ok" && !r.ocr).map(r => [baseName(r.name), r]));
results.filter(r => r.status === "ok" && r.ocr).forEach(r => {
  const t = textByBase.get(baseName(r.name));
  if (t && t !== r) { r.status = "duplicate"; r.dupOf = t.name; }
});
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
// drawings for every engine first, DRAWING_JOBS at a time (xargs -P keeps this synchronous)
const DRAWING_JOBS = 4;
const drawingsById = {};
// with --keep-drawings, an engine that HAS drawings from the last import keeps them; only the ones
// without any are extracted (D902, added 2026-10-06 — it never had any)
// ⚠ an EMPTY drawings object is "none": Z400 / D850 / D1402 had {} from an import that found no
//   drawing pages, and a truthiness test kept that {} forever (2026-10-06)
const hasKept = id => { try { const d = JSON.parse(fs.readFileSync(path.join(opt.out, id + ".json"), "utf8")).drawings;
  return !!d && Object.keys(d).length > 0; } catch (e) { return false; } };
if (opt.drawings || opt.keepDrawings) {
  const tmpD = fs.mkdtempSync(path.join(require("os").tmpdir(), "hqdrw-"));
  const tasks = [];
  for (const r of byModel.values()) {
    const id = slug(r.parsed.model), task = path.join(tmpD, id + ".task.json");
    if (!opt.drawings && hasKept(id)) continue;
    fs.writeFileSync(task, JSON.stringify({ pdf: r.file, parsed: r.parsed, outDir: opt.out, id, result: path.join(tmpD, id + ".out.json") }));
    tasks.push({ id, task, result: path.join(tmpD, id + ".out.json") });
  }
  console.log(`drawings: ${tasks.length} engine(s), ${DRAWING_JOBS} at a time …`);
  const t0 = Date.now();
  // ⚠ xargs runs its command ONCE even on empty input — a worker with no task crashed the whole
  //   import (2026-10-06, --keep-drawings with every engine kept). Nothing to draw → skip.
  if (tasks.length) execFileSync("xargs", ["-0", "-P", String(DRAWING_JOBS), "-n", "1", process.execPath, path.join(__dirname, "lib", "drawings.js"), "--worker"],
    { input: tasks.map(t => t.task).join("\0"), stdio: ["pipe", "inherit", "inherit"], maxBuffer: 64 << 20 });
  tasks.forEach(t => { try { drawingsById[t.id] = JSON.parse(fs.readFileSync(t.result, "utf8")); } catch (e) { console.log(`  ! drawings ${t.id}: worker failed`); } });
  fs.rmSync(tmpD, { recursive: true, force: true });
  console.log(`drawings done in ${Math.round((Date.now() - t0) / 60000)} min`);
}
for (const r of byModel.values()) {
  const p = r.parsed;
  const id = slug(p.model);
  const doc = {
    id, brand: p.brand, model: p.model, codeNo: p.codeNo, validity: p.validity, models: p.models,
    source: { file: r.name, pages: r.pages, parser: p.source || "kubota-book", ocrCheck: r.ocr ? r.ocr.check : 0 },
    sections: p.sections.map(s => ({ code: s.code, name: s.name, page: s.page, parts: s.parts })),
    flags: p.flags,
    drawings: null,
    importedAt: new Date().toISOString()
  };
  if (!drawingsById[id] && opt.keepDrawings) {
    try { const prev = JSON.parse(fs.readFileSync(path.join(opt.out, id + ".json"), "utf8")); if (prev.drawings && Object.keys(prev.drawings).length) drawingsById[id] = prev.drawings; } catch (e) { /* new engine: none to keep */ }
  }
  if (drawingsById[id]) {
    doc.drawings = drawingsById[id];
    const v = Object.values(doc.drawings).filter(x => !x.error);
    console.log(`  drawings ${p.model}: ${v.length} · ${v.reduce((a, x) => a + x.found, 0)}/${v.reduce((a, x) => a + x.refs, 0)} refs marked`);
  }
  fs.writeFileSync(path.join(opt.out, id + ".json"), JSON.stringify(doc));

  // stats for the report
  const pns = new Map();
  p.sections.forEach(s => s.parts.forEach(x => { if (!pns.has(x.pn)) pns.set(x.pn, { part: x, section: s }); }));
  let carried = 0;
  const perSection = p.sections.map(s => {
    const set = new Set(s.parts.map(x => x.pn));
    let c = 0;
    set.forEach(pn => { if (weCarry(pn)) c++; });
    return { code: s.code, name: s.name, total: set.size, carried: c };
  });
  pns.forEach((v, pn) => { if (weCarry(pn)) carried++; });
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
L.push(`${files.length} PDF(s) read · ${engines.length} engine(s) imported` + (miKeys ? ` · MI evidence: ${miSource}` : ""), "");
L.push("## Imported", "", "| Engine | Sections | Parts | We carry | Flags | File |", "|---|---|---|---|---|---|");
engines.forEach(e => L.push(`| ${e.model} | ${e.sections} | ${e.distinct} | ${miKeys ? e.carried : "—"} | ${e.flags.length} | ${e.file} |`));
L.push("");
engines.forEach(e => {
  L.push(`### ${e.model}`, "");
  if (miKeys) L.push("Per section (distinct part numbers · we carry):", "", e.perSection.map(s => `- ${s.code} ${s.name} — ${s.total} · **${s.carried}**`).join("\n"), "");
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
                 "duplicate": "Duplicates (kept the fuller copy)",
                 "ocr-below-bar": "OCR'd scans below the publish bar (part-number shape ≥ 95% · qty ≥ 85% · name ≥ 85%) — not published",
                 "scan-not-parts": "Scans that are not Kubota parts books (OCR found no parts table)", "unreadable": "Could not be read" };
Object.keys(groups).forEach(k => {
  const g = results.filter(r => r.status === k);
  if (!g.length) return;
  L.push(`## ${groups[k]} (${g.length})`, "");
  const sc = r => r.ocr && r.ocr.score ? ` · shape ${Math.round(100 * r.ocr.score.shape)}% · qty ${Math.round(100 * r.ocr.score.qty)}% · name ${Math.round(100 * r.ocr.score.name)}%` : "";
  g.forEach(r => L.push(`- ${r.name}${r.pages ? " · " + r.pages + " pp" : ""}${r.lines != null ? " · " + r.lines + " lines read" : ""}${k === "ocr-below-bar" ? sc(r) : ""}${r.dupOf ? " · duplicate of " + r.dupOf : ""}`));
  L.push("");
});
fs.writeFileSync(path.join(opt.out, "report.md"), L.join("\n"));

console.log(`imported ${engines.length} engine(s) → ${opt.out}`);
engines.forEach(e => console.log(`  ${e.model.padEnd(24)} ${String(e.distinct).padStart(4)} parts · carry ${miKeys ? e.carried : "—"} · ${e.flags.length} flag(s)`));
const counts = {};
results.forEach(r => { counts[r.status] = (counts[r.status] || 0) + 1; });
console.log("  by status:", JSON.stringify(counts));
