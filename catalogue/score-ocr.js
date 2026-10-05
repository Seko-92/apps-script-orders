// node score.js <dir>  — per-book OCR quality from the cached JSON
const fs = require("fs"), path = require("path"), dir = process.argv[2];
const SHAPE = /^(\d[A-Z0-9]\d{3}-\d{4}-\d|\d{5}-\d{5})$/;
const rows = [];
for (const f of fs.readdirSync(dir).filter(f => f.endsWith(".json")).sort()) {
  const b = JSON.parse(fs.readFileSync(path.join(dir, f)));
  const parts = b.sections.flatMap(s => s.parts), n = parts.length || 1;
  const pct = k => Math.round(100 * parts.filter(k).length / n);
  rows.push({ book: f.replace(/\.pdf\.json$/, "").replace(/^(KUBOTA ENGINE|Kubota Diesel Engine)\s*/i, "").slice(0, 34),
    sec: b.sections.length, untitled: b.sections.filter(s => s.code === "????").length, lines: parts.length,
    shape: pct(p => SHAPE.test(p.pn || "")), qty: pct(p => Array.isArray(p.qty) && p.qty.some(v => v != null)),
    name: pct(p => p.name && p.name.length > 2), ref: pct(p => p.ref), agree: pct(p => p.ocr && p.ocr.agree) });
}
console.log("book".padEnd(35) + "sec ???  lines shape  qty name  ref agree");
rows.forEach(r => console.log(r.book.padEnd(35) + String(r.sec).padStart(3) + String(r.untitled).padStart(4) + String(r.lines).padStart(7) +
  [r.shape, r.qty, r.name, r.ref, r.agree].map(v => String(v).padStart(5) + "%").join("")));
