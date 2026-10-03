// Quick look: node catalogue/try.js <pdf>
const { execFileSync } = require("child_process");
const { parseKubota } = require("./lib/kubota");
const { parseKpad } = require("./lib/kpad");
const f = process.argv[2];
const txt = execFileSync("pdftotext", ["-layout", f, "-"], { maxBuffer: 64 << 20 }).toString();
const r = /kpadweb\.kubota/.test(txt) ? parseKpad(txt, f) : parseKubota(txt, f);
const n = r.sections.reduce((a, s) => a + s.parts.length, 0);
const pns = new Set(r.sections.flatMap(s => s.parts.map(p => p.pn)));
console.log(r.model, r.codeNo, r.validity, JSON.stringify(r.models));
console.log(`${r.sections.length} sections · ${n} lines · ${pns.size} distinct PNs · ${r.flags.length} flags`);
if (process.argv[3]) {
  const s = r.sections.find(s => s.code === process.argv[3]);
  console.log(s.name, "p" + s.page); s.parts.forEach(p => console.log(` ${p.ref} ${p.pn.padEnd(14)} ${p.name.padEnd(22)} ${String(p.qty).padEnd(6)} ${p.remark}`));
}
r.flags.slice(0, 15).forEach(f => console.log(" FLAG", f.kind, "p" + f.page, f.line));
