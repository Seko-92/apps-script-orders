const fs = require("fs"), path = require("path"); const S = process.argv[2];
const picks = require(__dirname + "/picks.json");
let tot = 0, same = 0, filled = 0, emptied = 0, changed = 0; const lines = [];
picks.forEach((k, i) => {
  const fo = `${S}/reg/old-${i}.json`, fn = `${S}/reg/new-${i}.json`;
  if (!fs.existsSync(fo) || !fs.existsSync(fn)) { lines.push(`${path.basename(k.book)}: MISSING`); return; }
  const o = require(fo), n = require(fn); const key = r => r.page + "|" + r.pn;
  const om = new Map(); o.forEach(r => { const kk = key(r); om.set(kk, (om.get(kk) || []).concat([r])); });
  let b = { t: 0, s: 0, f: 0, e: 0, c: 0 }; const ex = [];
  n.forEach(r => {
    const l = om.get(key(r)); if (!l || !l.length) return; const p = l.shift();
    const a = JSON.stringify(p.qty), c = JSON.stringify(r.qty); b.t++;
    if (a === c) { b.s++; return; }
    const an = (p.qty || []).filter(v => v != null).length, cn = (r.qty || []).filter(v => v != null).length;
    const conflict = (p.qty || []).some((v, j) => v != null && r.qty && r.qty[j] != null && r.qty[j] !== v);
    if (conflict) b.c++; else if (cn > an) b.f++; else b.e++;
    ex.push(`   p${r.page} ${r.pn} ${a} → ${c}${conflict ? "  ⚠ CHANGED" : ""}`);
  });
  tot += b.t; same += b.s; filled += b.f; emptied += b.e; changed += b.c;
  lines.push(`${path.basename(k.book).slice(0, 44).padEnd(44)} rows ${b.t} · same ${b.s} · filled ${b.f} · emptied ${b.e} · changed ${b.c}`);
  ex.slice(0, 40).forEach(x => lines.push(x));
});
console.log(lines.join("\n"));
console.log(`TOTAL rows ${tot} · same ${same} · filled ${filled} · emptied ${emptied} · changed value ${changed}`);
