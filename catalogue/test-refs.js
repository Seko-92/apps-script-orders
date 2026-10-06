// node catalogue/test-refs.js — short-REF sequence repair (older Kubota books number REFs 1, 2, 3 …)
"use strict";
const { repairShortRefs } = require("./lib/kubota-scan");
let fail = 0;
function t(name, refs, want) {
  const p = refs.map(r => ({ ref: r })); repairShortRefs(p);
  const got = p.map(x => x.ref).join(","), ok = got === want.join(",");
  if (!ok) fail++;
  console.log((ok ? "ok   " : "FAIL ") + name + (ok ? "" : "\n     got  " + got + "\n     want " + want.join(",")));
}
t("Z400 E01 p7 read", ["1","2","2","","","3","3","4","5","6","71","8","8","9","10","11","","","14","15","","","","18","18","","20","","22","","24"],
                      ["1","2","2","","","3","3","4","5","6","7","8","8","9","10","11","12","13","14","15","","","","18","18","19","20","21","22","23","24"]);
t("2 _ _ 3 is undecidable — left blank", ["2","","","3"], ["2","","","3"]);
t("gap between equal REFs", ["3","","3","4"], ["3","3","3","4"]);
t("too wide a gap — no guess", ["1","","5","6"], ["1","","5","6"]);
t("3-digit book untouched", ["010","020","","030"], ["010","020","","030"]);
t("trailing gap — no next REF", ["3","4","",""], ["3","4","",""]);
t("3-digit noise in a short section", ["6","734","8"], ["6","7","8"]);
t("misread with two fitting digits — no guess", ["1","12","2"], ["1","12","2"]);
t("noise that fits nothing → blank, then filled", ["7","8","670","10"], ["7","8","9","10"]);
t("lone digit out of order → blank, then filled", ["1","9","3","4"], ["1","2","3","4"]);
t("a restart is left alone", ["11","12","1","1","1"], ["11","12","1","1","1"]);
t("digit with 15 fused on", ["2","415","5","6"], ["2","4","5","6"]);
t("3-digit at a page edge → blank", ["211","22","23"], ["","22","23"]);
// quantity vote: a fragment backs the column read (Z400 p15 row 7: col 10 · raw 19 · word 0 → 10)
const { voteQty } = require("./lib/scan-book");
const v = (name, got, want) => { const ok = got === want; if (!ok) fail++; console.log((ok ? "ok   " : "FAIL ") + name + (ok ? "" : " got " + got + " want " + want)); };
v("fragment backs the column read", voteQty(null, 19, 10, 0), 10);
v("no column read → unchanged (raw wins)", voteQty(null, 1, null, 12), 1);
v("plain majority unchanged", voteQty(null, 3, null, 3), 3);
// names: French spill cut even when OCR clipped the word; typewriter O read as 0
const { attachColumn } = require("./lib/kubota-scan");
const nm = t => { const rows = [{ y: 100, h: 40 }]; attachColumn(rows, [{ t, x: 0, y: 90, w: 10, h: 20, block: 1, par: 1, line: 1 }], "name"); return rows[0].name; };
v("clipped French word cut (CARTEI)", nm("COMP. CASE. GEAR CARTEI"), "COMP. CASE. GEAR");
v("clipped French word cut (BOUCH)", nm("PLUG BOUCH"), "PLUG");
v("0 RING → O RING", nm("0 RING"), "O RING");
v("English kept (GASKET,GEAR CASE)", nm("GASKET,GEAR CASE"), "GASKET,GEAR CASE");
console.log(fail ? fail + " FAILED" : "all passed"); process.exit(fail ? 1 : 0);
