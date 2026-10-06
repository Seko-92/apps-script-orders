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
console.log(fail ? fail + " FAILED" : "all passed"); process.exit(fail ? 1 : 0);
