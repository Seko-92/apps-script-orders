// node design-lab/test-catalogue-stock.js — Catalogue.js against the REAL MpnSearch.js helpers.
"use strict";
const fs = require("fs"), path = require("path"), vm = require("vm");
const ROOT = process.env.SRC || path.join(__dirname, "..");
const ctx = { console, MPN_SEARCH: null };
vm.createContext(ctx);
for (const f of ["MpnSearch.js", "Catalogue.js"]) vm.runInContext(fs.readFileSync(path.join(ROOT, f), "utf8"), ctx);

// fake MI: [sku, mpnMain, extraCol, title, active, avail, kit]
const ROWS = [
  ["167067", "1A091-21050", "", "Piston Rings STD For Kubota", true, 432, false],
  ["210176", "1G624-21112", "17331-21980, 1A091-21050", "Engine Overhaul Kit STD", true, 2, true],
  ["163044", "17331-21980", "", "Connecting Rod Bushing", true, 0, false],
  ["163045", "17331-21980", "", "Connecting Rod Bushing (ended)", false, 5, false],
  ["170001", "02102238", "", "Deutz part, zero kept", true, 3, false],
  ["170002", "2102238", "", "Deutz part, zero lost", true, 1, false]
];
let loads = 0;
ctx._mpnLoadMi = function () {
  loads++;
  return {
    r: { rows: ROWS.map(r => [r[1], r[2]]) },
    mpnCols: [{ name: "C:MPN", off: 0 }, { name: "C:Replace Part number", off: 1 }],
    describe: i => { const r = ROWS[i]; return { sku: r[0], title: r[3], location: "A-1", available: r[5], price: 10,
                                                 active: r[4], isKit: r[6], url: "" }; }
  };
};

let pass = 0, fail = 0;
const eq = (n, g, w) => { const G = JSON.stringify(g), W = JSON.stringify(w); if (G === W) pass++; else { fail++; console.log("FAIL " + n + "\n  got  " + G + "\n  want " + W); } };

const r = ctx.catalogueStock(["1A091-2105-0", "17331-2198-0", "0210 2238", "99999-9999-9", "1"]);
eq("ok", r.ok, true);
eq("one MI read for the whole engine", loads, 1);
eq("manual dashes ignored · part before kit", r.stock["1A091-2105-0"].map(x => x.sku), ["167067", "210176"]);
eq("kit flagged", r.stock["1A091-2105-0"][1].kit, true);
eq("ended listing dropped", r.stock["17331-2198-0"].some(x => x.sku === "163045"), false);
eq("non-kit first even at 0 on hand", r.stock["17331-2198-0"][0].sku, "163044");
eq("leading zero + space both match", r.stock["0210 2238"].map(x => x.sku).sort(), ["170001", "170002"]);
eq("not carried → empty list", r.stock["99999-9999-9"], []);
eq("too-short key → empty, never matches everything", r.stock["1"], []);
eq("empty request refused", ctx.catalogueStock([]).ok, false);
eq("cap", ctx.catalogueStock(new Array(1000).fill("1A091-2105-0")).ok, true);

console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
