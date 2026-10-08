// catalogueIndexPut + _catalogueEnginesFor (Catalogue.js) against a fake sheet.   node design-lab/test-catalogue-index.js
const fs = require("fs"), vm = require("vm"), path = require("path");
let pass = 0, fail = 0; const ok = (n, c, got) => { c ? pass++ : (fail++, console.log("✗ " + n + (got !== undefined ? " → got " + JSON.stringify(got) : ""))); };
function sheetStub() {
  let cells = []; const sh = { hidden: false,
    clearContents() { cells = []; }, getLastRow() { return cells.length; },
    getRange(r, c, nr, nc) { return { setNumberFormat() { return this; },
      setValues(v) { v.forEach((row, i) => { cells[r - 1 + i] = cells[r - 1 + i] || []; row.forEach((x, j) => cells[r - 1 + i][c - 1 + j] = x); }); return this; },
      getValues() { return cells.slice(r - 1, r - 1 + nr).map(row => row.slice(c - 1, c - 1 + nc)); } }; },
    isSheetHidden() { return this.hidden; }, hideSheet() { this.hidden = true; } };
  return sh;
}
let sheet = null;
const ctx = { console, SPREADSHEET_ID: "x", CATALOGUE_PUBLISH_KEY: "k3y", HQ_BOARD_API_URL: "https://hq.example/api/board",
  SpreadsheetApp: { openById: () => ({ getSheetByName: () => sheet, insertSheet: () => (sheet = sheetStub()) }) } };
vm.createContext(ctx);
const src = f => fs.readFileSync(path.join(__dirname, "..", f), "utf8");
vm.runInContext(src("MpnSearch.js") + "\n" + src("Catalogue.js"), ctx);

ok("no index yet → []", ctx._catalogueEnginesFor(["16423-21110"]).length === 0);
ok("wrong key refused", ctx.catalogueIndexPut("nope", { "1642321110": ["V2203"] }).ok === false);
ok("no key refused", ctx.catalogueIndexPut("", { "1642321110": ["V2203"] }).ok === false);
const r = ctx.catalogueIndexPut("k3y", { "1642321110": ["V2203", "D1105"], "2102238": ["D722"], "12": ["X"] }, { engines: 3 });
ok("put ok", r.ok && r.keys === 2, r);
ok("sheet hidden", sheet.hidden);
const hits = ctx._catalogueEnginesFor(["16423-21110", "02102238", "99999-9999-9"]);
ok("dashes + leading zero match", hits.length === 2, hits);
ok("printed number kept", hits[0].num === "16423-21110" && hits[0].engines.join() === "V2203,D1105", hits[0]);
ok("leading-zero number", hits[1] && hits[1].engines[0] === "D722", hits[1]);
ok("url", ctx._catalogueBaseUrl() === "https://hq.example/catalogue");
// the bundle's key must be the Parts Finder's key
const b = src("catalogue/bundle.js"); const pk = new Function(/function pnKey[\s\S]*?\n}/.exec(b)[0] + "; return pnKey;")();
["02102238", "16423-21110", "1G796 21050", "0000"].forEach(s => ok("bundle key = _mpnKey " + s, pk(s) === ctx._mpnKey(s), [pk(s), ctx._mpnKey(s)]));
console.log(`${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
