// node catalogue/test-ink.js — the "–" vs digit test that ink-fills empty variant qty cells.
// 2026-10-07: grey speckle around a dash spread total ink 28 px down the cell and every D1105
// alternator "–" read as a digit. The rule now measures the largest connected blob.
const { cellInkSpan } = require("./lib/scan-book");
let fail = 0;
const ok = (name, got, want) => { const p = got === want; if (!p) fail++; console.log((p ? "ok  " : "FAIL") + " " + name + (p ? "" : " → got " + got + ", want " + want)); };
function raster(W, H, draw) { const d = new Uint8Array(W * H).fill(255); draw((x, y) => { if (x >= 0 && y >= 0 && x < W && y < H) d[y * W + x] = 0; }); return { W, H, d }; }
const box = { x: 10, w: 100 }, lo = 0, hi = 72;
const isDigit = m => m.span >= Math.max(9, 0.12 * (hi - lo)) && m.ink >= 25;
const dash = p => { for (let y = 34; y < 38; y++) for (let x = 50; x < 62; x++) p(x, y); };
const speckle = p => { for (let k = 0; k < 30; k++) p(15 + (k * 37) % 90, 10 + (k * 13) % 52); };
const one = p => { for (let y = 25; y < 47; y++) for (let x = 55; x < 58; x++) p(x, y); };
ok("clean dash is not a digit", isDigit(cellInkSpan(raster(120, 72, dash), box, lo, hi)), false);
ok("dash + speckle is not a digit", isDigit(cellInkSpan(raster(120, 72, p => { dash(p); speckle(p); }), box, lo, hi)), false);
ok("a '1' is a digit", isDigit(cellInkSpan(raster(120, 72, one), box, lo, hi)), true);
ok("a '1' + speckle is a digit", isDigit(cellInkSpan(raster(120, 72, p => { one(p); speckle(p); }), box, lo, hi)), true);
ok("a vertical rule at the box edge is not a digit", isDigit(cellInkSpan(raster(120, 72, p => { for (let y = 0; y < 72; y++) { p(11, y); p(12, y); } }), box, lo, hi)), false);
ok("empty cell is not a digit", isDigit(cellInkSpan(raster(120, 72, () => {}), box, lo, hi)), false);
process.exit(fail ? 1 : 0);
