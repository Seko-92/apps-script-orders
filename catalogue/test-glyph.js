#!/usr/bin/env node
// lib/glyph.js against real word crops from V3600-T-E3B-EU-ZH4 (400 dpi), each with the answer from
// the manual page. `null` = the test must NOT decide (a wrong verdict is worse than none):
//   16010-5514-0 — the "1" sits outside tesseract's word box; every glyph index shifts by one
//   16010-2411-2 — a scrap of table rule inside the word box reads as an open glyph
// ⚠ Mutation-checked 2026-10-06: dropping the 5 or the J test fails one case each. The glyph-size and
//   "never fewer glyphs" guards are BACKUPS — on these crops the narrow-"1" check already catches both
//   null cases, so removing either guard alone still passes. Keep them; they cost nothing.
"use strict";
const fs = require("fs");
const path = require("path");
const { loadPgm, secondCharShape } = require("./lib/glyph");

const dir = path.join(__dirname, "test-fixtures", "glyph");
const cases = JSON.parse(fs.readFileSync(path.join(dir, "cases.json"), "utf8"));
let pass = 0, fail = 0;
for (const c of cases) {
  const R = loadPgm(path.join(dir, c.file));
  const v = secondCharShape(R, c.words, c.pn, c.pageH != null ? { pageH: c.pageH } : undefined);
  const got = v ? v.kind : null;
  const ok = got === c.want;
  ok ? pass++ : fail++;
  console.log(`${ok ? "✓" : "✗"} ${c.pn}  want ${c.want}  got ${got}`);
}
console.log(`${pass} passed · ${fail} failed`);
process.exit(fail ? 1 : 0);
