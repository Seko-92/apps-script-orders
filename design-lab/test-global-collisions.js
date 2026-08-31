/**
 * test-global-collisions.js — two root files may never declare the same top-level name.
 *
 * WHY THIS IS A TEST AND NOT A NOTE IN CLAUDE.md
 *   Apps Script concatenates every root .js file into ONE global scope, in an UNSPECIFIED
 *   order. So two top-level `function foo(){}` declarations are not a lint nit — one
 *   silently wins and the other vanishes, and which one is undefined. `var` behaves the
 *   same way.
 *
 *   ⚠⚠ THIS HAS BITTEN TWICE.
 *     • 2026-08-19  a second top-level `var BUDGET` would have clobbered Helpers.js's own,
 *                   caught only because the scan was run by hand that day.
 *     • 2026-08-30  `setN8nSheetsAccountNow` shipped in BOTH BrandTheme.js and
 *                   OwnerBridge.js. The two bodies DIFFERED — one refused safely when
 *                   unedited, the other returned "✅ …= PUT-THE-ACCOUNT-EMAIL-OR-none-HERE".
 *                   Someone editing the VALUE in one copy could have had the other run.
 *
 *   CLAUDE.md told the reader to run the scan before adding a top-level name. That is
 *   exactly the shape this project has already ruled against: "an enumeration that is not
 *   executable rots the day it is written." The 08-30 duplicate is the proof — the note
 *   existed, and the duplicate shipped anyway. So it runs in the suite now.
 *
 * ⚠ SECTION A PROVES THE DETECTOR CAN BITE. A scanner whose regex silently matches nothing
 *   reports a clean project forever. That is the vacuous-pass class this codebase keeps
 *   re-learning, so the planted fixtures come FIRST and the live verdict only counts after.
 */
'use strict';
const fs = require('fs'), path = require('path');

const ROOT = process.env.SRC || path.join(__dirname, '..');

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label +
              (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};
const section = n => console.log('\n' + n);

/**
 * Top-level declarations only — anchored at column 0, which is what "global" means once
 * the files are concatenated. An indented `var` is inside a function and is scoped there.
 *
 * ⚠ `const` / `let` are included even though a duplicate of those is a HARD error rather
 *   than a silent clobber: catching it here beats catching it as a red banner on the sheet.
 */
const DECL = /^(?:var|function|const|let)\s+([A-Za-z_$][\w$]*)/;

function declarationsIn(text) {
  const out = [];
  text.split('\n').forEach((line, i) => {
    const m = line.match(DECL);
    if (m) out.push({ name: m[1], line: i + 1 });
  });
  return out;
}

/** name → [ "file:line", … ] across every file given. */
function collisions(files) {
  const seen = new Map();
  files.forEach(({ file, text }) => {
    declarationsIn(text).forEach(d => {
      if (!seen.has(d.name)) seen.set(d.name, []);
      seen.get(d.name).push(file + ':' + d.line);
    });
  });
  const dupes = {};
  for (const [name, sites] of seen) if (sites.length > 1) dupes[name] = sites;
  return dupes;
}


// ═══════════════════════════════════════════════════════════════════════════════════════
section('A · the detector can actually bite  (run FIRST — a blind scanner passes forever)');
// ═══════════════════════════════════════════════════════════════════════════════════════

t('A1 a cross-file duplicate function is caught',
  collisions([
    { file: 'One.js', text: 'function foo() {}\n' },
    { file: 'Two.js', text: 'function foo() {}\n' }
  ]),
  { foo: ['One.js:1', 'Two.js:1'] });

t('A2 a cross-file duplicate var is caught',
  Object.keys(collisions([
    { file: 'One.js', text: 'var BUDGET = 1;\n' },
    { file: 'Two.js', text: 'var BUDGET = 2;\n' }
  ])),
  ['BUDGET']);

t('A3 the real 2026-08-30 shape is caught',
  Object.keys(collisions([
    { file: 'BrandTheme.js',  text: 'function setN8nSheetsAccountNow() {\n  return 1;\n}\n' },
    { file: 'OwnerBridge.js', text: 'function setN8nSheetsAccountNow() {\n  return 2;\n}\n' }
  ])),
  ['setN8nSheetsAccountNow']);

t('A4 a duplicate INSIDE one file is caught too',
  Object.keys(collisions([{ file: 'One.js', text: 'function a(){}\nfunction a(){}\n' }])),
  ['a']);

t('A5 ⚠ an INDENTED var is function-scoped and must NOT be flagged',
  collisions([
    { file: 'One.js', text: 'function f() {\n  var tmp = 1;\n}\n' },
    { file: 'Two.js', text: 'function g() {\n  var tmp = 2;\n}\n' }
  ]),
  {});

t('A6 distinct names are left alone',
  collisions([
    { file: 'One.js', text: 'function alpha() {}\n' },
    { file: 'Two.js', text: 'function beta() {}\n' }
  ]),
  {});

t('A7 a name that merely CONTAINS another is not a collision',
  collisions([
    { file: 'One.js', text: 'function run() {}\n' },
    { file: 'Two.js', text: 'function runNow() {}\n' }
  ]),
  {});

t('A8 a mention in a COMMENT is not a declaration',
  collisions([
    { file: 'One.js', text: 'function foo() {}\n' },
    { file: 'Two.js', text: '// function foo() is defined in One.js\n' }
  ]),
  {});


// ═══════════════════════════════════════════════════════════════════════════════════════
section('B · the live project');
// ═══════════════════════════════════════════════════════════════════════════════════════

const files = fs.readdirSync(ROOT)
  .filter(f => f.endsWith('.js'))
  .map(f => ({ file: f, text: fs.readFileSync(path.join(ROOT, f), 'utf8') }));

t('B1 the scan can see the project at all (not a vacuous pass)', files.length > 40, true);

const live = collisions(files);
t('B2 ⚠⚠ NO two root files declare the same top-level name', live, {});

const total = files.reduce((n, f) => n + declarationsIn(f.text).length, 0);
console.log('\n  scanned ' + files.length + ' root files · ' + total + ' top-level names');

console.log('\n' + (fail === 0 ? '✅' : '❌') +
            ' test-global-collisions: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail === 0 ? 0 : 1);
