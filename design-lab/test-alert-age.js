/**
 * test-alert-age.js — the two pure functions behind "checked 6d ago" (2026-09-16)
 *
 * WHAT THIS COVERS
 *   1. _priceAuditCellToMs  (PriceAudit.js)  — coerce a LAST_CHECKED cell to ms epoch
 *   2. _alertAgeText        (Sidebar.html)   — render an age, or admit it has none
 *
 * ⚠ BOTH ARE LOADED FROM THE REAL SHIPPED FILES, never re-typed here. A harness that
 * carries its own copy of the logic proves only that the copy works — the lesson this
 * project re-learned when a test stub could not reach the state the bug lived in.
 *
 * ⚠ SECTIONS FAIL SOFT. If one function cannot be extracted the others must still
 * report, or a single missing symbol hides every result (the `choosePicker` lesson).
 *
 * HOW TO PROVE IT BITES (do this after any edit to either function):
 *   • In _alertAgeText, change `if (!ts) return 'never checked'` to return
 *     'checked just now'      → section B must fail on the null case.
 *   • In _priceAuditCellToMs, drop the `v < 25000 || v > 80000` guard
 *                             → section A must fail on the stray-count case.
 *   • In _priceAuditCellToMs, delete the `typeof v === 'string'` branch
 *                             → section A must fail on the Gotcha-#16 string case.
 *
 * Run:  node design-lab/test-alert-age.js
 */

'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = process.env.SRC || path.join(__dirname, '..');

let pass = 0, fail = 0;
const failures = [];

function ok(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want);
  if (g === w) { pass++; return; }
  fail++;
  failures.push(`  ✗ ${name}\n      got  ${g}\n      want ${w}`);
}

function section(name, fn) {
  try { fn(); }
  catch (e) {
    fail++;
    failures.push(`  ✗ ${name} — SECTION THREW: ${e.message}`);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// Load _priceAuditCellToMs out of the real PriceAudit.js.
//
// ⚠ Comments are stripped before the function is located. This project has been
// bitten twice by asserting against source text that matched a COMMENT describing
// the thing rather than the code doing it.
// ───────────────────────────────────────────────────────────────────────────────
function loadCellToMs() {
  const src = fs.readFileSync(path.join(ROOT, 'PriceAudit.js'), 'utf8');
  const start = src.indexOf('function _priceAuditCellToMs');
  if (start === -1) throw new Error('_priceAuditCellToMs not found in PriceAudit.js');

  // Brace-match rather than slice a fixed window — a magic character count is what
  // silently truncated test-owner-bridge when a docblock grew (2026-09-02).
  let depth = 0, i = src.indexOf('{', start), end = -1;
  for (; i < src.length; i++) {
    if (src[i] === '{') depth++;
    else if (src[i] === '}') { depth--; if (depth === 0) { end = i + 1; break; } }
  }
  if (end === -1) throw new Error('could not brace-match _priceAuditCellToMs');

  const ctx = { Date };                       // real Date, injected — see below
  vm.createContext(ctx);
  vm.runInContext(src.slice(start, end) + '\n;this.fn = _priceAuditCellToMs;', ctx);
  return ctx.fn;
}

// ───────────────────────────────────────────────────────────────────────────────
// Load _alertAgeText out of the real Sidebar.html script block.
// ───────────────────────────────────────────────────────────────────────────────
function loadAgeText() {
  const html = fs.readFileSync(path.join(ROOT, 'Sidebar.html'), 'utf8');
  const start = html.indexOf('function _alertAgeText');
  if (start === -1) throw new Error('_alertAgeText not found in Sidebar.html');

  let depth = 0, i = html.indexOf('{', start), end = -1;
  for (; i < html.length; i++) {
    if (html[i] === '{') depth++;
    else if (html[i] === '}') { depth--; if (depth === 0) { end = i + 1; break; } }
  }
  if (end === -1) throw new Error('could not brace-match _alertAgeText');

  const ctx = { Date };
  vm.createContext(ctx);
  vm.runInContext(html.slice(start, end) + '\n;this.fn = _alertAgeText;', ctx);
  return ctx.fn;
}

// ═══════════════════════════════════════════════════════════════════════════════
// A · _priceAuditCellToMs — the Gotcha #16 surface
// ═══════════════════════════════════════════════════════════════════════════════
section('A · _priceAuditCellToMs', () => {
  const toMs = loadCellToMs();

  // ⚠ vm.createContext creates a SEPARATE REALM with its own Date, so a Date built
  // out here is NOT instanceof the sandbox's Date. Injecting the outer Date is what
  // makes `v instanceof Date` behave — the exact trap that made every Order Archive
  // section fail on its first run (2026-08-28).
  const d = new Date(2026, 8, 16, 14, 30, 0);
  ok('A1 a real Date returns its ms',            toMs(d),                d.getTime());
  ok('A2 an invalid Date is refused',            toMs(new Date('nope')), null);

  // The shape the column actually has today: written as a Date, formatted
  // 'M/d/yy h:mm am/pm'. A1 is that case.

  // Gotcha #16: a crossed number format changes the TYPE that comes back.
  const serial = 46266;                                   // ≈ 2026-09-16
  const asMs   = toMs(serial);
  ok('A3 a Sheets serial is recovered, not refused', asMs !== null, true);
  if (asMs !== null) {
    const y = new Date(asMs).getFullYear();
    ok('A3b and it lands in a plausible year',     y >= 2020 && y <= 2030, true);
  }

  // ⭐ THE GUARD THAT MATTERS. If a COUNT ever lands in this column (a drift total
  // written to the wrong cell), reading it as a date would silently report the badge
  // as last checked in 1900 — a confident wrong answer. Refuse instead.
  ok('A4 a stray small count is REFUSED, not read as 1900', toMs(54),    null);
  ok('A5 a stray large number is refused',                  toMs(999999), null);
  ok('A6 zero is refused',                                  toMs(0),      null);

  // A date that arrived as text (the other half of Gotcha #16).
  const parsed = toMs('9/16/26 2:30 PM');
  ok('A7 a date STRING is parsed',        parsed !== null, true);
  ok('A8 an unparseable string is refused', toMs('not a date'), null);
  ok('A9 an empty string is refused',       toMs(''),           null);
  ok('A10 whitespace is refused',           toMs('   '),        null);

  // Absent values — the "never checked" path.
  ok('A11 null is refused',      toMs(null),      null);
  ok('A12 undefined is refused', toMs(undefined), null);
  ok('A13 a bare object is refused', toMs({}),    null);
});

// ═══════════════════════════════════════════════════════════════════════════════
// B · _alertAgeText — the honesty rules
// ═══════════════════════════════════════════════════════════════════════════════
section('B · _alertAgeText', () => {
  const age = loadAgeText();
  const now = Date.now();
  const MIN = 60000, HR = 3600000, DAY = 86400000;

  // ⭐ THE HEADLINE. An unknown age must NOT read as a fresh one. This is the
  // reassuring-label-on-a-dangerous-state rule: the whole point of the row is to
  // say whether the count can be trusted, so "we don't know" has to say so.
  ok('B1 null  → never checked',      age(null),      'never checked');
  ok('B2 zero  → never checked',      age(0),         'never checked');
  ok('B3 undefined → never checked',  age(undefined), 'never checked');

  ok('B4 30 seconds → just now',      age(now - 30000),        'checked just now');
  ok('B5 12 minutes',                 age(now - 12 * MIN),     'checked 12m ago');
  ok('B6 59 minutes stays in minutes', age(now - 59 * MIN),    'checked 59m ago');
  ok('B7 60 minutes flips to hours',  age(now - 60 * MIN),     'checked 1h ago');
  ok('B8 5 hours',                    age(now - 5 * HR),       'checked 5h ago');
  ok('B9 23h stays in hours',         age(now - 23 * HR),      'checked 23h ago');
  ok('B10 24h flips to days',         age(now - 24 * HR),      'checked 1d ago');

  // ⭐ THE CASE THIS WHOLE FEATURE EXISTS FOR: the audit is weekly, so six days old
  // is the NORMAL state, and the row has to be able to say it out loud.
  ok('B11 six days — the weekly-audit norm', age(now - 6 * DAY), 'checked 6d ago');
  ok('B12 a month',                          age(now - 31 * DAY), 'checked 31d ago');

  // ⚠ Clock skew: the script's timezone and the browser's can disagree. We never
  // claim an audit ran in the future, and we never render a negative.
  ok('B13 a future timestamp degrades to just now', age(now + 5 * MIN), 'checked just now');
});

// ═══════════════════════════════════════════════════════════════════════════════
// C · The two together — a cell round-trips into a legible age
// ═══════════════════════════════════════════════════════════════════════════════
section('C · round trip', () => {
  const toMs = loadCellToMs();
  const age  = loadAgeText();

  const sixDaysAgo = new Date(Date.now() - 6 * 86400000);
  ok('C1 a Date cell six days old reads as 6d', age(toMs(sixDaysAgo)), 'checked 6d ago');

  // ⭐ AND THE FAILURE PATH STAYS HONEST END TO END. A cell we cannot read must
  // surface as "never checked" rather than collapsing to a fresh-looking age —
  // an unreadable value is not the same as no drift, and the row must not imply it is.
  ok('C2 an unreadable cell reads as never checked', age(toMs('garbage')), 'never checked');
  ok('C3 a stray count reads as never checked',      age(toMs(54)),        'never checked');
});

// ───────────────────────────────────────────────────────────────────────────────
console.log('');
console.log(failures.length ? failures.join('\n') : '  (no failures)');
console.log('');
console.log(`  ${pass} passed · ${fail} failed`);
process.exit(fail ? 1 : 0);
