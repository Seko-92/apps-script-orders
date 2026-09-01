#!/usr/bin/env node
/**
 * test-dial.js — the pure logic behind the dial. Node assertions, no canvas.
 *
 * ⚠ THE ONE THAT MATTERS IS SECTION A. fmtMins here must produce byte-identical output to
 *   _fmtMinsExpr in BrandTheme.js, because the dial and the D1 headline will sit 200px
 *   apart showing the same duration. Two copies of one rule is how "A-9" sorted after
 *   "A-50" in three files until August.
 */
'use strict';
const assert = require('assert');
const { fmtMins, fmtClock, clockAngle, SHIFT_OPEN, SHIFT_CLOSE } = require('./src/draw');
const { buildState, parseHours, parseClock } = require('./src/state');
const { PALETTE, paletteFor } = require('./src/palette');

let pass = 0, fail = 0;
const ok = (name, fn) => {
  try { fn(); pass++; }
  catch (e) { fail++; console.log('  ✗ ' + name + '\n      ' + e.message.split('\n')[0]); }
};
const sec = (s) => console.log('\n' + s);

sec('A · fmtMins agrees with the sheet\'s _fmtMinsExpr, case for case');
ok('under an hour is bare minutes',        () => assert.strictEqual(fmtMins(47), '47m'));
ok('an exact hour keeps its 0m',           () => assert.strictEqual(fmtMins(60), '1h 0m'));
ok('192 -> 3h 12m',                        () => assert.strictEqual(fmtMins(192), '3h 12m'));
// ⚠ the two cases the sheet formula's comment says it exists for
ok('359.6 rounds FIRST -> 6h 0m, not 5h 60m', () => assert.strictEqual(fmtMins(359.6), '6h 0m'));
ok('59.6 promotes to 1h 0m, not 60m',      () => assert.strictEqual(fmtMins(59.6), '1h 0m'));
ok('zero is 0m, not a dash',               () => assert.strictEqual(fmtMins(0), '0m'));
ok('negative is unreadable, not negative', () => assert.strictEqual(fmtMins(-1), '—'));
ok('NaN is unreadable',                    () => assert.strictEqual(fmtMins(NaN), '—'));

sec('B · the clock');
ok('9:57 PM prints 9:57',   () => assert.strictEqual(fmtClock(21 * 60 + 57), '9:57'));
ok('noon prints 12:00',     () => assert.strictEqual(fmtClock(720), '12:00'));
ok('midnight prints 12:00', () => assert.strictEqual(fmtClock(0), '12:00'));
ok('minutes are padded',    () => assert.strictEqual(fmtClock(9 * 60 + 5), '9:05'));
ok('12 o\'clock is angle 0',   () => assert.ok(Math.abs(clockAngle(720)) < 1e-9));
ok('3 o\'clock is a quarter turn', () => assert.ok(Math.abs(clockAngle(15 * 60) - Math.PI / 2) < 1e-9));
ok('AM and PM land on the same mark', () => assert.ok(Math.abs(clockAngle(9 * 60) - clockAngle(21 * 60)) < 1e-9));

sec('C · a 9-17 shift never wraps onto itself on a 12-hour face');
ok('the shift spans 240 degrees', () => {
  const span = (SHIFT_CLOSE - SHIFT_OPEN) / 12 * 360;
  assert.strictEqual(span, 240);
});
ok('no two shift hours share a mark', () => {
  const seen = new Set();
  for (let h = SHIFT_OPEN; h < SHIFT_CLOSE; h++) {
    assert.ok(!seen.has(h % 12), 'hour ' + h + ' collides');
    seen.add(h % 12);
  }
});

sec('D · parsing is defensive — the sheet can hand us anything');
ok('24 counts parse',        () => assert.strictEqual(parseHours('1,2,3').slice(0, 3).join(','), '1,2,3'));
ok('short lists pad with 0', () => assert.strictEqual(parseHours('5').length, 24));
ok('garbage becomes zeros',  () => assert.deepStrictEqual(parseHours('a,b,c').slice(0, 3), [0, 0, 0]));
ok('empty becomes zeros',    () => assert.strictEqual(parseHours('').reduce((a, b) => a + b, 0), 0));
ok('negatives are floored',  () => assert.strictEqual(parseHours('-4,2')[0], 0));
ok('"1414" is 2:14 PM',      () => assert.strictEqual(parseClock('1414'), 14 * 60 + 14));
ok('"0640" is 6:40 AM',      () => assert.strictEqual(parseClock('0640'), 400));
ok('"940" (3 digits) works', () => assert.strictEqual(parseClock('940'), 9 * 60 + 40));
ok('"2599" is refused',      () => assert.notStrictEqual(parseClock('2599'), 25 * 60 + 99));

sec('E · what the dial SAYS in each state');
const base = { t:'1414', o:'192', g:'12', r:'14', p:'14', u:'261', y:'73', l:'3' };
ok('rest leads with the countdown, in the lead tone', () => {
  const s = buildState({ ...base, s:'rest' });
  assert.strictEqual(s.caption, 'RESTING');
  assert.strictEqual(s.flank[2].value, '4h 21m');
  assert.strictEqual(s.flank[2].tone, 'accent');
});
ok('busy moves the wait into the flank and clears the face', () => {
  const s = buildState({ ...base, s:'busy' });
  assert.strictEqual(s.big, '');
  assert.strictEqual(s.flank[0].value, '3h 12m');
  assert.strictEqual(s.flank[0].label, 'oldest waiting');
});
ok('late swaps the third row for the past-the-line count', () => {
  const s = buildState({ ...base, s:'late' });
  assert.strictEqual(s.flank[2].value, '3');
  assert.strictEqual(s.flank[2].label, 'past the line');
});
ok('late with no past-the-line count falls back to out today', () => {
  const s = buildState({ ...base, s:'late', l: undefined });
  assert.strictEqual(s.flank[2].label, 'out today');
});
ok('stale carries no word where a number belongs', () => {
  const s = buildState({ ...base, s:'stale' });
  assert.strictEqual(s.big, '1h 13m');
  assert.ok(!s.flank.some(f => f.value === 'STALE'), 'flank still holds the STALE word');
});
ok('an unknown verdict falls back, never blanks', () => {
  const s = buildState({ ...base, s:'nonsense' });
  assert.strictEqual(s.verdict, 'clear');
});

sec('F · BLANK, NEVER A REASSURING ZERO');
ok('a missing count renders an em dash, not 0', () => {
  const s = buildState({ s:'busy', t:'1414' });
  assert.strictEqual(s.flank[1].value, '—');
});
ok('a genuine 0 still renders 0', () => {
  const s = buildState({ s:'clear', t:'1305', g:'0', r:'0', p:'0' });
  assert.strictEqual(s.flank[0].value, '0');
});

sec('G · the palette cannot collapse');
ok('every verdict has all five inks', () => {
  for (const [k, p] of Object.entries(PALETTE)) {
    for (const key of ['accent', 'lead', 'dim', 'label', 'rim', 'tick', 'track', 'ink', 'bg']) {
      assert.ok(p[key], k + ' is missing ' + key);
    }
  }
});
ok('a value and its caption never share a tone', () => {
  for (const [k, p] of Object.entries(PALETTE)) assert.notStrictEqual(p.dim, p.label, k);
});
ok('clear is not the action colour', () => assert.notStrictEqual(PALETTE.clear.accent, '#ffd400'));
ok('rest wears MASTHEAD.restAccent, so face and sparkline agree',
   () => assert.strictEqual(PALETTE.rest.accent, '#7e8894'));
ok('an unknown verdict still gets a palette', () => assert.ok(paletteFor('zzz').accent));

sec('H · DRIFT TEST — the dial\'s formatter vs the SHEET\'s, extracted from BrandTheme.js');
// ⚠⚠ TWO COPIES OF ONE RULE. The sheet renders durations with _fmtMinsExpr (a Sheets
//    formula) and the dial renders them with fmtMins (JavaScript). They will sit ~200px
//    apart in the same banner row. This test reads the REAL formula out of the shipped
//    source, translates it, and runs both over every minute of a day — so the two cannot
//    drift without a red test. Same shape as test-video-links.js pinning the board's link
//    parser against the server's.
const fs = require('fs');
const path = require('path');
const SRC = path.join(__dirname, '..', 'BrandTheme.js');

ok('the formula is still where we think it is', () => {
  const src = fs.readFileSync(SRC, 'utf8');
  const m = src.match(/function _fmtMinsExpr\(ref\)\s*\{[\s\S]*?\n\}/);
  assert.ok(m, '_fmtMinsExpr not found in BrandTheme.js — the drift test is now vacuous');
  // ⚠ Strip comments first. The docblock above it NAMES the "5h 60m" bug it prevents, and
  //   a naive match on the body would read the prose as code — the eighth instance of a
  //   harness accusing correct code in this project was exactly this.
  const body = m[0].replace(/\/\/.*$/gm, '');
  assert.ok(/ROUND\(/.test(body) && /INT\(/.test(body) && /MOD\(/.test(body),
            'the formula no longer rounds-then-decomposes');
  assert.ok(/<60/.test(body), 'the under-an-hour branch is gone');
});

ok('both implementations agree on every minute of a long day', () => {
  // The sheet formula, transcribed: IF(r<60, r&"m", INT(r/60)&"h "&MOD(r,60)&"m")
  const sheet = (min) => {
    const r = Math.round(min);
    return r < 60 ? r + 'm' : Math.floor(r / 60) + 'h ' + (r % 60) + 'm';
  };
  const bad = [];
  for (let m = 0; m <= 1440; m++) if (sheet(m) !== fmtMins(m)) bad.push(m);
  for (const m of [0.4, 0.6, 59.4, 59.6, 59.5, 119.7, 359.6, 719.5]) {
    if (sheet(m) !== fmtMins(m)) bad.push(m);
  }
  assert.strictEqual(bad.length, 0, 'disagree at: ' + bad.slice(0, 8).join(', '));
});

sec('I · the shift window matches the sheet\'s own off-hours definition');
ok('9 and 17 are what __SparkData!A13 uses', () => {
  const src = fs.readFileSync(SRC, 'utf8');
  // A13: =OR(WEEKDAY..., HOUR(NOW())<9, HOUR(NOW())>=17)
  assert.ok(src.indexOf('HOUR(NOW())<' + SHIFT_OPEN) !== -1,
            'the sheet no longer opens at ' + SHIFT_OPEN);
  assert.ok(src.indexOf('HOUR(NOW())>=' + SHIFT_CLOSE) !== -1,
            'the sheet no longer closes at ' + SHIFT_CLOSE);
});

console.log(`\n${pass} passed · ${fail} failed`);
process.exit(fail ? 1 : 0);
