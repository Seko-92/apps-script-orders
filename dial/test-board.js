#!/usr/bin/env node
/**
 * test-board.js — the split-flap mechanic and the board's geometry contract.
 *
 * ⚠ Mirrors test-dial.js: pure Node assertions, no canvas, loading the REAL modules so the
 *   tests cannot drift from what ships.
 *
 * ⚠⚠ THE ONES THAT MATTER ARE C AND G. C pins "only what changes moves" — the rule that
 *    makes the board beautiful AND cheap, since a clean tick has nothing to redraw. G pins
 *    the two image sizes against the sheet's OWN column widths, read out of BrandTheme.js:
 *    the plan carried 880 for the strip and the truth is 876, and a 4px overhang puts the
 *    board's right edge into column I on every install.
 */
'use strict';
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const B = require('./src/board');
const T = require('./src/board-terminal');

let pass = 0, fail = 0;
const ok = (name, fn) => {
  try { fn(); pass++; }
  catch (e) { fail++; console.log('  ✗ ' + name + '\n      ' + e.message.split('\n')[0]); }
};
const sec = (s) => console.log('\n' + s);

sec('A · the alphabet is the mechanic');
ok('space leads, so a blanked cell settles fastest', () => assert.strictEqual(B.ALPHABET[0], ' '));
ok('A -> B is one flap',        () => assert.strictEqual(B.distance('A', 'B'), 1));
// ⚠ FIVE, not four. The plan's prose says "R -> W passes S T U V" — four INTERMEDIATE
//   characters, which is five flaps. The first cut of this assertion read the prose as the
//   distance and failed against correct code: the 17th time a harness here has accused the
//   product. Suspect the harness first.
ok('R -> W travels five flaps',  () => assert.strictEqual(B.distance('R', 'W'), 5));
ok('it wraps, never reverses',  () => assert.ok(B.distance('Z', 'A') > 0));
ok('same char never moves',     () => assert.strictEqual(B.distance('Q', 'Q'), 0));
ok('case is folded',            () => assert.strictEqual(B.distance('a', 'A'), 0));
ok('an unmappable char becomes a blank flap, never an exception',
   () => assert.strictEqual(B.distance('%', ' '), 0));
ok('the em dash is NOT on the flaps — that was a real bug',
   () => assert.strictEqual(B.ALPHABET.indexOf('—'), -1));
ok('but the hyphen IS, which is what "no value" uses',
   () => assert.ok(B.ALPHABET.indexOf('-') > 0));

sec('B · the flap travels through the alphabet (move 4)');
ok('R -> W passes S, T, U on the way', () => {
  const seen = [];
  for (let t = 0; t < 1.2; t += 0.004) {
    const st = B.cellState(0, 'R', 'W', t);
    if (!seen.includes(st.cur)) seen.push(st.cur);
  }
  assert.deepStrictEqual(seen, ['R', 'S', 'T', 'U', 'V', 'W']);
});
ok('a longer journey lands LATER than a short one', () => {
  const short = B.loopSeconds('A', 'B'), long = B.loopSeconds('A', 'Z');
  assert.ok(long > short, `${long} should exceed ${short}`);
});

sec('C · ⭐ ONLY WHAT CHANGES MOVES (move 3) — the rule that makes it cheap');
ok('an unchanged cell is settled at every instant', () => {
  for (let t = 0; t < 5; t += 0.01) {
    const st = B.cellState(3, 'G', 'G', t);
    assert.strictEqual(st.moving, false);
    assert.strictEqual(st.phase, 0);
    assert.strictEqual(st.cur, 'G');
  }
});
ok('RESTING -> WORKING leaves I, N and G untouched for the whole loop', () => {
  const from = ' RESTING ', to = ' WORKING ';
  for (const i of [6, 7, 8]) {                       // the trailing "N G " cells
    for (let t = 0; t < 4; t += 0.05) {
      assert.strictEqual(B.cellState(i, from[i], to[i], t).moving, false);
    }
  }
});
ok('an identical state produces a loop with NOTHING moving', () => {
  const w = ' RESTING ';
  for (let i = 0; i < w.length; i++) {
    assert.strictEqual(B.distance(w[i], w[i]), 0);
  }
});

sec('D · the cascade is staggered, and every cell eventually lands');
ok('cell 0 starts before cell 5', () => {
  assert.strictEqual(B.cellState(5, 'A', 'B', 0.01).moving, false);
  assert.ok(B.cellState(0, 'A', 'B', 0.01).moving);
});
ok('every cell has landed by loopSeconds', () => {
  const from = 'ALL CLEAR', to = 'NO SIGNAL';
  const t = B.loopSeconds(from, to);
  for (let i = 0; i < from.length; i++) {
    const st = B.cellState(i, from[i], to[i], t);
    assert.strictEqual(st.moving, false);
    assert.strictEqual(st.cur, to[i].toUpperCase());
  }
});
ok('the loop stays under the 8s GIF cap (trap 7)', () => {
  let worst = 0;
  for (const a of Object.values(T.WORD)) {
    for (const b of Object.values(T.WORD)) {
      worst = Math.max(worst, B.loopSeconds(a.padEnd(9), b.padEnd(9)));
    }
  }
  assert.ok(worst < 8, `worst-case loop is ${worst.toFixed(2)}s`);
});

sec('E · the flap settles with a bounce, and is continuous');
ok('starts folded, ends flat', () => {
  assert.ok(Math.abs(B.flapScale(0)) < 1e-9);
  assert.ok(Math.abs(B.flapScale(1) - 1) < 1e-9);
});
ok('continuous at the 0.86 seam — the bounce must not jump', () => {
  const a = B.flapScale(0.8599), b = B.flapScale(0.8601);
  assert.ok(Math.abs(a - b) < 0.01, `${a} -> ${b}`);
});
ok('it OVERSHOOTS and comes back (a real flap bounces)', () => {
  let dipped = false;
  for (let u = 0.86; u <= 1; u += 0.005) if (B.flapScale(u) < 0.98) dipped = true;
  assert.ok(dipped, 'no settle bounce — it just stops');
});
ok('never inverts', () => {
  for (let u = 0; u <= 1; u += 0.01) assert.ok(B.flapScale(u) >= 0);
});

sec('F · what the board says, per verdict');
ok('every verdict has a word that FITS the fixed grid', () => {
  for (const [v, w] of Object.entries(T.WORD)) {
    assert.ok(w.length <= T.STATE_CELLS, `${v}: "${w}" is ${w.length} > ${T.STATE_CELLS}`);
  }
});
ok('the state word is centred in its grid', () => {
  assert.strictEqual(T.buildBoard({ s: 'rest' }).word, ' RESTING ');
});
ok('figures right-align, so blanks lead', () => {
  assert.strictEqual(T.buildBoard({ s: 'busy', g: '12' }).figA.value, '  12');
});
ok('⭐ YELLOW ONLY WHEN SOMEONE MUST MOVE', () => {
  for (const v of ['rest', 'clear', 'stale']) {
    assert.strictEqual(T.buildBoard({ s: v, o: '90' }).lead.act, false, v + ' must stay monochrome');
  }
  for (const v of ['busy', 'late']) {
    assert.strictEqual(T.buildBoard({ s: v, o: '90' }).lead.act, true, v + ' must call for a person');
  }
});
ok('exactly ONE field can ever be yellow', () => {
  const b = T.buildBoard({ s: 'late', o: '200', g: '9', p: '5' });
  assert.strictEqual([b.lead, b.figA, b.figB].filter((f) => f.act).length, 1);
});
ok('⚠ THE LABELS NEVER CHANGE (move 2) — an engraved word that moves is a lie', () => {
  const seen = { lead: new Set(), figA: new Set(), figB: new Set() };
  for (const v of ['rest', 'clear', 'busy', 'late', 'stale']) {
    const b = T.buildBoard({ s: v, o: '30', g: '4', p: '9' });
    seen.lead.add(b.lead.label); seen.figA.add(b.figA.label); seen.figB.add(b.figB.label);
  }
  for (const k of Object.keys(seen)) {
    assert.strictEqual(seen[k].size, 1, `${k} label changed across verdicts: ${[...seen[k]]}`);
  }
});
ok('a missing figure shows a hyphen the flaps HAVE, never an em dash', () => {
  const b = T.buildBoard({ s: 'stale' });
  for (const f of [b.lead, b.figA, b.figB]) {
    assert.ok(f.value.indexOf('—') === -1, 'em dash reached the flaps');
    assert.ok(f.value.trim() === '-', `expected a hyphen, got "${f.value}"`);
  }
});
ok('an over-long figure can never silently truncate to a plausible number', () => {
  // 4 cells hold 9999. If a bigger count ever arrives we want to SEE the clip, not read
  // a wrong number — this pins the width so a future change has to think about it.
  assert.strictEqual(T.FIG_CELLS, 4);
  assert.strictEqual(T.buildBoard({ s: 'busy', g: '9999' }).figA.value, '9999');
});

sec('G · ⭐⭐ THE GEOMETRY, READ OUT OF BrandTheme.js — not from the plan');
const BT = fs.readFileSync(path.join(__dirname, '..', 'BrandTheme.js'), 'utf8');
const widthOf = (col) => {
  const m = BT.match(new RegExp('setColumnWidth\\(Schema\\.cols\\.' + col + ',\\s*(\\d+)\\)'));
  assert.ok(m, 'no setColumnWidth for ' + col);
  return Number(m[1]);
};
ok('the anchor is exactly A + B + C', () => {
  const w = widthOf('SKU') + widthOf('QTY') + widthOf('LOCATION');
  assert.strictEqual(T.ANCHOR.w, w, `anchor ${T.ANCHOR.w} vs live ${w}`);
});
ok('⚠ the strip is 876 (D..H), NOT the 880 the plan carried', () => {
  const w = widthOf('SALES_ORDER') + widthOf('NOTE') + widthOf('STATUS') +
            widthOf('HAND') + widthOf('LEFT');
  assert.strictEqual(w, 876);
  assert.strictEqual(T.STRIP.w, w, `strip ${T.STRIP.w} vs live ${w}`);
});
ok('the strip is row 1 ONLY — covering row 2 would swallow the Pick ID dropdowns', () => {
  const m = BT.match(/rowHeight:\s*(\d+)/);
  assert.ok(m, 'MASTHEAD.rowHeight not found');
  assert.strictEqual(T.STRIP.h, Number(m[1]));
  assert.ok(T.STRIP.h < T.ANCHOR.h, 'the strip must not reach row 2');
});
ok('the window columns sum to the strip width', () => {
  assert.strictEqual(T.COL_D + T.COL_E + T.COL_FH, T.STRIP.w);
});
ok('the anchor matches MASTHEAD.imgW, which setupMasthead already asserts', () => {
  const m = BT.match(/imgW:\s*(\d+)/);
  assert.strictEqual(T.ANCHOR.w, Number(m[1]));
});

sec('H · the maker\'s plate is OURS');
ok('it says HQ MOTOR SERVICE', () => {
  const src = fs.readFileSync(path.join(__dirname, 'src', 'board-terminal.js'), 'utf8');
  assert.ok(src.indexOf("'HQ MOTOR SERVICE'") !== -1);
});
ok('⚠ SOLARI DI UDINE appears ONLY as design lineage in a comment, never as drawn text', () => {
  const src = fs.readFileSync(path.join(__dirname, 'src', 'board-terminal.js'), 'utf8');
  const drawn = src.match(/text:\s*'([^']*)'/g) || [];
  for (const d of drawn) {
    assert.ok(!/solari/i.test(d), 'another company name is being DRAWN: ' + d);
  }
});

console.log(`\n${pass} passed · ${fail} failed`);
process.exit(fail ? 1 : 0);
