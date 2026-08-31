/**
 * test-pickid-reset.js — the 4am reset must never write a value its own dropdown
 * rejects, and must never invent one.
 *
 * WHY THIS EXISTS. On 2026-08-31 describePickIdCells found that H2's option [0] is
 * "Pick ID for Adjustment " — with a TRAILING SPACE — while resetDailyPickIds wrote
 * "Pick ID for Adjustment" as an inline literal. Against allowInvalid:false,
 * setValue THROWS on a value the rule does not list (proven on this sheet
 * 2026-05-19). The throw landed in the function's own try/catch, at 4am, and
 * Logger.log wrote it where nobody reads. The Adjustment picker silently stopped
 * resetting, so yesterday's picker rolled forward — the exact failure the reset
 * exists to prevent.
 *
 * The fix DERIVES the placeholder: it is the option the _currentPicker gate regex
 * rejects. One rule read off the cell, instead of a second copy of a string that
 * lives in the sheet.
 *
 * Run:   node test-pickid-reset.js
 * HEAD:  SRC=/tmp/head node test-pickid-reset.js     (section A must fail there)
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');

const SRC  = process.env.SRC || path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(SRC, f), 'utf8');

let pass = 0, fail = 0;
const ok = (name, cond, got) => {
  if (cond) { pass++; console.log('  ✓ ' + name); }
  else { fail++; console.log('  ✗ ' + name + (got !== undefined ? '  → got ' + JSON.stringify(got) : '')); }
};
const section = t => console.log('\n' + t);
// ⚠ FAIL SOFT. Against HEAD these functions do not exist, and a bare TypeError would
//   abort the run before the later sections could report — but a before/after proof is
//   only useful if EVERY section reports (the choosePicker lesson, CLAUDE.md).
const block = fn => {
  try { fn(); }
  catch (e) { fail++; console.log('  ✗ ⚠ SECTION ABORTED: ' + e.message); }
};

// ── a cell that behaves like Sheets does, including the strict-validation throw ──
// ⚠ A stub cheaper than the real thing tests nothing (the audio FakeCtx lesson). The
//   whole bug lives in setValue REJECTING a value, so this stub must reject too.
function makeCell(opts, value, allowInvalid, kind) {
  const cell = {
    _v: value,
    _writes: 0,
    getValue: () => cell._v,
    getDataValidation() {
      if (opts === null) return null;
      return {
        getAllowInvalid: () => allowInvalid,
        getCriteriaType: () => kind || 'VALUE_IN_LIST',
        getCriteriaValues: () => [
          // ⚠ VALUE_IN_RANGE hands back a *Range*, not an Array — the shape that makes
          //   getBoardPickers read undefined.length and offer nobody.
          kind === 'VALUE_IN_RANGE' ? { getA1Notation: () => 'Z1:Z9' } : opts,
          true
        ]
      };
    },
    setValue(v) {
      if (opts && !allowInvalid && opts.indexOf(v) === -1) {
        throw new Error('The data you entered in cell violates the data validation rules set on this cell.');
      }
      cell._v = v; cell._writes++; return cell;
    }
  };
  return cell;
}

const SHIP_GATE = /^Shipping\s*-\s*/i;
const ADJ_GATE  = /^adjustment(?:s)?\s*[-:·]\s*/i;

// the REAL option lists, copied verbatim off the live sheet 2026-08-31
const LIVE_SHIP = ['Pick ID for Shipping', 'Shipping - AShamma 12343',
                   'Shipping - Hatem 21332', 'Shipping - YAwiss 1',
                   'Shipping - Turkmani 43122'];
const LIVE_ADJ  = ['Pick ID for Adjustment ', 'Adjustments - AShamma 12343',
                   'Adjustments - AHafiz 21332', 'Adjustments - YAwiss 1',
                   'Adjustments - AAlbri 43122'];

// ── load the real file ──────────────────────────────────────────────────────────
const sandbox = {
  console, JSON, String, Number, Boolean, Array, Object, RegExp, Math, Date,
  SPREADSHEET_ID: 'x', MAIN_SHEET_NAME: 'All orders',
  Schema: { cellEmployeeId: 'F2', cellAdjustmentId: 'H2' },
  Logger: { log: () => {} },
  Utilities: { formatDate: () => '' },
  SpreadsheetApp: null                       // replaced per-test
};
vm.createContext(sandbox);
vm.runInContext(read('ActivityLog.js'), sandbox, { filename: 'ActivityLog.js' });
const G = sandbox;

// ══════════════════════════════════════════════════════════════════════════════
section('A · the live bug — the Adjustment literal is not in its own list');

block(() => {
  // Proof the fixture is real, not invented: the trailing space is what breaks it.
  ok('A1 the live Adjustment list really does carry a trailing space',
     LIVE_ADJ[0] === 'Pick ID for Adjustment ' && LIVE_ADJ.indexOf('Pick ID for Adjustment') === -1);

  const cell = makeCell(LIVE_ADJ, 'Adjustments - YAwiss 1', false);
  const got  = G._pickIdPlaceholder(cell, ADJ_GATE, 'Pick ID for Adjustment');
  ok('A2 ⚠ the placeholder is DERIVED with the space, not the hardcoded literal',
     got === 'Pick ID for Adjustment ', got);

  const r = G._resetOnePickId({ getRange: () => cell }, 'H2', ADJ_GATE, 'Pick ID for Adjustment');
  ok('A3 ⚠⚠ so the 4am write LANDS instead of throwing', cell._v === 'Pick ID for Adjustment ', cell._v);
  ok('A4 and it reports the cell it reset', /H2/.test(r), r);

  // the counterfactual: the old hardcoded literal against the same live rule
  const old = makeCell(LIVE_ADJ, 'Adjustments - YAwiss 1', false);
  let threw = false;
  try { old.setValue('Pick ID for Adjustment'); } catch (e) { threw = true; }
  ok('A5 ⚠ the OLD literal would have thrown against this exact rule', threw);
});

section('B · the Shipping cell was fine, and must stay fine');
block(() => {
  const cell = makeCell(LIVE_SHIP, 'Shipping - YAwiss 1', false);
  const got  = G._pickIdPlaceholder(cell, SHIP_GATE, 'Pick ID for Shipping');
  ok('B1 derives the same value the literal used to hardcode', got === 'Pick ID for Shipping', got);
  G._resetOnePickId({ getRange: () => cell }, 'F2', SHIP_GATE, 'Pick ID for Shipping');
  ok('B2 and the write lands', cell._v === 'Pick ID for Shipping', cell._v);
});

section('C · it must never invent a picker');
block(() => {
  // A list with NO placeholder — every option is a real name.
  const cell = makeCell(['Shipping - A 1', 'Shipping - B 2'], 'Shipping - A 1', false);
  const got  = G._pickIdPlaceholder(cell, SHIP_GATE, 'Pick ID for Shipping');
  ok('C1 ⚠⚠ with no placeholder in the list it falls back to the literal —',
     got === 'Pick ID for Shipping', got);
  ok('C2    NEVER to a real picker name (that would attribute the day to them)',
     !SHIP_GATE.test(got), got);

  const r = G._resetOnePickId({ getRange: () => cell }, 'F2', SHIP_GATE, 'Pick ID for Shipping');
  ok('C3 ⚠ and the resulting throw is REPORTED, not swallowed', /FAILED/.test(r), r);
  ok('C4    while the cell keeps its old value rather than a wrong one',
     cell._v === 'Shipping - A 1', cell._v);
});

section('D · unreadable rules degrade to the literal');
block(() => {
  ok('D1 no validation at all', G._pickIdPlaceholder(makeCell(null, '', true), SHIP_GATE, 'LIT') === 'LIT');
  ok('D2 an empty option list',  G._pickIdPlaceholder(makeCell([], '', true), SHIP_GATE, 'LIT') === 'LIT');
  ok('D3 ⚠ a VALUE_IN_RANGE rule (getCriteriaValues()[0] is a Range, not an Array)',
     G._pickIdPlaceholder(makeCell(LIVE_SHIP, '', true, 'VALUE_IN_RANGE'), SHIP_GATE, 'LIT') === 'LIT');
});

section('E · index 0 wins, but a stray non-matching option still resolves');
block(() => {
  const odd = ['Shipping - A 1', 'Pick ID for Shipping', 'Shipping - B 2'];
  ok('E1 the placeholder is found even when it is not first',
     G._pickIdPlaceholder(makeCell(odd, '', true), SHIP_GATE, 'LIT') === 'Pick ID for Shipping');

  const two = ['Pick ID for Shipping', 'zzz junk', 'Shipping - A 1'];
  ok('E2 ⚠ with two non-matching entries, index 0 wins (what a reader expects)',
     G._pickIdPlaceholder(makeCell(two, '', true), SHIP_GATE, 'LIT') === 'Pick ID for Shipping');
});

section('F · idempotent, and the cells fail independently');
block(() => {
  const cell = makeCell(LIVE_SHIP, 'Pick ID for Shipping', false);
  const r = G._resetOnePickId({ getRange: () => cell }, 'F2', SHIP_GATE, 'Pick ID for Shipping');
  ok('F1 already-reset writes nothing', cell._writes === 0, cell._writes);
  ok('F2 and says so', /already reset/.test(r), r);

  // ⚠⚠ THE REGRESSION THAT MATTERS: a throw on the FIRST cell must not skip the second.
  const bad  = makeCell(['Shipping - A 1'], 'Shipping - A 1', false);   // no placeholder → throws
  const good = makeCell(LIVE_ADJ, 'Adjustments - YAwiss 1', false);
  const sheet = { getRange: a1 => (a1 === 'F2' ? bad : good) };
  G.SpreadsheetApp = { openById: () => ({ getSheetByName: () => sheet }) };

  const out = G.resetDailyPickIds();
  ok('F3 ⚠⚠ the second cell still reset after the first threw',
     good._v === 'Pick ID for Adjustment ', good._v);
  ok('F4 and the failure is named in the returned line', /FAILED/.test(out), out);
  ok('F5 while the line still reports the one that worked', /H2/.test(out), out);
});

console.log('\n' + (fail ? '✗ ' : '✅ ') + 'test-pickid-reset: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
