// ============================================================================
// ORDER LOOK — the quieter DIRECT table (2026-09-30). Loads the REAL files.
//
// A · the three rules: marker, ranges never on STATUS, HOLD regex really \b
// B · every CF stripper keeps them (the 2026-08-30 lesson: two strippers deleted
//     the identity rules on the first edit, and a display layer leaves no trace)
// C · _applyAllConditionalFormatting rebuilds them when ON, drops them when OFF,
//     and appends them AFTER the identity rules (first match wins)
// D · the box painter: finished order pale + drawn first; OFF = today's drawing
// E · nameplate: "· N ready" only when a ready cell is passed
//
// Usage: node test-order-look.js      (SRC=/path/to/old/tree to run against HEAD)
// ============================================================================
'use strict';
const fs = require('fs'), vm = require('vm'), path = require('path');
const SRC = process.env.SRC || path.join(__dirname, '..');
let fails = 0;
const ok = (n, c, got) => { console.log((c ? '  ✓ ' : '  ✗ ') + n + (c ? '' : '  → got ' + JSON.stringify(got))); if (!c) fails++; };

// ---- mocks -----------------------------------------------------------------
const props = {};
function mkRange(r, c, n, w) { return { r, c, n: n || 1, w: w || 1,
  getRow: () => r, getColumn: () => c, getNumRows: () => n || 1, getNumColumns: () => w || 1 }; }
function mkRule(o) { return {
  _o: o,
  getBooleanCondition: () => o.formula == null ? null : ({ getCriteriaType: () => 'CUSTOM_FORMULA', getCriteriaValues: () => [o.formula] }),
  getRanges: () => o.ranges || [] }; }
function builder() {
  const o = {};
  const b = new Proxy({}, { get: (t, k) => {
    if (k === 'build') return () => mkRule(o);
    if (k === 'whenFormulaSatisfied') return f => { o.formula = f; return b; };
    if (k === 'setRanges') return rs => { o.ranges = rs; return b; };
    if (k === 'setFontColor') return c => { o.font = c; return b; };
    return () => b;
  } });
  return b;
}
let cfRules = [];
const borderCalls = [];
function makeSheet(cfg) {
  return {
    getConditionalFormatRules: () => cfRules.slice(),
    setConditionalFormatRules: rs => { cfRules = rs.slice(); },
    getMaxRows: () => cfg.maxRows || 100,
    getRange: (r, c, n, w) => {
      const R = mkRange(r, c, n, w);
      R.getValues = () => Array.from({ length: n || 1 }, (_, i) => [cfg.col(c, r + i)]);
      R.getBackgrounds = () => Array.from({ length: n || 1 }, () => Array(w || 1).fill(null));
      R.setBackgrounds = () => R;
      R.setBorder = function () { borderCalls.push({ r, n, args: [...arguments] }); return R; };
      return R;
    }
  };
}
const ctx = {
  console, Math, JSON, String, Array, Object, Number, RegExp,
  SpreadsheetApp: {
    newConditionalFormatRule: builder,
    BooleanCriteria: { CUSTOM_FORMULA: 'CUSTOM_FORMULA' },
    BorderStyle: { SOLID: 'SOLID', SOLID_MEDIUM: 'SOLID_MEDIUM', SOLID_THICK: 'SOLID_THICK' }
  },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null,
    setProperty: (k, v) => { props[k] = v; }, deleteProperty: k => { delete props[k]; } }) },
  SPREADSHEET_ID: 'x', MAIN_SHEET_NAME: 'All orders'
};
vm.createContext(ctx);
for (const f of ['Config.js', 'Schema.js', 'Helpers.js', 'RowManagement.js', 'BrandTheme.js', 'IdentityGuard.js', 'Holds.js', 'OrderLook.js']) {
  const p = path.join(SRC, f);
  if (!fs.existsSync(p)) { console.log('  (missing ' + f + ' — that is HEAD)'); continue; }
  try { vm.runInContext(fs.readFileSync(p, 'utf8'), ctx, { filename: f }); }
  catch (e) { console.log('  load ' + f + ': ' + e.message); }
}
const has = n => typeof ctx[n] === 'function';
const S = ctx.Schema;

// ---- A ---------------------------------------------------------------------
console.log('\nA · the three rules');
let rules = [];
try { rules = ctx._buildOrderLookRules(makeSheet({ col: () => '' })); } catch (e) {}
ok('three rules built', rules.length === 3, rules.length);
const fx = rules.map(r => r._o.formula || '');
ok('every rule carries the hq-look marker', fx.length === 3 && fx.every(f => f.indexOf('N("hq-look")') !== -1), fx);
ok('no rule range touches STATUS (F)', rules.length === 3 && rules.every(r => r._o.ranges.every(g => !(g.c <= S.cols.STATUS && g.c + g.w - 1 >= S.cols.STATUS))));
ok('all rules only set a quiet font', rules.length === 3 && rules.every(r => r._o.font === '#a39b86'));
const note = fx[2] || '';
ok('HOLD regex reaches Sheets as \\b (not a backspace)', note.indexOf('(?i)\\bhold\\b') !== -1 && note.indexOf('\b') === -1, note);
ok('kit (↳) and Zoho-flag (⚠) notes exempt', /LEFT\(\$E4,1\)<>"↳"/.test(note) && /LEFT\(\$E4,1\)<>"⚠"/.test(note));
ok('SO rule avoids COUNTIF / TRIM( (col-D stripper words)', !/COUNTIF|TRIM\(/.test(fx[1] || 'COUNTIF'), fx[1]);
ok('done rule needs this line terminal AND no open line', /OR\(\$F4="SHIPPED",\$F4="CANCELED"\)/.test(fx[0]) && /"PENDING"\)\+COUNTIFS/.test(fx[0]));

// ---- B ---------------------------------------------------------------------
console.log('\nB · the strippers keep them');
const legacyD = mkRule({ formula: '=COUNTIF($D:$D,$D4)>1', ranges: [mkRange(4, S.cols.SALES_ORDER, 997, 1)] });
const legacyA = mkRule({ formula: '=UPPER(TRIM(A4))="X"', ranges: [mkRange(4, S.cols.SKU, 1000, 1)] });
cfRules = rules.concat([legacyD, legacyA]);
try { ctx.removeLegacySalesOrderCFRules(makeSheet({ col: () => '' })); } catch (e) { console.log('  ' + e.message); }
ok('col-D stripper keeps all three', rules.every(r => cfRules.indexOf(r) !== -1));
ok('…and still removes a real legacy col-D rule', cfRules.indexOf(legacyD) === -1);
try { ctx.removeDuplicateHighlightRules(makeSheet({ col: () => '' })); } catch (e) { console.log('  ' + e.message); }
ok('col-A stripper keeps all three', rules.every(r => cfRules.indexOf(r) !== -1));
ok('…and still removes a real duplicate-SKU rule', cfRules.indexOf(legacyA) === -1);

// ---- C ---------------------------------------------------------------------
console.log('\nC · a theme re-apply');
function themeRun() {
  cfRules = rules.slice();
  try { ctx._applyAllConditionalFormatting(makeSheet({ col: () => '' })); } catch (e) { console.log('  ' + e.message); }
  return cfRules.filter(r => String(r._o.formula || '').indexOf('hq-look') !== -1);
}
props.ORDER_LOOK_ON = 'on';
let look = themeRun();
ok('ON: exactly three look rules after a re-apply (no duplicates, none lost)', look.length === 3, look.length);
const idIdx = cfRules.findIndex(r => /INDIRECT\("'__Identity'|__Identity/.test(r._o.formula || ''));
const firstLook = cfRules.findIndex(r => String(r._o.formula || '').indexOf('hq-look') !== -1);
ok('ON: look rules come AFTER the identity rules', idIdx !== -1 && firstLook > idIdx, [idIdx, firstLook]);
ok('ON: look rules are the very last rules', firstLook === cfRules.length - 3, [firstLook, cfRules.length]);
delete props.ORDER_LOOK_ON;
look = themeRun();
ok('OFF: a re-apply removes them', look.length === 0, look.length);

// ---- D ---------------------------------------------------------------------
console.log('\nD · the order boxes');
// DIRECT divider at 10 → data from 12. A live · B finished · C live.
const SO = { 12: 'SO-A', 13: 'SO-A', 14: 'SO-B', 15: 'SO-B', 16: 'SO-C' };
const ST = { 12: 'PREPARING', 13: 'PENDING', 14: 'SHIPPED', 15: 'CANCELED', 16: 'PREPARING' };
const sheetD = makeSheet({ maxRows: 30, col: (c, r) => c === S.cols.STATUS ? (ST[r] || '') : c === S.cols.NOTE ? '' : (SO[r] || '') });
const colD = []; for (let r = S.dataStartRow; r <= 16; r++) colD.push([SO[r] || '']);
function paint() {
  borderCalls.length = 0;
  ctx._paintDirectOrderDividers(sheetD, 10, colD, 16, 0);
  return borderCalls.filter(b => b.args.length >= 7);   // the per-block box writes
}
props.ORDER_LOOK_ON = 'on';
let boxes = paint();
const byRow = r => boxes.find(b => b.r === r);
ok('three boxes drawn', boxes.length === 3, boxes.length);
ok('finished SO-B: pale colour', byRow(14) && byRow(14).args[6] === '#ddd3b0', byRow(14) && byRow(14).args[6]);
ok('finished SO-B: thin line', byRow(14) && byRow(14).args[7] === 'SOLID', byRow(14) && byRow(14).args[7]);
ok('live SO-A and SO-C stay gold', byRow(12) && byRow(12).args[6] === '#c9a227' && byRow(16) && byRow(16).args[6] === '#c9a227');
ok('finished box drawn FIRST (live owns the shared edge)', boxes.length === 3 && boxes[0].r === 14, boxes.map(b => b.r));
delete props.ORDER_LOOK_ON;
boxes = paint();
ok('OFF: every box gold, top-down, as today', boxes.map(b => b.r + ':' + b.args[6]).join(',') === '12:#c9a227,14:#c9a227,16:#c9a227', boxes.map(b => b.r + ':' + b.args[6]));

// ---- E ---------------------------------------------------------------------
console.log('\nE · the band nameplate');
const plain = ctx._nameplateFormula('DIRECT ORDERS', 'A18', 'A29');
const withReady = ctx._nameplateFormula('DIRECT ORDERS', 'A18', 'A29', 'A31');
ok('no ready cell → no "ready" text', plain.indexOf('ready') === -1);
ok('ready cell → "· N ready", hidden at 0 / blank', /OR\('__SparkData'!A31="",'__SparkData'!A31=0\)/.test(withReady) && withReady.indexOf('" ready"') !== -1, withReady);

console.log('\n' + (fails ? `✗ ${fails} FAILED` : '✓ ALL PASSED'));
process.exit(fails ? 1 : 0);
