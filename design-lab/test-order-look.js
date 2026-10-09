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
      R.getValue = () => cfg.col(c, r);
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
ok('four rules built', rules.length === 4, rules.length);
const fx = rules.slice(1).map(r => r._o.formula || '');   // [0] is kit-ready, checked in G
ok('every rule carries the hq-look marker', rules.length === 4 && rules.every(r => (r._o.formula || '').indexOf('N("hq-look")') !== -1));
ok('no rule range touches STATUS (F)', rules.length === 4 && rules.every(r => r._o.ranges.every(g => !(g.c <= S.cols.STATUS && g.c + g.w - 1 >= S.cols.STATUS))));
ok('the three quiet rules only set a quiet font', rules.slice(1).every(r => r._o.font === '#a39b86'));
const note = fx[2] || '';
ok('HOLD regex reaches Sheets as \\b (not a backspace)', note.indexOf('(?i)\\bhold\\b') !== -1 && note.indexOf('\b') === -1, note);
ok('kit (↳) and Zoho-flag (⚠) notes exempt', /ISERROR\(FIND\("↳",\$E4&""\)\)/.test(note) && /LEFT\(\$E4,1\)<>"⚠"/.test(note));
ok('SO rule avoids COUNTIF / TRIM( (col-D stripper words)', !/COUNTIF|TRIM\(/.test(fx[1] || 'COUNTIF'), fx[1]);
ok('done rule needs this line terminal AND no open line', /OR\(\$F4="SHIPPED",\$F4="CANCELED"\)/.test(fx[0]) && /"PENDING"\)\+COUNTIFS/.test(fx[0]));

// ---- B ---------------------------------------------------------------------
console.log('\nB · the strippers keep them');
const legacyD = mkRule({ formula: '=COUNTIF($D:$D,$D4)>1', ranges: [mkRange(4, S.cols.SALES_ORDER, 997, 1)] });
const legacyA = mkRule({ formula: '=UPPER(TRIM(A4))="X"', ranges: [mkRange(4, S.cols.SKU, 1000, 1)] });
cfRules = rules.concat([legacyD, legacyA]);
try { ctx.removeLegacySalesOrderCFRules(makeSheet({ col: () => '' })); } catch (e) { console.log('  ' + e.message); }
ok('col-D stripper keeps all four', rules.every(r => cfRules.indexOf(r) !== -1));
ok('…and still removes a real legacy col-D rule', cfRules.indexOf(legacyD) === -1);
try { ctx.removeDuplicateHighlightRules(makeSheet({ col: () => '' })); } catch (e) { console.log('  ' + e.message); }
ok('col-A stripper keeps all four (incl. the col-A kit rule with COUNTIFS)', rules.every(r => cfRules.indexOf(r) !== -1));
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
ok('ON: exactly four look rules after a re-apply (no duplicates, none lost)', look.length === 4, look.length);
const idIdx = cfRules.findIndex(r => /INDIRECT\("'__Identity'|__Identity/.test(r._o.formula || ''));
const firstLook = cfRules.findIndex(r => String(r._o.formula || '').indexOf('hq-look') !== -1);
ok('ON: look rules come AFTER the identity rules', idIdx !== -1 && firstLook > idIdx, [idIdx, firstLook]);
ok('ON: look rules are the very last rules', firstLook === cfRules.length - 4, [firstLook, cfRules.length]);
delete props.ORDER_LOOK_ON;
look = themeRun();
ok('OFF: a re-apply removes them', look.length === 0, look.length);

// ---- D ---------------------------------------------------------------------
console.log('\nD · the order boxes');
// DIRECT divider at 10 → data from 12. A live · B finished · C live.
const SO = { 12: 'SO-A', 13: 'SO-A', 14: 'SO-B', 15: 'SO-B', 16: 'SO-C' };
const ST = { 12: 'PREPARING', 13: 'PENDING', 14: 'SHIPPED', 15: 'CANCELED', 16: 'PREPARING' };
const sheetD = makeSheet({ maxRows: 30, col: (c, r) => (c === S.cols.SKU && r === 10) ? 'DIRECT' : c === S.cols.STATUS ? (ST[r] || '') : c === S.cols.NOTE ? '' : (SO[r] || '') });
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
borderCalls.length = 0;
ctx._paintDirectOrderDividers(sheetD, 10, colD, 16, 0, { noBoxes: true });
ok('noBoxes (Amazon): clears but draws no per-order box', borderCalls.filter(b => b.args.length >= 7).length === 0 && borderCalls.length > 0, borderCalls.length);
borderCalls.length = 0;
ctx._paintDirectOrderDividers(sheetD, 11, colD, 16, 0);   // divider "moved" — row 11 holds no marker
ok('rows moved under the paint: nothing written', borderCalls.length === 0, borderCalls.length);
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

// ---- F ---------------------------------------------------------------------
console.log('\nF · K numbering (the one rule the print will share)');
const kits = new Set(['217205', '157644', '158679']);
const R = (sku, so, note) => ({ sku, so, note: note || '' });
const plan = ctx._kitTagPlan ? ctx._kitTagPlan([
  R('164979', 'SO-1', '↳ from KIT-217205 · Miguel'),   // 0 part of 217205
  R('217205', 'SO-1', 'Miguel'),                        // 1 parent → K2 (157644 < 217205)
  R('157644', 'SO-1', ''),                              // 2 parent → K1
  R('172539', 'SO-1', '↳ from KIT-157644'),             // 3 part of 157644
  R('199999', 'SO-1', '↳ added to KIT-157644'),         // 4 custom add → K1
  R('158679', 'SO-1', ''),                              // 5 kit NOT expanded here → plain
  R('171111', 'SO-1', 'Miguel'),                        // 6 loose
  R('157644', 'SO-2', ''),                              // 7 other order: its OWN number, K3
  R('160000', 'SO-2', '⚠️ QTY: 2 → 1 IN ZOHO\n↳ from KIT-157644'),  // 8 flagged part → K3
  R('155555', 'SO-3', '↳ from KIT-217205')              // 9 orphan part: parent not in SO-3
], kits) : [];
const fmt = e => e ? (e.parent ? 'P' : 'C') + e.k : '-';
ok('numbered by order, then kit SKU lowest first; unique across the sheet', plan.map(fmt).join(' ') === 'C2 P2 P1 C1 C1 - - P3 C3 -', plan.map(fmt).join(' '));
// 2026-10-09 floor report: two orders, one kit each, both read "K1".
const floor = ctx._kitTagPlan ? ctx._kitTagPlan([
  R('164988', '12-15269-54269', '↳ from KIT-157563 · deploy 3 total'), R('157563', '12-15269-54269', ''),
  R('194568', '12-15269-54269', 'HOLD . ↳ from KIT-157563 · deploy 3 total'),
  R('173817', '06-15279-95092', '↳ from KIT-159093'), R('159093', '06-15279-95092', '')
], new Set(['157563', '159093'])) : [];
ok('two kits on two orders never share a K number', floor.length && floor[1].k !== floor[4].k && floor[0].k === floor[1].k && floor[2].k === floor[1].k && floor[3].k === floor[4].k, floor.map(fmt));
// sticky: a kit keeps its number; a new kit takes the lowest FREE number
const st = ctx._kitTagPlan ? ctx._kitTagPlan([
  R('159093', 'SO-9', ''), R('1', 'SO-9', '↳ from KIT-159093'),
  R('157563', 'SO-1', ''), R('2', 'SO-1', '↳ from KIT-157563')
], new Set(['157563', '159093']), { 'SO-1|157563': 2 }) : [];
ok('a kit keeps the number it already had (K2 stays K2)', st.length && st[2].k === 2 && st[3].k === 2, st.map(fmt));
ok('a new kit takes the lowest free number (K1)', st.length && st[0].k === 1, st.map(fmt));
// the sticky wrapper remembers between runs
const store = {};
ctx.PropertiesService.getDocumentProperties = () => ({ getProperty: k => store[k] || null, setProperty: (k, v) => { store[k] = v; } });
const kitsB = new Set(['157563', '159093']);
const run1 = ctx._kitTagPlanSticky ? ctx._kitTagPlanSticky([R('157563', 'SO-5', ''), R('9', 'SO-5', '↳ from KIT-157563')], kitsB) : [];
const run2 = ctx._kitTagPlanSticky ? ctx._kitTagPlanSticky([R('159093', 'SO-0', ''), R('8', 'SO-0', '↳ from KIT-159093'),
                                                            R('157563', 'SO-5', ''), R('9', 'SO-5', '↳ from KIT-157563')], kitsB) : [];
ok('remembered: an earlier kit is not renumbered when a new order sorts above it', run1.length && run1[0].k === 1 && run2[2].k === 1 && run2[0].k === 2, [run1.map(fmt), run2.map(fmt)]);
ok('parent format "▣ K1 "@, part "K1 "@', ctx._kitTagFormat && ctx._kitTagFormat({ k: 1, parent: true }) === '"▣ K1 "@' && ctx._kitTagFormat({ k: 3, parent: false }) === '"K3 "@');
const re = ctx._kitTagPlan ? ctx._kitTagPlan([2,0,1,3,4,5,6,7,8,9].map(i => [
  R('164979','SO-1','↳ from KIT-217205 · Miguel'), R('217205','SO-1','Miguel'), R('157644','SO-1',''),
  R('172539','SO-1','↳ from KIT-157644'), R('199999','SO-1','↳ added to KIT-157644'), R('158679','SO-1',''),
  R('171111','SO-1','Miguel'), R('157644','SO-2',''), R('160000','SO-2','⚠️ QTY: 2 → 1 IN ZOHO\n↳ from KIT-157644'),
  R('155555','SO-3','↳ from KIT-217205')][i]), kits) : [];
ok('re-sorting the rows never renumbers a kit', re.length && re[0].k === 1 && re[1].k === 2 && re[2].k === 2, re.map(fmt));

// ---- G ---------------------------------------------------------------------
console.log('\nG · kit ready + the marker routine');
const kitRule = rules[0] || { _o: {} };
const kf = kitRule._o.formula || '';
ok('kit-ready sits on the SKU column only', kitRule._o.ranges && kitRule._o.ranges.length === 1 && kitRule._o.ranges[0].c === S.cols.SKU && kitRule._o.ranges[0].w === 1);
ok('matches "…KIT-<sku>" alone or followed by a space (no prefix collision)', kf.indexOf('"*↳ * KIT-"&$A4)') !== -1 && kf.indexOf('"*↳ * KIT-"&$A4&" *"') !== -1, kf);
ok('green only while parent is live and no part PENDING', /\$F4<>"SHIPPED"/.test(kf) && /"PENDING"\)=0\)$/.test(kf));
// refreshKitSkuMarkers end to end on a fake sheet
const colVals = { 1: ['157644','172539','171111','217205','164979'], 4: ['SO-1','SO-1','SO-1','SO-1','SO-1'],
                  5: ['', '↳ from KIT-157644', 'x', '', '↳ from KIT-217205'] };
let written = null;
const sheetK = { getLastRow: () => S.dataStartRow + 4, getRange: (r, c, n) => ({
  getValues: () => (colVals[c] || []).slice(0, n).map(v => [v]),
  getNumberFormats: () => Array.from({ length: n }, () => ['@']),
  setNumberFormats: f => { written = f.map(x => x[0]); } }) };
ctx.SpreadsheetApp.openById = () => ({ getSheetByName: () => sheetK });
ctx.SpreadsheetApp.flush = () => {};
ctx._obIsOwner = () => true;
ctx.buildKitMap = () => new Map([['157644', {}], ['217205', {}]]);
props.ORDER_LOOK_ON = 'on';
try { ctx.refreshKitSkuMarkers(); } catch (e) { console.log('  ' + e.message); }
ok('ON: parents and parts wear their K tags', JSON.stringify(written) === JSON.stringify(['"▣ K1 "@', '"K1 "@', '@', '"▣ K2 "@', '"K2 "@']), written);
delete props.ORDER_LOOK_ON;
try { ctx.refreshKitSkuMarkers(); } catch (e) {}
ok('OFF: exactly today\'s marks (▣ on parents, parts plain)', JSON.stringify(written) === JSON.stringify(['"▣ "@', '@', '@', '"▣ "@', '@']), written);

console.log('\n' + (fails ? `✗ ${fails} FAILED` : '✓ ALL PASSED'));
process.exit(fails ? 1 : 0);
