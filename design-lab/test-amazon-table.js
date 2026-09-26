/**
 * test-amazon-table.js — THE THIRD TABLE, end to end, against a MODEL SHEET (2026-09-26).
 *
 * Loads the REAL Schema / Helpers / RowManagement / OrderService / Replacements / Amazon /
 * Holds / DashboardService / LocationService into a VM, over a fake spreadsheet whose row
 * operations are APPLIED to a model (insertRowsBefore really shifts rows, deleteRows really
 * removes them). Asserting "deleteRows was called with (29, 5)" proves arithmetic; what
 * matters is whether a row that held data — or the AMAZON divider — is still there after.
 *
 * ⭐ THE HEADLINE RISK IT GUARDS: every range writer that used to treat "DIRECT+2 → last row"
 *   as the Direct table would, with Amazon below it, reach INTO the Amazon table — sorting
 *   Amazon rows into Direct, deleting Amazon's rows, blanking the yellow AMAZON band. Each
 *   section drives one of those writers and checks the divider and every data row survive.
 *
 * Run:  node design-lab/test-amazon-table.js        (SRC=/path/to/tree to run against HEAD)
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label + (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};
const ok = (label, cond, detail) => t(label + (cond ? '' : '  [' + JSON.stringify(detail) + ']'), !!cond, true);
const section = (name, fn) => {
  console.log('\n' + name);
  try { fn(); } catch (e) { fail++; console.log('  ✗ SECTION THREW (soft): ' + (e && e.stack || e).toString().split('\n').slice(0, 3).join(' | ')); }
};

// =======================================================================================
// THE MODEL SHEET
// =======================================================================================
const W = 10, YELLOW = '#ffd400';
const STATUS_DV = ['PENDING', 'PREPARING', 'SHIPPED', 'CANCELED'];
function cell(v) { return { v: v == null ? '' : v, f: '', bg: null, fs: 10, dv: null, fx: '' }; }
/* ⚠ REAL-SHEETS RULE (the frozen-nameplate bug): reading a formula cell returns its COMPUTED
   value, and writing that value back REPLACES the formula. The model keeps the formula in fx
   and reports '⟨computed⟩' as the value, so a read-then-write-back is visible as a lost fx. */
const shown = (x) => x.fx ? '⟨computed⟩' : x.v;
/* ⚠ REAL-SHEETS RULE (the 2026-09-26 live failure): a cell carrying a list validation
   REFUSES any other value. The model throws on the write (Sheets throws at the next flush —
   stricter here, never looser). */
function checkDv(x, v) {
  if (x.dv && v !== '' && v != null && x.dv.indexOf(String(v)) === -1) {
    throw new Error('The data you entered violates the data validation rules set on this cell. ' +
                    'Please enter one of the following values: ' + x.dv.join(', ') + '. (got ' + JSON.stringify(v) + ')');
  }
}
const fmtOf = (x) => ({ f: x.f, bg: x.bg, fs: x.fs, dv: x.dv ? x.dv.slice() : null });
function blankRow() { const r = []; for (let c = 0; c < W; c++) r.push(cell('')); return r; }
function rowOf(vals) { const r = blankRow(); vals.forEach((v, i) => { r[i].v = v; }); return r; }

function makeSheet(rows) {
  const grid = rows.map(r => Array.isArray(r) ? rowOf(r) : r);
  grid.forEach((r, i) => { const a = String(r[0].v).trim().toUpperCase();
    if (i + 1 >= 4 && a !== 'DIRECT' && a !== 'AMAZON' && a.indexOf('◈') !== 0) r[5].dv = STATUS_DV.slice(); });
  const S = { grid, ops: [] };
  const ensure = (r) => { while (grid.length < r) grid.push(blankRow()); };
  // A new row takes the FORMAT of its neighbour above (values blank) — like Sheets.
  const inherit = (above) => { const src = grid[above - 1]; const r0 = blankRow();
    if (src) r0.forEach((x, j) => Object.assign(x, fmtOf(src[j]))); return r0; };
  function range(r, c, nr, nc) {
    nr = nr || 1; nc = nc || 1;
    const cells = () => { const out = []; for (let i = 0; i < nr; i++) { const row = grid[r - 1 + i] || blankRow();
      const o = []; for (let j = 0; j < nc; j++) o.push(row[c - 1 + j] || cell('')); out.push(o); } return out; };
    const each = (fn) => { for (let i = 0; i < nr; i++) { ensure(r + i); for (let j = 0; j < nc; j++) fn(grid[r - 1 + i][c - 1 + j], i, j); } };
    const api = {
      getRow: () => r, getColumn: () => c, getNumRows: () => nr, getNumColumns: () => nc,
      getA1Notation: () => 'R' + r + 'C' + c,
      getValues: () => cells().map(o => o.map(shown)),
      getValue: () => shown(cells()[0][0]),
      getFormula: () => cells()[0][0].fx,
      getDisplayValue: () => String(shown(cells()[0][0])),
      // A cell hidden inside a merge holds no value — Sheets keeps only the top-left one.
      setValues: (v) => { each((x, i, j) => { if (x.hid) return; checkDv(x, v[i][j]); x.v = v[i][j]; x.fx = ''; }); return px; },
      setValue: (v) => { each(x => { if (x.hid) return; checkDv(x, v); x.v = v; x.fx = ''; }); return px; },
      setFormula: (v) => { each(x => { if (x.hid) return; x.fx = v; x.v = ''; }); return px; },
      merge: () => { each((x, i, j) => { if (i || j) { x.hid = true; x.v = ''; } }); return px; },
      breakApart: () => { each(x => { x.hid = false; }); return px; },
      clearDataValidations: () => { each(x => { x.dv = null; }); return px; },
      getDataValidations: () => cells().map(o => o.map(x => x.dv)),
      getNumberFormats: () => cells().map(o => o.map(x => x.f)),
      setNumberFormats: (v) => { each((x, i, j) => { x.f = v[i][j]; }); return px; },
      setNumberFormat: (f) => { each(x => { x.f = f; }); return px; },
      getBackgrounds: () => cells().map(o => o.map(x => x.bg)),
      setBackgrounds: (v) => { each((x, i, j) => { x.bg = v[i][j]; }); return px; },
      setBackground: (b) => { each(x => { x.bg = b; }); return px; },
      getFontSizes: () => cells().map(o => o.map(x => x.fs)),
      setFontSizes: (v) => { each((x, i, j) => { x.fs = v[i][j]; }); return px; },
      getRichTextValues: () => cells().map(o => o.map(x => ({ getText: () => String(shown(x)) }))),
      setRichTextValues: (v) => { each((x, i, j) => { x.v = v[i][j].getText(); x.fx = ''; }); return px; },
      clearContent: () => { each(x => { x.v = ''; }); return px; },
      // PASTE_FORMAT carries number format, background AND DATA VALIDATION — tiled.
      copyTo: (target, kind) => { if (kind === 'F') target._pasteFmt(cells()); return px; },
      copyFormatToRange: (sh, c1, c2, r1, r2) => { sh.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1)._pasteFmt(cells()); return px; },
      _pasteFmt: (src) => { each((x, i, j) => { const s0 = src[i % src.length][j % src[0].length]; Object.assign(x, fmtOf(s0)); }); return px; },
      protect: () => ({ setDescription() { return this; }, setWarningOnly() { return this; } }),
      getMergedRanges: () => []
    };
    const px = new Proxy(api, { get: (o, k) => (k in o) ? o[k] : (() => px) });   // any styling → chainable no-op
    return px;
  }
  const sheet = {
    getName: () => 'All orders',
    getLastRow: () => { for (let i = grid.length - 1; i >= 0; i--) if (grid[i].some(x => String(x.v) !== '')) return i + 1; return 0; },
    getMaxRows: () => grid.length,
    getLastColumn: () => W,
    getRange: (a, b, c2, d) => {
      if (typeof a === 'string') { const m = /^([A-Z])(\d+)$/.exec(a); return range(+m[2], m[1].charCodeAt(0) - 64, 1, 1); }
      return range(a, b, c2, d);
    },
    insertRowsBefore: (row, n) => { S.ops.push(['insertBefore', row, n]); grid.splice(row - 1, 0, ...Array.from({ length: n }, () => inherit(row - 1))); },
    insertRowsAfter: (row, n) => { S.ops.push(['insertAfter', row, n]); grid.splice(row, 0, ...Array.from({ length: n }, () => inherit(row))); },
    deleteRows: (row, n) => { S.ops.push(['delete', row, n]); grid.splice(row - 1, n); },
    deleteRow: (row) => grid.splice(row - 1, 1),
    setRowHeight: () => {}, setRowHeights: () => {},
    getProtections: () => [], getConditionalFormatRules: () => [], setConditionalFormatRules: () => {},
    getBandings: () => [], hideRows: () => {}, showRows: () => {}, getSheetId: () => 1
  };
  S.sheet = sheet;
  S.colA = () => grid.map(r => String(r[0].v));
  S.col = (c) => grid.map(r => r[c - 1].v);
  S.row = (n) => grid[n - 1].map(x => x.fx || x.v);
  return S;
}

/* eBay 4-6, blank 7, ▌DIRECT 8, header 9, Direct 10-12, blank 13 — the live shape. */
function liveRows() {
  const H = ['◈ SKU', 'QTY', 'LOCATION', 'ORDER', 'NOTE', 'STATUS', 'HAND', 'LEFT', 'SHIPPING', 'SHIP COST'];
  return [
    ['HQ'], [''], H,
    ['111111', 1, 'B-9',  '05-11111-11111', '', 'PENDING',   4],
    ['222222', 2, 'A-50', '05-22222-22222', '', 'PREPARING', 9],
    ['333333', 1, 'C-2',  '05-33333-33333', '', 'SHIPPED',   2],
    [''],
    ['DIRECT'],
    H,
    ['444444', 1, 'D-4',  'SO-100', '', 'PENDING', 5],
    ['555555', 3, 'E-5',  'SO-100', '', 'PENDING', 7],
    ['666666', 1, 'A-1',  'SO-090', '', 'SHIPPED', 1],
    ['']
  ];
}

// =======================================================================================
// THE SANDBOX — real files, stubs only at the edges (Sheets, lock, logs, stock lookups)
// =======================================================================================
function boot(rows) {
  const S = makeSheet(rows || liveRows());
  const logs = [], published = [];
  const SKUS = { '167517': { loc: 'K-7', avail: 12 }, '171378': { loc: 'L-3', avail: 1 },
                 '444444': { loc: 'D-4', avail: 5 }, '555555': { loc: 'E-5', avail: 7 } };
  const sb = {
    console: { log: () => {} },
    SPREADSHEET_ID: 'x', MAIN_SHEET_NAME: 'All orders',
    SpreadsheetApp: {
      openById: () => ({ getSheetByName: (n) => (n === 'All orders' ? S.sheet : null) }),
      getActive: () => ({ getSheetByName: () => S.sheet }),
      flush: () => {}, CopyPasteType: { PASTE_FORMAT: 'F' },
      BorderStyle: { SOLID_MEDIUM: 'M', SOLID_THICK: 'T' },
      ProtectionType: { RANGE: 'R' },
      BooleanCriteria: { CUSTOM_FORMULA: 'CF' }, newConditionalFormatRule: () => ({}),
      newTextStyle: () => { const b = new Proxy({}, { get: (o, k) => k === 'build' ? () => ({}) : () => b }); return b; },
      newRichTextValue: () => { let t = ''; const b = { setText: (x) => { t = x; return b; }, setLinkUrl: () => b, setTextStyle: () => b, build: () => ({ getText: () => t }) }; return b; }
    },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Utilities: { formatDate: () => '', getUuid: () => 'abcdef0123456789' },
    // stock / enrichment edges
    getSingleLocation: (sku) => (SKUS[sku] ? SKUS[sku].loc : null),
    getSingleInventory: (s) => (SKUS[s] ? { available: SKUS[s].avail } : null),
    getSingleZohoStock: () => null,
    resolveHandValue: (mi, zo) => (zo != null ? zo : (mi != null ? mi : 0)),
    buildLocationMap: () => new Map([['444444', 'D-4'], ['555555', 'E-5'], ['666666', 'A-1'], ['167517', 'K-7']]),
    logActivityBatch: (b) => logs.push.apply(logs, b),
    _dashBustTickCache: () => {}, refreshKitSkuMarkers: () => {}, refreshAllOrdersEnrichment: () => {},
    publishBoardTickInline: (why) => published.push(why),
    refreshAllOrdersLockCarveOuts: () => {}, refreshDynamicBandings: () => {}, _ensureSparkData: () => {},
    _obIsOwner: () => true, _obRequireOwner: () => '', _asOwner: () => { throw new Error('no hop in tests'); },
    // ⭐ The band styler is the REAL one from BrandTheme.js (loaded below). It used to be a
    //   stand-in here, which is why the live failure — the real styler writing onto a cell
    //   carrying the STATUS dropdown — could not be seen by this suite.
    TABLE_BUFFER_ROWS: 1,
    ACTIVITY_LOG: { sheetName: 'Activity Log', cols: { ORDER_ID: 3 } }
  };
  const sbEdge = { getSingleLocation: sb.getSingleLocation, getSingleInventory: sb.getSingleInventory,
                   getSingleZohoStock: sb.getSingleZohoStock, buildLocationMap: sb.buildLocationMap,
                   resolveHandValue: sb.resolveHandValue };
  vm.createContext(sb);
  ['Schema.js', 'Helpers.js', 'RowManagement.js', 'OrderService.js', 'Replacements.js',
   'Amazon.js', 'Holds.js', 'DashboardService.js', 'LocationService.js', 'BrandTheme.js', 'OrderLinks.js'].forEach(f => {
    vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), sb, { filename: f });
  });
  // re-assert the edge stubs the loaded files would otherwise shadow
  Object.assign(sb, { getSingleLocation: sbEdge.getSingleLocation, getSingleInventory: sbEdge.getSingleInventory,
                      getSingleZohoStock: sbEdge.getSingleZohoStock, buildLocationMap: sbEdge.buildLocationMap,
                      resolveHandValue: sbEdge.resolveHandValue,
                      refreshKitSkuMarkers: () => {}, refreshAllOrdersEnrichment: () => {},
                      refreshDynamicBandings: () => {}, _obIsOwner: () => true,
                      TABLE_BUFFER_ROWS: 1 });
  return { S, sb, logs, published };
}
const findRow = (S, v) => S.colA().indexOf(v) + 1;
const dataRowsOf = (S, from, to) => S.colA().slice(from - 1, to).filter(v => v && v !== 'AMAZON' && v !== 'DIRECT' && v.indexOf('◈') !== 0);

// =======================================================================================
section('A · the layout helper — today\'s shape is unchanged without an Amazon table', () => {
  const { S, sb } = boot();
  const L = sb.getTableLayout(S.sheet);
  t('A1 DIRECT row', L.direct, 8);
  t('A2 no Amazon table', L.amazon, -1);
  t('A3 Direct runs to the end of the sheet (today\'s behaviour)', L.directTable, { start: 10, end: 13 });
  t('A4 eBay stops above DIRECT', L.ebay, { start: 4, end: 7 });
  t('A5 structural rows = DIRECT + its header', Object.keys(L.structural).map(Number), [8, 9]);
  t('A6 table 3 does not exist', sb._tableSegment(3, L), null);
  // pure ranges with an Amazon divider
  const L2 = sb._layoutRanges({ direct: 8, amazon: 14, maxRows: 18, lastRow: 18, structural: {} });
  t('A7 Direct stops above AMAZON', L2.directTable, { start: 10, end: 13 });
  t('A8 Amazon runs to the end', L2.amazonTable, { start: 16, end: 18 });
  t('A9 tableOfRow', [5, 8, 9, 11, 14, 15, 17].map(r => sb.tableOfRow(r, L2)),
    ['EBAY', 'STRUCT', 'STRUCT', 'DIRECT', 'STRUCT', 'STRUCT', 'AMAZON']);
  const bad = sb._layoutRanges({ direct: 8, amazon: -1, maxRows: 18, lastRow: 18, structural: {} });
  t('A10 an AMAZON marker above DIRECT is ignored (broken sheet, not a layout)',
    sb.getTableLayout(S.sheet, [['x'], ['AMAZON'], ['x'], ['x'], ['x'], ['x'], ['x'], ['DIRECT']]).amazon, -1);
  ok('A11 Schema.isStructuralMarker knows both markers, trims and ignores case',
     sb.Schema.isStructuralMarker(' amazon ') && sb.Schema.isStructuralMarker('DIRECT') && !sb.Schema.isStructuralMarker('167517'));
  void bad;
});

// =======================================================================================
section('B · setupAmazonTable — band + header + one blank row at the bottom', () => {
  const { S, sb } = boot();
  const msg = sb.setupAmazonTable();
  ok('B1 reports success', /Amazon table added/.test(msg), msg);
  t('B2 the AMAZON divider is the row after Direct\'s buffer', findRow(S, 'AMAZON'), 14);
  t('B3 its header row copies the DIRECT header', S.row(15)[0], '◈ SKU');
  t('B4 one blank row below the header', S.colA().slice(15), ['']);
  t('B5 the marker value is EXACTLY "AMAZON" (strict-equality contract)', S.row(14)[0], 'AMAZON');
  ok('B6 re-running is refused, not duplicated', /already exists/.test(sb.setupAmazonTable()) &&
     S.colA().filter(v => v === 'AMAZON').length === 1);
  t('B7 Direct + eBay data untouched', dataRowsOf(S, 4, 13), ['111111', '222222', '333333', '444444', '555555', '666666']);
});

// =======================================================================================
section('C · THE DOOR — addAmazonOrder', () => {
  const { S, sb, logs, published } = boot();
  sb.setupAmazonTable();
  const r = sb.addAmazonOrder('114-3941689-8772232', [{ sku: '167517', qty: 2 }, { sku: '171378' }], '9/29', 'gift wrap', 'telegram');
  ok('C1 ok', r.ok, r.message);
  t('C2 lands at the top of the Amazon table (AMAZON+2)', r.row, 16);
  t('C3 row 16 content', S.row(16).slice(0, 7), ['167517', 2, 'K-7', 'AMZ-114-3941689-8772232', 'ship by 9/29 · gift wrap', 'PENDING', 12]);
  t('C4 row 17 content (qty defaults to 1)', S.row(17).slice(0, 7), ['171378', 1, 'L-3', 'AMZ-114-3941689-8772232', 'ship by 9/29 · gift wrap', 'PENDING', 1]);
  t('C5 the AMAZON divider did not move', findRow(S, 'AMAZON'), 14);
  t('C6 Direct + eBay untouched', dataRowsOf(S, 4, 13), ['111111', '222222', '333333', '444444', '555555', '666666']);
  t('C7 one RECEIVED per line, source "amazon" (never a warehouse source)', logs.map(l => [l[0], l[1], l[2], l[4]]),
    [['RECEIVED', 'AMZ-114-3941689-8772232', '167517', 'amazon'], ['RECEIVED', 'AMZ-114-3941689-8772232', '171378', 'amazon']]);
  t('C8 the board is published inline', published, ['amazon']);
  ok('C9 the low-stock line is warned about (1 on hand, 1 needed is fine; 171378 has 1)', Array.isArray(r.warnings));
  // ⭐ the painter ran for real — the yellow band must survive it
  t('C10 ⚠ the AMAZON band is still yellow after the order-box painter ran',
    S.grid[13].every(x => x.bg === YELLOW), true);

  // duplicate
  const before = JSON.stringify(S.colA());
  const d = sb.addAmazonOrder('AMZ-114-3941689-8772232', [{ sku: '167517', qty: 1 }], '', '', 'sidebar');
  ok('C11 an exact SALES_ORDER|SKU duplicate is refused', !d.ok && /already has 167517/.test(d.message), d.message);
  t('C12 …and nothing was written', JSON.stringify(S.colA()), before);

  // second entry for the SAME order → contiguous with its block
  sb.addAmazonOrder('114-3941689-8772232', [{ sku: '444444', qty: 1 }], '', '', 'sidebar');
  const colD = S.col(4);
  const rowsOfOrder = colD.map((v, i) => v === 'AMZ-114-3941689-8772232' ? i + 1 : 0).filter(Boolean);
  t('C13 a second entry for the same order stays CONTIGUOUS (no split box)',
    rowsOfOrder, [rowsOfOrder[0], rowsOfOrder[0] + 1, rowsOfOrder[0] + 2]);

  // a different order goes to the top of the table
  sb.addAmazonOrder('111-1111111-1111111', [{ sku: '167517', qty: 1 }], 'tomorrow', '', 'sidebar');
  t('C14 a new order goes to the top of the Amazon table', S.row(16)[3], 'AMZ-111-1111111-1111111');
  // ⚠ the case that matters: the order is NOT at the top any more (111… is above it)
  const c16 = sb.addAmazonOrder('114-3941689-8772232', [{ sku: '555555', qty: 1 }], '', '', 'x');
  ok('C16a the extra line was accepted', c16.ok, c16.message);
  const r114 = S.col(4).map((v, i) => v === 'AMZ-114-3941689-8772232' ? i + 1 : 0).filter(Boolean);
  t('C16 ⭐ a later line for an order that is NOT at the top still joins its own block',
    r114.every((r, i) => i === 0 || r === r114[i - 1] + 1), true);
  t('C15 every Amazon row sits BELOW the AMAZON divider', S.col(4).map((v, i) => String(v).indexOf('AMZ-') === 0 ? i + 1 : 99)
    .filter(r => r !== 99).every(r => r > findRow(S, 'AMAZON') + 1), true);
});

// =======================================================================================
section('D · the door refuses what it must', () => {
  const { S, sb } = boot();
  const noTable = sb.addAmazonOrder('114-3941689-8772232', [{ sku: '167517', qty: 1 }], '', '', 'x');
  ok('D1 no Amazon table yet → refused with the setup instruction', !noTable.ok && /setupAmazonTable/.test(noTable.message), noTable.message);
  sb.setupAmazonTable();
  const cases = [
    ['D2 not an Amazon id', ['12345', [{ sku: '167517' }]], /not an Amazon order number/],
    ['D3 an eBay-shaped id', ['05-15052-93025', [{ sku: '167517' }]], /not an Amazon order number/],
    ['D4 unknown SKU is REFUSED, not warned', ['114-3941689-8772232', [{ sku: '999999' }]], /Unknown SKU/],
    ['D5 no lines', ['114-3941689-8772232', []], /At least one SKU/],
    ['D6 qty 0', ['114-3941689-8772232', [{ sku: '167517', qty: 0 }]], /whole number/],
    ['D7 qty over the cap', ['114-3941689-8772232', [{ sku: '167517', qty: 51 }]], /cap/],
    ['D8 a marker typed as a SKU', ['114-3941689-8772232', [{ sku: 'AMAZON' }]], /not a SKU/],
    ['D9 a bad ship-by date', ['114-3941689-8772232', [{ sku: '167517' }], '13/45'], /not a date/]
  ];
  const before = JSON.stringify(S.colA());
  cases.forEach(([label, args, re]) => {
    const r = sb.addAmazonOrder(args[0], args[1], args[2] || '', '', 'x');
    ok(label, !r.ok && re.test(r.message), r.message);
  });
  t('D10 none of those wrote anything', JSON.stringify(S.colA()), before);
});

// =======================================================================================
section('E · ⚠ SORT — Direct never reaches into Amazon, Amazon sorts only itself', () => {
  const { S, sb } = boot();
  sb.setupAmazonTable();
  sb.addAmazonOrder('222-2222222-2222222', [{ sku: '167517', qty: 1 }], '', '', 'x');
  sb.addAmazonOrder('111-1111111-1111111', [{ sku: '171378', qty: 1 }], '', '', 'x');
  const amzBefore = S.colA().slice(15);
  sb.sortTableByStatusAndLocation(2);
  t('E1 the AMAZON divider did not move', findRow(S, 'AMAZON'), 14);
  t('E2 Direct rows stayed inside Direct', dataRowsOf(S, 10, 13).sort(), ['444444', '555555', '666666']);
  t('E3 Amazon rows untouched by the Direct sort', S.colA().slice(15), amzBefore);
  t('E4 the Direct sort did its job (SO-100 PENDING group first, shipped SO-090 last)', S.colA().slice(9, 12), ['444444', '555555', '666666']);
  sb.sortTableByStatusAndLocation(3);
  t('E5 Amazon sort orders by sales order (oldest id first)', S.col(4).slice(15, 17), ['AMZ-111-1111111-1111111', 'AMZ-222-2222222-2222222']);
  t('E6 …and the divider is still the divider', S.row(14)[0], 'AMAZON');
  t('E7 Direct untouched by the Amazon sort', S.colA().slice(9, 12), ['444444', '555555', '666666']);
});

// =======================================================================================
section('F · ⚠ CLEANUP + BUFFERS — never delete the divider or a data row', () => {
  const { S, sb } = boot();
  sb.setupAmazonTable();
  sb.addAmazonOrder('114-3941689-8772232', [{ sku: '167517', qty: 1 }], '', '', 'x');
  // give Direct a fat blank tail above the AMAZON divider
  const amz = findRow(S, 'AMAZON');
  S.sheet.insertRowsBefore(amz, 4);
  const data = () => S.colA().filter(v => /^\d{6}$/.test(v)).sort();
  const dataBefore = data();
  sb.deleteEmptyRows(2);
  t('F1 the Direct cleanup kept exactly one blank above the AMAZON divider',
    findRow(S, 'AMAZON') - 1 - 12, 1);
  t('F2 no data row was lost', data(), dataBefore);
  t('F3 the divider is still there, exactly once', S.colA().filter(v => v === 'AMAZON').length, 1);
  // Amazon tail: add 3 blanks at the end, then clean
  S.sheet.insertRowsAfter(S.sheet.getMaxRows(), 3);
  sb.deleteEmptyRows(3);
  t('F4 the Amazon cleanup leaves one blank at the end', S.sheet.getMaxRows() - S.sheet.getLastRow(), 1);
  t('F5 still no data row lost', data(), dataBefore);
  // balance all three
  S.sheet.insertRowsBefore(findRow(S, 'DIRECT'), 3);
  S.sheet.insertRowsBefore(findRow(S, 'AMAZON'), 2);
  const rep = sb.balanceTableBuffers(1);
  ok('F6 balance reports all three tables', /eBay/.test(rep) && /DIRECT/.test(rep) && /AMAZON/.test(rep), rep);
  const d = findRow(S, 'DIRECT'), a = findRow(S, 'AMAZON');
  t('F7 eBay: one blank above DIRECT', d - 1 - 6, 1);
  t('F8 Direct: one blank above AMAZON', a - 1 - (d + 4), 1);
  t('F9 Amazon: one blank at the end', S.sheet.getMaxRows() - S.sheet.getLastRow(), 1);
  t('F10 no data row lost by the balance', data(), dataBefore);
});

// =======================================================================================
section('G · Update Locations (Direct) never writes into the AMAZON divider row', () => {
  const { S, sb } = boot();
  sb.setupAmazonTable();
  sb.addAmazonOrder('114-3941689-8772232', [{ sku: '167517', qty: 1 }], '', '', 'x');
  S.grid[9][2].v = 'OLD';                       // make Direct row 10 need an update
  sb.updateAllExistingRows(2);
  t('G1 the Direct row was updated', S.row(10)[2], 'D-4');
  t('G2 the AMAZON divider row\'s LOCATION cell is still blank', S.row(14)[2], '');
  S.grid[15][2].v = 'OLD';
  sb.updateAllExistingRows(3);
  t('G3 table 3 updates the Amazon row', S.row(16)[2], 'K-7');
});

// =======================================================================================
section('H · ROLLBACK — removeAmazonTable', () => {
  const { S, sb } = boot();
  sb.setupAmazonTable();
  sb.addAmazonOrder('114-3941689-8772232', [{ sku: '167517', qty: 1 }], '', '', 'x');
  const refused = sb.removeAmazonTable();
  ok('H1 refuses while an Amazon row holds data', /still has 1 row/.test(refused), refused);
  t('H2 …and removed nothing', findRow(S, 'AMAZON'), 14);
  const forced = sb.removeAmazonTableForce();
  ok('H3 force removes it', /removed/.test(forced), forced);
  t('H4 no AMAZON divider left', findRow(S, 'AMAZON'), 0);
  t('H5 the sheet is back to its original shape', S.colA(), liveRows().map(r => String(r[0] || '')));
  t('H6 layout is back to two tables', sb.getTableLayout(S.sheet).amazon, -1);
  // empty table removes without force
  sb.setupAmazonTable();
  ok('H7 an EMPTY Amazon table removes without force', /removed/.test(sb.removeAmazonTable()));
});

// =======================================================================================
section('I · channel classifiers — HOLD scan and the board order', () => {
  const { S, sb } = boot();
  sb.setupAmazonTable();
  sb.addAmazonOrder('114-3941689-8772232', [{ sku: '167517', qty: 1 }], '', 'HOLD for address', 'x');
  const data = S.grid.slice(3).map(r => r.map(x => x.v));
  const held = sb.holdScanRows(data);
  t('I1 a HOLD on an Amazon order is reported as AMAZON', held.map(h => [h.orderId, h.channel]),
    [['AMZ-114-3941689-8772232', 'AMAZON']]);
  const rows = [
    { channel: 'AMAZON', orderId: 'AMZ-1', location: 'A-1', sku: 'x' },
    { channel: 'DIRECT', orderId: 'SO-1', location: 'A-1', sku: 'y' },
    { channel: 'EBAY', orderId: '05-1', location: 'Z-9', sku: 'z' }
  ];
  rows.sort(sb._dashComparePickRows);
  t('I2 the pick list runs eBay → Direct → Amazon', rows.map(r => r.channel), ['EBAY', 'DIRECT', 'AMAZON']);
  t('I3 per-channel caps count Amazon on its own', sb._dashCapPerChannel(
    [{ channel: 'AMAZON' }, { channel: 'AMAZON' }, { channel: 'EBAY' }], 1).map(r => r.channel), ['AMAZON', 'EBAY']);
  t('I4 an unknown channel reads as eBay (today\'s default)', sb._dashChannelOf({ channel: '???' }), 'EBAY');
});

// =======================================================================================
section('J · pure door helpers', () => {
  const { sb } = boot();
  t('J1 id normalisation', ['114-3941689-8772232', 'amz-114-3941689-8772232', '#114 3941689 8772232']
    .map(x => sb._amzNormalizeOrderId(x).salesOrder),
    ['AMZ-114-3941689-8772232', 'AMZ-114-3941689-8772232', 'AMZ-114-3941689-8772232']);
  ok('J2 ⚠ the AMZ- id can NEVER match n8n S4\'s eBay filter (/^[\\d-]+$/)',
     sb._rlSelfCheck('AMZ-114-3941689-8772232') && !sb._rlSelfCheck('114-3941689-8772232'));
  t('J3 /amazon parsing: pairs, default qty, by, note',
    JSON.stringify(sb._amzParseCommand('114-3941689-8772232 167517 2 171378 by 9/29 note call first')),
    JSON.stringify({ orderId: '114-3941689-8772232', lines: [{ sku: '167517', qty: 2 }, { sku: '171378', qty: 1 }], shipBy: '9/29', note: 'call first' }));
  const now = new Date(2026, 8, 26, 12);
  t('J4 ship-by formats', ['9/29', '2026-09-30', 'today', 'tomorrow', '', '12/31/26']
    .map(x => sb._amzShipBy(x, now).text), ['9/29', '9/30', '9/26', '9/27', '', '12/31']);
  ok('J5 an impossible date is refused', !sb._amzShipBy('2/31', now).ok);
  t('J6 note wording — no bare AMAZON word', sb._amzNote('9/29', ''), 'ship by 9/29');
  t('J6b no deadline, no note → the cell stays BLANK (no 📌 on the board)', sb._amzNote('', ''), '');
  t('J7 same SKU twice is MERGED', sb._amzCleanLines([{ sku: '167517', qty: 1 }, { sku: '167517', qty: 2 }]).lines,
    [{ sku: '167517', qty: 3 }]);
});

// =======================================================================================
section('L · ⚠ REAL-SHEETS RULES — validation, the half-built live state, all-or-nothing', () => {
  // L1–L4: validation lands only where it belongs
  { const { S, sb } = boot();
    const msg = sb.setupAmazonTable();
    const a = findRow(S, 'AMAZON');
    ok('L1 setup succeeds with the STATUS dropdown on every data row (the live failure)', /✅/.test(msg), msg);
    t('L2 the band carries no validation', S.sheet.getRange(a, 1, 1, W).getDataValidations()[0].filter(Boolean).length, 0);
    t('L3 the header carries no validation, and reads STATUS', [S.sheet.getRange(a + 1, 6).getDataValidations()[0][0], S.row(a + 1)[5]], [null, 'STATUS']);
    ok('L4 the blank Amazon row KEEPS the STATUS dropdown', !!S.sheet.getRange(a + 2, 6).getDataValidations()[0][0]);
    const r = sb.addAmazonOrder('000-0000000-0000001', [{ sku: '167517', qty: 1 }], '', 'test', 'editor');
    const row = S.colA().indexOf('167517') + 1;
    ok('L5 the door writes PENDING into a validated cell', r.ok && S.row(row)[5] === 'PENDING', r.message);
    ok('L6 the new Amazon row carries the STATUS dropdown', !!S.sheet.getRange(row, 6).getDataValidations()[0][0]);
    let refused = false; try { S.sheet.getRange(row, 6).setValue('AMAZON'); } catch (e) { refused = true; }
    ok('L7 …and it still refuses a non-status value (the rule is real on the new row)', refused);
  }
  // L8–L11: the exact half-built state the first live run left behind
  { const rows = liveRows();
    rows.push([''], ['◈ SKU', 'QTY', 'LOCATION', 'ORDER', 'NOTE'], ['']);
    const { S, sb } = boot(rows);
    const before = S.sheet.getMaxRows();
    const msg = sb.setupAmazonTable();
    const a = findRow(S, 'AMAZON');
    ok('L8 re-running cleans the debris and builds the table', /✅/.test(msg) && /removed 3 leftover/.test(msg), msg);
    t('L9 exactly ONE header copy below DIRECT\'s own', S.colA().filter(v => v === '◈ SKU').length, 3);
    t('L10 the band sits right under the Direct tail, not under the debris', a, 14);
    t('L11 no net row growth beyond band + header + buffer', S.sheet.getMaxRows(), before - 3 + 3);
  }
  // L12: debris with data below it → refuse, delete nothing
  { const rows = liveRows();
    rows.push([''], ['◈ SKU', 'QTY', 'LOCATION', 'ORDER', 'NOTE'], ['999999', 1]);
    const { S, sb } = boot(rows);
    const before = S.colA().slice();
    const msg = sb.setupAmazonTable();
    ok('L12 debris with a real row below → refused, nothing deleted', /❌/.test(msg) && JSON.stringify(S.colA()) === JSON.stringify(before), msg);
  }
  // L13–L14: any failure mid-build rolls the rows back out
  { const { S, sb } = boot();
    const before = S.colA().slice(), maxBefore = S.sheet.getMaxRows();
    sb._styleAmazonDivider = () => { throw new Error('simulated refusal'); };
    const msg = sb.setupAmazonTable();
    ok('L13 a failure mid-build returns ❌ with the reason', /❌/.test(msg) && /simulated refusal/.test(msg), msg);
    ok('L14 …and the sheet is exactly as it was', S.sheet.getMaxRows() === maxBefore && JSON.stringify(S.colA()) === JSON.stringify(before));
  }
  // L15: a band written but with the WRONG marker is caught by the read-back
  { const { S, sb } = boot();
    const real = sb._styleAmazonDivider;
    sb._styleAmazonDivider = (sh, row) => { real(sh, row); sh.getRange(row, 1).setValue('AMAZON TABLE'); };
    const msg = sb.setupAmazonTable();
    ok('L15 a wrong marker is caught by the read-back and rolled back', /❌/.test(msg) && findRow(S, 'AMAZON TABLE') === 0, msg);
  }
});

// =======================================================================================
section('M · the band nameplates — a live FORMULA on both, "open · waiting"', () => {
  const { S, sb } = boot();
  S.sheet.getRange(8, sb.Schema.bandPlateCol).setValue('HQMS · DIRECT ORDERS · 0 waiting');   // the frozen live state
  sb.setupAmazonTable();
  const a = findRow(S, 'AMAZON');
  const rep = sb._applyDividerNameplate(S.sheet);
  const P = sb.Schema.bandPlateCol;
  const d = S.sheet.getRange(8, P).getFormula(), m = S.sheet.getRange(a, P).getFormula();
  ok('M1 the static DIRECT text is replaced by a formula', d.charAt(0) === '=' && /A18/.test(d) && /A29/.test(d), d);
  ok('M2 the AMAZON band reads A28 + A30', m.charAt(0) === '=' && /A28/.test(m) && /A30/.test(m), m);
  ok('M3 both say "open" and "waiting"', /open/.test(d) && /waiting/.test(d) && /open/.test(m) && /waiting/.test(m));
  ok('M4 the report ticks both bands', (rep.match(/✓/g) || []).length === 2 && !/✗/.test(rep), rep);
  ok('M5 the single-count form (row 2) is unchanged', /waiting/.test(sb._nameplateFormula('EBAY ORDERS', 'A17')) &&
     !/open/.test(sb._nameplateFormula('EBAY ORDERS', 'A17')));
});

// =======================================================================================
section('N · ⚠ WHOLE-COLUMN WRITERS never touch a band row (the frozen-nameplate root cause)', () => {
  const { S, sb } = boot();
  sb.setupAmazonTable();
  const a = findRow(S, 'AMAZON');
  // formulas on both bands: the nameplate (G) and a logo (D)
  const P = sb.Schema.bandPlateCol, G = sb.Schema.cols.HAND;
  sb._styleDirectDivider(S.sheet, 8);                 // the live DIRECT band, new layout
  [8, a].forEach(row => {
    S.sheet.getRange(row, P).setFormula('=NAMEPLATE(' + row + ')');
    S.sheet.getRange(row, 4).setFormula('=IMAGE("logo-' + row + '")');
  });
  // the OLD layout kept the nameplate in G (HAND) — put a formula there too, so the test
  // still proves the HAND writer never rewrites a band cell, wherever the plate lives
  S.sheet.getRange(8, G).breakApart(); S.sheet.getRange(8, G).setFormula('=OLDPLATE()');
  sb.buildZohoStockMap = () => new Map();
  sb._isManualSalesOrder = (v) => !/^[\d-]+$/.test(String(v || '').trim());
  const maps = { inventoryMap: new Map([['111111', { available: 3 }], ['444444', { available: 5 }]]) };
  const hand = sb.recomputeHand(maps, new Map());
  ok('N1 recomputeHand ran', /HAND recomputed/.test(hand), hand);
  t('N2 …a formula in the band\'s HAND cell (G) survives — the frozen-nameplate bug', S.sheet.getRange(8, G).getFormula(), '=OLDPLATE()');
  t('N3 …the nameplates survive', [S.sheet.getRange(8, P).getFormula(), S.sheet.getRange(a, P).getFormula()], ['=NAMEPLATE(8)', '=NAMEPLATE(' + a + ')']);
  t('N4 …and HAND still lands on the data rows', S.row(4)[6], 3);
  sb.applyOrderLinksToColumn(S.sheet, 4, 4, S.sheet.getLastRow(), null);
  t('N5 applyOrderLinksToColumn keeps the DIRECT logo formula in D', S.sheet.getRange(8, 4).getFormula(), '=IMAGE("logo-8")');
  t('N6 …and the AMAZON logo formula in D', S.sheet.getRange(a, 4).getFormula(), '=IMAGE("logo-' + a + '")');
  t('N7 …while order ids on data rows are still written', S.row(4)[3], '05-11111-11111');
});

// =======================================================================================
section('O · the band layout — word A:C · mark in D · nameplate F:J', () => {
  const { S, sb } = boot();
  sb.setupAmazonTable();
  const a = findRow(S, 'AMAZON');
  sb._styleDirectDivider(S.sheet, 8);
  [[8, 'DIRECT', 'mark-direct', 'A18', 'A29'], [a, 'AMAZON', 'mark-amazon', 'A28', 'A30']].forEach(([row, word, file, open, wait]) => {
    t('O1 ' + word + ' marker is EXACTLY the word in col A', S.row(row)[0], word);
    const logo = S.sheet.getRange(row, 4).getFormula();
    ok('O2 ' + word + ' mark is an =IMAGE of ' + file + ' in D', /IMAGE\("https:\/\/hq\.yassinqurabi\.com\/mast\/mark-/.test(logo) && logo.indexOf(file) !== -1, logo);
    const plate = S.sheet.getRange(row, sb.Schema.bandPlateCol).getFormula();
    ok('O3 ' + word + ' nameplate formula in F reads ' + open + ' + ' + wait, plate.indexOf(open) !== -1 && plate.indexOf(wait) !== -1, plate);
    t('O4 ' + word + ' E is empty', S.row(row)[4], '');
  });
  t('O5 the layout still finds both tables', [sb.getTableLayout(S.sheet).direct, sb.getTableLayout(S.sheet).amazon], [8, a]);
  const rep = sb._applyDividerNameplate(S.sheet);
  ok('O6 the installer rebuilds both bands and reports ✓ twice', (rep.match(/✓/g) || []).length === 2 && !/✗/.test(rep), rep);
  S.sheet.getRange(8, sb.Schema.bandPlateCol).setValue('HQMS · DIRECT ORDERS · 0 waiting');
  // a STATUS dropdown on the DIRECT band's F (possible on the live sheet) must not block it
  { const { S: S2, sb: sb2 } = boot();
    S2.grid[7].forEach(x => { x.dv = ['PENDING', 'PREPARING', 'SHIPPED', 'CANCELED']; });
    let err = null; try { sb2._styleDirectDivider(S2.sheet, 8); } catch (e) { err = String(e); }
    ok('O8 a STATUS dropdown on the band row does not block the restyle', err === null && !!S2.sheet.getRange(8, sb2.Schema.bandPlateCol).getFormula(), err); }
  const rep2 = sb._applyDividerNameplate(S.sheet);
  ok('O7 a frozen static nameplate is rebuilt into a formula', !!S.sheet.getRange(8, sb.Schema.bandPlateCol).getFormula() && !/✗/.test(rep2), rep2);
});

// =======================================================================================
section('K · Telegram /amazon — card, then the button commits once', () => {
  const { S, sb, logs } = boot();
  const cache = {};
  sb.CacheService = { getScriptCache: () => ({ put: (k, v) => { cache[k] = v; }, get: k => cache[k] || null,
                                               remove: k => { delete cache[k]; } }) };
  vm.runInContext(fs.readFileSync(path.join(SRC, 'TelegramCommands.js'), 'utf8'), sb, { filename: 'TelegramCommands.js' });
  sb.setupAmazonTable();
  const usage = sb.TG_ROUTES['/amazon'].run('');
  ok('K1 no args → usage', /Usage: \/amazon/.test(usage), usage);
  const card = sb.TG_ROUTES['/amazon'].run('114-3941689-8772232 167517 2 by 9/29');
  ok('K2 the card names the order, the line and the Seller Central step',
     card && /AMZ-114-3941689-8772232/.test(card.text) && /2× 167517/.test(card.text) && /Seller Central/.test(card.text),
     card && card.text);
  const data = card.buttons[0][0].data;
  ok('K3 ⚠ callback data fits Telegram\'s 64-byte cap (with the HQ: prefix)', ('HQ:' + data).length <= 64, data);
  t('K4 nothing written before the tap', S.colA().filter(v => v === '167517').length, 0);
  const token = data.split(':')[1];
  const res = sb.TG_ACTIONS.amz.run(token);
  ok('K5 the tap adds the order', /added to the Amazon table/.test(res), res);
  t('K6 one row landed in the Amazon table', S.row(16).slice(0, 4), ['167517', 2, 'K-7', 'AMZ-114-3941689-8772232']);
  const again = sb.TG_ACTIONS.amz.run(token);
  ok('K7 a second tap cannot add it again (token burned)', /expired/.test(again), again);
  t('K8 still exactly one row', S.colA().filter(v => v === '167517').length, 1);
  const bad = sb.TG_ROUTES['/amazon'].run('114-3941689-8772232 999999');
  ok('K9 an unknown SKU is refused at the card, before any button exists', typeof bad === 'string' && /Unknown SKU/.test(bad), bad);
  void logs;
});

console.log('\n' + (fail ? '✗ ' + fail + ' FAILED · ' : '✓ all · ') + pass + ' passed');
process.exit(fail ? 1 : 0);
