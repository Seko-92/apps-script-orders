// Test Supplies.js against the REAL file, with the Apps Script globals stubbed.
// Loading the real source (not a re-typed copy) is the point — a test against a
// paraphrase proves nothing.
//
//   node test-supplies.js
//   SRC=/tmp/mutated-Supplies.js node test-supplies.js     ← mutation proof
//
// THE ASSERTIONS THAT MATTER MOST are section C: adjustSupplyCount must refuse a
// real 6-digit part SKU and must NEVER reach pushSingleStockAdjustBySku, because
// that function will happily adjust any SKU present in the Zoho mirror.
'use strict';
const fs   = require('fs');
const path = require('path');
const vm   = require('vm');

const SRC = fs.readFileSync(process.env.SRC || path.join(__dirname, '..', 'Supplies.js'), 'utf8');

let pass = 0, fail = 0;
const ok = (name, cond, got) => {
  if (cond) { pass++; console.log('  ok  ' + name); }
  else { fail++; console.log('  FAIL ' + name + (got !== undefined ? '  → got ' + JSON.stringify(got) : '')); }
};
const section = n => console.log('\n' + n);

// ---------------------------------------------------------------- fake sheet --
// Records every write so a test can assert that NOTHING was written, which is the
// only way to prove the unreadable-source bail actually bails.
function makeSheet(rows) {
  const grid = rows.map(r => r.slice());
  const writes = [];
  const ensure = (r, c) => {
    while (grid.length < r) grid.push(new Array(12).fill(''));
    grid.forEach(row => { while (row.length < c) row.push(''); });
  };

  // Every styling setter returns `this` so the production chains work untouched.
  const styling = ['setBackground','setFontColor','setFontFamily','setFontWeight',
    'setFontSize','setFontLine','setFontStyle','setVerticalAlignment','setBorder',
    'setHorizontalAlignment','setNumberFormat','setWrap','setNote','clearNote'];

  function Range(row, col, nR, nC) {
    this.getRow = () => row;
    this.getColumn = () => col;
    this.getNumRows = () => nR;
    this.getValue = () => { ensure(row, col); return grid[row - 1][col - 1]; };
    this.getValues = () => {
      ensure(row + nR - 1, col + nC - 1);
      const out = [];
      for (let i = 0; i < nR; i++) out.push(grid[row - 1 + i].slice(col - 1, col - 1 + nC));
      return out;
    };
    this.setValue = v => {
      ensure(row, col); grid[row - 1][col - 1] = v;
      writes.push({ row, col, value: v }); return this;
    };
    this.setValues = vs => {
      ensure(row + vs.length - 1, col + vs[0].length - 1);
      vs.forEach((r, i) => r.forEach((v, j) => { grid[row - 1 + i][col - 1 + j] = v; }));
      writes.push({ row, col, values: vs }); return this;
    };
    this.clearContent = () => {
      ensure(row, col); grid[row - 1][col - 1] = '';
      writes.push({ row, col, cleared: true }); return this;
    };
    styling.forEach(m => { this[m] = () => this; });
  }

  return {
    _grid: grid,
    _writes: writes,
    getName: () => 'Supplies',
    getRange: (r, c, nR, nC) => new Range(r, c, nR === undefined ? 1 : nR, nC === undefined ? 1 : nC),
    getLastRow: () => {
      for (let i = grid.length - 1; i >= 0; i--) {
        if (grid[i].some(v => v !== '' && v !== null && v !== undefined)) return i + 1;
      }
      return 0;
    },
    getBandings: () => [],
    setFrozenRows: () => {}, setRowHeight: () => {}, setColumnWidth: () => {},
    hideColumns: () => {},
    getConditionalFormatRules: () => [],
    setConditionalFormatRules: () => {},
    insertRowsBefore: () => {}
  };
}

// Build a sheet with the title band + header + the given data rows.
const sheetWith = (dataRows) => makeSheet(
  [new Array(12).fill(''), new Array(12).fill('')].concat(dataRows));

// row helper: [sku, item, kind, location, onHand, reorderAt, status, lastCounted, by, note]
const R = (sku, onHand, reorderAt, status) =>
  [sku, 'name', 'box', 'S-1', onHand === undefined ? '' : onHand,
   reorderAt === undefined ? '' : reorderAt, status || '', '', '', ''];

// Zoho mirror fixture. ⚠ The KEY is derived from the SKU, never typed: the real
// buildZohoStockMap keys on sku.toLowerCase(), and a hand-written key drifts the
// moment the SKU scheme changes — which is exactly what happened on the
// PKG- → SUP- rename, and what these five assertions caught.
const Z = (sku, name, onHand) =>
  [String(sku).toLowerCase(), { skuOriginal: sku, itemName: name, onHand: onHand }];

// ------------------------------------------------------------------ sandbox --
function load(opts) {
  opts = opts || {};
  const pushCalls = [];
  const logged = [];
  const sheet = opts.sheet || sheetWith([]);

  const sandbox = {
    SPREADSHEET_ID: 'x',
    Date, Map, Array, Math, String, Number, parseFloat, isNaN, JSON,
    console: { log: () => {} },
    Logger: { log: () => {} },
    SHEET_PULSE: { supplies: { sheetName: 'Supplies', chip: 'J1', stamp: 'L1', inBand: true } },
    stampSheetPulse: () => {},
    _installPulseChip: () => {},
    Utilities: { formatDate: () => '9/16/26 3:04 PM' },
    SpreadsheetApp: {
      openById: () => ({ getSheetByName: n => (n === 'Supplies' ? sheet : null), insertSheet: () => sheet }),
      getActive: () => ({ getSheetByName: () => sheet, setActiveSheet: () => {} }),
      BandingTheme: { LIGHT_GREY: 'g' },
      BorderStyle: { SOLID_THICK: 't' },
      newConditionalFormatRule: () => {
        const b = {};
        ['whenFormulaSatisfied','setBackground','setFontColor','setBold','setRanges']
          .forEach(m => { b[m] = () => b; });
        b.build = () => ({});
        return b;
      }
    },
    buildZohoStockMap: () => (opts.zoho !== undefined ? opts.zoho : new Map()),
    getCurrentPicker: () => (opts.picker !== undefined ? opts.picker : 'Yassin 1'),
    logActivity: (...a) => logged.push(a),
    pushSingleStockAdjustBySku: (sku, target, o) => {
      pushCalls.push({ sku, target, opts: o });
      return opts.pushResult !== undefined
        ? opts.pushResult
        : { ok: true, before: 10, target, delta: target - 10, adjustmentId: 'adj1', message: 'ok' };
    }
  };
  vm.createContext(sandbox);
  vm.runInContext(SRC, sandbox);
  return { sandbox, pushCalls, logged, sheet };
}

// =============================================================== A · _supStatus
section('A · _supStatus — the verdict, and what it refuses to guess');
{
  const { sandbox } = load();
  const st = sandbox._supStatus;

  ok('0 on hand is OUT',                st(0, 50) === 'OUT', st(0, 50));
  ok('negative is OUT',                 st(-3, 50) === 'OUT', st(-3, 50));
  ok('at the reorder point is LOW',     st(50, 50) === 'LOW', st(50, 50));
  ok('below the reorder point is LOW',  st(12, 50) === 'LOW', st(12, 50));
  ok('above it is OK',                  st(400, 50) === 'OK', st(400, 50));

  // ⚠ THE TWO HONESTY CASES. Both return "" and NEITHER may return "OK" — a row
  // nobody could judge must not read as a row that is fine.
  ok('no reading from Zoho → blank, not OK',   st('', 50) === '', st('', 50));
  ok('null reading → blank, not OK',           st(null, 50) === '', st(null, 50));
  ok('undefined reading → blank',              st(undefined, 50) === '', st(undefined, 50));
  ok('no reorder point → blank, not OK',       st(400, '') === '', st(400, ''));
  ok('no reorder point on a big count → blank', st(9999, null) === '', st(9999, null));

  // ⚠ ...but OUT does not need a reorder point. Zero is zero.
  ok('OUT wins even with no reorder point', st(0, '') === 'OUT', st(0, ''));

  ok('a Date reading (Gotcha #16) → blank', st(new Date(), 50) === '', st(new Date(), 50));
  ok('numeric strings still work',          st('12', '50') === 'LOW', st('12', '50'));
}

// ========================================================== B · isSupplySku
section('B · SUPPLIES.isSupplySku — the hard boundary');
{
  const { sandbox } = load();
  const is = sandbox.SUPPLIES.isSupplySku;
  const P  = sandbox.SUPPLIES.skuPrefix;   // derive, so a prefix change cannot rot the test

  // ONE literal pin, so a silent prefix change is still visible here.
  ok('the prefix is SUP-', P === 'SUP-', P);

  ok('prefixed SKU accepted',       is(P + '104') === true);
  ok('lowercase accepted',          is(P.toLowerCase() + '104') === true, P.toLowerCase() + '104');
  ok('surrounding space tolerated', is('  ' + P + '205 ') === true);

  // ⚠⚠ THE ONE THAT MATTERS. Real part SKUs here are 6 digits, which is exactly
  // why a letter prefix can never collide with one.
  ok('a 6-digit part SKU is REFUSED',   is('167517') === false, is('167517'));
  ok('another part SKU refused',        is('194244') === false);
  ok('a kit SKU refused',               is('217475') === false);
  ok('empty refused',                   is('') === false);
  ok('null refused',                    is(null) === false);
  ok('a SKU merely CONTAINING the prefix refused',
     is('BOX' + P + '1') === false, is('BOX' + P + '1'));
  // Sequential numbering means these are all digits after the prefix — make sure
  // nothing about that shape confuses the guard.
  ok('a bare number is still refused',  is('104') === false, is('104'));
}

// =================================================== C · adjustSupplyCount guards
section('C · adjustSupplyCount — refuses twice before it can reach real stock');
{
  // C1 — a real part SKU must never reach the Zoho write
  {
    const t = load({ sheet: sheetWith([R('167517', 5, 2)]) });   // even if it IS on the sheet
    const res = t.sandbox.adjustSupplyCount('167517', 99);
    ok('C1 a part SKU is refused', res.ok === false, res);
    ok('C1 ...and the Zoho write was NEVER called', t.pushCalls.length === 0, t.pushCalls.length);
    ok('C1 ...and the refusal names the prefix', /SUP-/.test(res.error || ''), res.error);
  }

  // C2 — a SUP- SKU that is not declared on the sheet
  {
    const t = load({ sheet: sheetWith([R('SUP-104', 40, 50)]) });
    const res = t.sandbox.adjustSupplyCount('SUP-201', 200);
    ok('C2 an undeclared supply is refused', res.ok === false, res);
    ok('C2 ...and the Zoho write was NEVER called', t.pushCalls.length === 0, t.pushCalls.length);
  }

  // C3 — the happy path
  {
    const t = load({ sheet: sheetWith([R('SUP-104', 40, 50)]) });
    const res = t.sandbox.adjustSupplyCount('SUP-104', 400);
    ok('C3 a declared supply is accepted', res.ok === true, res);
    ok('C3 ...and the Zoho write ran once', t.pushCalls.length === 1, t.pushCalls.length);
    ok('C3 ...targeting the right SKU', t.pushCalls[0].sku === 'SUP-104', t.pushCalls[0].sku);
    ok('C3 ...with the count as the target', t.pushCalls[0].target === 400, t.pushCalls[0].target);
  }

  // C4 — the supplies-specific delta ceiling. STOCK_ADJUST.maxDelta (50) would
  // refuse an opening count of 0 → 500; force:true would remove the guard entirely.
  {
    const t = load({ sheet: sheetWith([R('SUP-201', 0, 200)]) });
    t.sandbox.adjustSupplyCount('SUP-201', 500);
    const o = t.pushCalls[0].opts;
    ok('C4 passes an explicit maxDelta', o.maxDelta === t.sandbox.SUPPLIES.maxDelta, o.maxDelta);
    ok('C4 ...which is generous enough for a pack', o.maxDelta >= 500, o.maxDelta);
    ok('C4 ...and does NOT use force', !o.force, o.force);
  }

  // C5 — Zoho's `reason` is a MANAGED DROPDOWN: every distinct string becomes a
  // permanent entry humans pick from. One constant, and the picker in description.
  {
    const a = load({ sheet: sheetWith([R('SUP-104', 40, 50)]), picker: 'Hatem 2' });
    a.sandbox.adjustSupplyCount('SUP-104', 100);
    const b = load({ sheet: sheetWith([R('SUP-104', 40, 50)]), picker: 'Yassin 1' });
    b.sandbox.adjustSupplyCount('SUP-104', 100);
    ok('C5 the reason is one constant string',
       a.pushCalls[0].opts.reason === b.pushCalls[0].opts.reason, a.pushCalls[0].opts.reason);
    ok('C5 ...and carries NO picker name',
       !/Hatem|Yassin/.test(a.pushCalls[0].opts.reason), a.pushCalls[0].opts.reason);
    ok('C5 ...while the picker rides separately',
       a.pushCalls[0].opts.picker === 'Hatem 2', a.pushCalls[0].opts.picker);
  }

  // C6 — input validation
  {
    const t = load({ sheet: sheetWith([R('SUP-104', 40, 50)]) });
    ok('C6 a blank count is refused',    t.sandbox.adjustSupplyCount('SUP-104', '').ok === false);
    ok('C6 a negative count is refused', t.sandbox.adjustSupplyCount('SUP-104', -5).ok === false);
    ok('C6 zero IS allowed (a shelf can be empty)',
       t.sandbox.adjustSupplyCount('SUP-104', 0).ok === true);
    ok('C6 an empty SKU is refused',     t.sandbox.adjustSupplyCount('', 10).ok === false);
  }

  // C7 — a failed Zoho write must not be reported as success
  {
    const t = load({ sheet: sheetWith([R('SUP-104', 40, 50)]),
                     pushResult: { ok: false, message: 'refusing to adjust blind' } });
    const res = t.sandbox.adjustSupplyCount('SUP-104', 100);
    ok('C7 a Zoho failure surfaces as a failure', res.ok === false, res);
    ok('C7 ...carrying Zoho\'s own words', /blind/.test(res.error || ''), res.error);
  }
}

// ====================================================== D · refreshSupplies
section('D · refreshSupplies — never blanks a shelf on an unreadable source');
{
  // D1 — empty mirror must bail without touching the sheet
  {
    const sheet = sheetWith([R('SUP-104', 40, 50, 'LOW')]);
    const t = load({ sheet, zoho: new Map() });
    const msg = t.sandbox.refreshSupplies();
    ok('D1 an empty mirror bails', /empty|untouched/i.test(msg), msg);
    ok('D1 ...and wrote NOTHING', sheet._writes.length === 0, sheet._writes.length);
    ok('D1 ...leaving the last known count', sheet._grid[2][4] === 40, sheet._grid[2][4]);
  }

  // D2 — a throwing mirror is the same story
  {
    const sheet = sheetWith([R('SUP-104', 40, 50, 'LOW')]);
    const t = load({ sheet });
    t.sandbox.buildZohoStockMap = () => { throw new Error('sheet gone'); };
    const msg = t.sandbox.refreshSupplies();
    ok('D2 an unreadable mirror bails', /unreadable|untouched/i.test(msg), msg);
    ok('D2 ...and wrote NOTHING', sheet._writes.length === 0, sheet._writes.length);
  }

  // D3 — the normal path fills ON HAND and derives STATUS
  {
    const sheet = sheetWith([R('SUP-104', '', 50), R('SUP-201', '', 200)]);
    const zoho = new Map([
      Z('SUP-104', 'Medium box', 12),
      Z('SUP-201', 'Poly bag', 900)
    ]);
    const t = load({ sheet, zoho });
    const msg = t.sandbox.refreshSupplies();
    ok('D3 ON HAND filled from the mirror', sheet._grid[2][4] === 12, sheet._grid[2][4]);
    ok('D3 under the reorder point reads LOW', sheet._grid[2][6] === 'LOW', sheet._grid[2][6]);
    ok('D3 above it reads OK', sheet._grid[3][6] === 'OK', sheet._grid[3][6]);
    ok('D3 the summary counts what it tracked', /2 tracked/.test(msg), msg);
    ok('D3 ...and names the LOW one', /1 LOW/.test(msg), msg);
  }

  // D4 — a supply that vanished from the mirror keeps its row, blank reading
  {
    const sheet = sheetWith([R('SUP-104', 40, 50, 'LOW')]);
    const zoho = new Map([Z('SUP-902', 'x', 5)]);
    const t = load({ sheet, zoho });
    t.sandbox.refreshSupplies();
    ok('D4 the row survives', sheet._grid[2][0] === 'SUP-104', sheet._grid[2][0]);
    ok('D4 ON HAND goes blank, not 0', sheet._grid[2][4] === '', sheet._grid[2][4]);
    ok('D4 STATUS goes blank, not OK', sheet._grid[2][6] === '', sheet._grid[2][6]);
  }

  // D5 — auto-discovery of SUP- items created in Zoho, and ONLY those
  {
    const sheet = sheetWith([R('SUP-104', '', 50)]);
    const zoho = new Map([
      Z('SUP-104', 'Medium box', 80),
      Z('SUP-301', 'Tape', 30),
      Z('167517', 'Piston Kit', 12)
    ]);
    const t = load({ sheet, zoho });
    const msg = t.sandbox.refreshSupplies();
    const skus = sheet._grid.slice(2).map(r => r[0]).filter(Boolean);
    ok('D5 the new SUP- item was discovered', skus.indexOf('SUP-301') >= 0, skus);
    // ⚠⚠ THE LINE THAT KEEPS THE TWO WORLDS APART — a real part in the mirror must
    // never be pulled onto the supplies registry, because being on the registry is
    // what makes a SKU adjustable from this door.
    ok('D5 a real part SKU was NOT discovered', skus.indexOf('167517') === -1, skus);
    ok('D5 the summary says so', /1 newly discovered/.test(msg), msg);
  }
}

// ================================================= E · getLowSupplies ordering
section('E · getLowSupplies — worst first, and silent about what it cannot judge');
{
  const sheet = sheetWith([
    R('SUP-104',   40,  50, 'LOW'),
    R('SUP-201',  0, 200, 'OUT'),
    R('SUP-301', 10, 100, 'LOW'),
    R('SUP-401',  900, 100, 'OK'),
    R('SUP-501',    12,  '', '')     // no reorder point — unjudgeable
  ]);
  const t = load({ sheet });
  const low = t.sandbox.getLowSupplies();

  ok('E1 only OUT and LOW are listed', low.length === 3, low.length);
  ok('E2 OUT comes first', low[0].status === 'OUT', low[0].status);
  ok('E3 ...and it is the right one', low[0].sku === 'SUP-201', low[0].sku);
  // biggest gap under the reorder point first: TAPE is 90 short, BOX is 10
  ok('E4 the deepest shortfall leads the LOW rows', low[1].sku === 'SUP-301', low[1].sku);
  ok('E5 shortBy is computed', low[1].shortBy === 90, low[1].shortBy);
  // ⚠ a row with no reorder point can never alert — it must not silently pass as fine
  ok('E6 an unjudgeable row is NOT reported as low',
     low.every(r => r.sku !== 'SUP-501'), low.map(r => r.sku));
  ok('E7 an OK row is not reported', low.every(r => r.sku !== 'SUP-401'));
  ok('E8 the count helper agrees', t.sandbox.getSuppliesLowCount() === 3,
     t.sandbox.getSuppliesLowCount());
}

// ============================================================ F · schema shape
section('F · schema — the things other files depend on');
{
  const { sandbox } = load();
  const S = sandbox.SUPPLIES;
  ok('F1 dataWidth matches the header count', S.headers.length === S.dataWidth, S.headers.length);
  ok('F2 every column is inside dataWidth',
     Object.keys(S.cols).every(k => S.cols[k] >= 1 && S.cols[k] <= S.dataWidth));
  ok('F3 idx is 0-based', S.idx('SKU') === 0 && S.idx('NOTE') === S.dataWidth - 1);
  ok('F4 the title band sits above the header', S.titleRow < S.headerRow);
  ok('F5 data starts below the header', S.dataStartRow === S.headerRow + 1);
  ok('F6 the prefix is letter-led (cannot collide with a 6-digit SKU)',
     /^[A-Z]/.test(S.skuPrefix), S.skuPrefix);
}

// ============================ G · the StockAdjust change is a no-op for everyone else
// Supplies added opts.maxDelta to pushSingleStockAdjustBySku, which IS reachable from
// doPost (boardAdjust → boardAdjustStock → here). This section is the evidence for the
// claim that no New Version is needed: for every caller that does NOT pass maxDelta —
// which is all of them today — the ceiling must be byte-identical to before.
section('G · pushSingleStockAdjustBySku — the ±50 board guard is untouched');
{
  const SA = fs.readFileSync(process.env.SA_SRC || path.join(__dirname, '..', 'StockAdjust.js'), 'utf8');
  const seen = [];
  const box = {
    console: { log: () => {} }, Date, Map, String, Number, parseFloat, isNaN,
    buildZohoStockMap: () => new Map([['167517', { itemId: 'i1', onHand: 10 }]]),
    triggerZohoStockAdjust: p => { seen.push(p); return { ok: true, data: { before: 10, delta: 1 } }; }
  };
  vm.createContext(box);
  vm.runInContext(SA, box);

  // The Floor Board's own call shape, copied from boardAdjustStock.
  box.pushSingleStockAdjustBySku('167517', 11, { reason: 'x', picker: 'p' });
  ok('G1 no maxDelta → the ±50 part guard still applies',
     seen[0].max_delta === box.STOCK_ADJUST.maxDelta, seen[0].max_delta);
  ok('G2 ...and that is still 50', box.STOCK_ADJUST.maxDelta === 50, box.STOCK_ADJUST.maxDelta);

  box.pushSingleStockAdjustBySku('167517', 11, { force: true });
  ok('G3 force:true still lifts it to maxQty',
     seen[1].max_delta === box.STOCK_ADJUST.maxQty, seen[1].max_delta);

  box.pushSingleStockAdjustBySku('167517', 11, { maxDelta: 2000 });
  ok('G4 an explicit maxDelta is honoured', seen[2].max_delta === 2000, seen[2].max_delta);
}

// ================================================== H · the seeded shelf survey
// 61 hand-transcribed rows are exactly where a typo hides, and a duplicate SKU would
// only surface as a Zoho creation error partway through a 61-item session.
section('H · SUPPLIES_SEED — the picker\'s survey, checked before it reaches Zoho');
{
  const { sandbox } = load();
  const seed = sandbox.SUPPLIES_SEED;
  const S = sandbox.SUPPLIES;

  ok('H1 all 61 surveyed items are present', seed.length === 61, seed.length);

  const skus = seed.map(r => r[0]);
  const dupes = skus.filter((s, i) => skus.indexOf(s) !== i);
  // ⚠ Zoho enforces unique SKUs, so a duplicate here fails on CREATION — partway
  // through, after some items already exist. Catch it now, not at item 40.
  ok('H2 every SKU is unique', dupes.length === 0, dupes);

  ok('H3 every SKU passes the guard', seed.every(r => S.isSupplySku(r[0])),
     seed.filter(r => !S.isSupplySku(r[0])).map(r => r[0]));

  // The label constraint: fixed width, digits only after the prefix, and NOT four
  // digits (which would collide visually with the picker's own box codes).
  const shape = /^SUP-\d{3}$/;
  ok('H4 every SKU is SUP- plus exactly three digits',
     skus.every(s => shape.test(s)), skus.filter(s => !shape.test(s)));

  ok('H5 every row has a title', seed.every(r => r[1] && r[1].length > 2),
     seed.filter(r => !r[1] || r[1].length <= 2).map(r => r[0]));
  ok('H6 every row has a kind', seed.every(r => r[2] && r[2].length > 0));
  ok('H7 every row carries its opening count',
     seed.every(r => /opening count/.test(r[3])),
     seed.filter(r => !/opening count/.test(r[3])).map(r => r[0]));

  // ⚠ ON HAND must NOT be seeded — it comes from Zoho and only from Zoho.
  ok('H8 the seed row is 4 fields, so ON HAND cannot be written',
     seed.every(r => r.length === 4), seed.find(r => r.length !== 4));

  // Family blocks, so a shelf of labelled containers groups by eye.
  const inBlock = (lo, hi) => skus.filter(s => {
    const n = parseInt(s.slice(4), 10); return n >= lo && n <= hi;
  }).length;
  ok('H9 45 boxes in the 1xx block',      inBlock(101, 145) === 45, inBlock(101, 145));
  ok('H10 4 ring boxes at 151+',          inBlock(151, 154) === 4,  inBlock(151, 154));
  ok('H11 9 shipping bags at 201+',       inBlock(201, 209) === 9,  inBlock(201, 209));
  ok('H12 3 poly bags at 211+',           inBlock(211, 213) === 3,  inBlock(211, 213));

  // The two codes he uses twice must survive as SEPARATE rows — that was the
  // collision that would have broken a bare-code SKU scheme.
  ok('H13 both Hq and Brown 1003 are present',
     seed.some(r => r[1] === 'Box (Hq) 1003') && seed.some(r => r[1] === 'Box (Brown) 1003'));
  ok('H14 both Hq and Brown 1006 are present',
     seed.some(r => r[1] === 'Box (Hq) 1006') && seed.some(r => r[1] === 'Box (Brown) 1006'));

  // The seeder must never write over an existing registry.
  {
    const t = load({ sheet: sheetWith([R('SUP-101', 12, 5)]) });
    const res = t.sandbox.seedSuppliesRegistry();
    ok('H15 seeding a non-empty sheet is REFUSED', /refusing/i.test(res), res);
    ok('H16 ...and nothing was written', t.sheet._writes.length === 0, t.sheet._writes.length);
  }

  // ...and on an empty sheet it lands all 61.
  {
    const t = load({ sheet: sheetWith([]) });
    t.sandbox.seedSuppliesRegistry();
    const written = t.sheet._grid.slice(2).map(r => r[0]).filter(Boolean);
    ok('H17 an empty sheet receives all 61', written.length === 61, written.length);
    ok('H18 ON HAND is left blank for Zoho to fill',
       t.sheet._grid.slice(2, 63).every(r => r[S.idx('ON_HAND')] === ''),
       t.sheet._grid[2][S.idx('ON_HAND')]);
    ok('H19 REORDER AT is left blank for the user',
       t.sheet._grid.slice(2, 63).every(r => r[S.idx('REORDER_AT')] === ''));
  }
}

console.log('\n' + (fail === 0 ? 'OK' : 'FAILED') + '  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail === 0 ? 0 : 1);
