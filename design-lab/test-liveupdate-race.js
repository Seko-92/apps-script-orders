/*
 * test-liveupdate-race.js — the stale-row-coordinates race in liveUpdateTrigger.
 * ---------------------------------------------------------------------------
 * Loads the REAL LiveSync.js + the REAL Schema.js in a VM.
 *
 * ⚠ THE WHOLE TRICK: the mutation is injected INSIDE the fake
 *   buildLocationAndInventoryMaps(). That is the ~2-5s call the trigger makes
 *   AFTER capturing e.range's coordinates, so mutating there reproduces
 *   precisely what n8n does to the sheet during the real window.
 *
 * Two failures, and the SILENT one is the dangerous half:
 *   · rows DELETED  → coordinates overflow → the crash that got emailed
 *   · rows INSERTED → LOCATION/HAND written to a DIFFERENT order's row, no error
 *
 * This is the 2026-05-08 row-shift class. Ruling: never write to a row number
 * captured before a slow operation without re-checking it still means what you
 * think it means.
 *
 *   node test-liveupdate-race.js
 *   SRC=/tmp/head node test-liveupdate-race.js     # prove it bites
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = process.env.SRC || path.join(__dirname, '..');
let pass = 0, fail = 0;
function t(name, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? (pass++, console.log('  ✓ ' + name))
     : (fail++, console.log('  ✗ ' + name + '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
}

/* ---- a sheet that behaves like Sheets: out-of-bounds getRange THROWS ---- */
function FakeSheet(rows) {
  this.rows = rows;                        // array of arrays, 10 wide
  this.name = 'All orders';
  this.writes = [];                        // {row, col, values}
}
FakeSheet.prototype.getName    = function () { return this.name; };
FakeSheet.prototype.getMaxRows = function () { return this.rows.length; };
FakeSheet.prototype.getRange   = function (row, col, nR, nC) {
  nR = nR || 1; nC = nC || 1;
  if (row < 1 || col < 1 || row + nR - 1 > this.rows.length)
    throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
  const self = this;
  return {
    getRow: () => row,
    getColumn: () => col,
    getValues: function () {
      const out = [];
      for (let r = 0; r < nR; r++) {
        const line = [];
        for (let c = 0; c < nC; c++) line.push(self.rows[row - 1 + r][col - 1 + c]);
        out.push(line);
      }
      return out;
    },
    getValue: function () { return self.rows[row - 1][col - 1]; },
    setValues: function (v) {
      self.writes.push({ row: row, col: col, values: JSON.parse(JSON.stringify(v)) });
      for (let r = 0; r < v.length; r++)
        for (let c = 0; c < v[r].length; c++) self.rows[row - 1 + r][col - 1 + c] = v[r][c];
      return this;
    },
    setBackground: function () { return this; },
    setFontColor: function () { return this; },
    setFontWeight: function () { return this; },
    setNumberFormat: function () { return this; }
  };
};

/* A small realistic sheet: 3 banner rows, then eBay rows. */
function makeRows() {
  const blank = () => ['', '', '', '', '', '', '', '', '', ''];
  const rows = [blank(), blank(), blank()];                       // 1-3 banner
  rows.push(['111111', 1, '', '24-AAA', '', 'PENDING', '', '', '', '']);   // 4
  rows.push(['222222', 1, '', '24-BBB', '', 'PENDING', '', '', '', '']);   // 5
  rows.push(['333333', 1, '', '24-CCC', '', 'PENDING', '', '', '', '']);   // 6
  rows.push(['444444', 1, '', '24-DDD', '', 'PENDING', '', '', '', '']);   // 7
  return rows;
}

function makeEnv(opts) {
  opts = opts || {};
  const sheet = new FakeSheet(opts.rows || makeRows());
  const logs = [];
  // ⚠ THE INJECTION POINT — this stands in for the ~2-5s map build, so mutating
  //   the sheet inside it reproduces exactly what n8n does during the real window.
  const slowMapBuilder = function () {
    if (opts.duringSlowWork) opts.duringSlowWork(sheet);
    return {
      locationMap:  new Map([['111111','A-1'],['222222','B-2'],['333333','C-3'],
                             ['444444','D-4'],['999999','Z-9']]),
      inventoryMap: new Map([['111111',{available:11}],['222222',{available:22}],
                             ['333333',{available:33}],['444444',{available:44}],
                             ['999999',{available:99}]])
    };
  };
  const sandbox = {
    console: { log: m => logs.push(String(m)), error: m => logs.push(String(m)) },
    MAIN_SHEET_NAME: 'All orders',
    getLiveUpdateState: () => 'ON',
    getBoundaryRow: () => (opts.boundary === undefined ? -1 : opts.boundary),

    buildZohoStockMap: () => new Map(),
    SpreadsheetApp: { openById: () => { throw new Error('the real map builder must never run in this harness'); } },
    SPREADSHEET_ID: 'test',
    _isManualSalesOrder: () => false,
    resolveHandValue: (mi, zo, preferZoho) => (preferZoho ? (zo == null ? mi : zo) : (mi == null ? zo : mi)),
    updateOrderStatsInSheet: () => {}
  };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'Schema.js'), 'utf8'), sandbox, { filename: 'Schema.js' });
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'LiveSync.js'), 'utf8'), sandbox, { filename: 'LiveSync.js' });

  /* ⚠ AFTER loading, not before. LiveSync.js DEFINES buildLocationAndInventoryMaps,
     so a stub set beforehand is overwritten by the real one — which then reaches for
     SpreadsheetApp and the whole harness reports a phantom failure. First run of this
     test lost eleven assertions to exactly that. Suspect the harness first. */
  sandbox.buildLocationAndInventoryMaps = slowMapBuilder;

  return {
    sheet, logs, sandbox,
    edit: function (row, skus) {
      const range = sheet.getRange(row, 1, skus.length, 1);
      // the event carries the values AS THEY WERE AT EDIT TIME
      const captured = skus.map(s => [s]);
      const ev = { range: Object.assign({}, range, {
        getRow: () => row, getColumn: () => 1, getSheet: () => sheet,
        getValues: () => JSON.parse(JSON.stringify(captured))
      })};
      let threw = null;
      try { sandbox.liveUpdateTrigger(ev); } catch (e) { threw = e.message; }
      return threw;
    },
    locOf: r => sheet.rows[r - 1][2],
    handOf: r => sheet.rows[r - 1][6],
    skuOf:  r => sheet.rows[r - 1][0]
  };
}

console.log('THE NORMAL CASE — must be untouched (regression net)');
{
  const e = makeEnv();
  const threw = e.edit(5, ['222222']);
  t('no throw', threw, null);
  t('LOCATION filled on the edited row', e.locOf(5), 'B-2');
  t('HAND filled on the edited row',     e.handOf(5), 22);
  t('  … and the neighbour above is untouched', e.locOf(4), '');
  t('  … and the neighbour below is untouched', e.locOf(6), '');
}

console.log('\nA MULTI-ROW EDIT still works (regression net)');
{
  const e = makeEnv();
  const threw = e.edit(4, ['111111', '222222']);
  t('no throw', threw, null);
  t('row 4 located', e.locOf(4), 'A-1');
  t('row 5 located', e.locOf(5), 'B-2');
}

console.log('\n⚠ ROWS DELETED DURING THE SLOW WORK — the crash that got emailed');
{
  // n8n's E-chain sweeps shipped rows while the maps are being built.
  const e = makeEnv({
    duringSlowWork: sheet => { sheet.rows.splice(5, 2); }   // drop rows 6 and 7
  });
  const threw = e.edit(7, ['444444']);                       // row 7 no longer exists
  t('does NOT throw the coordinates error', threw, null);
  t('  … and wrote nothing rather than guessing', e.sheet.writes.length, 0);
}

console.log('\n⚠⚠ ROWS INSERTED ABOVE — the SILENT half, and the dangerous one');
{
  // An n8n arrival inserts a new order at the top of the eBay table while the
  // maps build. Everything the picker edited slides DOWN by one.
  const e = makeEnv({
    duringSlowWork: sheet => {
      sheet.rows.splice(3, 0, ['777777', 1, '', '24-NEW', '', 'PENDING', '', '', '', '']);
    }
  });
  const threw = e.edit(5, ['222222']);   // was row 5; after the insert it is row 6
  t('no throw', threw, null);
  // Row 5 now holds the SKU that used to be at row 4 — writing there would be a LIE.
  t('the row that MOVED IN was not stamped', e.locOf(5), '');
  t('  … and nothing was written at all', e.sheet.writes.length, 0);
  t('  … the moved row keeps its own identity', e.skuOf(6), '222222');
}

console.log('\nBOUNDARY + EMPTY SKU still behave (regression nets)');
{
  const rows = makeRows();
  rows[5][0] = 'DIRECT';                       // row 6 is the divider
  const e = makeEnv({ rows: rows, boundary: 6 });
  const threw = e.edit(6, ['DIRECT']);
  t('the divider row does not throw', threw, null);
  t('  … and its LOCATION is preserved, not overwritten', e.locOf(6), '');
}
{
  const e = makeEnv();
  const threw = e.edit(5, ['']);
  t('an emptied SKU writes blanks, not NOT FOUND', [threw, e.locOf(5)], [null, '']);
}

console.log('\n' + (fail ? '❌ ' : '✅ ') + pass + ' passed · ' + fail + ' failed');
process.exit(fail ? 1 : 0);
