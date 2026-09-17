// test-prep-sort-guard.js — the structural guard on _sortCompactPrepSegment
//
// Regression test for the 2026-09-17 INCOMING corruption: the sorter treated a
// photo HEADER row as data, sorted it to row 418 by its literal "LOCATION" text,
// and dragged 466 photo rows into the INCOMING table behind it.
//
//   node design-lab/test-prep-sort-guard.js
//   SRC=/tmp/head node design-lab/test-prep-sort-guard.js     # before/after proof

const fs = require('fs'), path = require('path'), vm = require('vm');
const ROOT = process.env.SRC || path.join(__dirname, '..');

let pass = 0, fail = 0;
const ok = (name, got, want) => {
  const good = JSON.stringify(got) === JSON.stringify(want);
  good ? pass++ : fail++;
  console.log(`${good ? '  ok  ' : '  FAIL'} ${name}` +
              (good ? '' : `\n        got  ${JSON.stringify(got)}\n        want ${JSON.stringify(want)}`));
};

// ---- a sheet model that actually applies the row ops -------------------------------
function makeSheet(rows) {
  const grid = rows.map(r => r.slice());
  const api = {
    _grid: grid,
    getLastRow: () => grid.length,
    getMaxRows: () => grid.length,
    getRange(r, c, nr = 1, nc = 1) {
      return {
        getValues: () => {
          const out = [];
          for (let i = 0; i < nr; i++) {
            const row = grid[r - 1 + i] || [];
            out.push(Array.from({ length: nc }, (_, j) => row[c - 1 + j] ?? ''));
          }
          return out;
        },
        setValues(vals) {
          vals.forEach((row, i) => {
            while (grid.length < r + i) grid.push([]);
            row.forEach((v, j) => { grid[r - 1 + i][c - 1 + j] = v; });
          });
          return this;
        },
        clearContent() {
          for (let i = 0; i < nr; i++)
            for (let j = 0; j < nc; j++)
              if (grid[r - 1 + i]) grid[r - 1 + i][c - 1 + j] = '';
          return this;
        },
        removeCheckboxes: () => api.getRange(r, c, nr, nc),
        setDataValidation: () => api.getRange(r, c, nr, nc)
      };
    },
    deleteRows(start, n) { grid.splice(start - 1, n); },
    insertRowsAfter(after, n) {
      for (let i = 0; i < n; i++) grid.splice(after, 0, []);
    }
  };
  return api;
}

// ---- load the real source ----------------------------------------------------------
const sandbox = {
  console,
  SpreadsheetApp: {
    newDataValidation: () => ({ requireCheckbox: () => ({ build: () => ({}) }) }),
    BorderStyle: { SOLID_THICK: 'SOLID_THICK', SOLID_MEDIUM: 'SOLID_MEDIUM', SOLID: 'SOLID' },
    flush: () => {}
  },
  applySkuLinksToColumn: () => {},
  buildSkuEnrichmentMap: () => new Map(),
  _refreshPrepQueueDuplicates: () => {},
  stampSheetPulse: () => {},
  SHEET_PULSE: { prepQueue: {}, outOfStock: {} },
  _installPulseChip: () => {},
  getSinglePhotoCount: () => 0,
  Utilities: { formatDate: () => '' },
  Session: { getScriptTimeZone: () => 'UTC' }
};
sandbox.globalThis = sandbox;
vm.createContext(sandbox);
for (const f of ['Helpers.js', 'PhotoQueue.js', 'PrepQueue.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sandbox, { filename: f });
}
const { PREP_QUEUE, PREP_PHOTO, _sortCompactPrepSegment } = sandbox;

const item = (sku, loc) => [sku, '', loc, 1, 'note', new Date(2026, 8, 1, 9, 0), false];
const blank = () => ['', '', '', '', '', '', ''];

function run(label, blockRows) {
  // rows 1-2 are chrome; block starts at row 3
  const sheet = makeSheet([['CURRENT'], ['◈ SKU'], ...blockRows]);
  const before = JSON.stringify(sheet._grid);
  try {
    const n = _sortCompactPrepSegment(sheet, 3, blockRows.length, 0);
    return { threw: false, n, sheet, changed: JSON.stringify(sheet._grid) !== before };
  } catch (e) {
    return { threw: true, msg: String(e.message || e), sheet,
             changed: JSON.stringify(sheet._grid) !== before };
  }
}

console.log('\nA · a clean block still sorts (the regression net)');
{
  const r = run('clean', [item('300', 'A-50'), item('100', 'A-9'), item('200', 'B-2')]);
  ok('does not throw', r.threw, false);
  ok('returns the real count', r.n, 3);
  ok('natural aisle order, A-9 before A-50',
     [r.sheet._grid[2][0], r.sheet._grid[3][0], r.sheet._grid[4][0]], ['100', '300', '200']);
}

console.log('\nB · the NEEDS PHOTOS divider inside the block');
{
  const r = run('photo divider',
    [item('100', 'A-9'), [PREP_PHOTO.marker, '', '', '', 'ONLY THE LOGO', '', ''], item('200', 'B-2')]);
  ok('REFUSES', r.threw, true);
  ok('names the divider', /NEEDS PHOTOS divider/.test(r.msg || ''), true);
  ok('names the offending row', /row 4\b/.test(r.msg || ''), true);
  ok('sheet left completely untouched', r.changed, false);
}

console.log('\nC · the photo HEADER row — the exact shape of the incident');
{
  const r = run('photo header',
    [item('100', 'A-9'), PREP_PHOTO.headers.slice(), item('200', 'B-2')]);
  ok('REFUSES', r.threw, true);
  ok('names it a header row', /column-header row/.test(r.msg || ''), true);
  ok('sheet left completely untouched', r.changed, false);
}

console.log('\nD · the INCOMING divider inside the block');
{
  const r = run('incoming divider',
    [item('100', 'A-9'), [PREP_QUEUE.boundaryMarker, '', '', '', '', '', ''], item('200', 'B-2')]);
  ok('REFUSES', r.threw, true);
  ok('names the INCOMING divider', /INCOMING divider/.test(r.msg || ''), true);
  ok('sheet left completely untouched', r.changed, false);
}

console.log('\nE · a real SKU that merely looks odd is NOT structural');
{
  const r = run('odd skus', [item('Timing Belt', 'NOT FOUND'), item('N/A', 'NOT FOUND'), item('111111', 'A-1')]);
  ok('does not throw', r.threw, false);
  ok('keeps all three', r.n, 3);
}

console.log('\nF · refreshPhotoQueue refuses to write above the INCOMING table');
{
  // The photo write covers ~466 rows from divider+2 to the bottom of the sheet.
  // If the divider ever resolved above INCOMING, that lands squarely in the
  // picker's prep list — the shape of damage the 2026-09-17 incident left.
  let wrote = false;
  // Chainable range: every Sheets call returns the range, so the healthy path
  // can run to completion without the stub having to enumerate the API.
  const range = new Proxy({}, {
    get: (_, k) => {
      if (k === 'getValues')          return () => [['', '', '', '', '', '', '']];
      if (k === 'getDataValidation')  return () => null;
      if (k === 'setValues' || k === 'setValue') return () => { wrote = true; return range; };
      return () => range;
    }
  });
  const fakeSheet = {
    getRange: () => range, getLastRow: () => 500, getMaxRows: () => 500,
    setRowHeight: () => {}, insertRowsAfter: () => {}, getParent: () => ({})
  };
  sandbox.SPREADSHEET_ID = 'x';
  sandbox.SpreadsheetApp.openById = () => ({ getSheetByName: () => fakeSheet });
  sandbox._scanItemsNeedingPhotos = () => ({ ok: true, items: [], reason: '' });
  sandbox.compareLocations = sandbox.compareLocations || (() => 0);

  const call = (photoDiv, incDiv) => {
    wrote = false;
    sandbox._ensurePhotoDivider = () => photoDiv;
    sandbox._getPrepBoundaryRow = () => incDiv;
    let msg = '';
    try { msg = String(sandbox.refreshPhotoQueue() || ''); }
    catch (e) { msg = 'THREW: ' + e.message; if(process.env.DBG) console.log('        (throw: ' + e.message + ')'); }
    return { msg, wrote };
  };
  const refused = m => /not below the INCOMING table/.test(m);

  let r = call(10, 16);              // photo divider ABOVE the INCOMING divider
  ok('refuses when the photo divider is above INCOMING', refused(r.msg), true);
  ok('names both rows', /row 10/.test(r.msg) && /row 16/.test(r.msg), true);
  ok('writes nothing', r.wrote, false);

  r = call(17, 16);                  // landing on the INCOMING header row
  ok('refuses when it lands on the INCOMING header', refused(r.msg), true);
  ok('writes nothing', r.wrote, false);

  r = call(42, 16);                  // the healthy layout — must NOT be blocked
  ok('a correct layout is not refused', refused(r.msg), false);
  ok('and it does not throw', /^THREW/.test(r.msg), false);
}

console.log(`\n${fail === 0 ? '✅' : '❌'}  ${pass} passed, ${fail} failed\n`);
process.exit(fail === 0 ? 0 : 1);
