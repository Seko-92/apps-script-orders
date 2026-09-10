/**
 * test-photo-queue-guard.js — refreshPhotoQueue must REFUSE, not wipe, when it
 * cannot read Master Inventory. Loads the REAL PhotoQueue.js so the tests cannot
 * drift from what ships.
 *
 * WHY THIS EXISTS. Everything from the NEEDS PHOTOS divider to the bottom of the
 * Prep Queue sheet is machine-owned, and refreshPhotoQueue clears it
 * UNCONDITIONALLY before rewriting. Until 2026-09-10 the scan returned a bare
 * array, and FIVE different failures produced an empty one — MI sheet missing,
 * MI empty, `sku` header renamed, a mid-read throw, and the one legitimate case
 * (backlog clear). So an unreadable MI cleared every ✔ DONE tick and FIRST SEEN
 * date in the table, and that state exists in NO other sheet, backup or log.
 * The hourly housekeeping pass runs this, so the window was one hour.
 *
 * ⚠⚠ SECTION E IS THE ONE THAT KEEPS THE GUARD HONEST. The lazy fix — bail on
 *    `items.length === 0` — would also refuse the case we WANT to act on (the
 *    backlog genuinely reaching zero), leaving stale rows on screen forever.
 *    That is why the scan returns a VERDICT and not a count.
 *
 * ⚠ The mock APPLIES clearContent/setValues to a model region rather than only
 *   recording the calls. "clearContent was called with (12,1,988,7)" proves
 *   arithmetic; what we need to know is whether a picker's ✔ DONE is still there
 *   afterwards.
 *
 * PROVE BY MUTATION (each should turn the run red):
 *   a) delete the `if (!scan.ok) return …` bail        → A, B, C, D fail
 *   b) change the bail to `if (!scan.items.length)`    → E fails (surgical net)
 *   c) make the catch return `{ok: true, …}`           → D fails
 *   d) drop `p.firstSeen ||` from the row build        → F fails (smart-merge)
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');
const CODE = fs.readFileSync(path.join(SRC, 'PhotoQueue.js'), 'utf8');

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label +
    (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};
const section = (name, fn) => {
  console.log('\n' + name);
  try { fn(); } catch (e) { fail++; console.log('  ✗ SECTION THREW (soft): ' + e.message); }
};

const DIVIDER = 10, DATA_START = 12;   // divider row 10 → header 11 → data 12+

/**
 * @param existing  rows already in the NEEDS PHOTOS region (7 wide)
 * @param mi        {headers:[], rows:[[]]} — or null to simulate "sheet missing",
 *                  or the string 'throw' to simulate a mid-read failure.
 */
function build(existing, mi) {
  // The photo region as a model: index 0 === sheet row DATA_START.
  let region = existing.map(r => r.slice());
  const MAXROWS = DATA_START + Math.max(existing.length, 40) - 1;

  const cleared = [];
  const inert = new Proxy({}, { get: (_, k) =>
    k === 'getValues' ? () => [[]] : () => inert });
  function regionRange(row, col, nRows, nCols) {
    return {
      getValues: () => {
        const out = [];
        for (let i = 0; i < nRows; i++) {
          const src = region[row - DATA_START + i] || new Array(7).fill('');
          out.push(src.slice(col - 1, col - 1 + nCols));
        }
        return out;
      },
      setValues: (vals) => {
        for (let i = 0; i < vals.length; i++) {
          const idx = row - DATA_START + i;
          if (!region[idx]) region[idx] = new Array(7).fill('');
          for (let c = 0; c < vals[i].length; c++) region[idx][col - 1 + c] = vals[i][c];
        }
      },
      clearContent: function () {
        cleared.push([row, nRows]);
        for (let i = 0; i < nRows; i++) region[row - DATA_START + i] = new Array(7).fill('');
        return this;
      },
      setBackground: function () { return this; },
      setBorder: function () { return this; },
      removeCheckboxes: function () { return this; },
      setDataValidation: function () { return this; },
      setNumberFormat: function () { return this; },
      setFontWeight: function () { return this; },
      setFontFamily: function () { return this; },
      setFontSize: function () { return this; },
      setFontColor: function () { return this; },
      setHorizontalAlignment: function () { return this; },
      setVerticalAlignment: function () { return this; },
      merge: function () { return this; },
      setNote: function () { return this; }
    };
  }

  const sheet = {
    getName: () => 'Prep Queue',
    getMaxRows: () => MAXROWS,
    getLastRow: () => {
      for (let i = region.length - 1; i >= 0; i--) {
        if (region[i] && String(region[i][0] || '').trim()) return DATA_START + i;
      }
      return DIVIDER + 1;                       // header row, no data
    },
    getLastColumn: () => 7,
    getRange: (row, col, nRows, nCols) => {
      if (row >= DATA_START) return regionRange(row, col, nRows || 1, nCols || 1);
      return inert;   // ⚠ structural rows (band, header) must be a REAL no-op —
                      // an early version handed back a regionRange anchored at
                      // DATA_START, so writing the header row silently overwrote
                      // the first data row of the model and ate a survivor.
    },
    insertRowsAfter: () => {},
    setRowHeight: () => {},
    setColumnWidth: () => {},
    hideColumns: () => {},
    getConditionalFormatRules: () => [],
    setConditionalFormatRules: () => {}
  };

  // ---- Master Inventory stub -------------------------------------------------
  const miSheet = mi === 'throw'
    ? { getLastRow: () => { throw new Error('boom — MI read failed mid-flight'); },
        getLastColumn: () => 7 }
    : (mi ? {
        getLastRow: () => mi.rows.length + 1,
        getLastColumn: () => mi.headers.length,
        getRange: (row, col, nRows, nCols) => ({
          getValues: () => row === 1
            ? [mi.headers]
            : mi.rows.slice(row - 2, row - 2 + nRows).map(r => r.slice(col - 1, col - 1 + nCols))
        })
      } : null);

  const sandbox = {
    console: { log: () => {} },
    SPREADSHEET_ID: 'x',
    DB_SHEET_NAME: 'Master Inventory',
    DB_SKU_HEADER: 'sku', DB_TITLE_HEADER: 'title', DB_LOCATION_HEADER: 'C:Model Year',
    DB_LISTING_STATUS_HEADER: 'listingStatus',
    DB_QUANTITY_HEADER: 'quantity', DB_QUANTITY_SOLD_HEADER: 'quantitySold',
    PREP_QUEUE: {
      sheetName: 'Prep Queue',
      cols: { SKU: 1, QTY: 2, LOCATION: 3, HAND: 4, NOTE: 5, DATE_ADDED: 6, DONE: 7 },
      idx: function (n) { return this.cols[n] - 1; },
      dataWidth: 7, titleRow: 1, headerRow: 2, dataStartRow: 3,
      boundaryMarker: 'INCOMING', bufferRows: 4,
      headers: ['◈ SKU', '# QTY', 'LOCATION', '◫ HAND', 'NOTE', 'DATE ADDED', '✔ DONE']
    },
    SHEET_PULSE: { prepQueue: { stamp: 'I1' } },
    SpreadsheetApp: {
      openById: () => ({ getSheetByName: (n) => n === 'Master Inventory' ? miSheet : sheet }),
      getActive: () => ({ getSheetByName: () => sheet }),
      newDataValidation: () => ({ requireCheckbox: function () { return this; }, build: () => ({}) }),
      flush: () => {}
    },
    Session: { getScriptTimeZone: () => 'America/Chicago' },
    Utilities: { formatDate: () => '9/10/26' },
    compareLocations: (a, b) => String(a).localeCompare(String(b)),
    applySkuLinksToColumn: () => {},
    buildSkuEnrichmentMap: () => ({}),
    stampSheetPulse: () => {},
    _stylePrepBand: () => {},
    _stylePrepHeaderRow: () => {},
    _getPhotoBoundaryRow: () => DIVIDER
  };
  vm.createContext(sandbox);
  vm.runInContext(CODE, sandbox);
  // _ensurePhotoDivider does structural work we don't model — pin the row.
  sandbox._ensurePhotoDivider = () => DIVIDER;

  return {
    run: () => sandbox.refreshPhotoQueue(),
    scan: () => sandbox._scanItemsNeedingPhotos(),
    survivors: () => region.filter(r => r && String(r[0] || '').trim())
                           .map(r => ({ sku: r[0], first: r[5], done: r[6] })),
    cleared
  };
}

/* A picker's in-progress state: two SKUs, one already ticked ✔ DONE. */
const EXISTING = [
  ['165447', 0, 'A-9',  3, 'Piston Kit',   '8/1/26',  true],
  ['172827', 1, 'B-12', 7, 'Gasket Set',   '8/14/26', false]
];
// ⚠ maxImages is 1 — ONE image is the logo placeholder and still needs shooting.
//   A 'has real photos' fixture therefore needs TWO urls, not one.
const MI_HEADERS = ['sku', 'title', 'C:Model Year', 'listingStatus',
                    'quantity', 'quantitySold', 'pictureUrl1', 'pictureUrl2'];

// ============================================================================
section('A · MI SHEET MISSING — the region must survive', () => {
  const h = build(EXISTING, null);
  const msg = h.run();
  t('refuses and says so', /NOT refreshed/.test(msg), true);
  t('names the cause', /sheet not found/.test(msg), true);
  t('nothing was cleared', h.cleared.length, 0);
  t('both rows survive', h.survivors().map(r => r.sku), ['165447', '172827']);
  t('the ✔ DONE tick survives', h.survivors()[0].done, true);
  t('FIRST SEEN survives', h.survivors()[0].first, '8/1/26');
});

section('B · MI PRESENT BUT EMPTY — survive', () => {
  const h = build(EXISTING, { headers: MI_HEADERS, rows: [] });
  const msg = h.run();
  t('refuses', /NOT refreshed/.test(msg), true);
  t('names the cause', /no data rows/.test(msg), true);
  t('nothing was cleared', h.cleared.length, 0);
  t('both rows survive', h.survivors().length, 2);
});

section('C · `sku` HEADER RENAMED — survive', () => {
  const h = build(EXISTING, { headers: ['SKU_ID', 'title'], rows: [['165447', 'x']] });
  const msg = h.run();
  t('refuses', /NOT refreshed/.test(msg), true);
  t('names the column', /'sku' not found/.test(msg), true);
  t('nothing was cleared', h.cleared.length, 0);
  t('both rows survive', h.survivors().length, 2);
});

section('D · MI READ THROWS MID-FLIGHT — survive', () => {
  const h = build(EXISTING, 'throw');
  const msg = h.run();
  t('refuses', /NOT refreshed/.test(msg), true);
  t('names it a read failure', /read failed/.test(msg), true);
  t('nothing was cleared', h.cleared.length, 0);
  t('both rows survive', h.survivors().length, 2);
  t('scan verdict is ok:false', h.scan().ok, false);
});

// ============================================================================
section('E · REGRESSION NET — a genuinely CLEAR backlog still clears the region', () => {
  // Every MI row has a real photo → nothing needs shooting. ok:true, items:[].
  const h = build(EXISTING, {
    headers: MI_HEADERS,
    rows: [['165447', 'Piston Kit', 'A-9',  'Active', 5, 2, 'http://a/1.jpg', 'http://a/2.jpg'],
           ['172827', 'Gasket Set', 'B-12', 'Active', 9, 2, 'http://b/1.jpg', 'http://b/2.jpg']]
  });
  const msg = h.run();
  t('does NOT refuse', /NOT refreshed/.test(msg), false);
  t('reports zero remaining', /0 item/.test(msg), true);
  t('scan verdict is ok:true', h.scan().ok, true);
  t('scan found nothing', h.scan().items.length, 0);
  t('the region WAS cleared', h.cleared.length > 0, true);
  t('no stale rows left behind', h.survivors().length, 0);
});

section('F · NORMAL RUN — writes, and the smart-merge preserves picker state', () => {
  // 165447 still needs a photo (logo only). 172827 got real photos → drops off.
  const h = build(EXISTING, {
    headers: MI_HEADERS,
    rows: [['165447', 'Piston Kit', 'A-9',  'Active', 5, 2, '', ''],
           ['172827', 'Gasket Set', 'B-12', 'Active', 9, 2, 'http://b/1.jpg', 'http://b/2.jpg']]
  });
  const msg = h.run();
  t('does NOT refuse', /NOT refreshed/.test(msg), false);
  t('reports one remaining', /1 item/.test(msg), true);
  const s = h.survivors();
  t('only the photo-less SKU remains', s.map(r => r.sku), ['165447']);
  t('its ✔ DONE tick is preserved', s[0].done, true);
  t('its FIRST SEEN is preserved, not re-stamped', s[0].first, '8/1/26');
});

section('G · a NEW item gets today\'s date and an unticked box', () => {
  const h = build(EXISTING, {
    headers: MI_HEADERS,
    rows: [['165447', 'Piston Kit', 'A-9',  'Active', 5, 2, '', ''],
           ['199999', 'Brand New',  'C-40', 'Active', 4, 0, '', '']]
  });
  h.run();
  const brandNew = h.survivors().filter(r => r.sku === '199999')[0];
  t('new SKU is listed', !!brandNew, true);
  t('stamped with today', brandNew.first, '9/10/26');
  t('not pre-ticked', brandNew.done, false);
});

console.log('\n' + (fail ? '✗' : '✓') + ' ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
