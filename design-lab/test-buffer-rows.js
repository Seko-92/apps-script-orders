/**
 * test-buffer-rows.js — the trailing blank-row buffer, loaded from the REAL
 * RowManagement.js so the tests cannot drift from what ships.
 *
 * WHY THIS EXISTS. deleteEmptyRows calls sheet.deleteRows() on the sheet the whole
 * warehouse works from, and it had no coverage at all. Until 2026-09-09 the buffer size
 * lived in THREE places with three different answers — TABLE_BUFFER_ROWS, a literal
 * `last + 4` for eBay (keep 3) and MAX_EMPTY_ROWS_TO_KEEP for DIRECT (keep 5) — which is
 * exactly what the live sheet showed, and it meant the two Cleanup buttons SILENTLY UNDID
 * any change to the constant.
 *
 * ⚠⚠ THE ONE THAT MATTERS IS B: `keep >= 1` is a SAFETY property, not a preference.
 *    delStart = last + keep + 1, so a keep of 0 puts the delete ON the last data row.
 *
 * ⚠ The mock APPLIES the row operations to a model sheet rather than only recording them —
 *   asserting "deleteRows was called with (29, 5)" proves arithmetic, not survival. What we
 *   need to know is whether any row that held data is still there afterwards.
 *
 * PROVE BY MUTATION (each should turn the run red):
 *   a) put `last + 4` back for t===1                    → A fails
 *   b) drop the Math.max(1, …) clamp and set the const 0 → B fails
 *   c) make balanceTableBuffers use a literal 3          → C fails
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');
const CODE = fs.readFileSync(path.join(SRC, 'RowManagement.js'), 'utf8');

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

/* A model sheet: colA[i] is row i+1's column A. Row ops are APPLIED, not just recorded. */
function build(colA) {
  const rows = colA.slice();
  const noop = { copyTo: () => {}, setValues: () => {}, getValues: () => [[]],
                 setBorder: () => {}, setBackground: () => {}, getA1Notation: () => '' };
  const sheet = {
    getMaxRows: () => rows.length,
    getLastRow: () => { for (let i = rows.length - 1; i >= 0; i--) if (rows[i]) return i + 1; return 0; },
    deleteRows: (start, n) => rows.splice(start - 1, n),
    insertRowsBefore: (row, n) => rows.splice(row - 1, 0, ...Array(n).fill('')),
    insertRowsAfter:  (row, n) => rows.splice(row, 0, ...Array(n).fill('')),
    getRange: () => noop,
    setRowHeights: () => {}
  };
  const sandbox = {
    console: { log: () => {} },
    SPREADSHEET_ID: 'x', MAIN_SHEET_NAME: 'All orders',
    Schema: { dataStartRow: 4, headerRow: 3, dataWidth: 10, cols: { SKU: 1 } },
    SpreadsheetApp: {
      openById: () => ({ getSheetByName: () => sheet }),
      CopyPasteType: { PASTE_FORMAT: 'F' },
      flush: () => {}
    },
    // The two Helpers.js globals RowManagement leans on, modelled off column A.
    getBoundaryRow: () => { const i = rows.findIndex(v => String(v).trim().toUpperCase() === 'DIRECT'); return i === -1 ? -1 : i + 1; },
    findLastDataRowInSegment: (start, end) => {
      for (let r = Math.min(end, rows.length); r >= start; r--) if (rows[r - 1]) return r;
      return start - 1;
    },
    // ⚠ Config.js's global, injected so the suite can RUN against HEAD. Without it
    //   sections A and B threw instead of reporting, and a before/after proof is only
    //   useful if every section reports (the choosePicker lesson). The current code does
    //   not read it; HEAD does, and that is exactly what A3 is there to catch.
    MAX_EMPTY_ROWS_TO_KEEP: 5,
    verifyAndRestoreHeaders: () => {},
    setupDuplicateSalesOrderHighlighting: () => {},
    _obIsOwner: () => true
  };
  vm.createContext(sandbox);
  vm.runInContext(CODE, sandbox, { filename: 'RowManagement.js' });
  return { B: sandbox, rows };
}

/* The live shape from the 2026-09-09 screenshot: eBay 4..22, 3 blanks, DIRECT band at 26,
   header 27, one DIRECT row at 28, 5 blanks, maxRows 33. */
const LIVE = () => {
  const r = Array(33).fill('');
  r[0] = 'HQ'; r[2] = '◈ SKU';
  for (let i = 4; i <= 22; i++) r[i - 1] = 'SKU' + i;
  r[25] = 'DIRECT'; r[26] = '◈ SKU'; r[27] = '164664';
  return r;
};

// ===============================================================================
section('A · ⭐ BOTH tables trim to the SAME constant', () => {
  const { B, rows } = build(LIVE());
  t('A0 the constant is the one the sheet should carry', B.TABLE_BUFFER_ROWS, 1);

  B.deleteEmptyRows(2);                       // DIRECT first: it sits at the bottom
  t('A1 DIRECT keeps exactly TABLE_BUFFER_ROWS blanks',
    rows.length - 28, B.TABLE_BUFFER_ROWS);

  B.deleteEmptyRows(1);
  // ⚠ Row numbers are stale the moment a trim runs — the eBay cut shifts everything below
  //   it up. Re-locate by CONTENT, the same reason the 05-08 row-shift class exists.
  const b = rows.findIndex(v => v === 'DIRECT') + 1;
  const eLast = rows.lastIndexOf('SKU22') + 1;
  const dLast = rows.lastIndexOf('164664') + 1;
  t('A2 eBay keeps exactly TABLE_BUFFER_ROWS blanks',
    (b - 1) - eLast, B.TABLE_BUFFER_ROWS);

  // ⚠⚠ THE BUG: 3 for eBay, 5 for DIRECT, from two hardcoded literals.
  t('A3 ⭐ and the two tables AGREE — they did not before',
    (b - 1) - eLast, rows.length - dLast);
});

// ===============================================================================
section('B · ⚠⚠ it can NEVER delete a row that holds data', () => {
  [1, 0, -4, null].forEach(function (n) {
    const { B, rows } = build(LIVE());
    B.TABLE_BUFFER_ROWS = n;                  // including values the clamp must absorb
    B.deleteEmptyRows(2);
    B.deleteEmptyRows(1);
    const kept = rows.filter(v => v && v !== 'DIRECT' && v !== '◈ SKU' && v !== 'HQ');
    t('B1 with TABLE_BUFFER_ROWS=' + JSON.stringify(n) + ' every data row survives',
      kept.length, 20);                       // 19 eBay + 1 DIRECT
    t('B2 ...and the DIRECT marker is still there (getBoundaryRow depends on it)',
      rows.filter(v => v === 'DIRECT').length, 1);
    t('B3 ...and at least one blank row remains below DIRECT',
      rows.length - rows.lastIndexOf('164664') - 1 >= 1, true);
  });
});

// ===============================================================================
section('C · balanceTableBuffers reads the SAME constant', () => {
  const { B, rows } = build(LIVE());
  B.balanceTableBuffers();
  const b = rows.findIndex(v => v === 'DIRECT') + 1;
  t('C1 eBay tail', (b - 1) - 22, B.TABLE_BUFFER_ROWS);
  t('C2 DIRECT tail', rows.length - rows.lastIndexOf('164664') - 1, B.TABLE_BUFFER_ROWS);

  // It must GROW as readily as it trims, or a sheet that ran dry stays dry.
  const wide = build(LIVE());
  wide.B.balanceTableBuffers(4);
  const wb = wide.rows.findIndex(v => v === 'DIRECT') + 1;
  t('C3 it grows too, on an explicit target', (wb - 1) - 22, 4);

  // ⚠ Same clamp, same reason.
  const zero = build(LIVE());
  zero.B.balanceTableBuffers(0);
  t('C4 ⚠⚠ 0 is clamped to 1 here as well',
    zero.rows.length - zero.rows.lastIndexOf('164664') - 1, 1);
});

// ===============================================================================
section('D · the dead buffer code is gone', () => {
  const { B } = build(LIVE());
  t('D1 ensureDirectTableBuffer no longer exists', typeof B.ensureDirectTableBuffer, 'undefined');
  t('D2 ...and nothing reads MAX_EMPTY_ROWS_TO_KEEP', /MAX_EMPTY_ROWS_TO_KEEP/.test(
    CODE.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '')), false);
  t('D3 ⚠ the tombstone stays, so the next reader does not re-derive why',
    /REMOVED 2026-09-09 — ensureDirectTableBuffer/.test(CODE), true);
});

console.log('\n' + (fail === 0 ? '✅' : '❌') +
  ' test-buffer-rows: ' + pass + ' passed, ' + fail + ' failed\n');
process.exit(fail === 0 ? 0 : 1);
