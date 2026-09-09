/**
 * test-mi-schema.js — MI columns resolved by NAME, loaded from the REAL MiSchema.js
 * and the REAL KitRegistry.js so the tests cannot drift from what ships.
 *
 * WHY THIS EXISTS. `_buildMasterInventoryMap` read MI by POSITION:
 *     getRange(2, 2, lastRow-1, 39)  →  data[i][38]   // "col 40 = C:Model Year"
 * MI's column order is not ours to control — the `C:` columns are minted by n8n's
 * autoMapInputData, one per distinct eBay item specific, in first-seen order. Column 40
 * was an accident.
 *
 * ⚠⚠ AND THE FAILURE WAS SILENT. A shifted column feeds a NEIGHBOURING FIELD into the
 *    /^K[-\s]/ aisle test, so every kit quietly classifies as MANUAL — a READY kit gets
 *    expanded into components instead of shipping as its pre-assembled box. No exception,
 *    no wrong-looking number, nothing to notice. That is why section A is the headline:
 *    it does not assert an index, it asserts that REORDERING THE SHEET CHANGES NOTHING.
 *
 * ⚠ The fixture uses the REAL 216-column header row and REAL rows from the 2026-09-09
 *   export, so "column 40" is genuinely where it is in production, not a stand-in.
 *
 * PROVE BY MUTATION (each should turn the run red):
 *   a) restore the positional read in _buildMasterInventoryMap  → A2/A3 fail
 *   b) make resolve() return -1 instead of throwing             → B fails
 *   c) drop the `!lookup.hasOwnProperty(h)` first-wins guard     → D2 fails
 *   d) make readColumns() read the full sheet width             → C2 fails
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');
const FIX = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'mi-header.json'), 'utf8'));

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

/* ---------------------------------------------------------------------------
 * A model MI sheet built from the REAL header row. `order` is a permutation of
 * column indices, so the same logical data can be laid out any way we like.
 * getRange RECORDS its calls, so we can assert the read stayed bounded.
 * ------------------------------------------------------------------------- */
function buildSheet(order) {
  const headers = order.map(i => FIX.headers[i]);
  const dataRows = FIX.rows.map(src => order.map(i => (src[i] === undefined ? '' : src[i])));
  const calls = [];
  const sheet = {
    getSheetName: () => 'Master Inventory',
    getLastColumn: () => headers.length,
    getLastRow: () => dataRows.length + 1,
    getRange: (r, c, nr, nc) => {
      calls.push({ r, c, nr, nc });
      return {
        getValues: () => {
          if (r === 1) return [headers.slice(c - 1, c - 1 + nc)];
          const out = [];
          for (let i = r - 2; i < r - 2 + nr && i < dataRows.length; i++) {
            out.push(dataRows[i].slice(c - 1, c - 1 + nc));
          }
          return out;
        }
      };
    }
  };
  return { sheet, calls, headers };
}

function load(sheet) {
  const sandbox = {
    console: { log: () => {} },
    SPREADSHEET_ID: 'x',
    DB_SHEET_NAME: 'Master Inventory',
    DB_SKU_HEADER: 'sku',
    DB_TITLE_HEADER: 'title',
    DB_QUANTITY_HEADER: 'quantity',
    DB_LOCATION_HEADER: 'C:Model Year',
    DB_QUANTITY_SOLD_HEADER: 'quantitySold',
    DB_LISTING_STATUS_HEADER: 'listingStatus',
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => sheet }) }
  };
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(SRC, 'MiSchema.js'), 'utf8'), sandbox);
  vm.runInContext(fs.readFileSync(path.join(SRC, 'KitRegistry.js'), 'utf8'), sandbox);
  sandbox.MiSchema.clearCache();
  return sandbox;
}

/* The production layout, and two rearrangements of it. */
const N = FIX.headers.length;
const IDENTITY = Array.from({ length: N }, (_, i) => i);
const LOC_I = FIX.headers.indexOf('C:Model Year');

// Move C:Model Year from 40 to the front — the exact shape of a rebuild reordering.
const MOVED = (() => {
  const rest = IDENTITY.filter(i => i !== LOC_I);
  return [rest[0], LOC_I, ...rest.slice(1)];
})();
// Full reverse — nothing lands where it was.
const REVERSED = IDENTITY.slice().reverse();

section('A · REORDERING THE SHEET MUST CHANGE NOTHING  ⚠ the headline', () => {
  const base = load(buildSheet(IDENTITY).sheet)._buildMasterInventoryMap();
  t('A1 baseline map is built from the real header layout', Object.keys(base).length, FIX.rows.length);

  const moved = load(buildSheet(MOVED).sheet)._buildMasterInventoryMap();
  t('A2 C:Model Year moved off column 40 → identical map', moved, base);

  const rev = load(buildSheet(REVERSED).sheet)._buildMasterInventoryMap();
  t('A3 every column reversed → identical map', rev, base);

  t('A4 ...and the aisle really is the aisle, not a neighbour',
    Object.keys(base).map(k => base[k].location).filter(v => /^K-/.test(String(v))).length > 0, true);
});

section('B · A MISSING COLUMN FAILS LOUD, NOT SILENTLY', () => {
  const noLoc = IDENTITY.filter(i => i !== LOC_I);
  const s = load(buildSheet(noLoc).sheet);
  let msg = '(did not throw)';
  try { s._buildMasterInventoryMap(); } catch (e) { msg = e.message; }
  t('B1 dropping C:Model Year throws', msg !== '(did not throw)', true);
  t('B2 ...and the message names the missing column', /C:Model Year/.test(msg), true);
  t('B3 ...and offers the closest headers, so a RENAME is diagnosable',
    /Closest headers|No similar header/.test(msg), true);
});

section('C · THE READ STAYS BOUNDED (no getDataRange on 216 columns)', () => {
  const b = buildSheet(IDENTITY);
  const s = load(b.sheet);
  s._buildMasterInventoryMap();
  const dataReads = b.calls.filter(c => c.r !== 1);
  t('C1 exactly ONE data read', dataReads.length, 1);
  t('C2 ...and it is narrower than the sheet', dataReads[0].nc < N, true);
  t('C3 ...spanning sku(2)..C:Model Year(40) = 39 cols, derived not hardcoded',
    { c: dataReads[0].c, nc: dataReads[0].nc }, { c: 2, nc: 39 });

  // The span must FOLLOW the columns, not stay at 39.
  const b2 = buildSheet(MOVED);
  const s2 = load(b2.sheet);
  s2._buildMasterInventoryMap();
  const d2 = b2.calls.filter(c => c.r !== 1)[0];
  t('C4 ...and it SHRINKS when the columns move closer together', d2.nc < 39, true);
});

section('D · resolver mechanics', () => {
  const s = load(buildSheet(IDENTITY).sheet);
  const M = s.MiSchema, sh = s.SpreadsheetApp.openById().getSheetByName();

  t('D1 case/whitespace tolerant', M.resolve(sh, ['  SKU  '])['  SKU  '], FIX.headers.indexOf('sku'));

  // MI genuinely carries near-duplicate specifics (C:Model vs C:Model2, C:Bobcat vs C:bobcat).
  const dupes = {};
  FIX.headers.forEach(h => { const k = String(h).trim().toLowerCase(); if (k) dupes[k] = (dupes[k] || 0) + 1; });
  const collided = Object.keys(dupes).filter(k => dupes[k] > 1);
  t('D2 a duplicated header resolves to the FIRST occurrence', (() => {
    if (!collided.length) return 'first';
    const name = collided[0];
    const first = FIX.headers.findIndex(h => String(h).trim().toLowerCase() === name);
    return M.resolve(sh, [name])[name] === first ? 'first' : 'later';
  })(), 'first');

  t('D3 optional misses yield -1 instead of throwing',
    M.resolve(sh, [], ['definitelyNotAColumn']).definitelyNotAColumn, -1);

  t('D4 absCol is 1-based for writers',
    M.readColumns(sh, ['sku']).absCol.sku, FIX.headers.indexOf('sku') + 1);
});

section('E · the header row is read ONCE per execution', () => {
  const b = buildSheet(IDENTITY);
  const s = load(b.sheet);
  s._buildMasterInventoryMap();
  s._buildMasterInventoryMap();
  s._buildMasterInventoryMap();
  t('E1 three map builds → one header read', b.calls.filter(c => c.r === 1).length, 1);
  s.MiSchema.clearCache();
  s._buildMasterInventoryMap();
  t('E2 ...and clearCache re-reads it', b.calls.filter(c => c.r === 1).length, 2);
});

section('F · the positional read is really gone from the shipped file', () => {
  const code = fs.readFileSync(path.join(SRC, 'KitRegistry.js'), 'utf8');
  // ⚠ Start at the DOCBLOCK, not the `function` keyword — the tombstone lives in the
  //   comment above it, and slicing from the keyword silently excludes what F3 checks.
  const at = code.indexOf('function _buildMasterInventoryMap');
  const doc = code.lastIndexOf('/**', at);
  const body = code.slice(doc);
  const fn = body.slice(0, body.indexOf('\n}') + 2);
  // ⚠ STRIP COMMENTS FIRST. The tombstone QUOTES the old positional read, so a naive
  //   match finds the documentation and reports the bug as still present. (CLAUDE.md
  //   records this trap; it caught this very test on its second run.)
  const fnCode = fn.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '');
  t('F1 no hardcoded 39-column read', /getRange\(\s*2\s*,\s*2\s*,[^)]*39\s*\)/.test(fnCode), false);
  t('F2 it resolves through MiSchema', /MiSchema\.readColumns/.test(fn), true);
  t('F3 ⚠ the tombstone stays, so nobody re-derives why', /USED TO READ BY POSITION/.test(fn), true);
});

console.log('\n' + (fail === 0 ? '✅' : '❌') +
  ' test-mi-schema: ' + pass + ' passed, ' + fail + ' failed\n');
process.exit(fail === 0 ? 0 : 1);
