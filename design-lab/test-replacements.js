/**
 * test-replacements.js — the PURE CORE of Replacements.js, loaded from the REAL
 * file so the tests cannot drift from what ships.
 *
 * What is worth proving here:
 *   1. ⚠ THE ANCHORED-FILTER CONTRACT. A composed SALES_ORDER must never match
 *      n8n S4's /^[\d-]+$/ filter. If it did, the row rejoins the shipped-check and
 *      gets auto-flipped to SHIPPED behind everyone's back — the exact failure the
 *      prefix exists to prevent. This is the assertion the whole feature rests on.
 *   2. Composition for BOTH kinds, and every refusal path, with its wording.
 *   3. Original-order validation: found on the sheet, found in the log, and MISSED —
 *      because an unvalidated wrong digit is the silent failure this replaces.
 *   4. ⭐ The 2026-08-28 incident, replayed as a fixture. A documented incident beats
 *      an invented one (same reason the archive tests use 24-14979-87359).
 *
 * ⚠ EVERY SECTION FAILS SOFT — the choosePicker lesson: a harness that throws in
 *   section C tells you nothing about D-H.
 *
 * PROVE BY MUTATION (each should turn the run red):
 *   a) REPLACEMENT.kinds.missing.prefix = ""        → A + H fail (self-check catches it)
 *   b) drop the !_rlSelfCheck guard in _rlValidate   → A3 fails
 *   c) _rlFindOriginal: delete the Activity Log stage → G2 fails
 *   d) _rlValidate: accept qty 0                     → D5 fails
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');

const SRC  = process.env.SRC || path.join(__dirname, '..');
const CODE = fs.readFileSync(path.join(SRC, 'Replacements.js'), 'utf8');

// ---- the sheet layer, stubbed -------------------------------------------------
// Only _rlFindOriginal touches it. Everything else is pure.
// SHEET_ROWS entries may be a bare string (col D only) or {sku, so}. _rlFindOriginal
// reads A..D in ONE block, so the stub must hand back 4-wide rows or the duplicate
// check silently sees undefined SKUs and never fires.
let SHEET_ROWS = [];
let LOG_IDS    = [];    // {id, ts} rows in the Activity Log

const asRow = r => (typeof r === 'string' ? { sku: '', so: r } : r);

function fakeRange(vals) {
  return { getValues: () => vals.map(v => [v]) };
}
const SpreadsheetApp = {
  openById: () => ({
    getSheetByName(name) {
      if (name === 'All orders') {
        return {
          getLastRow: () => 3 + SHEET_ROWS.length,
          getRange: (row, col, n) => ({
            // A..D: [SKU, QTY, LOCATION, SALES_ORDER]
            getValues: () => SHEET_ROWS.slice(0, n).map(r => {
              const o = asRow(r);
              return [o.sku, '', '', o.so];
            })
          })
        };
      }
      if (name === 'Activity Log') {
        return {
          getLastRow: () => 1 + LOG_IDS.length,
          getRange: (row, col, n) => col === 3
            ? fakeRange(LOG_IDS.map(r => r.id))
            : fakeRange(LOG_IDS.map(r => r.ts))
        };
      }
      return null;
    }
  })
};

// ---- stock resolvers, stubbed FAITHFULLY -------------------------------------
// ⚠ "A stub cheaper than the real thing tests nothing" (the audio FakeCtx lesson).
//    resolveHandValue is the REAL routing logic copied verbatim, and the lookups
//    return the same shapes the production functions do — including null for an
//    unknown SKU, which is what drives the `knownSku` warning path.
let STOCK = {};   // sku -> { location, mi, zo }

const getSingleLocation  = sku => (STOCK[String(sku).toLowerCase()] || {}).location || "NOT FOUND";
const getSingleInventory = sku => {
  const r = STOCK[String(sku).toLowerCase()];
  return (r && r.mi != null) ? { available: r.mi } : null;
};
const getSingleZohoStock = sku => {
  const r = STOCK[String(sku).toLowerCase()];
  return (r && r.zo != null) ? { available: r.zo, onHand: r.zo } : null;
};
function resolveHandValue(miAvail, zoAvail, preferZoho) {
  if (preferZoho) {
    if (zoAvail != null) return zoAvail;
    if (miAvail != null) return miAvail;
    return 0;
  }
  if (miAvail != null) return miAvail;
  if (zoAvail != null) return zoAvail;
  return 0;
}

const sandbox = {
  console,
  getSingleLocation, getSingleInventory, getSingleZohoStock, resolveHandValue,
  // ⚠ VM realms have their OWN Date constructor, and _rlFindOriginal does
  //   `ts instanceof Date` on the log timestamps. Inject the outer Date or that
  //   check silently fails and the lookback guard never engages.
  //   (The rule this project banked on 2026-08-28 — 14th harness-accuses-code case.)
  Date,
  SpreadsheetApp,
  SPREADSHEET_ID: 'x',
  MAIN_SHEET_NAME: 'All orders',
  Schema: { cols: { SKU: 1, SALES_ORDER: 4 }, dataStartRow: 4 },
  ACTIVITY_LOG: { sheetName: 'Activity Log', cols: { TIMESTAMP: 1, ORDER_ID: 3 }, dataStartRow: 2 }
};
vm.createContext(sandbox);
vm.runInContext(CODE, sandbox, { filename: 'Replacements.js' });
const R = sandbox;

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label +
    (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};
const has = (label, str, needle) => {
  const ok = String(str).indexOf(needle) !== -1;
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label +
    (ok ? '' : '  → ' + JSON.stringify(str) + ' has no ' + JSON.stringify(needle)));
};
const section = (name, fn) => {
  console.log('\n' + name);
  try { fn(); }
  catch (e) { fail++; console.log('  ✗ SECTION THREW (soft): ' + e.message); }
};

// ===============================================================================
section('A · THE ANCHORED-FILTER CONTRACT (n8n S4 must exclude these rows)', () => {
  // A real eBay order id is exactly what S4 accepts, so it must FAIL our self-check.
  t('A1 a clean eBay id is NOT safe to write raw', R._rlSelfCheck('09-15094-35132'), false);
  t('A2 both composed kinds are safe', [
    R._rlSelfCheck('Missing #: 05-15052-93025'),
    R._rlSelfCheck('Replacement #: 19-14597-26309')
  ], [true, true]);

  // A3 is the assertion the whole feature rests on: if the prefix is ever dropped or
  // mangled, _rlValidate must REFUSE rather than write a row that can auto-ship.
  const saved = R.REPLACEMENT.kinds.missing.prefix;
  R.REPLACEMENT.kinds.missing.prefix = '';
  const v = R._rlValidate('missing', '05-15052-93025', '212498', 1, '');
  R.REPLACEMENT.kinds.missing.prefix = saved;
  t('A3 a dropped prefix is REFUSED, not written', v.ok, false);
  has('A3b the refusal names the shipped-check', v.error, 'shipped-check');

  t('A4 blank is not safe', R._rlSelfCheck(''), false);
  t('A5 a digits-only string with no dash is safe (not an eBay id shape)',
    R._rlSelfCheck('12345'), true);
});

// ===============================================================================
section('B · order-id normalization (substring matching)', () => {
  t('B1 strips punctuation and cases up',
    R._rlNormalizeOrderId('05-15052-93025'), '051505293025');
  t('B2 a pasted label still resolves',
    R._rlNormalizeOrderId('Order # 05-15052-93025'), 'ORDER051505293025');
  t('B3 an existing replacement row CONTAINS the original id',
    R._rlNormalizeOrderId('Missing #: 05-15052-93025')
      .indexOf(R._rlNormalizeOrderId('05-15052-93025')) !== -1, true);
  t('B4 null-safe', R._rlNormalizeOrderId(null), '');
});

// ===============================================================================
section('C · composition, both kinds', () => {
  const m = R._rlValidate('missing', '05-15052-93025', '212498', 2, 'ship separately');
  t('C1 missing composes', m.clean.salesOrder, 'Missing #: 05-15052-93025');
  t('C2 label + qty + sku carried', [m.clean.label, m.clean.qty, m.clean.sku],
    ['MISSING', 2, '212498']);

  const r = R._rlValidate('replacement', '19-14597-26309', '171378', 1,
                          'ship the gaskets, studs and nuts only');
  t('C3 replacement composes', r.clean.salesOrder, 'Replacement #: 19-14597-26309');

  t('C4 kind is case-insensitive', R._rlValidate('MISSING', '05-15052-93025', '1', 1, '').ok, true);
  t('C5 qty defaults to 1', R._rlValidate('missing', '05-15052-93025', '1', '', '').clean.qty, 1);
  t('C6 qty accepts a padded string',
    R._rlValidate('missing', '05-15052-93025', '1', ' 3 ', '').clean.qty, 3);
});

// ===============================================================================
section('D · refusals — and the wording, because a refusal has to teach', () => {
  const bad = (args) => R._rlValidate.apply(null, args);

  t('D1 unknown kind', bad(['oops', '05-15052-93025', '1', 1, '']).ok, false);
  t('D2 missing original order', bad(['missing', '', '1', 1, '']).ok, false);
  has('D2b names what it wants', bad(['missing', '', '1', 1, '']).error, 'Original order id');
  t('D3 too-short original', bad(['missing', '123', '1', 1, '']).ok, false);
  t('D4 missing sku', bad(['missing', '05-15052-93025', '', 1, '']).ok, false);
  t('D5 qty 0 refused', bad(['missing', '05-15052-93025', '1', 0, '']).ok, false);
  t('D6 negative qty refused', bad(['missing', '05-15052-93025', '1', -2, '']).ok, false);
  t('D7 fractional qty refused', bad(['missing', '05-15052-93025', '1', 1.5, '']).ok, false);
  t('D8 over the cap refused', bad(['missing', '05-15052-93025', '1', 500, '']).ok, false);
  t('D9 non-numeric qty refused', bad(['missing', '05-15052-93025', '1', 'two', '']).ok, false);

  // A pasted already-composed value must not be silently re-prefixed into
  // "Missing #: Missing #: 05-…" — that lands on the sheet and reads as a bug.
  const dbl = bad(['missing', 'Missing #: 05-15052-93025', '1', 1, '']);
  t('D10 an already-prefixed value is refused', dbl.ok, false);
  has('D10b it says to pass the original', dbl.error, 'ORIGINAL order id');
});

// ===============================================================================
section('E · note handling', () => {
  const long = 'x'.repeat(400);
  const v = R._rlValidate('missing', '05-15052-93025', '1', 1, long);
  t('E1 truncated to the cap', v.clean.note.length, R.REPLACEMENT.maxNoteChars);
  t('E2 ends with an ellipsis so truncation is visible',
    v.clean.note.slice(-1), '…');
  t('E3 whitespace-only note becomes empty',
    R._rlValidate('missing', '05-15052-93025', '1', 1, '   ').clean.note, '');
});

// ===============================================================================
section('F · Activity Log DETAIL wording', () => {
  const c = R._rlValidate('missing', '05-15052-93025', '212498', 1, '').clean;
  t('F1 names kind + original', R._rlDetail(c, ''), 'MISSING line for 05-15052-93025');
  t('F2 carries where the original was found',
    R._rlDetail(c, 'in the Activity Log'),
    'MISSING line for 05-15052-93025 (original in the Activity Log)');
});

// ===============================================================================
section('G · original-order validation — the silent failure this replaces', () => {
  const now = Date.now();

  SHEET_ROWS = ['09-15094-35132', '02-14623-46718'];
  LOG_IDS  = [];
  t('G1 found on the sheet', R._rlFindOriginal('09-15094-35132').where, 'on the sheet');

  // The order shipped yesterday and n8n swept the row at ~1 AM. The log is the only
  // place that still remembers it — and a "missing item" report is EXACTLY the case
  // that arrives after the order is gone from the sheet.
  SHEET_ROWS = ['02-14623-46718'];
  LOG_IDS  = [{ id: '05-15052-93025', ts: new Date(now - 2 * 86400000) }];
  t('G2 found in the Activity Log after the row was swept',
    R._rlFindOriginal('05-15052-93025').where, 'in the Activity Log');

  SHEET_ROWS = ['02-14623-46718'];
  LOG_IDS  = [];
  t('G3 a wrong digit is CAUGHT, not silently accepted',
    R._rlFindOriginal('05-15052-93026').found, false);

  // An existing replacement row proves the original is one of ours.
  SHEET_ROWS = ['Missing #: 05-15052-93025'];
  LOG_IDS  = [];
  t('G4 an existing replacement row counts as proof',
    R._rlFindOriginal('05-15052-93025').found, true);

  // Beyond the lookback the walk stops — a 2-year-old order is not a live case.
  SHEET_ROWS = [];
  LOG_IDS  = [{ id: '05-15052-93025', ts: new Date(now - 400 * 86400000) }];
  t('G5 outside the lookback window it is not found',
    R._rlFindOriginal('05-15052-93025').found, false);

  SHEET_ROWS = []; LOG_IDS = [];
  t('G6 empty everything is a clean miss, not a throw',
    R._rlFindOriginal('05-15052-93025').found, false);
});

// ===============================================================================
section('H · ⭐ THE 2026-08-28 INCIDENT, REPLAYED', () => {
  // What happened: row 09-15094-35132 / SKU 212498 was picked at 12:15:30 and
  // shelf-counted at 12:15:48. At 12:21:40 its col D was overwritten by hand with
  // "Missing #: 05-15052-93025" — destroying the pick and the count, and breaking
  // the dedupe signature so n8n re-inserted the line at 12:25:30.
  //
  // What SHOULD have happened is one call to this door.
  SHEET_ROWS = ['09-15094-35132'];
  LOG_IDS  = [{ id: '05-15052-93025', ts: new Date(Date.now() - 86400000) }];

  const v = R._rlValidate('missing', '05-15052-93025', '212498', 1, '');
  t('H1 the intended line composes cleanly', v.clean.salesOrder, 'Missing #: 05-15052-93025');
  t('H2 the original order validates', R._rlFindOriginal('05-15052-93025').found, true);
  t('H3 the composed row is excluded from the shipped-check',
    R._rlSelfCheck(v.clean.salesOrder), true);

  // ⭐ The property that makes the incident impossible: the value that was DESTROYED
  //    is not an input to this path at all. There is no parameter that can address an
  //    existing row, so 09-15094-35132 cannot be reached, let alone overwritten.
  t('H4 the overwritten order is not addressable by any argument',
    Object.keys(v.clean).some(k => String(v.clean[k]) === '09-15094-35132'), false);
  t('H5 the victim row still reads as a clean eBay id (untouched by construction)',
    R._rlSelfCheck(SHEET_ROWS[0]), false);
});

// ===============================================================================
section('I · Telegram arg parsing — qty is detected by SHAPE, not position', () => {
  const P = R._rlParseCommandArgs;

  t('I1 full form', P('05-15052-93025 212498 2 ship separately'),
    { originalOrder: '05-15052-93025', sku: '212498', qty: 2, note: 'ship separately' });

  // ⭐ The case that makes shape-detection worth it: without it, "ship" would be
  //    swallowed as a quantity and the note would silently lose its first word.
  t('I2 a non-numeric 3rd token is the NOTE, not a qty',
    P('05-15052-93025 212498 ship urgently'),
    { originalOrder: '05-15052-93025', sku: '212498', qty: '', note: 'ship urgently' });

  t('I3 order + sku only', P('05-15052-93025 212498'),
    { originalOrder: '05-15052-93025', sku: '212498', qty: '', note: '' });
  t('I4 order only', P('05-15052-93025').sku, '');
  t('I5 empty', P('').originalOrder, '');
  t('I6 collapses runs of whitespace', P('  05-15052-93025   212498   3   two  words  '),
    { originalOrder: '05-15052-93025', sku: '212498', qty: 3, note: 'two words' });
  t('I7 a decimal 3rd token is a note, not a qty (qty must be a bare integer)',
    P('05-15052-93025 212498 1.5 units').qty, '');

  // End-to-end: parse then validate, the way the route will.
  const a = P('05-15052-93025 212498 2 ship separately');
  const v = R._rlValidate('missing', a.originalOrder, a.sku, a.qty, a.note);
  t('I8 parse → validate round-trips', [v.ok, v.clean.salesOrder, v.clean.qty, v.clean.note],
    [true, 'Missing #: 05-15052-93025', 2, 'ship separately']);
});

// ===============================================================================
section('J · the duplicate guard — double-tap must be safe BY CONSTRUCTION', () => {
  const now = Date.now();
  const originalOnSheet = { sku: '999999', so: '05-15052-93025' };

  // Baseline: nothing to collide with, so the line is allowed.
  SHEET_ROWS = [originalOnSheet];
  LOG_IDS = [];
  t('J1 first add is allowed',
    R.previewReplacementLine('missing', '05-15052-93025', '212498', 1, '').ok, true);

  // The line now exists. A second identical add — the double-tap — must refuse.
  SHEET_ROWS = [originalOnSheet, { sku: '212498', so: 'Missing #: 05-15052-93025' }];
  const dup = R.previewReplacementLine('missing', '05-15052-93025', '212498', 1, '');
  t('J2 an exact repeat is REFUSED', dup.ok, false);
  has('J2b the refusal names the existing row', dup.error, 'row 5');
  has('J2c and says what to do instead', dup.error, 'raise the qty');

  // ⚠ The guard must be NARROW. A different SKU on the same original order is a
  //   legitimate second missing item, not a duplicate.
  t('J3 a DIFFERENT sku on the same order is still allowed',
    R.previewReplacementLine('missing', '05-15052-93025', '171378', 1, '').ok, true);

  // ...and the same SKU under the other KIND is a different line entirely
  // (an item was missing AND a replacement was sent — both are real).
  t('J4 the same sku under the other kind is allowed',
    R.previewReplacementLine('replacement', '05-15052-93025', '212498', 1, '').ok, true);

  // The original-order lookup must still work when the only trace is the log.
  SHEET_ROWS = [{ sku: '212498', so: 'Missing #: 05-15052-93025' }];
  LOG_IDS = [{ id: '05-15052-93025', ts: new Date(now - 86400000) }];
  const dupLog = R.previewReplacementLine('missing', '05-15052-93025', '212498', 1, '');
  t('J5 the guard fires even when the original is only in the log', dupLog.ok, false);
});

// ===============================================================================
section('K · stock resolution — ZOHO-FIRST, because a composed SO is a manual row', () => {
  const now = Date.now();
  SHEET_ROWS = [{ sku: '999999', so: '05-15052-93025' }];
  LOG_IDS = [];

  // ⭐ THE ROUTING CLAIM the file header makes. A composed SO is not a clean eBay
  //    order id → _isManualSalesOrder is true → Zoho wins over MI. If this ever
  //    flipped, the new row would show one number and the next recomputeHand pass
  //    would silently overwrite it with another.
  STOCK = { '212498': { location: 'E-57', mi: 3, zo: 7 } };
  const p1 = R.previewReplacementLine('missing', '05-15052-93025', '212498', 1, '');
  t('K1 Zoho (7) wins over MI (3)', p1.stock.hand, 7);
  t('K2 location comes from MI', p1.stock.location, 'E-57');
  t('K3 no warnings when stock is healthy', p1.warnings, []);

  STOCK = { '212498': { location: 'E-57', mi: 3, zo: null } };
  t('K4 falls back to MI when Zoho does not know the sku',
    R.previewReplacementLine('missing', '05-15052-93025', '212498', 1, '').stock.hand, 3);

  // An unknown SKU must be STATED, not silently rendered as a confident 0.
  STOCK = {};
  const p2 = R.previewReplacementLine('missing', '05-15052-93025', '212498', 1, '');
  t('K5 an unknown sku still ALLOWS the line', p2.ok, true);
  t('K6 ...and warns rather than pretending', p2.warnings.length, 1);
  has('K6b the warning names both sources', p2.warnings[0], 'neither Zoho nor Master Inventory');

  // Short stock is a fact to weigh, never a block — the person at the shelf decides.
  STOCK = { '212498': { location: 'E-57', mi: null, zo: 1 } };
  const p3 = R.previewReplacementLine('missing', '05-15052-93025', '212498', 5, '');
  t('K7 short stock warns but does NOT block', p3.ok, true);
  has('K7b and states the gap', p3.warnings.join(' | '), 'Only 1 on hand against a qty of 5');

  STOCK = { '212498': { location: 'NOT FOUND', mi: 4, zo: 4 } };
  has('K8 a missing shelf is called out',
    R.previewReplacementLine('missing', '05-15052-93025', '212498', 1, '').warnings.join(' | '),
    'No shelf location');
});

console.log('\n' + (fail === 0 ? '✅' : '❌') +
  ' test-replacements: ' + pass + ' passed, ' + fail + ' failed\n');
process.exit(fail === 0 ? 0 : 1);
