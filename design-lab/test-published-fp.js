// ============================================================================
// THE CACHE-INVALIDATION NET — does the fingerprint actually notice?
//
// The dirty flag depends on every write path remembering to call
// _dashBustTickCache(). That enumeration was missed FIVE TIMES IN TWO DAYS, so
// the publisher now checks for itself: on the run where it would otherwise
// skip, it fingerprints the All Orders data range and republishes if that moved
// without anyone announcing it.
//
// This tests _pubFingerprintOf — the pure fold that does the actual deciding —
// loaded from the REAL Published.js, never a re-typed copy. Every case below is
// a change the flag has ALREADY failed to report at least once:
//
//   · a row inserted      → kit expansion, the Zoho Pull insert  (08-14)
//   · a STATUS flipped    → the Zoho Pull cancel branch          (08-14)
//   · a NOTE edited       → a human typing on the sheet          (08-14)
//   · a row deleted       → a human deleting rows                (08-14)
//
// ⚠ The row-count-only net I first proposed would have caught the first and
// last of those and MISSED the middle two. That is why this folds content.
//
// Usage: node test-published-fp.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC = path.join(__dirname, '..', 'Published.js');

// Load the REAL file with just enough of Apps Script stubbed that it parses.
// Only the pure fold is exercised — nothing here touches a stub.
const sandbox = {
  SpreadsheetApp: {}, PropertiesService: {}, Utilities: {}, CacheService: {},
  Schema: { dataStartRow: 4, dataWidth: 10 },
  SPREADSHEET_ID: 'x', MAIN_SHEET_NAME: 'All orders',
  console: { log() {} }, Date
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(SRC, 'utf8'), sandbox);
const fp = sandbox._pubFingerprintOf;

let failed = 0;
function check(label, cond, extra) {
  console.log(`  ${cond ? '✓' : '✗'} ${label}` + (cond ? '' : `   ${extra || ''}`));
  if (!cond) failed++;
}

// A believable slice of All Orders: SKU QTY LOC SO NOTE STATUS HAND LEFT SHIP COST
const base = () => ([
  ['158346', 8, 'F-23',  'SO-24853', 'Miguel',                        'PENDING', 110, '', '', ''],
  ['159012', 1, 'K-56',  'SO-24853', 'Miguel',                        'PENDING', 5,   '', '', ''],
  ['167400', 4, 'E-37',  'SO-24853', '↳ from KIT-159012 · Miguel',    'PENDING', 329, '', '', ''],
  ['173421', 1, 'G-22',  'SO-24853', 'Miguel',                        'PENDING', 19,  '', '', '']
]);

console.log('\n' + '='.repeat(70));
console.log('  the fold notices what the dirty flag has missed');
console.log('='.repeat(70));

const b = fp(base());

// The production call excludes HAND (col G, 0-based 6) — see _pubFingerprint.
const HAND = 6;
const fpH = v => fp(v, [HAND]);
const bH = fpH(base());

check('identical data → identical fingerprint', fp(base()) === b);
check('fingerprint is short enough for a Script Property', b.length < 64, `len ${b.length}`);

// ── the four real misses ──────────────────────────────────────────────────
let v = base(); v.splice(3, 0, ['172593', 4, 'F-20', 'SO-24853', '↳ from KIT-159012 · Miguel', 'PENDING', 328, '', '', '']);
check('ROW INSERTED noticed  (kit expansion / Pull insert)', fp(v) !== b);

v = base(); v[1][5] = 'CANCELED';
check('STATUS FLIPPED noticed  (Pull cancel — row count unchanged)', fp(v) !== b);

v = base(); v[0][4] = 'Miguel · HOLD do not ship';
check('NOTE EDITED noticed  (human typing)', fp(v) !== b);

v = base(); v.splice(1, 1);
check('ROW DELETED noticed  (human deleting)', fp(v) !== b);

// ── other changes the board renders ───────────────────────────────────────
v = base(); v[0][1] = 9;
check('QTY changed noticed', fp(v) !== b);

v = base(); v[0][2] = 'F-24';
check('LOCATION changed noticed', fp(v) !== b);

// ⚠ HAND IS EXCLUDED ON PURPOSE — see _pubFingerprint. recomputeHand rewrites
// column G on every Zoho stock push (every 2 min, 9–5) and does not bust the
// cache, so folding it in would make the sheet look permanently modified and
// turn the publisher into a near-constant rebuilder. HAND's freshness is
// unchanged by this: it already rides the 8-minute keep-fresh.
v = base(); v[0][6] = 111;
check('HAND churn does NOT force a rebuild (scheduled sync)', fpH(v) === bH);

v = base(); v[0][7] = 5;
check('LEFT (shelf count) changed noticed', fp(v) !== b);

// ...and with HAND excluded, everything else must STILL be caught — otherwise
// the exclusion would have quietly blinded the net.
v = base(); v[0][5] = 'CANCELED';   check('with HAND excluded · STATUS still caught', fpH(v) !== bH);
v = base(); v[0][4] = 'HOLD';       check('with HAND excluded · NOTE still caught',   fpH(v) !== bH);
v = base(); v[0][2] = 'F-24';       check('with HAND excluded · LOCATION still caught', fpH(v) !== bH);
v = base(); v.splice(1, 1);         check('with HAND excluded · row delete still caught', fpH(v) !== bH);

// ── the traps a naive fold falls into ─────────────────────────────────────
v = base(); v[0][0] = '15834'; v[0][1] = 68;   // "158346"+8 vs "15834"+68
check('cell boundaries respected (no run-together collision)', fp(v) !== b);

v = base(); var t = v[0]; v[0] = v[1]; v[1] = t;
check('ROW ORDER matters (a re-sort is a change)', fp(v) !== b);

v = base(); v[0][4] = ''; v[0][3] = 'SO-24853Miguel';
check('field boundaries respected', fp(v) !== b);

// ── stability: things that must NOT churn ─────────────────────────────────
const d = new Date(2026, 7, 14, 12, 0, 0);
const withDate = () => { const x = base(); x[0][8] = d; return x; };
check('a DATE cell is stable across reads', fp(withDate()) === fp(withDate()));

const nullish = () => { const x = base(); x[0][8] = null; x[0][9] = undefined; return x; };
check('null and undefined fold stably', fp(nullish()) === fp(nullish()));
// null, undefined and "" all mean "this cell is blank" — folding them alike is
// CORRECT, and it is what stops an empty cell churning the fingerprint.
check('blank / null / undefined all read as empty', fp(nullish()) === b);

check('empty sheet has its own stable fingerprint', fp([]) === fp([]));
check('empty differs from populated', fp([]) !== b);

console.log('\n' + '='.repeat(70));
if (failed) { console.log(`  ${failed} FAILURE(S)`); process.exit(1); }
console.log('  ALL CLEAR');
