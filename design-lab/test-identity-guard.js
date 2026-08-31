/**
 * test-identity-guard.js — the identity guard, loaded from the REAL files.
 *
 * Loads IdentityGuard.js and OrderService.js in a VM with the same Schema the sheet
 * uses, so the tests cannot drift from what ships.
 *
 * THE THREE THINGS WORTH PROVING:
 *
 *   1. ⭐ THE ALIBI. Before 2026-08-30, hand-editing a live SALES_ORDER logged RECEIVED
 *      for the NEW value — and the guard builds its known-good set from RECEIVED events.
 *      So the 08-28 corruption legitimised itself a minute before the guard looked, and
 *      the verdict came back "ok". Section A reproduces both worlds.
 *
 *   2. The copied row. GONE and UNKNOWN are both blind to it (the pair is complete and
 *      genuinely received, on the source row), and Zoho Pull's delta twin must NOT flag.
 *
 *   3. The CF rules are built from the constants, not typed twice.
 *
 * Run:  node test-identity-guard.js
 * HEAD: SRC=/tmp/head node test-identity-guard.js
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');

const SRC = process.env.SRC || path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(SRC, f), 'utf8');

const cols = { SKU:1, QTY:2, LOCATION:3, SALES_ORDER:4, NOTE:5, STATUS:6,
               HAND:7, LEFT:8, SHIPPING:9, SHIP_COST:10 };
const LOGCOLS = { TIMESTAMP:1, EVENT:2, ORDER_ID:3, SKU:4, QTY:5, SOURCE:6,
                  DETAIL:7, NOTE:8, PICKER:9 };

// ── the fake spreadsheet ────────────────────────────────────────────────────────────
let LOG_ROWS = [];            // [ts, event, orderId, sku, qty, source, detail, note, picker]
let IDENT_ROWS = [];          // the published mismatch list
let IDENT_WRITES = 0;

function fakeRange(rows, startRow, startCol, nRows, nCols) {
  // ⚠ COLUMN-ADDRESSED ON PURPOSE. The first version ignored startCol in setValues and
  //   clearContent, so a write to column B landed in column A and K10 accused correct
  //   code. A stub cheaper than the real thing tests nothing.
  const cell = (r, c) => { while (rows[r].length <= c) rows[r].push(''); return rows[r]; };
  const self = () => fakeRange(rows, startRow, startCol, nRows, nCols);
  return {
    getValues: () => rows.slice(startRow - 1, startRow - 1 + nRows)
                          .map(r => { const o = []; for (let c = 0; c < nCols; c++) o.push(r[startCol - 1 + c] === undefined ? '' : r[startCol - 1 + c]); return o; }),
    setValues: v => {
      for (let i = 0; i < v.length; i++) {
        while (rows.length <= startRow - 1 + i) rows.push([]);
        for (let c = 0; c < v[i].length; c++) { cell(startRow - 1 + i, startCol - 1 + c)[startCol - 1 + c] = v[i][c]; }
      }
      return self();
    },
    setNumberFormat: () => self(),
    clearContent: () => {
      for (let i = 0; i < nRows; i++) {
        if (!rows[startRow - 1 + i]) continue;
        for (let c = 0; c < nCols; c++) { if (rows[startRow - 1 + i][startCol - 1 + c] !== undefined) rows[startRow - 1 + i][startCol - 1 + c] = ''; }
      }
    },
    getBackgrounds: () => rows.slice(startRow - 1, startRow - 1 + nRows).map(() => new Array(nCols).fill('#ffffff')),
    setBackground: () => {}, clearNote: () => {}, getNote: () => '', getBackground: () => '#ffffff'
  };
}
const logSheet = {
  getLastRow: () => LOG_ROWS.length,
  getRange: (r, c, nr, nc) => fakeRange(LOG_ROWS, r, c, nr, nc)
};
const identSheet = {
  getLastRow: () => IDENT_ROWS.filter(r => r && r.some(v => v)).length,
  getRange: (r, c, nr, nc) => {
    const base = fakeRange(IDENT_ROWS, r, c, nr, nc);
    return Object.assign({}, base, {
      setValues: v => { IDENT_WRITES++; return base.setValues(v); }
    });
  },
  hideSheet: () => {}
};
let MAIN_ROWS = [];
const mainSheet = {
  getLastRow: () => MAIN_ROWS.length,
  getRange: (r, c, nr, nc) => fakeRange(MAIN_ROWS, r, c, nr, nc)
};
const ss = {
  getSheetByName: n => n === 'Activity Log' ? logSheet
                     : n === '__Identity'   ? identSheet
                     : n === 'All orders'   ? mainSheet : null,
  insertSheet: () => identSheet
};

let PROPS = {};
const sandbox = {
  console, Date, JSON, Math, Object, String, Number, Array, parseInt, parseFloat, isNaN,
  Schema: {
    cols, idx: n => cols[n] - 1,
    status: { PENDING:'PENDING', PREPARING:'PREPARING', SHIPPED:'SHIPPED', CANCELED:'CANCELED' },
    validStatuses: ['PENDING','PREPARING','SHIPPED','CANCELED'],
    isValidStatus: s => ['PENDING','PREPARING','SHIPPED','CANCELED']
                          .indexOf(String(s||'').trim().toUpperCase()) !== -1,
    boundaryMarker: 'DIRECT', dataStartRow: 4, dataWidth: 10, bannerRows: 3
  },
  ACTIVITY_LOG: { sheetName:'Activity Log', cols: LOGCOLS, idx: n => LOGCOLS[n]-1,
                  dataWidth: 9, dataStartRow: 2 },
  MAIN_SHEET_NAME: 'All orders', SPREADSHEET_ID: 'x',
  SpreadsheetApp: { openById: () => ss, getActive: () => ss,
    BooleanCriteria: { CUSTOM_FORMULA: 'CUSTOM_FORMULA' } },
  PropertiesService: { getScriptProperties: () => ({
    getProperty: k => (k in PROPS ? PROPS[k] : null),
    setProperty: (k, v) => { PROPS[k] = String(v); },
    deleteProperty: k => { delete PROPS[k]; }
  })},
  getBoundaryRow: () => BOUNDARY,
  UrlFetchApp: { fetch: () => { SENT++; return {}; } },
  TELEGRAM_ADMIN_CHAT_ID: 'chat', TELEGRAM_BOT_TOKEN: 'tok',
  Logger: { log: () => {} },
  logActivityBatch: b => { BATCHES.push(b); }
};
let BOUNDARY = -1, SENT = 0, BATCHES = [];
vm.createContext(sandbox);
vm.runInContext(read('IdentityGuard.js'), sandbox, { filename: 'IdentityGuard.js' });

// OrderService is loaded only for _mrClassify. It is huge and full of Sheets calls, but
// nothing at top level executes, so the declarations land and the rest stays unused.
let ORDER_LOADED = true;
try { vm.runInContext(read('OrderService.js'), sandbox, { filename: 'OrderService.js' }); }
catch (e) { ORDER_LOADED = false; console.log('  (OrderService.js did not load: ' + e.message + ')'); }

let ROWMGMT_LOADED = true;
try { vm.runInContext(read('RowManagement.js'), sandbox, { filename: 'RowManagement.js' }); }
catch (e) { ROWMGMT_LOADED = false; console.log('  (RowManagement.js did not load: ' + e.message + ')'); }

const G = sandbox;

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label +
              (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};
const section = name => console.log('\n' + name);
const soft = (name, fn) => { try { fn(); } catch (e) { fail++; console.log('  ✗ SECTION THREW (soft): ' + e.message); } };

// ── builders ────────────────────────────────────────────────────────────────────────
const row = (sku, order, status, note, qty) => {
  const r = new Array(10).fill('');
  r[cols.SKU-1] = sku; r[cols.SALES_ORDER-1] = order;
  r[cols.STATUS-1] = status || ''; r[cols.NOTE-1] = note || '';
  r[cols.QTY-1] = qty == null ? 1 : qty;
  return r;
};
const received = (orderId, sku, qty) => ['2026-08-28', 'RECEIVED', orderId, sku, qty == null ? 1 : qty, 'n8n', '', '', ''];
const identEdit = (orderId, sku) => ['2026-08-28', 'IDENTITY_EDIT', orderId, sku, 1, 'manual', 'identity changed', '', ''];
const setLog = rows => { LOG_ROWS = [['ts','EVENT','ORDER ID','SKU','QTY','SOURCE','DETAIL','NOTE','PICKER']].concat(rows); };
const known = () => G._igKnownFromLog(ss);

// The real incident.
const REAL_SO = '09-15094-35132', BOGUS_SO = 'Missing #: 05-15052-93025', SKU = '165447';


// =====================================================================================
section('A · THE ALIBI — the reason the shipped guard could not see 2026-08-28');
// =====================================================================================
soft('A', () => {
  // The world BEFORE the edit: the row was legitimately received.
  setLog([received(REAL_SO, SKU)]);
  t('A1 the untouched row reconciles',
    G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PREPARING' }, known(), {}).verdict, 'ok');

  // ── HEAD's behaviour: manualReceiveOnEdit logged RECEIVED for EVERY change, so the
  //    corrupted pair entered the known-good set before the guard ever looked.
  setLog([received(REAL_SO, SKU), received(BOGUS_SO, SKU)]);
  t('A2 ⚠ under HEAD the corruption legitimises itself → "ok" (the bug)',
    G._igVerdict({ orderId: BOGUS_SO, sku: SKU, status: 'PREPARING' }, known(), {}).verdict, 'ok');

  // ── The fix: a replacement logs IDENTITY_EDIT, which is not evidence.
  setLog([received(REAL_SO, SKU), identEdit(BOGUS_SO, SKU)]);
  t('A3 ⭐ with IDENTITY_EDIT the corruption is caught',
    G._igVerdict({ orderId: BOGUS_SO, sku: SKU, status: 'PREPARING' }, known(), {}).verdict, 'mismatch');
  t('A4 …and it names the reason',
    G._igVerdict({ orderId: BOGUS_SO, sku: SKU, status: 'PREPARING' }, known(), {}).reason,
    'this order/SKU pair was never received');
  t('A5 IDENTITY_EDIT never enters the known-good set',
    known().pairs[G._igSig(BOGUS_SO, SKU)] === undefined, true);
  t('A6 the original identity still matches, so Ctrl+Z restores "ok"',
    G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PREPARING' }, known(), {}).verdict, 'ok');
});


// =====================================================================================
section('B · _mrClassify — a receipt or an alteration?');
// =====================================================================================
soft('B', () => {
  if (!ORDER_LOADED || typeof G._mrClassify !== 'function') {
    fail++; console.log('  ✗ _mrClassify missing — HEAD, or OrderService failed to load');
    return;
  }
  const ev = old => ({ oldValue: old });
  const SO = cols.SALES_ORDER, SK = cols.SKU;

  t('B1 SO typed onto a row that already had a SKU → RECEIVED',
    G._mrClassify(ev(undefined), true, SO, '165447', 'SO-1').event, 'RECEIVED');
  t('B2 ⭐ SKU typed second, SO already there → RECEIVED (the hole that logged nothing)',
    G._mrClassify(ev(undefined), true, SK, '165447', 'SO-1').event, 'RECEIVED');
  t('B3 templated fill EXTENDS the old value → still a receipt',
    G._mrClassify(ev('Replacement for #:'), true, SO, '165447',
                  'Replacement for #: 26-14551-63163').event, 'RECEIVED');
  t('B4 ⚠ a live identity REPLACED → IDENTITY_EDIT',
    G._mrClassify(ev(REAL_SO), true, SO, SKU, BOGUS_SO).event, 'IDENTITY_EDIT');
  t('B5 …and the old value is recoverable from DETAIL',
    G._mrClassify(ev(REAL_SO), true, SO, SKU, BOGUS_SO).detail.indexOf(REAL_SO) !== -1, true);
  t('B6 a SKU swapped on a live row is an alteration too',
    G._mrClassify(ev('111111'), true, SK, '222222', 'SO-1').event, 'IDENTITY_EDIT');
  t('B7 no real change → nothing logged',
    G._mrClassify(ev(REAL_SO), true, SO, SKU, REAL_SO), null);
  t('B8 a paste cannot tell creation from mutation → stays optimistic',
    G._mrClassify({}, false, SO, SKU, BOGUS_SO).event, 'RECEIVED');
});


// =====================================================================================
section('C · GONE — half an identity on a live row');
// =====================================================================================
soft('C', () => {
  setLog([received(REAL_SO, SKU)]);
  t('C1 SALES ORDER deleted',
    G._igVerdict({ orderId: '', sku: SKU, status: 'PREPARING' }, known(), {}).verdict, 'gone');
  t('C2 SKU deleted',
    G._igVerdict({ orderId: REAL_SO, sku: '', status: 'PENDING' }, known(), {}).verdict, 'gone');
  t('C3 both gone but the status remains',
    G._igVerdict({ orderId: '', sku: '', status: 'SHIPPED' }, known(), {}).verdict, 'gone');
  t('C4 ⚠ a row with NO status is being typed — never judged',
    G._igVerdict({ orderId: '', sku: SKU, status: '' }, known(), {}).verdict, 'skip');
  t('C5 the DIRECT header row is not a status',
    G._igVerdict({ orderId: 'SALES ORDER', sku: '◈ SKU', status: 'STATUS' }, known(), {}).verdict, 'skip');
});


// =====================================================================================
section('D · DUPLICATED — the copied row, and the delta twin that must stay quiet');
// =====================================================================================
soft('D', () => {
  setLog([received(REAL_SO, SKU)]);
  BOUNDARY = -1;

  // A whole row copied onto an empty row.
  let data = [row(SKU, REAL_SO, 'PENDING'), row(SKU, REAL_SO, 'PENDING')];
  let scan = G._igScanRows(data, -1);
  t('D1 ⭐ the copy is counted twice', scan.pairCounts[G._igSig(REAL_SO, SKU)], 2);
  t('D2 …and both twins flag',
    G._igVerdict(scan.rows[0], known(), scan.pairCounts).verdict, 'duplicate');
  t('D3 GONE and UNKNOWN are both blind to it — the pair IS legitimate',
    G._igVerdict(scan.rows[0], known(), {}).verdict, 'ok');

  // Zoho Pull's insert_delta: a legitimate second row, and it says so.
  data = [row(SKU, REAL_SO, 'PENDING'),
          row(SKU, REAL_SO, 'PENDING', '↳ delta from Zoho · was 2, now 5')];
  scan = G._igScanRows(data, -1);
  t('D4 ⭐ the delta twin is excluded from the count', scan.pairCounts[G._igSig(REAL_SO, SKU)], 1);
  t('D5 …so neither row flags',
    G._igVerdict(scan.rows[0], known(), scan.pairCounts).verdict, 'ok');

  // A delta twin AND a genuine copy.
  data = [row(SKU, REAL_SO, 'PENDING'), row(SKU, REAL_SO, 'PENDING'),
          row(SKU, REAL_SO, 'PENDING', '↳ delta from Zoho · was 2, now 5')];
  scan = G._igScanRows(data, -1);
  t('D6 a delta AND a copy still flags', scan.pairCounts[G._igSig(REAL_SO, SKU)], 2);

  // Different orders for the same SKU are not duplicates.
  data = [row(SKU, REAL_SO, 'PENDING'), row(SKU, 'SO-99', 'PENDING')];
  scan = G._igScanRows(data, -1);
  t('D7 the same SKU on two DIFFERENT orders is normal',
    scan.pairCounts[G._igSig(REAL_SO, SKU)], 1);

  // The boundary rows never count.
  const div = new Array(10).fill(''); div[0] = 'DIRECT';
  data = [row(SKU, REAL_SO, 'PENDING'), div, row('◈ SKU', 'SALES ORDER', 'STATUS')];
  scan = G._igScanRows(data, 5);
  t('D8 divider + DIRECT header are skipped', scan.rows.length, 1);
});


// =====================================================================================
section('E · THE EVIDENCE GUARD — never accuse on an absence');
// =====================================================================================
soft('E', () => {
  setLog([received('SO-OTHER', '999999')]);
  t('E1 neither the order nor the SKU is in the tail → skip, not accuse',
    G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PENDING' }, known(), {}).verdict, 'skip');

  setLog([received('SO-OTHER', SKU)]);
  t('E2 the SKU is known but this pair is not → mismatch',
    G._igVerdict({ orderId: BOGUS_SO, sku: SKU, status: 'PENDING' }, known(), {}).verdict, 'mismatch');
  t('E3 …and it names where the SKU legitimately came from',
    G._igOrdersForSku(known(), SKU), ['SO-OTHER']);

  t('E4 an unreadable log yields no verdict rather than a wrong one',
    G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PENDING' }, null, {}).verdict, 'skip');
  t('E5 …but duplication is still knowable from the sheet alone',
    G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PENDING' }, null,
                 { [G._igSig(REAL_SO, SKU)]: 2 }).verdict, 'duplicate');
});


// =====================================================================================
section('F · THE PUBLISHED LIST — idempotent, and the count that reads it');
// =====================================================================================
soft('F', () => {
  IDENT_ROWS = []; IDENT_WRITES = 0;
  G._igWriteMismatchList(ss, ['a|1', 'b|2']);
  t('F1 the list is written', G._igReadMismatchList(ss), ['a|1', 'b|2']);
  const after = IDENT_WRITES;
  G._igWriteMismatchList(ss, ['a|1', 'b|2']);
  t('F2 ⭐ an unchanged set writes nothing', IDENT_WRITES, after);
  G._igWriteMismatchList(ss, ['a|1']);
  t('F3 a shrinking set rewrites', G._igReadMismatchList(ss), ['a|1']);
  G._igWriteMismatchList(ss, []);
  t('F4 clearing works', G._igReadMismatchList(ss), []);

  // the sidebar count mirrors the CF exactly
  const data = [row(SKU, REAL_SO, 'PENDING'),          // listed → UNKNOWN
                row('', 'SO-2', 'PENDING'),            // GONE
                row('777', 'SO-3', 'PENDING'),         // fine
                row('888', 'SO-4', 'PENDING'),         // duplicate pair
                row('888', 'SO-4', 'PENDING')];
  const rows = G._igIssueRows(data, -1, [G._igSig(REAL_SO, SKU)]);
  t('F5 the badge counts all three states', rows, [4, 5, 7, 8]);
  t('F6 ⭐ with an empty list nothing UNKNOWN flags — fail-safe',
    G._igIssueRows([row(SKU, REAL_SO, 'PENDING')], -1, []), []);
});


// =====================================================================================
section('G · THE CF RULES — built from the constants, not typed twice');
// =====================================================================================
soft('G', () => {
  const bt = read('BrandTheme.js');
  const m = bt.match(/function _identityFormulas\(anchorRow\)\s*\{[\s\S]*?\n\}/);
  t('G1 _identityFormulas is present', !!m, true);
  if (!m) return;
  const body = m[0];
  // ⭐ ONE builder, used by both the installer and the self-test — so the formula the
  //    diagnostic evaluates is byte-identical to the one that ships.
  t('G1b _buildIdentityRules uses it rather than a second copy',
    /_identityFormulas\(Schema\.dataStartRow\)/.test(bt), true);
  t('G1c …and so does the self-test',
    /_identityFormulas\(target\.row\)/.test(read('IdentityGuard.js')), true);

  t('G2 it reads the status list from Schema, never a literal',
    /Schema\.validStatuses\.map/.test(body), true);
  t('G3 the helper sheet name comes from IDENTITY_GUARD',
    /IDENTITY_GUARD\.sheetName/.test(body), true);
  t('G4 the list window comes from IDENTITY_GUARD',
    /IDENTITY_GUARD\.listMax/.test(body), true);
  t('G5 ⭐ the delta token comes from IDENTITY_GUARD, matching ZohoPull',
    /IDENTITY_GUARD\.deltaNoteToken/.test(body), true);
  t('G6 …and that token is the string ZohoPull actually writes',
    read('ZohoPull.js').indexOf(G.IDENTITY_GUARD.deltaNoteToken) !== -1, true);
  const rulesFn = (bt.match(/function _buildIdentityRules\(sheet\)\s*\{[\s\S]*?\n\}/) || [''])[0];
  t('G7 it paints only the two identity columns',
    /Schema\.cols\.SKU/.test(rulesFn) && /Schema\.cols\.SALES_ORDER/.test(rulesFn) &&
    !/Schema\.cols\.STATUS,/.test(rulesFn), true);
  t('G8 ⚠ the UNKNOWN lookup is IFERROR-wrapped so a missing sheet fails silent',
    /IFERROR\(ISNUMBER\(MATCH/.test(body), true);
  t('G9 ⚠ the duplicate count uses OPEN-ENDED ranges, which survive row inserts',
    /:\$A,/.test(body) && /:\$D,/.test(body) && !/\$A\$\d+:\$A\$\d+/.test(body), true);
  t('G10 no static fill is ever written by the guard',
    /setBackground\(IDENTITY_GUARD/.test(read('IdentityGuard.js')), false);

  // the theme must rebuild them, or applyBrandTheme silently deletes the feature
  t('G11 ⚠⚠ _applyAllConditionalFormatting rebuilds them',
    /keep\.push\.apply\(keep, _buildIdentityRules\(sheet\)\)/.test(bt), true);
  t('G12 a standalone installer exists', /function setupIdentityHighlighting\(\)/.test(bt), true);
  t('G13 …and a matching stripper', /function _stripIdentityRules\(rules\)/.test(bt), true);

  // ⚠⚠ An UNINSTALLED feature and a clean sheet look identical, because the mark is a
  //    display layer. The diagnostic must say so before anything else.
  const ig = read('IdentityGuard.js');
  t('G14 ⚠⚠ a reader exists for "are the rules actually on the sheet?"',
    /function _igCountInstalledRules\(sheet\)/.test(ig), true);
  t('G15 …and the diagnostic leads with it',
    ig.indexOf('CF rules installed:') !== -1 &&
    ig.indexOf('CF rules installed:') < ig.indexOf('established rows judged:'), true);
  t('G16 …counting by the same range signature the stripper uses',
    /cols\[0\] === Schema\.cols\.SKU && cols\[1\] === Schema\.cols\.SALES_ORDER/.test(ig), true);
});


// =====================================================================================
section('H · THE RECONCILE — cheap when clean, and it writes no data row');
// =====================================================================================
soft('H', () => {
  PROPS = {}; SENT = 0; IDENT_ROWS = [];
  t('H1 ⭐ nothing edited → one property read and out', G.runIdentityReconcile(), 'clean');

  PROPS[G.IDENTITY_GUARD.dirtyKey] = '1';
  setLog([received(REAL_SO, SKU)]);
  MAIN_ROWS = [[],[],[], row(SKU, REAL_SO, 'PREPARING')];
  BOUNDARY = -1;
  const out = G.runIdentityReconcile();
  t('H2 a healthy sheet flags nothing', out, '0 flagged');
  t('H3 …and the dirty flag is cleared', PROPS[G.IDENTITY_GUARD.dirtyKey], undefined);

  PROPS[G.IDENTITY_GUARD.dirtyKey] = '1';
  MAIN_ROWS = [[],[],[], row(SKU, BOGUS_SO, 'PREPARING')];
  const out2 = G.runIdentityReconcile();
  t('H4 ⭐ the corrupted row is flagged and reported', out2, '1 flagged · 1 reported');
  t('H5 …the pair is published for the CF rule to find',
    G._igReadMismatchList(ss), [G._igSig(BOGUS_SO, SKU)]);
  t('H6 …and Telegram went out exactly once', SENT, 1);

  PROPS[G.IDENTITY_GUARD.dirtyKey] = '1';
  SENT = 0;
  G.runIdentityReconcile();
  t('H7 ⚠ alert-once-per-crossing — the same problem does not re-send', SENT, 0);

  // the fix
  PROPS[G.IDENTITY_GUARD.dirtyKey] = '1';
  MAIN_ROWS = [[],[],[], row(SKU, REAL_SO, 'PREPARING')];
  G.runIdentityReconcile();
  t('H8 ⭐ fixing the row empties the published list', G._igReadMismatchList(ss), []);
});


// =====================================================================================
section('I · SHOULD-RECONCILE — no heartbeat sweep any more');
// =====================================================================================
soft('I', () => {
  t('I1 dirty runs', G._igShouldReconcile('123'), true);
  t('I2 ⭐ clean does not — the CF clears itself, so there is nothing to sweep for',
    G._igShouldReconcile(null), false);
  t('I3 STATUS is watched: a row becomes judgeable when it gets one',
    G.IDENTITY_GUARD.watch.indexOf('STATUS') !== -1, true);
  t('I4 QTY is deliberately NOT watched',
    G.IDENTITY_GUARD.watch.indexOf('QTY'), -1);
});


// =====================================================================================
section('J · ⚠⚠ THE OTHER DOORS — two CF strippers that deleted these rules on every edit');
// =====================================================================================
soft('J', () => {
  // 2026-08-30, found on real data: setupIdentityHighlighting() reported success, and the
  // FIRST edit afterwards silently removed all three rules. Both strippers asked "does ANY
  // range touch my column?" — and the identity rules span cols A AND D, with a status guard
  // containing both TRIM( and UPPER(TRIM(. Nothing marked, ever, and because the mark is a
  // display layer there was no residue to notice.
  if (!ROWMGMT_LOADED) { fail++; console.log('  ✗ RowManagement.js did not load'); return; }

  const mkRule = (formula, rangeCols) => ({
    getBooleanCondition: () => ({
      getCriteriaType: () => 'CUSTOM_FORMULA',
      getCriteriaValues: () => [formula]
    }),
    getRanges: () => rangeCols.map(c => ({ getColumn: () => c, getNumColumns: () => 1 }))
  });

  // the real thing: two ranges (A + D), and the status guard's TRIM(
  const identity = mkRule(
    '=AND(OR(UPPER(TRIM($F4))="PENDING"), OR($A4="", $D4=""))',
    [cols.SKU, cols.SALES_ORDER]);
  // genuine legacy rules the strippers exist to remove
  const legacySO  = mkRule('=COUNTIF($D$4:$D,$D4)>1', [cols.SALES_ORDER]);
  const legacySku = mkRule('=UPPER(TRIM($A4))="165447"', [cols.SKU]);

  let written = null;
  const fakeSheet = {
    getConditionalFormatRules: () => [identity, legacySO, legacySku],
    setConditionalFormatRules: r => { written = r; }
  };

  written = null;
  G.removeLegacySalesOrderCFRules(fakeSheet);
  let kept = written || [identity, legacySO, legacySku];
  t('J1 ⭐ the identity rules SURVIVE the SALES ORDER stripper', kept.indexOf(identity) !== -1, true);
  t('J2 …and a genuine legacy col-D rule is still removed', kept.indexOf(legacySO), -1);

  written = null;
  G.removeDuplicateHighlightRules(fakeSheet);
  kept = written || [];
  t('J3 ⭐ the identity rules SURVIVE the SKU stripper', kept.indexOf(identity) !== -1, true);
  t('J4 …and a genuine duplicate-SKU rule is still removed', kept.indexOf(legacySku), -1);

  // the property that makes both safe, stated directly
  const rm = read('RowManagement.js');
  t('J5 ⚠ both strippers require EVERY range to be their own column',
    (rm.match(/var isOrderColumn = ranges\.length > 0;/g) || []).length === 1 &&
    (rm.match(/var isSkuColumn = ranges\.length > 0;/g) || []).length === 1, true);
});


// =====================================================================================
section('K · QTY — the identity is right and the quantity is not');
// =====================================================================================
soft('K', () => {
  setLog([received(REAL_SO, SKU, 1)]);
  const k = known();

  t('K1 the received qty reconciles',
    G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PENDING', qty: 1 }, k, {}).verdict, 'ok');
  t('K2 ⭐ a hand-typed qty is caught',
    G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PENDING', qty: 7 }, k, {}).verdict, 'qty');
  t('K3 …and it says what we actually received',
    G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PENDING', qty: 7 }, k, {}).reason,
    'we received 1 of this, not 7');
  t('K4 a number and its string form are the same qty',
    G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PENDING', qty: '1' }, k, {}).verdict, 'ok');
  t('K5 a blank qty is not judged',
    G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PENDING', qty: '' }, k, {}).verdict, 'ok');

  // ⚠ ONE PROBLEM PER ROW — a row we cannot vouch for at all is never ALSO accused of a
  //   wrong qty. The more fundamental verdict wins.
  t('K6 ⚠ an unknown identity outranks a qty complaint',
    G._igVerdict({ orderId: BOGUS_SO, sku: SKU, status: 'PENDING', qty: 7 }, k, {}).verdict, 'mismatch');

  // a pair legitimately received twice at different quantities — a re-entry, or a delta
  setLog([received(REAL_SO, SKU, 1), received(REAL_SO, SKU, 4)]);
  const k2 = known();
  t('K7 ⭐ EITHER received qty is accepted',
    [G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PENDING', qty: 1 }, k2, {}).verdict,
     G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PENDING', qty: 4 }, k2, {}).verdict],
    ['ok', 'ok']);
  t('K8 …but a third value is still wrong',
    G._igVerdict({ orderId: REAL_SO, sku: SKU, status: 'PENDING', qty: 9 }, k2, {}).verdict, 'qty');

  // the key the sheet builds must equal the key JS publishes
  t('K9 the qty signature is identity + qty',
    G._igQtySig(REAL_SO, SKU, 7), G._igSig(REAL_SO, SKU) + '|7');

  // end to end through the reconcile
  PROPS = {}; SENT = 0; IDENT_ROWS = [];
  setLog([received(REAL_SO, SKU, 1)]);
  PROPS[G.IDENTITY_GUARD.dirtyKey] = '1';
  MAIN_ROWS = [[],[],[], row(SKU, REAL_SO, 'PREPARING', '', 7)];
  BOUNDARY = -1;
  G.runIdentityReconcile();
  t('K10 ⭐ the qty key is published to column B, not column A',
    [G._igReadMismatchList(ss, 1), G._igReadMismatchList(ss, 2)],
    [[], [G._igQtySig(REAL_SO, SKU, 7)]]);

  PROPS[G.IDENTITY_GUARD.dirtyKey] = '1';
  MAIN_ROWS = [[],[],[], row(SKU, REAL_SO, 'PREPARING', '', 1)];
  G.runIdentityReconcile();
  t('K11 ⭐ correcting the qty empties the list', G._igReadMismatchList(ss, 2), []);

  // the badge sees it
  t('K12 the sidebar count includes a wrong qty',
    G._igIssueRows([row(SKU, REAL_SO, 'PENDING', '', 7)], -1, [], [G._igQtySig(REAL_SO, SKU, 7)]),
    [4]);

  // the CF rule
  const bt = read('BrandTheme.js');
  const F = (bt.match(/function _identityFormulas\(anchorRow\)\s*\{[\s\S]*?\n\}/) || [''])[0];
  t('K13 the qty rule reads column B of the helper', /listRefFor\('B'\)/.test(F), true);
  t('K14 …and builds the same three-part key JS publishes',
    /LOWER\(TRIM\(' \+ d \+ '\)\)&"\|"&LOWER\(TRIM\(' \+ a \+ '\)\)&"\|"&TRIM\(' \+ b \+ '\)/.test(F), true);
  const rulesFn = (bt.match(/function _buildIdentityRules\(sheet\)\s*\{[\s\S]*?\n\}/) || [''])[0];
  t('K15 ⚠ it paints column B alone — the identity is fine on such a row',
    /Schema\.cols\.QTY/.test(rulesFn), true);
  t('K16 the rule count constant matches what is built', G.IDENTITY_GUARD.ruleCount, 4);
});


console.log('\n' + (fail === 0 ? '✅' : '❌') +
            ' test-identity-guard: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail === 0 ? 0 : 1);
