/**
 * test-owner-bridge.js — the hop that lets a LOCKED All Orders keep a working sidebar.
 *
 * Loads the REAL OwnerBridge.js (and StatusService.js for the pairs target shape) in a VM,
 * so the tests cannot drift from what ships.
 *
 * THE THINGS WORTH PROVING:
 *   1. ⚠⚠ the allowlist is a real boundary — an unlisted name, and every prototype name,
 *      are refused. A `this[fnName]` dispatch would hand a token-holder every global
 *      function in the project plus constructor / __proto__ / toString.
 *   2. ⚠ an UNKNOWN identity HOPS. The hop works for everybody; the direct path only
 *      spares the owner ~2s. Guessing "probably the owner" reintroduces the blocked-write
 *      failure the bridge exists to remove.
 *   3. ⚠⚠ VALUES cross the hop, never row numbers — the 2026-05-08 / 2026-08-21 row-shift
 *      class, twice bitten here.
 *   4. ⚠ Apps Script answers HTTP 200 with an error PAGE. The body decides, not the status.
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');

const SRC = process.env.SRC || path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(SRC, f), 'utf8');

const cols = { SKU:1, QTY:2, LOCATION:3, SALES_ORDER:4, NOTE:5, STATUS:6,
               HAND:7, LEFT:8, SHIPPING:9, SHIP_COST:10 };

let FETCHED = [], REPLY = () => JSON.stringify({ ok:true, result:'did it' });
let EMAIL = '';
let PROPS = {};

const sandbox = {
  console, JSON, String, Number, Array, Object, Boolean, eval,
  Schema: { cols, idx: n => cols[n]-1, dataStartRow: 4, dataWidth: 10,
            status:{PENDING:'PENDING',PREPARING:'PREPARING',SHIPPED:'SHIPPED',CANCELED:'CANCELED'},
            validStatuses:['PENDING','PREPARING','SHIPPED','CANCELED'],
            normalize:s=>String(s||'').trim().toUpperCase(),
            isValidStatus:s=>['PENDING','PREPARING','SHIPPED','CANCELED'].indexOf(String(s||'').trim().toUpperCase())!==-1,
            isTerminal:s=>['SHIPPED','CANCELED'].indexOf(String(s||'').trim().toUpperCase())!==-1 },
  WEB_APP_URL: 'https://script.google.com/macros/s/EXEC/exec',
  APP_SECRET_TOKEN: 'tok', OWNER_EMAIL: 'owner@example.com',
  MAIN_SHEET_NAME: 'All orders', SPREADSHEET_ID: 'x',
  Session: { getEffectiveUser: () => ({ getEmail: () => EMAIL }) },
  PropertiesService: { getScriptProperties: () => ({
    getProperty: k => (k in PROPS ? PROPS[k] : null),
    setProperty: (k,v) => { PROPS[k] = String(v); },
    deleteProperty: k => { delete PROPS[k]; }
  })},
  UrlFetchApp: { fetch: (url, opts) => {
    FETCHED.push({ url, payload: JSON.parse(opts.payload), opts });
    return { getContentText: () => REPLY() };
  }},
  SpreadsheetApp: { openById: () => ({ getSheetByName: () => null }) },
  Logger: { log: () => {} },
  ALL_ORDERS_LOCK: { tag: 'HQ-LOCK', n8nAccountKey: 'N8N_SHEETS_ACCOUNT', noneSentinel: 'none' },
  protectAllOrdersSheet: () => { PROTECTED = true; return '✅ locked'; }
};
let PROTECTED = false;
vm.createContext(sandbox);
vm.runInContext(read('OwnerBridge.js'), sandbox, { filename: 'OwnerBridge.js' });
let STATUS_OK = true;
try { vm.runInContext(read('StatusService.js'), sandbox, { filename:'StatusService.js' }); }
catch (e) { STATUS_OK = false; console.log('  (StatusService.js did not load: ' + e.message + ')'); }
const G = sandbox;

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label +
              (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};
const section = n => console.log('\n' + n);
const soft = (n, f) => { try { f(); } catch (e) { fail++; console.log('  ✗ SECTION THREW (soft): ' + e.message); } };
const reset = () => { FETCHED = []; PROPS = {}; EMAIL = ''; REPLY = () => JSON.stringify({ ok:true, result:'did it' }); };


section('A · ⚠⚠ THE ALLOWLIST IS THE SECURITY BOUNDARY');
soft('A', () => {
  reset();
  t('A1 an unlisted function is refused',
    G._obRunAsOwner('deleteEverything', []).ok, false);
  ['constructor','__proto__','toString','hasOwnProperty','valueOf'].forEach((n,i) => {
    t('A2.' + (i+1) + ' prototype name refused: ' + n, G._obRunAsOwner(n, []).ok, false);
  });
  t('A3 …and the refusal names what was asked for',
    G._obRunAsOwner('nope', []).error.indexOf('nope') !== -1, true);
  t('A4 _asOwner refuses before it ever reaches the network', (() => {
    try { G._asOwner('deleteEverything', []); return 'no throw'; }
    catch (e) { return FETCHED.length === 0 ? 'refused, no fetch' : 'REFUSED BUT FETCHED'; }
  })(), 'refused, no fetch');

  const bridge = read('OwnerBridge.js');
  t('A5 ⚠⚠ dispatch is a literal map, never this[name]',
    /this\s*\[\s*fnName\s*\]/.test(bridge) || /global\s*\[/.test(bridge), false);
  t('A6 every allowlisted name has an entry in the dispatch map', (() => {
    const map = (bridge.match(/var map = \{[\s\S]*?\n  \};/) || [''])[0];
    return G.OWNER_BRIDGE.actionNames.filter(n => map.indexOf(n + ':') === -1);
  })(), []);
  // ⚠ SCAN EVERY FILE, never a hardcoded list. The first version listed nine files and
  //   passed vacuously the moment an action landed in a tenth — which is exactly what
  //   happened when addReplacementFromSidebar was added from Replacements.js.
  t('A7 …and every allowlisted name is a function that actually exists', (() => {
    const all = fs.readdirSync(SRC).filter(f => f.endsWith('.js')).map(read).join('\n');
    return G.OWNER_BRIDGE.actionNames.filter(n => all.indexOf('function ' + n + '(') === -1);
  })(), []);
});


section('B · WHO HOPS — and an unknown identity always does');
soft('B', () => {
  reset(); EMAIL = 'owner@example.com';
  t('B1 the owner runs directly', G._obIsOwner(), true);

  reset(); EMAIL = 'staff@example.com';
  t('B2 a signed-in staff member hops', G._obIsOwner(), false);

  reset(); EMAIL = '';
  t('B3 ⚠ an ANONYMOUS user hops — the safe default', G._obIsOwner(), false);

  reset(); EMAIL = 'OWNER@EXAMPLE.COM';
  t('B4 the comparison is case-insensitive', G._obIsOwner(), true);

  reset(); EMAIL = 'owner@example.com';
  PROPS['OWNER_EMAIL'] = 'someoneelse@example.com';
  t('B5 a Script Property overrides the Secrets constant', G._obIsOwner(), false);

  reset();
  t('B6 ⚠⚠ inside doPost the owner context wins outright — recursion is impossible', (() => {
    EMAIL = 'staff@example.com';
    const before = G._obIsOwner();
    const r = G._obRunAsOwner('sortEbayTable', []);   // sets _OB_IN_OWNER_CONTEXT
    return [before, r.ok !== undefined];
  })(), [false, true]);
});


section('C · ⚠⚠ WHAT CROSSES THE HOP — values, never row numbers');
soft('C', () => {
  reset(); EMAIL = 'staff@example.com';
  G._asOwner('markPreparingByValues', [[{ orderId: 'SO-1', sku: '165447' }]]);
  const p = FETCHED[0].payload;
  t('C1 the action is runAsOwner', p.action, 'runAsOwner');
  t('C2 the token rides server-side, never in a browser', p.token, 'tok');
  t('C3 ⚠ the payload carries VALUES', p.args[0][0], { orderId: 'SO-1', sku: '165447' });
  t('C4 …and no row number anywhere in it',
    /"(row|startRow|numRows|rowNumber)"/.test(JSON.stringify(p)), false);
  t('C5 it posts to our own /exec', FETCHED[0].url, G.WEB_APP_URL);
  t('C6 ⚠ redirects are followed — /exec answers 302',
    FETCHED[0].opts.followRedirects === false, false);
});


section('D · ⚠ THE REPLY — Apps Script answers 200 with an error PAGE');
soft('D', () => {
  reset(); EMAIL = 'staff@example.com';
  REPLY = () => '<!DOCTYPE html><html>Sorry, the file you requested does not exist</html>';
  t('D1 ⚠⚠ an HTML page is a FAILURE, whatever the status code', (() => {
    try { G._asOwner('sortEbayTable', []); return 'accepted the page'; }
    catch (e) { return /New Version|page instead/.test(e.message) ? 'refused, and says why' : e.message; }
  })(), 'refused, and says why');

  reset(); EMAIL = 'staff@example.com';
  REPLY = () => JSON.stringify({ ok: false, error: 'Kit row not found' });
  t('D2 a genuine refusal surfaces its own reason', (() => {
    try { G._asOwner('commitKitFromModal', ['s']); return 'no throw'; }
    catch (e) { return e.message; }
  })(), 'Kit row not found');

  reset(); EMAIL = 'staff@example.com';
  REPLY = () => JSON.stringify({ ok: true, result: '✅ Sorted.' });
  t('D3 a success returns the inner result', G._asOwner('sortEbayTable', []), '✅ Sorted.');

  reset(); EMAIL = 'staff@example.com';
  REPLY = () => { throw new Error('network down'); };
  t('D4 an unreachable web app fails loudly, never silently', (() => {
    try { G._asOwner('sortEbayTable', []); return 'silently succeeded'; }
    catch (e) { return /Could not reach/.test(e.message) ? 'loud' : e.message; }
  })(), 'loud');
});


section('E · THE pairs TARGET SHAPE — resolved by value, under the lock');
soft('E', () => {
  if (!STATUS_OK) { fail++; console.log('  ✗ StatusService.js did not load'); return; }
  const rows = [
    ['165447', 1, 'I-6', 'SO-1', '', 'PENDING', '', '', '', ''],
    ['999999', 1, 'A-1', 'SO-2', '', 'PENDING', '', '', '', ''],
    ['165447', 1, 'I-6', 'SO-3', '', 'PENDING', '', '', '', '']
  ];
  const sheet = { getRange: () => ({ getValues: () => rows }) };
  const last = 6;
  t('E1 a pair resolves to its CURRENT row',
    G._resolveStatusTargetRows(sheet, { pairs: [{ orderId:'SO-2', sku:'999999' }] }, last), [5]);
  t('E2 several pairs resolve in one pass',
    G._resolveStatusTargetRows(sheet, { pairs: [{ orderId:'SO-1', sku:'165447' },
                                                { orderId:'SO-3', sku:'165447' }] }, last), [4, 6]);
  t('E3 ⚠ the same SKU on a DIFFERENT order is not swept up',
    G._resolveStatusTargetRows(sheet, { pairs: [{ orderId:'SO-1', sku:'165447' }] }, last), [4]);
  t('E4 an unknown pair matches nothing rather than guessing',
    G._resolveStatusTargetRows(sheet, { pairs: [{ orderId:'SO-9', sku:'000' }] }, last), []);
  t('E5 an empty or malformed pair list is refused',
    [G._resolveStatusTargetRows(sheet, { pairs: [] }, last),
     G._resolveStatusTargetRows(sheet, { pairs: [{ orderId:'', sku:'' }] }, last)], [[], []]);
});


section('F · THE GUARDS — every protected writer actually has one');
soft('F', () => {
  // ⚠ every file, for the same reason as A7
  const all = fs.readdirSync(SRC).filter(f => f.endsWith('.js')).map(read).join('\n');
  const guarded = G.OWNER_BRIDGE.actionNames.filter(n =>
    new RegExp("_asOwner\\('" + n + "'").test(all));
  // markPreparingByValues is the WRITE half — it is hopped TO, never FROM
  const expect = G.OWNER_BRIDGE.actionNames.filter(n => n !== 'markPreparingByValues');
  t('F1 every allowlisted action except the write-half has a guard',
    expect.filter(n => guarded.indexOf(n) === -1), []);
  t('F2 ⚠ markSelectedPreparing resolves the selection BEFORE hopping',
    /_asOwner\('markPreparingByValues', \[pairs\]\)/.test(read('FulfillmentService.js')), true);
  t('F3 …and the write half never sees a row number',
    /function markPreparingByValues\(pairs\)/.test(read('FulfillmentService.js')), true);
  t('F4 runAsOwner is wired into doPost',
    /payload\.action === 'runAsOwner'/.test(read('OrderService.js')), true);
  // ⚠ ASSERT ON THE DATA STRUCTURE, NOT THE SOURCE TEXT. The first version of this
  //   matched /DOPOST_LOCK_FREE[\s\S]{0,400}runAsOwner/ and failed against correct code —
  //   because the COMMENT above the handler says "It is NOT in DOPOST_LOCK_FREE". The
  //   documentation describes exactly what the pattern was hunting for. Ninth instance of
  //   a harness accusing working code in this project; the standing rule is to strip
  //   comments, or better, read the value.
  const lockFree = (read('OrderService.js').match(/var DOPOST_LOCK_FREE = \{[\s\S]*?\};/) || [''])[0];
  t('F5 ⚠ runAsOwner is NOT lock-free — these are writes',
    lockFree.indexOf('runAsOwner') !== -1, false);
  t('F5b ⚠⚠ doPost declares the owner context up front, so ARRIVALS never hop',
    /_OB_IN_OWNER_CONTEXT = true;/.test(read('OrderService.js').slice(0, read('OrderService.js').indexOf('runAsOwner'))), true);
  t('F6 …and the lock-free list is still only reads',
    /board(Tick|Radio|Pickers|Print|Part|PartLite|Order)/.test(lockFree), true);
});


section('G · ⚠⚠ THE ROLLOUT GATES — refusing is the safety');
soft('G', () => {
  const armed = () => JSON.stringify({ ok:false, error:'Not an allowlisted owner action: __probe__' });

  // gate 1 — no bridge, no lock
  reset(); PROTECTED = false; EMAIL = 'owner@example.com';
  REPLY = () => '<!DOCTYPE html>Sorry, the file you requested does not exist';
  const r1 = G.installAllOrdersLock();
  t('G1 ⚠⚠ a dead bridge REFUSES the lock', PROTECTED, false);
  t('G2 …and says why, in terms of the consequence',
    /REFUSED/.test(r1) && /staff sidebar write/i.test(r1), true);

  // gate 2 — no n8n carve-out, no lock
  reset(); PROTECTED = false; EMAIL = 'owner@example.com'; REPLY = armed;
  const r2 = G.installAllOrdersLock();
  t('G3 ⚠ a missing n8n account REFUSES the lock', PROTECTED, false);
  t('G4 …and names the ~1 AM sweep it would silently stop',
    /N8N_SHEETS_ACCOUNT/.test(r2), true);

  // both gates satisfied
  reset(); PROTECTED = false; EMAIL = 'owner@example.com'; REPLY = armed;
  PROPS['N8N_SHEETS_ACCOUNT'] = 'n8n@example.com';
  const r3 = G.installAllOrdersLock();
  t('G5 ⭐ with both gates passed it installs', PROTECTED, true);
  t('G6 …and warns against the incognito test that broke 2026-08-29',
    /incognito/i.test(r3), true);
  t('G7 …and names the rollback', /unprotectAllOrdersSheet/.test(r3), true);

  // the end-to-end verifier forces a REAL hop even for the owner
  reset(); EMAIL = 'owner@example.com'; REPLY = armed;
  G.verifyOwnerBridge();
  t('G8 ⭐ verifyOwnerBridge does a real round-tripped WRITE, not just a ping',
    FETCHED.some(f => f.payload.fn === 'refreshKitSkuMarkers'), true);

  // ⚠⚠ the lock controls must NEVER be reachable through the bridge
  t('G10 ⚠⚠ unprotectAllOrdersSheet is NOT allowlisted — staff could otherwise unlock',
    G.OWNER_BRIDGE.actionNames.indexOf('unprotectAllOrdersSheet'), -1);
  t('G11 ⚠⚠ …and neither is protectAllOrdersSheet or the installer',
    ['protectAllOrdersSheet','installAllOrdersLock'].filter(
      n => G.OWNER_BRIDGE.actionNames.indexOf(n) !== -1), []);
  t('G12 the lock controls are owner-gated in the source', (() => {
    const bt = read('BrandTheme.js'), ob = read('OwnerBridge.js');
    return /_obRequireOwner\("Locking All Orders"\)/.test(bt) &&
           /_obRequireOwner\("Unlocking All Orders"\)/.test(bt) &&
           /_obRequireOwner\("Locking All Orders"\)/.test(ob);
  })(), true);
  reset(); EMAIL = 'staff@example.com';
  t('G13 a staff member gets a sentence, not a stack trace',
    /owner-only/.test(G._obRequireOwner('Locking All Orders') || ''), true);
  reset(); EMAIL = 'owner@example.com';
  t('G14 …and the owner is not obstructed', G._obRequireOwner('Locking All Orders'), null);

  // the setter cannot fire empty
  reset(); PROPS = {};
  G.setN8nSheetsAccountNow();
  t('G9 ⚠ the setter refuses while its value is still blank',
    PROPS['N8N_SHEETS_ACCOUNT'], undefined);
});


section('H · ⚠⚠ onOpen — the sidebar must never be hostage to a background refresh');
soft('H', () => {
  // 2026-08-30, reported from a real employee PC: with All Orders locked, an employee
  // opening the sheet got NO SIDEBAR AT ALL. onOpen runs as the USER who opened the file,
  // setupDuplicateSalesOrderHighlighting writes borders and number formats to the protected
  // sheet, that throw killed the rest of the function — and showSidebar() sat after it.
  // Same shape as the 2026-05-01 onEditInstallable bug: nothing a user depends on may sit
  // downstream of something that can fail.
  // ⚠⚠ THE BODY IS BRACE-MATCHED, NOT SLICED AT A MAGIC NUMBER. This used to take a fixed
  //    3000 characters after `function onOpen()`, so the moment the function grew a longer
  //    comment block the window cut off before showSidebar() and the suite accused
  //    CORRECT CODE of the exact regression it exists to prevent. Found 2026-09-02 when
  //    the menu rewrite landed; H1, H3 and H4 had already been failing that way for a
  //    while and nobody could see it, because a truncated window fails quietly.
  //    Fifteenth instance in this project of the harness being the bug. Suspect it first.
  //
  // ⚠ AND COMMENTS ARE STRIPPED FIRST. This function's own docblock says "showSidebar()
  //   used to sit after it" — so a naive indexOf finds the PROSE before the CALL, and H1's
  //   ordering check passes for entirely the wrong reason. Documentation describes the
  //   bug, which is exactly what these patterns are hunting for.
  const main = read('Main.js');
  const bodyOf = (src, decl) => {
    const start = src.indexOf(decl);
    if (start < 0) return '';
    let i = src.indexOf('{', start), depth = 0;
    for (let j = i; j < src.length; j++) {
      if (src[j] === '{') depth++;
      else if (src[j] === '}' && --depth === 0) return src.slice(start, j + 1);
    }
    return src.slice(start);
  };
  const stripComments = (src) =>
    src.replace(/\/\*[\s\S]*?\*\//g, ' ').replace(/^[ \t]*\/\/.*$/gm, ' ');
  const body = stripComments(bodyOf(main, 'function onOpen()'));

  const iSidebar = body.indexOf('showSidebar()');
  const iDup     = body.indexOf('setupDuplicateSalesOrderHighlighting()');
  const iHand    = body.indexOf('setupHandConditionalFormatting()');
  t('H1 ⭐ showSidebar runs BEFORE the sheet-writing maintenance',
    iSidebar > -1 && iDup > iSidebar && iHand > iSidebar, true);
  t('H2 ⚠ showSidebar is wrapped, so even IT cannot take the rest down',
    /try \{\s*showSidebar\(\);\s*\} catch/.test(body), true);
  t('H3 every background step is isolated in its own try/catch',
    (body.match(/try \{[^}]*\}\s*catch \(err\)/g) || []).length >= 4, true);
  t('H4 ⚠ the sheet writers are owner-gated, so staff never trigger them at all',
    /if \(maintainer\) \{/.test(body) && /_obIsOwner\(\)/.test(body), true);

  // and the admin/setup writers refuse rather than error
  const bt = read('BrandTheme.js');
  const adminGated = ['setupBuyerNoteHighlighting','setupEbayLogo','setupIdentityHighlighting',
                      'setupKitRowHighlighting','protectSheetStructure','unprotectSheetStructure']
    .filter(n => {
      // ⚠ MATCH THE NAME, NOT A SIGNATURE SHAPE. This looked for the literal
      //   `function NAME()` with empty parens, so setupEbayLogo(plate) — which IS
      //   owner-gated, and has been — was silently dropped from the list and the count
      //   came back 5 of 6. The suite then reported a missing owner gate on a function
      //   that had one. Sixteenth instance of the harness accusing correct code here;
      //   the tell, as ever, was that the "missing" thing was plainly there when read.
      // ⚠ Anchored with \s*\( so setupEbayLogoSomethingElse could never satisfy it.
      const m = new RegExp('function ' + n + '\\s*\\(').exec(bt);
      return !!m && bt.slice(m.index, m.index + 500).indexOf('_obRequireOwner') > -1;
    });
  t('H5 every admin writer on All Orders is owner-gated', adminGated.length, 6);
  t('H6 ⚠⚠ …and NONE of them is allowlisted in the bridge',
    ['setupBuyerNoteHighlighting','setupEbayLogo','setupIdentityHighlighting',
     'setupKitRowHighlighting','protectSheetStructure','unprotectSheetStructure']
      .filter(n => G.OWNER_BRIDGE.actionNames.indexOf(n) !== -1), []);
});


section('I · ⚠⚠ THE SWEEP — no sidebar action may write a LOCKED column unguarded');
soft('I', () => {
  // 2026-08-30, reported from use: markSelectedPreparing worked but adding a
  // missing/replacement line was refused. The bridge was fine — addReplacementFromSidebar
  // simply was not allowlisted. It was missed because the first sweep read only the first
  // ~900 chars after `google.script.run`, and that call sits past a long success handler.
  //
  // ⭐ SO THE SWEEP IS A TEST NOW, NOT A ONE-OFF. Patching callers one at a time is what
  //   cost the CF-stripper round; an enumeration that is not executable rots the day it is
  //   written.
  const files = fs.readdirSync(SRC).filter(f => f.endsWith('.js'));
  const bodies = Object.create(null), known = new Set();
  files.forEach(f => {
    const s = read(f);
    const re2 = /^function (\w+)\s*\(/gm; let m;
    while ((m = re2.exec(s))) {
      const n = m.group === undefined ? m[1] : m[1];
      known.add(n);
      const st = m.index + m[0].length;
      const nx = s.indexOf('\nfunction ', st);
      bodies[n] = s.slice(st, nx > 0 ? nx : s.length)
                   .replace(/\/\/.*|\/\*[\s\S]*?\*\//g, '');   // ⚠ strip comments
    }
  });

  const sb = read('Sidebar.html');
  // ⚠ ANY `.serverFn(` anywhere counts. A windowed scan is what missed the bug.
  const entry = [...known].filter(n => new RegExp('[.\'"]' + n + '\\s*[(\'"]').test(sb));

  const allsrc = files.map(read).join('\n');
  const guarded = new Set([...allsrc.matchAll(/_asOwner\('(\w+)'/g)].map(m => m[1]));
  const ownerOnly = new Set(Object.keys(bodies).filter(n => bodies[n].includes('_obRequireOwner')));

  const WRITE = /setValues?\(|insertRows?Before|insertRows?After|deleteRows?\(|setBackground|setNumberFormats?|setBorder|clearContent|setRichTextValues?|setConditionalFormatRules|copyFormatToRange|setDataValidation|setFormula/;
  // the lock's carve-outs — a function touching ONLY these is fine for staff
  const OPEN = new Set(['NOTE', 'STATUS', 'LEFT']);

  const writesLocked = (fn, seen = new Set(), d = 0) => {
    if (seen.has(fn) || d > 6 || !bodies[fn]) return null;
    seen.add(fn);
    const b = bodies[fn];
    if (WRITE.test(b) && b.includes('MAIN_SHEET_NAME')) {
      const cols = [...b.matchAll(/Schema\.cols\.(\w+)/g)].map(m => m[1]);
      const structural = /insertRows?Before|insertRows?After|deleteRows?\(|setConditionalFormatRules/.test(b);
      if (structural || cols.some(c => !OPEN.has(c))) return fn;
    }
    for (const m of b.matchAll(/\b(\w+)\s*\(/g)) {
      const n = m[1];
      if (guarded.has(n) || ownerOnly.has(n)) continue;
      const r = writesLocked(n, seen, d + 1);
      if (r) return r;
    }
    return null;
  };

  const leaks = entry
    .filter(n => !guarded.has(n) && !ownerOnly.has(n))
    .map(n => [n, writesLocked(n)])
    .filter(([, w]) => w);

  t('I1 the sweep can see enough of the sidebar to be meaningful', entry.length > 80, true);
  t('I2 ⚠⚠ NO sidebar action writes a locked column unguarded',
    leaks.map(([n, w]) => n + ' → ' + w), []);
  t('I3 the two found in production are now covered',
    ['addReplacementFromSidebar', 'recomputeHandFromZohoStock'].filter(n => !guarded.has(n)), []);
  t('I4 ⭐ hold acknowledgement is NOT bridged — it writes only NOTE, a carve-out',
    guarded.has('acknowledgeSelectedHold') || guarded.has('boardAckHold'), false);
});


console.log('\n' + (fail === 0 ? '✅' : '❌') +
            ' test-owner-bridge: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail === 0 ? 0 : 1);
