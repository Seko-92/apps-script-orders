// ============================================================================
// test-kit-build.js — building a kit nobody bought
//
// Loads the REAL KitBuild.js, the REAL expandKit (KitExpansion.js) and the REAL
// compareLocations (Helpers.js) in one VM. Only the DATA SOURCES are stubbed —
// the registry, the MI maps and the Zoho mirror. Everything under test is code
// that ships.
//
// WHAT IS WORTH PROVING
//
//   1. ⭐ THE ROLL-UP. Two kits sharing a piston must become ONE gather line and
//      ONE walk. That is the entire reason this is a window and not a better
//      hand-typed row — on the sheet each kit expands under its own parent, so
//      the picker walks to that shelf twice.
//
//   2. ⚠ AISLE ORDER, NOT ALPHABETICAL. "A-9" sorts AFTER "A-50" under
//      localeCompare. That bug shipped in three files and was only caught when
//      a new surface printed the list in a column someone read top to bottom.
//
//   3. ⭐ THE HONESTY RULES. Bundled parts start unchecked (a picker sent after
//      a gasket already inside the gasket set walks for nothing); a kit with
//      unreadable Purchase Description lines says so instead of printing a list
//      that looks complete; an unknown SKU comes back with a reason.
//
//   4. ⚠⚠ IT WRITES NOTHING. The whole safety argument. Asserted against the
//      source, because a future edit that adds one setValues would be silent.
//
// Run:  node test-kit-build.js
// HEAD: SRC=/tmp/head node test-kit-build.js
// ============================================================================
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');

const SRC = process.env.SRC || path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(SRC, f), 'utf8');

let failed = 0, passed = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` +
    (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  ok ? passed++ : failed++;
};
const section = n => console.log('\n' + n);
const soft = (name, fn) => { try { fn(); }
  catch (e) { console.log(`  ✗ ${name} THREW: ${e.message}`); failed++; } };

// ── THE FIXTURE REGISTRY ────────────────────────────────────────────────────
// Shaped like the real thing: two kits that SHARE a piston, one READY kit, one
// with an unreadable PD line, one component that ships inside another's box.
const REGISTRY = new Map([
  ['158679', { sku:'158679', name:'Engine Overhaul Kit STD', type:'READY',
    location:'K-55', engine:'V2203', salesDescription:'', unparsedLines:[],
    components:[
      { sku:'155394', name:'Full Gasket Set',  qty:1, bundled:false, bundledInto:'' },
      { sku:'167517', name:'Piston Kit STD',   qty:2, bundled:false, bundledInto:'' },
      { sku:'162198', name:'Head Gasket',      qty:1, bundled:true,  bundledInto:'155394' },
      { sku:'171018', name:'Main Bearing Set', qty:1, bundled:false, bundledInto:'' }
    ]}],
  ['217205', { sku:'217205', name:'Repair Kit 0.50', type:'MANUAL',
    location:'NOT FOUND', engine:'', salesDescription:'', unparsedLines:[],
    components:[
      { sku:'167517', name:'Piston Kit STD',   qty:3, bundled:false, bundledInto:'' },
      { sku:'173763', name:'Thrust Washer',    qty:2, bundled:false, bundledInto:'' }
    ]}],
  ['999001', { sku:'999001', name:'Half-read Kit', type:'MANUAL',
    location:'K-9', engine:'', salesDescription:'',
    unparsedLines:['-1 Head Gasket (161262)'],
    components:[
      { sku:'155394', name:'Full Gasket Set',  qty:1, bundled:false, bundledInto:'' }
    ]}]
]);

// Shelves chosen to expose a lexical sort: A-9 must precede A-50.
const LOCATIONS = new Map([
  ['155394','A-50'], ['167517','A-9'], ['162198','L-226'],
  ['171018','E-17'], ['173763','B-4']
]);
const MI_AVAIL   = new Map([['155394',12], ['167517',22], ['162198',7], ['171018',3]]);
const ZOHO_AVAIL = new Map([['167517',18]]);   // Zoho wins where present

const sandbox = {
  console, JSON, Math, Object, String, Number, Array, parseInt, parseFloat, isNaN, Date,
  SPREADSHEET_ID: 'x', MAIN_SHEET_NAME: 'All orders',
  Logger: { log(){} },
  Schema: { pickIdA1: () => 'F2', dataStartRow: 4, dataWidth: 10,
            idx: () => 0, cols: {} },
  HtmlService: { createTemplateFromFile: () => ({ evaluate: () => ({ setWidth(){return this;}, setHeight(){return this;} }) }) },
  SpreadsheetApp: { openById: () => ({ getSheetByName: () => ({
      getRange: () => ({ getValue: () => 'Shipping - Yassin 1' }) }) }),
    getUi: () => ({ showModalDialog(){} }) },
  Utilities: { formatDate: () => '9/3/26 10:14 AM' },
  // ── data sources, the only stubs ──
  buildKitMap: () => REGISTRY,
  getKitInfo: sku => REGISTRY.get(String(sku).trim()) || null,
  buildLocationAndInventoryMaps: () => ({
    locationMap: LOCATIONS,
    inventoryMap: new Map([...MI_AVAIL].map(([k,v]) => [k, { available: v }]))
  }),
  buildZohoStockMap: () => new Map([...ZOHO_AVAIL].map(([k,v]) => [k, { available: v }])),
  _extractPickIdData: raw => String(raw).replace(/^Shipping\s*-\s*/i,'').replace(/\s+(\d+)$/,' · $1')
};
vm.createContext(sandbox);

// REAL compareLocations + its parser, REAL expandKit — never re-typed.
function lift(file, fnName) {
  const src = read(file);
  const i = src.indexOf('function ' + fnName);
  if (i === -1) throw new Error(fnName + ' not found in ' + file);
  let d = 0, started = false;
  for (let j = i; j < src.length; j++) {
    if (src[j] === '{') { d++; started = true; }
    else if (src[j] === '}') { d--; if (started && d === 0)
      return vm.runInContext(src.slice(i, j + 1), sandbox, { filename: file + '#' + fnName }); }
  }
  throw new Error('unbalanced ' + fnName);
}
lift('Helpers.js', '_parseShelfLocation');
lift('Helpers.js', 'compareLocations');
lift('KitExpansion.js', 'expandKit');
vm.runInContext(read('KitBuild.js'), sandbox, { filename: 'KitBuild.js' });
const B = sandbox;

console.log('\n' + '='.repeat(74));
console.log('  kit build — roll-up, aisle order, and the honesty rules');
console.log('='.repeat(74));

// ============================================================================
section('A · ONE KIT — the simple case');
// ============================================================================
soft('A', () => {
  const r = B.getKitBuildPlan({ kits: [{ sku: '158679', qty: 1 }] });
  t('A1 it plans', r.ok, true);
  t('A2 the kit is found', r.kits[0].found, true);

  // ⭐ the bundled Head Gasket starts EXCLUDED, so it is not gathered
  const gathered = r.gather.map(g => g.sku).sort();
  t('A3 ⭐ the bundled part is NOT on the gather list', gathered, ['155394','167517','171018']);
  t('A4 …and it is still SHOWN, flagged, so the picker can re-check it',
    r.kits[0].components.filter(c => c.bundled).map(c => [c.sku, c.excluded, c.bundledInto]),
    [['162198', true, '155394']]);

  t('A5 component qty comes straight from the registry at ×1',
    r.gather.filter(g => g.sku === '167517')[0].totalQty, 2);
});

// ============================================================================
section('B · QUANTITY — build three, gather three times the parts');
// ============================================================================
soft('B', () => {
  const r = B.getKitBuildPlan({ kits: [{ sku: '158679', qty: 3 }] });
  t('B1 ⭐ a 2-per-kit part × 3 kits is 6',
    r.gather.filter(g => g.sku === '167517')[0].totalQty, 6);
  t('B2 a 1-per-kit part × 3 kits is 3',
    r.gather.filter(g => g.sku === '155394')[0].totalQty, 3);
  t('B3 the totals report kit UNITS, not kit lines', r.totals.kitUnits, 3);
  t('B4 …and pieces is the sum of the gather list',
    r.totals.pieces, r.gather.reduce((n,g) => n + g.totalQty, 0));
});

// ============================================================================
section('C · ⭐⭐ THE ROLL-UP — the reason this is a window');
// ============================================================================
soft('C', () => {
  const r = B.getKitBuildPlan({ kits: [{ sku:'158679', qty:2 }, { sku:'217205', qty:3 }] });

  const piston = r.gather.filter(g => g.sku === '167517');
  t('C1 ⭐ the shared piston is ONE line, not two', piston.length, 1);
  t('C2 ⭐ …carrying the summed quantity (2×2 + 3×3)', piston[0].totalQty, 13);
  t('C3 ⭐ …and naming both kits it serves',
    piston[0].usedBy.map(u => u.kitSku).sort(), ['158679','217205']);
  t('C4 the per-kit breakdown survives for the assemble section',
    piston[0].usedBy.map(u => [u.kitSku, u.kits, u.subtotal]).sort(),
    [['158679', 2, 4], ['217205', 3, 9]]);
  t('C5 the totals count it as one shared line', r.totals.shared, 1);

  // ⚠ THE POINT: one walk to A-9, not two.
  const shelves = r.gather.map(g => g.location);
  t('C6 ⭐ A-9 appears exactly once across both kits',
    shelves.filter(l => l === 'A-9').length, 1);
});

// ============================================================================
section('D · ⚠ AISLE ORDER — natural, never lexical');
// ============================================================================
soft('D', () => {
  const r = B.getKitBuildPlan({ kits: [{ sku:'158679', qty:1 }, { sku:'217205', qty:1 }] });
  const order = r.gather.map(g => g.location);
  t('D1 ⭐ A-9 comes BEFORE A-50 — localeCompare puts it after',
    order.indexOf('A-9') < order.indexOf('A-50'), true);
  t('D2 the whole walk is in aisle order', order, ['A-9','A-50','B-4','E-17']);
  t('D3 distinct shelves are counted for the walk', r.totals.shelves, 4);
});

// ============================================================================
section('E · ⭐ THE HONESTY RULES');
// ============================================================================
soft('E', () => {
  // (a) an unreadable Purchase Description line
  const r1 = B.getKitBuildPlan({ kits: [{ sku:'999001', qty:1 }] });
  t('E1 ⭐ an unparsed PD line raises a warning', r1.warnings.length, 1);
  t('E2 …named as incomplete, not silently dropped', r1.warnings[0].kind, 'unparsed');
  t('E3 …and it quotes the offending line',
    r1.warnings[0].lines, ['-1 Head Gasket (161262)']);

  // (b) a SKU that is not a kit
  const r2 = B.getKitBuildPlan({ kits: [{ sku:'000000', qty:1 }] });
  t('E4 an unknown SKU is refused WITH a reason, not an empty list',
    [r2.kits[0].found, r2.warnings[0].kind], [false, 'unknown']);
  t('E5 …and it does not poison the gather list', r2.gather.length, 0);

  // (c) a component with no shelf
  const r3 = B.getKitBuildPlan({ kits: [{ sku:'217205', qty:1 }] });
  const tw = r3.gather.filter(g => g.sku === '173763')[0];
  t('E6 a component MI does not know says NOT FOUND, never blank',
    [tw.location, tw.missing], ['B-4', false]);

  // (d) stock routing
  const r4 = B.getKitBuildPlan({ kits: [{ sku:'158679', qty:1 }] });
  t('E7 ⭐ availability is ZOHO-FIRST where Zoho knows the SKU',
    r4.gather.filter(g => g.sku === '167517')[0].available, 18);
  t('E8 …and falls back to MI where it does not',
    r4.gather.filter(g => g.sku === '155394')[0].available, 12);
  t('E9 ⚠ short-against-hand is flagged, informationally',
    r4.kits[0].components.filter(c => c.sku === '171018')[0].short, false);
});

// ============================================================================
section('F · EXCLUSIONS — the picker decides');
// ============================================================================
soft('F', () => {
  const r = B.getKitBuildPlan({
    kits: [{ sku:'158679', qty:1 }],
    excluded: { '158679': ['167517'] }
  });
  t('F1 an unchecked component leaves the gather list',
    r.gather.map(g => g.sku).indexOf('167517'), -1);
  t('F2 …but is still listed for review, marked',
    r.kits[0].components.filter(c => c.sku === '167517')[0].excluded, true);

  // Re-checking a bundled part puts it back.
  const r2 = B.getKitBuildPlan({ kits: [{ sku:'158679', qty:1 }], excluded: { '158679': [] } });
  t('F3 ⭐ a bundled part can be re-checked and then IS gathered',
    r2.gather.map(g => g.sku).indexOf('162198') !== -1, true);

  // ⚠⚠ THE CONTRACT THE CLIENT MUST HONOUR. Once a kit HAS a key, the server stops
  //   proposing and does exactly what it is told. So a client that seeds only the
  //   toggled part loses the bundled default on the very first click — the gasket
  //   that ships inside the gasket set quietly comes back onto the walk.
  const r3 = B.getKitBuildPlan({ kits: [{ sku:'158679', qty:1 }],
                                 excluded: { '158679': ['167517'] } });
  t('F4 ⚠ a partial list means the picker HAS an opinion — bundled is no longer proposed',
    r3.gather.map(g => g.sku).indexOf('162198') !== -1, true);
  t('F5 ⭐ …which is why the modal seeds the bundled skus into `excluded` on first plan',
    /seedBundled|bundled[\s\S]{0,80}excluded\[/.test(read('KitBuildModal.html')), true);

  // ⚠⚠ AND THE OTHER HALF OF THE SAME CONTRACT (bug found 2026-09-03).
  //   The client built its payload with `if (list.length) payload.excluded[k] = list`,
  //   so an EMPTY list was dropped — the server then read "untouched" and re-applied
  //   the bundled default. Re-checking a bundled part silently did not stick.
  //   Presence of the key IS the opinion, so the key must always be sent.
  const MODALSRC = read('KitBuildModal.html')
    .replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '');
  t('F6 ⚠ the client never drops an EMPTY exclusion list',
    /if\s*\(list\.length\)\s*payload\.excluded/.test(MODALSRC), false);
  t('F7 ⭐ …it sends the key for every kit it has an opinion about',
    /payload\.excluded\[k\]\s*=\s*Object\.keys/.test(MODALSRC), true);

  // Prove the server side does the right thing with that empty list.
  const r4 = B.getKitBuildPlan({ kits: [{ sku:'158679', qty:1 }], excluded: { '158679': [] } });
  t('F8 ⭐ an empty list re-includes the bundled part — the re-check sticks',
    r4.gather.map(g => g.sku).sort(), ['155394','162198','167517','171018']);
});

// ============================================================================
section('G · BOUNDS + GUARDS');
// ============================================================================
soft('G', () => {
  t('G1 no kits is a refusal, not an empty sheet',
    B.getKitBuildPlan({ kits: [] }).ok, false);
  const many = [];
  for (let i = 0; i < B.KIT_BUILD.maxKits + 1; i++) many.push({ sku:'158679', qty:1 });
  t('G2 too many kits is refused with the limit named',
    B.getKitBuildPlan({ kits: many }).ok, false);
  t('G3 a zero/negative qty is clamped to 1',
    B.getKitBuildPlan({ kits: [{ sku:'158679', qty:0 }] }).totals.kitUnits, 1);
  t('G4 an absurd qty is clamped to the cap',
    B.getKitBuildPlan({ kits: [{ sku:'158679', qty:9999 }] }).totals.kitUnits,
    B.KIT_BUILD.maxQtyPerKit);
});

// ============================================================================
section('H · ⚠⚠ IT WRITES NOTHING — the whole safety argument');
// ============================================================================
soft('H', () => {
  const src = read('KitBuild.js')
    .replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '');
  ['setValue', 'setValues', 'appendRow', 'insertRows', 'deleteRow', 'setProperty',
   'logActivity', 'clearContent', 'CacheService'].forEach(bad => {
    t('H · no ' + bad + '(', src.indexOf(bad) === -1, true);
  });
  t('H10 ⭐ and the picker gate is NOT applied — nothing is stamped, so nothing is owed',
    /needsPicker|No picker set/.test(src), false);
});

console.log('\n' + '='.repeat(74));
if (failed) { console.log(`❌ test-kit-build: ${passed} passed, ${failed} failed`); process.exit(1); }
console.log(`✅ test-kit-build: ${passed} passed, 0 failed`);
