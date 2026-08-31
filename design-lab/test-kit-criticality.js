// =====================================================================================
// test-kit-criticality.js — the lens the ripple deliberately cannot provide.
//
// analyzeRestockRipple returns early on any kit that is not already blocked
// (`if (build.buildable > 0) return;`). That is right for its question and it is
// exactly why a component TWELVE healthy kits depend on is invisible until the
// morning it hits zero. This proves the other lens sees it.
//
// Loads the REAL Ripple.js in a VM with stubs, so it cannot drift from shipped code.
//
// Prove against HEAD:
//   mkdir -p /tmp/headkc && git show HEAD:Ripple.js > /tmp/headkc/Ripple.js \
//     && SRC=/tmp/headkc node design-lab/test-kit-criticality.js
// =====================================================================================
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');

let pass = 0, fail = 0;
const ok = (n, c, x) => {
  if (c) { pass++; return; }
  fail++; console.log('  ✗ ' + n + (x !== undefined ? '  → ' + JSON.stringify(x) : ''));
};

/**
 * @param kits  [{ sku, name, comps:[[sku,qty]], bad?:true }]
 * @param stock { skuLower: available }
 */
function run(kits, stock, opts) {
  opts = opts || {};
  const calls = { buildKitMap: 0, buildMaps: 0, buildZoho: 0 };

  const kitMap = new Map();
  kits.forEach(k => kitMap.set(String(k.sku).toLowerCase(), {
    sku: k.sku, name: k.name || '',
    components: (k.comps || []).map(c => ({ sku: c[0], qty: c[1], name: 'part ' + c[0] })),
    _bad: !!k.bad
  }));

  const ctx = {
    console: { log: () => {}, error: () => {} },
    Map, JSON, String, Number, Math, Array, Object, Date, parseInt, parseFloat, isNaN, RegExp,
    buildLocationAndInventoryMaps: () => { calls.buildMaps++;
      return { inventoryMap: new Map(), locationMap: new Map([['aaa', 'E-37']]) }; },
    buildZohoStockMap: () => { calls.buildZoho++; return new Map(); },
    buildKitMap: () => { calls.buildKitMap++; return kitMap; },
    _oosResolveAvailFactory: () => (key) => (key in stock ? stock[key] : null),
    // ⚠ the authoritative verdict is STUBBED so the test controls which kits are
    //   assessable — the point here is the tally, not the buildability engine.
    _oosComputeKitBuild: (kit) => kit._bad
      ? { buildable: '⚠', limitedBy: '⚠ PD unreadable' }
      : { buildable: 1, limitedBy: '' }
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(SRC, 'Ripple.js'), 'utf8'), ctx, { filename: 'Ripple.js' });

  /* ⚠ FAIL SOFT, NEVER THROW (the choosePicker lesson, re-learned 2026-08-21).
     Against HEAD this function does not exist. Returning undefined made every
     later section throw on the first property access, so the run aborted and the
     before/after proof showed nothing about the sections that DO pass on both.
     A before/after proof is only useful if every section reports. */
  if (!ctx.analyzeKitCriticality) {
    return { missing: true, calls, ctx,
             d: { parts: [], atRisk: [], totalKits: 0, totalParts: 0, skipped: 0 } };
  }
  const d = opts.shared
    ? ctx.analyzeKitCriticality({ inventoryMap: new Map(), locationMap: new Map() }, new Map(), kitMap)
    : ctx.analyzeKitCriticality();
  return { d, calls, ctx };
}
// Never undefined — a missing part yields a blank record so assertions FAIL
// with a readable value instead of throwing on a property access.
const find = (d, sku) => d.parts.filter(p => p.sku === sku)[0] || {};

console.log('\n  test-kit-criticality.js — what the catalogue leans on\n');

// ── A · THE WHOLE POINT: a part nothing is short of still ranks ──────────────────────
{
  // 'aaa' feeds 5 HEALTHY kits and is in plentiful stock. The ripple cannot see it
  // at all — every kit using it is buildable. This lens must rank it first.
  const kits = [
    { sku: 'K1', comps: [['aaa', 1], ['zzz', 1]] },
    { sku: 'K2', comps: [['aaa', 1]] },
    { sku: 'K3', comps: [['aaa', 1]] },
    { sku: 'K4', comps: [['aaa', 1]] },
    { sku: 'K5', comps: [['aaa', 1]] },
    { sku: 'K6', comps: [['zzz', 1]] }
  ];
  const { d } = run(kits, { aaa: 99, zzz: 99 });
  ok('A1 · a part nothing is short of is still ranked', !!find(d, 'aaa').sku);
  ok('A2 · fan-out decides the order — aaa (5) above zzz (2)',
     !!d.parts[0] && d.parts[0].sku === 'aaa', d.parts.map(p => p.sku + ':' + p.kits));
  ok('A3 · the headline count is kits-that-depend-on-it', find(d, 'aaa').kits === 5,
     find(d, 'aaa').kits);
  ok('A4 · nothing is flagged short when everything is covered',
     find(d, 'aaa').cannotCover === 0, find(d, 'aaa').cannotCover);
  ok('A5 · it names the kits, for the "→" line',
     (find(d, 'aaa').kitSkus || []).length === 5);
}

// ── B · depth is measured against the HUNGRIEST recipe, not an average ──────────────
{
  const kits = [
    { sku: 'K1', comps: [['aaa', 1]] },
    { sku: 'K2', comps: [['aaa', 6]] }   // the hungry one
  ];
  const { d } = run(kits, { aaa: 12 });
  const p = find(d, 'aaa');
  ok('B1 · maxQtyPer takes the largest requirement', p.maxQtyPer === 6, p.maxQtyPer);
  ok('B2 · coverKits = floor(12 / 6) = 2, not floor(12/1)', p.coverKits === 2, p.coverKits);
  ok('B3 · both kits still covered, so nothing reads short', p.cannotCover === 0, p.cannotCover);
}

// ── C · "already short" is present tense, counted per recipe ────────────────────────
{
  const kits = [
    { sku: 'K1', comps: [['aaa', 1]] },   // covered by 3
    { sku: 'K2', comps: [['aaa', 4]] },   // NOT covered by 3
    { sku: 'K3', comps: [['aaa', 9]] }    // NOT covered by 3
  ];
  const { d } = run(kits, { aaa: 3 });
  const p = find(d, 'aaa');
  ok('C1 · counts only the recipes it cannot cover', p.cannotCover === 2, p.cannotCover);
  ok('C2 · fan-out still counts all three', p.kits === 3, p.kits);
}

// ── D · untrustworthy kits are EXCLUDED and COUNTED, never silently folded in ───────
{
  const kits = [
    { sku: 'K1', comps: [['aaa', 1]] },
    { sku: 'K2', comps: [['aaa', 1]], bad: true }   // PD unreadable
  ];
  const { d } = run(kits, { aaa: 5 });
  ok('D1 · the unreadable kit does not inflate fan-out', find(d, 'aaa').kits === 1,
     find(d, 'aaa').kits);
  ok('D2 · but it IS counted, so the exclusion is visible', d.skipped === 1, d.skipped);
  ok('D3 · totalKits still reports everything walked', d.totalKits === 2, d.totalKits);
}

// ── E · the alert set: load-bearing AND thin, both conditions required ──────────────
{
  const kits = [];
  for (let i = 0; i < 6; i++) kits.push({ sku: 'K' + i, comps: [['aaa', 1], ['bbb', 1]] });
  kits.push({ sku: 'Ksolo', comps: [['ccc', 1]] });
  // aaa: 6 kits, 1 left  → load-bearing AND thin  → alert
  // bbb: 6 kits, 50 left → load-bearing, healthy  → no alert
  // ccc: 1 kit,  0 left  → thin but not load-bearing → no alert
  const { d } = run(kits, { aaa: 1, bbb: 50, ccc: 0 });
  const names = d.atRisk.map(p => p.sku);
  ok('E1 · a thin load-bearing part is flagged', names.indexOf('aaa') !== -1, names);
  ok('E2 · a healthy load-bearing part is NOT flagged', names.indexOf('bbb') === -1, names);
  ok('E3 · an empty part only one kit uses is NOT flagged', names.indexOf('ccc') === -1, names);
  ok('E4 · so the alert set is exactly the intersection', d.atRisk.length === 1, names);
}

// ── F · REGRESSION NETS — shared maps, and callers that pass nothing ────────────────
{
  const kits = [{ sku: 'K1', comps: [['aaa', 1]] }];
  const bare = run(kits, { aaa: 5 });
  ok('F1 · a caller passing nothing builds its own maps', bare.calls.buildMaps === 1,
     bare.calls);
  const shared = run(kits, { aaa: 5 }, { shared: true });
  ok('F2 · a caller handing over maps does NOT rebuild them', shared.calls.buildMaps === 0,
     shared.calls);
  ok('F3 · nor the kit map — the hourly pass builds it once for both lenses',
     shared.calls.buildKitMap === 0, shared.calls);
}

// ── G · REGRESSION NET — the ripple still works, and still ignores healthy kits ─────
{
  const kits = [
    { sku: 'K1', comps: [['aaa', 1]] },      // healthy
    { sku: 'K2', comps: [['ddd', 1]] }       // ddd is empty → blocked
  ];
  const ctxRun = run(kits, { aaa: 9, ddd: 0 });
  const r = ctxRun.ctx.analyzeRestockRipple
    ? ctxRun.ctx.analyzeRestockRipple(
        { inventoryMap: new Map(), locationMap: new Map() }, new Map())
    : null;
  ok('G1 · analyzeRestockRipple still runs unchanged', !!r);
  // ⚠ Its stub reports every kit buildable:1, so nothing is blocked — which is the
  //   point: the ripple sees nothing here while criticality still ranks both parts.
  ok('G2 · and criticality reports parts the ripple has no opinion on',
     ctxRun.d.parts.length === 2, ctxRun.d.parts.map(p => p.sku));
}

console.log('\n  ' + pass + ' passed, ' + fail + ' failed\n');
process.exit(fail ? 1 : 0);
