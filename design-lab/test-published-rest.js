// =====================================================================================
// test-published-rest.js — the resting panel's `rest` key must survive the PUBLISHED tier.
//
// WHY THIS EXISTS. The two dark rows shipped 2026-08-20 with their server half proven
// (test-rest-snapshot.js, 35 assertions) — but only through getSidebarTick, the LIVE
// call. The Control Panel is normally served from __Published!A1 instead, and that copy
// carried api + alerts and NOT rest. So the rows worked on the fallback tier and
// silently never on the tier that actually answers: the feature looked broken exactly
// when everything else was healthy.
//
// Loads the REAL Published.js into a VM with stubs, so this cannot drift from shipped
// code. Captures the JSON handed to setValue — i.e. the bytes a reader gets.
//
// Prove against HEAD:
//   mkdir -p /tmp/headpr && git show HEAD:Published.js > /tmp/headpr/Published.js \
//     && SRC=/tmp/headpr node design-lab/test-published-rest.js
// =====================================================================================
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');

let pass = 0, fail = 0;
const ok = (n, c, x) => {
  if (c) { pass++; return; }
  fail++; console.log('  ✗ ' + n + (x !== undefined ? '  → ' + JSON.stringify(x) : ''));
};

const REST = { cantBuild: { kits: 26, units: 117 }, ripple: [{ sku: '195306', sole: 4 }] };

/**
 * @param o.slow      what _sidebarSlowParts()/the cache holds (null = cold)
 * @param o.cached    what the raw script cache holds on the inline path
 * @param o.bulk      pad the tick so it blows the 45,000-char cap
 */
function run(reason, o) {
  o = o || {};
  const written = {};
  const calls = { slowParts: 0, cacheGet: 0 };

  const sheet = {
    getRange: addr => ({
      setValue: v => { written[addr] = v; },
      getValue: () => written[addr] || ''
    }),
    setColumnWidth: () => {}, hideSheet: () => {}
  };

  const ctx = {
    console: { log: () => {}, error: () => {} },
    JSON, String, Number, Date, Math, Array, Object, parseInt, parseFloat, isNaN, RegExp,
    SPREADSHEET_ID: 'x',
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => sheet, insertSheet: () => sheet }) },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: () => null, setProperty: () => {}, deleteProperty: () => {}
    }) },
    CacheService: { getScriptCache: () => ({
      get: () => { calls.cacheGet++; return o.cached ? JSON.stringify(o.cached) : null; },
      put: () => {}
    }) },
    Utilities: { formatDate: () => '' },
    SIDEBAR_SLOW_CACHE_KEY: 'hqSidebarSlow_v1',
    _sidebarSlowParts: () => { calls.slowParts++; return o.slow || null; },
    _buildDashboardTick: () => ({
      cockpit: { timeline: o.bulk ? new Array(400).fill({ t: 12.5, e: 'SHIPPED', o: '24-15008-33107' }) : [] },
      alerts: { paidShipping: { count: 0, rows: [] } },
      api: null, picker: 'Yassin · 1', pickers: [], lastSync: '4:24 PM', paceCar: null,
      openOrders: o.bulk ? new Array(900).fill({ sku: '195306', loc: 'A-14', qty: 1, note: 'x'.repeat(40) }) : [],
      openOrdersTotal: 0, openOrdersBy: { EBAY: 0, DIRECT: 0 }, kits: [],
      serverTime: new Date().toISOString()
    })
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(SRC, 'Published.js'), 'utf8'), ctx, { filename: 'Published.js' });

  const res = ctx.publishBoardTick(reason);
  let tick = null;
  try { tick = JSON.parse(written['A1']); } catch (e) {}
  return { res, tick, calls, written };
}

const SLOW = { api: { worstPct: 12 }, alerts: { outOfStock: { count: 107, rows: [1, 2, 3] } },
               rest: REST, _builtAt: Date.now() };

console.log('\n  test-published-rest.js — the published tier must carry `rest`\n');

// ── A · scheduled publish: rest lands TOP-LEVEL, where the panel reads it ────────────
{
  const { tick } = run('changed', { slow: SLOW });
  ok('A1 · tick.rest is present on a scheduled publish', !!(tick && tick.rest), tick && tick.rest);
  ok('A2 · it is the real snapshot, not a placeholder',
     !!(tick && tick.rest && tick.rest.cantBuild && tick.rest.cantBuild.units === 117),
     tick && tick.rest);
  ok('A3 · the ripple rows survive the trip',
     !!(tick && tick.rest && Array.isArray(tick.rest.ripple) && tick.rest.ripple[0].sku === '195306'));
  // The client reads `full.rest` where full is the whole tick — NOT full.sidebar.rest.
  ok('A4 · it is TOP-LEVEL, matching getSidebarTick\'s own shape',
     !!(tick && Object.prototype.hasOwnProperty.call(tick, 'rest')));
}

// ── B · not paid for twice: the cell has a 45,000-char budget the board shares ───────
{
  const { tick } = run('changed', { slow: SLOW });
  ok('B1 · rest is NOT duplicated inside tick.sidebar',
     !!(tick && tick.sidebar && !Object.prototype.hasOwnProperty.call(tick.sidebar, 'rest')),
     tick && tick.sidebar && Object.keys(tick.sidebar));
}

// ── C · REGRESSION NET: the shipped api/alerts behaviour is untouched ────────────────
{
  const { tick } = run('changed', { slow: SLOW });
  ok('C1 · sidebar.api still carried', !!(tick && tick.sidebar && tick.sidebar.api));
  ok('C2 · sidebar.alerts still carried', !!(tick && tick.sidebar && tick.sidebar.alerts));
  ok('C3 · alerts stay LEAN — count only, rows stripped',
     !!(tick && tick.sidebar.alerts.outOfStock &&
        tick.sidebar.alerts.outOfStock.count === 107 &&
        !('rows' in tick.sidebar.alerts.outOfStock)),
     tick && tick.sidebar.alerts.outOfStock);
  ok('C4 · client validation would accept this copy (cockpit + sidebar both present)',
     !!(tick && tick.cockpit && tick.sidebar));
}

// ── D · INLINE publish (a ✓ Pick): carries rest, and must NEVER open eight sheets ────
{
  const { tick, calls } = run('pick', { cached: SLOW });
  ok('D1 · an inline publish still carries rest from the warm cache', !!(tick && tick.rest),
     tick && tick.rest);
  ok('D2 · REGRESSION NET — it never rebuilds the slow half inside a pick',
     calls.slowParts === 0, { slowParts: calls.slowParts });
  ok('D3 · it read the cache directly', calls.cacheGet > 0);
}

// ── E · cold cache on an inline publish degrades, never throws ───────────────────────
{
  const { res, tick } = run('pick', { cached: null });
  ok('E1 · publish still succeeds on a cold cache', !!(res && res.ok), res && res.message);
  ok('E2 · rest is null rather than undefined-shaped garbage',
     !!(tick && tick.rest === null), tick && tick.rest);
  ok('E3 · REGRESSION NET — sidebar object still emitted (client falls back on its emptiness)',
     !!(tick && tick.sidebar));
}

// ── F · over the cap: sidebar is shed first, rest is small enough to survive ─────────
{
  const { res, tick } = run('changed', { slow: SLOW, bulk: true });
  ok('F1 · an over-cap payload is still written, not refused', !!(res && res.ok), res && res.message);
  ok('F2 · it did trim', !!(tick && tick._trimmed));
  ok('F3 · sidebar is the first thing shed', !!(tick && !tick.sidebar), tick && !!tick.sidebar);
  ok('F4 · rest survives the sidebar trim', !!(tick && tick.rest), tick && tick.rest);
  ok('F5 · the pick list is still there — the board is why this cell exists',
     !!(tick && tick.openOrders && tick.openOrders.length > 0));
}

console.log('\n  ' + pass + ' passed, ' + fail + ' failed\n');
process.exit(fail ? 1 : 0);
