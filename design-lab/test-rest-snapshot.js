// =====================================================================================
// test-rest-snapshot.js — the SERVER half of the resting panel's two dark rows.
//
// The CLIENT half is already proven by shoot-rest-panel.js case B, which asserts the
// exact wording and "117 units" against a hand-fed tick. This proves the other end:
// that the tick's `rest` key is produced correctly, cheaply, and once an hour.
//
// Loads the REAL KitHealth.js / Ripple.js / UIService.js into a VM with stubs, so the
// test cannot drift from the shipped code.
//
// Prove against HEAD:  mkdir -p /tmp/headrs && for f in KitHealth.js Ripple.js UIService.js; do \
//                        git show HEAD:$f > /tmp/headrs/$f; done && SRC=/tmp/headrs node test-rest-snapshot.js
// =====================================================================================
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');

let pass = 0, fail = 0;
const ok = (n, c, x) => { c ? pass++ : (fail++, console.log('  ✗ ' + n + (x !== undefined ? '  → ' + JSON.stringify(x) : ''))); };

// ── the Apps Script surface these functions actually touch ──────────────────────────
function makeCtx(opts) {
  opts = opts || {};
  const props = opts.props || {};
  const calls = { getRange: [], built: { mi: 0, zoho: 0, kit: 0 } };

  const sheet = opts.rows === null ? null : {
    getLastRow: () => (opts.lastRow !== undefined ? opts.lastRow : 2 + (opts.rows || []).length),
    getRange: (r, c, n, w) => { calls.getRange.push({ r, c, n, w });
      return { getValues: () => (opts.rows || []).slice(0, n) }; }
  };

  const ctx = {
    console: { log: () => {}, error: () => {} },
    Map, JSON, String, Number, Date, parseInt, parseFloat, isNaN, Math, Array, Object, RegExp,
    SPREADSHEET_ID: 'x',
    SpreadsheetApp: { getActive: () => ({ getSheetByName: () => sheet }), openById: () => ({ getSheetByName: () => sheet }) },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => (k in props ? props[k] : null),
      setProperty: (k, v) => { props[k] = v; }
    }) },
    CacheService: { getScriptCache: () => null },
    // cross-file collaborators
    buildLocationAndInventoryMaps: () => { calls.built.mi++;   return { inventoryMap: new Map(), locationMap: new Map() }; },
    buildZohoStockMap:             () => { calls.built.zoho++; return new Map(); },
    buildKitMap:                   () => { calls.built.kit++;  return new Map(); },
    _oosResolveAvailFactory: () => () => 0,
    _oosComputeKitBuild: () => ({ buildable: 5, limitedBy: '' }),
    getLatestApiMetrics: () => null,
    getActionableAlerts: () => null,
    getDashboardSnapshot: () => null,
    getLastSyncFromSheet: () => '',
    getCurrentPicker: () => ''
  };
  ctx.calls = calls; ctx.props = props;
  vm.createContext(ctx);
  for (const f of ['KitHealth.js', 'Ripple.js', 'UIService.js']) {
    try { vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f }); }
    catch (e) { console.log('  ! could not load ' + f + ': ' + e.message); }
  }
  return ctx;
}

const S = makeCtx({}).KIT_HEALTH.stock;
const HEAD = (n) => ['(missing)', n];   // fail-soft label

// ═══ 1b · THE 2026-08-21 PRODUCTION BUG — a stale DATE format on column Q ════════════
//
// Live symptom: "30 can't build (0 units)". The values were never missing — AT_RISK
// was inserted at column 17, exactly where the old 17-wide schema had LAST_CHECKED,
// so it INHERITED a date number format (Sheets keeps formats through clearContent).
// getValues() therefore returned Date OBJECTS, parseFloat(Date) was NaN, and a
// customer-facing exposure was silently reported as zero.
// Sheets' epoch is 1899-12-30 → 1899-12-31 IS 1, and 1900-01-01 IS 2.
console.log('\n── a stale date format must not silently score zero');
{
  const rows = [
    [S.CANT_BUILD, new Date(1899, 11, 31)],   // serial 1 — verbatim from the live log
    [S.CANT_BUILD, new Date(1900, 0, 1)],     // serial 2 — verbatim from the live log
    [S.CANT_BUILD, 4]                         // a plain number still reads normally
  ];
  const c = makeCtx({ rows });
  const g = c.getKitOversellSnapshot ? c.getKitOversellSnapshot() : { kits: -1, units: -1 };
  ok('the kits still count', g.kits === 3, g.kits);
  ok('a Date is recovered to its serial → 1 + 2 + 4 = 7', g.units === 7, g.units);
  ok('and nothing is reported unreadable once recovered', g.unreadable === 0, g.unreadable);
}
{
  // The other half of the rule: when it genuinely CANNOT be read, say so rather
  // than scoring zero. A reassuring number on a dangerous state is a bug here.
  const c = makeCtx({ rows: [[S.CANT_BUILD, 'n/a'], [S.CANT_BUILD, '']] });
  const g = c.getKitOversellSnapshot ? c.getKitOversellSnapshot() : { unreadable: -1 };
  ok('an unreadable AT RISK is COUNTED, never silently zero', g.unreadable === 2, g.unreadable);
}

// ═══ 1 · getKitOversellSnapshot — the cheap sheet read ═══════════════════════════════
console.log('\n── the number the row actually claims');
{
  // P = STOCK_STATUS, Q = AT_RISK, in that order (the 2-wide read)
  const rows = [
    [S.CANT_BUILD,  3],
    [S.OVER_LISTED, 9],          // ⚠ advertised beyond capacity but PARTLY covered
    [S.CANT_BUILD,  1],
    [S.STOCK_BUILD, ''],
    [S.OOS,         ''],
    [S.CANT_BUILD,  ''],         // flagged but no size recorded
    [S.UNKNOWN,     '']
  ];
  const c = makeCtx({ rows });
  const g = c.getKitOversellSnapshot ? c.getKitOversellSnapshot() : { kits: '(missing)', units: '(missing)' };
  ok("counts CAN'T BUILD only — 3 of them", g.kits === 3, g.kits);
  ok('⚠ OVER-LISTED is NOT counted (it still ships)', g.kits !== 4, g.kits);
  ok('sums AT RISK across them → 4 units', g.units === 4, g.units);
  ok("a blank AT RISK doesn't poison the sum", !isNaN(g.units), g.units);

  const r = c.calls.getRange[0];
  ok('starts at STOCK_STATUS (P/16)', r && r.c === 16, r && r.c);
  ok('ONE read, exactly 2 columns wide (P..Q)', r && r.w === 2 && c.calls.getRange.length === 1,
     { w: r && r.w, reads: c.calls.getRange.length });
  ok('starts at the first DATA row, not the header', r && r.r === 3, r && r.r);
}

console.log('\n── it never invents a number');
{
  const c1 = makeCtx({ rows: null });                       // no Kit Health sheet at all
  const g1 = c1.getKitOversellSnapshot ? c1.getKitOversellSnapshot() : null;
  ok('no sheet → zeros, not a throw', g1 && g1.kits === 0 && g1.units === 0, g1);

  const c2 = makeCtx({ rows: [], lastRow: 2 });             // audited never / empty
  const g2 = c2.getKitOversellSnapshot ? c2.getKitOversellSnapshot() : null;
  ok('never audited → zeros', g2 && g2.kits === 0 && g2.units === 0, g2);
  ok('...and it does not read a range at all', c2.calls.getRange.length === 0, c2.calls.getRange.length);

  const c3 = makeCtx({ rows: [[S.CANT_BUILD, 2]] });
  const g3 = c3.getKitOversellSnapshot ? c3.getKitOversellSnapshot() : null;
  ok('a single blocked kit reports itself', g3 && g3.kits === 1 && g3.units === 2, g3);
}

// ═══ 2 · analyzeRestockRipple — the ~6s function learns to share ═════════════════════
console.log('\n── the expensive maps are handed over, not rebuilt');
{
  const c = makeCtx({});
  if (!c.analyzeRestockRipple) { ok('analyzeRestockRipple present', false, HEAD('analyzeRestockRipple')); }
  else {
    const maps = { inventoryMap: new Map(), locationMap: new Map() };
    c.analyzeRestockRipple(maps, new Map());
    ok('given both, it builds NEITHER map itself', c.calls.built.mi === 0 && c.calls.built.zoho === 0, c.calls.built);
    ok('...but still reads the kit map (not shareable)', c.calls.built.kit === 1, c.calls.built.kit);

    const c2 = makeCtx({});
    c2.analyzeRestockRipple();
    ok('given nothing, it builds them — unchanged for every old caller',
       c2.calls.built.mi === 1 && c2.calls.built.zoho === 1, c2.calls.built);

    // ⚠⚠ THE REGRESSION NET. A time trigger hands its target an EVENT OBJECT as
    // the first argument. Positional trust would silently use it as `maps`.
    const c3 = makeCtx({});
    c3.analyzeRestockRipple({ triggerUid: '99', authMode: 'FULL' }, undefined);
    ok('an EVENT OBJECT is refused and the maps get built', c3.calls.built.mi === 1, c3.calls.built);

    const c4 = makeCtx({});
    c4.analyzeRestockRipple(maps);                     // maps but no zoho
    ok('half-shared: takes the MI maps, builds only Zoho',
       c4.calls.built.mi === 0 && c4.calls.built.zoho === 1, c4.calls.built);

    const c5 = makeCtx({});
    c5.analyzeRestockRipple({ inventoryMap: new Map() });   // missing locationMap
    ok('a HALF-BUILT maps object is refused', c5.calls.built.mi === 1, c5.calls.built);
  }
}

// ═══ 3 · the parked blob ════════════════════════════════════════════════════════════
console.log('\n── the snapshot is written once an hour and read from there');
{
  const c = makeCtx({ rows: [[S.CANT_BUILD, 117]] });
  if (!c.refreshRestSnapshot) { ok('refreshRestSnapshot present', false, HEAD('refreshRestSnapshot')); }
  else {
    c.getRippleTop = (n) => [{ sku: '195306', sole: 7, name: 'Thrust Washer Set', location: 'E-12',
                               avail: 0, shortIn: 9 }];
    const msg = c.refreshRestSnapshot({ inventoryMap: new Map(), locationMap: new Map() }, new Map());
    const snap = JSON.parse(c.props.HQ_REST_SNAPSHOT);

    ok('cantBuild lands in the blob', snap.cantBuild.units === 117, snap.cantBuild);
    ok('ripple lands in the blob', snap.ripple.length === 1 && snap.ripple[0].sole === 7, snap.ripple);
    ok('⚠ ONLY {sku,sole} is stored — a Properties value is capped at 9KB',
       Object.keys(snap.ripple[0]).sort().join(',') === 'sku,sole', Object.keys(snap.ripple[0]));
    ok('it stamps when it was built', typeof snap._builtAt === 'number' && snap._builtAt > 0, snap._builtAt);
    ok('and it reports itself to the housekeeping summary', /can't build/.test(msg), msg);

    const back = c._restSnapshot();
    ok('reads back what it wrote', back && back.cantBuild.units === 117, back && back.cantBuild);
  }
}

console.log('\n── a stale figure is never asserted as current');
{
  const fresh = { cantBuild: { kits: 1, units: 2 }, ripple: [], _builtAt: Date.now() - 3600e3 };
  const c1 = makeCtx({ props: { HQ_REST_SNAPSHOT: JSON.stringify(fresh) } });
  ok('an hour old → served (the pass runs hourly)', c1._restSnapshot && c1._restSnapshot() !== null);

  const old = { cantBuild: { kits: 1, units: 2 }, ripple: [], _builtAt: Date.now() - 25 * 3600e3 };
  const c2 = makeCtx({ props: { HQ_REST_SNAPSHOT: JSON.stringify(old) } });
  ok('over a day old → null, so the row draws nothing', c2._restSnapshot && c2._restSnapshot() === null,
     c2._restSnapshot && c2._restSnapshot());

  const c3 = makeCtx({ props: {} });
  ok('never written → null', c3._restSnapshot && c3._restSnapshot() === null);

  const c4 = makeCtx({ props: { HQ_REST_SNAPSHOT: '{not json' } });
  ok('corrupt → null, never a throw into the tick', c4._restSnapshot && c4._restSnapshot() === null);

  const c5 = makeCtx({ props: { HQ_REST_SNAPSHOT: JSON.stringify({ cantBuild: { kits: 1, units: 2 }, ripple: [] }) } });
  ok('no _builtAt at all → treated as ancient, not as fresh',
     c5._restSnapshot && c5._restSnapshot() === null, c5._restSnapshot && c5._restSnapshot());
}

console.log('\n── one half failing must not cost the other');
{
  const c = makeCtx({ rows: [[S.CANT_BUILD, 42]] });
  if (c.refreshRestSnapshot) {
    c.getRippleTop = () => { throw new Error('ripple exploded'); };
    c.refreshRestSnapshot();
    const snap = JSON.parse(c.props.HQ_REST_SNAPSHOT);
    ok('ripple throws → cantBuild still lands', snap.cantBuild.units === 42, snap.cantBuild);
    ok('...and ripple degrades to empty, not missing', Array.isArray(snap.ripple) && snap.ripple.length === 0, snap.ripple);

    const c2 = makeCtx({ rows: null });
    c2.getRippleTop = () => [{ sku: 'A', sole: 3 }];
    c2.refreshRestSnapshot();
    const s2 = JSON.parse(c2.props.HQ_REST_SNAPSHOT);
    ok('no Kit Health sheet → ripple still lands', s2.ripple.length === 1, s2.ripple);
    ok('...and cantBuild is zero, so the row stays silent', s2.cantBuild.units === 0, s2.cantBuild);
  } else ok('refreshRestSnapshot present', false, HEAD('refreshRestSnapshot'));
}

// ═══ 4 · the tick carries it ════════════════════════════════════════════════════════
console.log('\n── the sidebar tick carries `rest`, on the SLOW clock');
{
  const c = makeCtx({ props: { HQ_REST_SNAPSHOT: JSON.stringify(
    { cantBuild: { kits: 26, units: 117 }, ripple: [{ sku: '195306', sole: 7 }], _builtAt: Date.now() }) } });
  if (!c._sidebarSlowParts) ok('_sidebarSlowParts present', false, HEAD('_sidebarSlowParts'));
  else {
    const slow = c._sidebarSlowParts();
    ok('the slow half carries it', slow.rest && slow.rest.cantBuild.units === 117, slow.rest);
    const tick = c.getSidebarTick(true);
    ok('and getSidebarTick hands it to the panel', !!(tick.rest && tick.rest.cantBuild.units === 117), tick.rest);
    // ⚠ FAIL SOFT. On an older UIService.js there is no `rest` key at all, and a
    // throw here would abort the run before this section could report — the point
    // of a before/after proof is seeing WHICH parts changed and which did not.
    const rip = (tick.rest && tick.rest.ripple) || [];
    ok('shape is exactly what the panel reads',
       !!(rip[0] && rip[0].sku === '195306' && rip[0].sole === 7), rip);
  }
}

console.log('\n' + (fail ? `✗ ${fail} failed, ${pass} passed` : `✓ all ${pass} passed`));
process.exit(fail ? 1 : 0);
