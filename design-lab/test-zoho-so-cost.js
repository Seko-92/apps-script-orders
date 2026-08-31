/*
 * test-zoho-so-cost.js — what the Zoho SO webhook actually COSTS.
 * ---------------------------------------------------------------------------
 * Loads the REAL ZohoSalesOrders.js + the REAL Config.js + Schema.js in a VM.
 *
 * THE BUG THIS PINS (2026-08-25). The n8n Zoho SO proxy kept failing with
 * "The connection was aborted, perhaps the server is offline" while sales
 * orders kept mirroring correctly — the signature of work that COMPLETES after
 * the client has given up on the socket. Cause: upsertPendingSalesOrder called
 * _computePriceCheck on every fire, inside doPost's script lock, and that did
 * TWO full-width Master Inventory reads (~1.4M cells) — one of which was DEAD
 * (`maps` was assigned and never read).
 *
 * So most assertions here are about COST, not output:
 *   · the dead buildLocationAndInventoryMaps() call must never happen
 *   · a terminal SO (closed / void / shipped) must not touch MI at all
 *   · …and must not LOSE the flag it already earned
 *
 * ⚠ THE COUNTER IS FAITHFUL ON PURPOSE. The stubbed buildLocationAndInventoryMaps
 *   increments miBulkReads too, because in production it IS a full-width MI read
 *   (getDataRange, ~720,000 cells). A stub that read nothing would let the
 *   "one MI read, not two" assertion pass against HEAD and prove nothing.
 *
 * ⚠ EVERY SECTION MUST REPORT even against a build that lacks the new function,
 *   or the before/after proof only tells you about section A. First run of this
 *   test threw on HEAD at `moot is not a function` and aborted — the
 *   `choosePicker` lesson, again.
 *
 *   node test-zoho-so-cost.js
 *   SRC=/tmp/…/headzoho node test-zoho-so-cost.js    # prove it bites
 */
const fs   = require('fs');
const path = require('path');
const vm   = require('vm');

const ROOT = process.env.SRC || path.join(__dirname, '..');
let pass = 0, fail = 0;
function t(name, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? (pass++, console.log('  ✓ ' + name))
     : (fail++, console.log('  ✗ ' + name + '  → got ' + JSON.stringify(got) +
                            ', want ' + JSON.stringify(want)));
}
/* Fail SOFT: one section blowing up must not silence the ones after it. */
function section(label, body) {
  console.log('\n' + label);
  try { body(); }
  catch (e) { fail++; console.log('  ✗ SECTION ABORTED → ' + e.message); }
}

/* ------------------------------------------------------------------ fakes */
const MI_HEADERS = ['itemId', 'sku', 'currentPrice', 'startPrice', 'quantity'];
const MI_ROWS = [
  ['1', '111111', 100, 100, 5],
  ['2', '222222',  50,  50, 5]
];
function pendingHeaderRow() { return ['SO #','CUSTOMER','DATE','ORDER','PAYMENT','SHIPMENT',
                                      'ITEMS','TOTAL','UPDATED','PULLED?','PULLED AT',
                                      '_PAYLOAD','INVOICE #','PRICE CHECK']; }

function makeSheet(name, rows, counters) {
  return {
    getName: () => name,
    getLastRow: () => rows.length,
    getLastColumn: () => (rows[0] ? rows[0].length : 0),
    insertRowsBefore: function (before, howMany) {
      for (let i = 0; i < howMany; i++) rows.splice(before - 1, 0, new Array(14).fill(''));
      return this;
    },
    getRange: function (row, col, nR, nC) {
      nR = nR || 1; nC = nC || 1;
      // Count only the BULK data read on Master Inventory — the ~720,000-cell
      // round trip we are trying to stop paying for.
      if (name === 'Master Inventory' && nR > 1) counters.miBulkReads++;
      return {
        getValues: () => {
          const out = [];
          for (let r = 0; r < nR; r++) {
            const line = [];
            for (let c = 0; c < nC; c++) {
              const src = rows[row - 1 + r] || [];
              line.push(src[col - 1 + c] === undefined ? '' : src[col - 1 + c]);
            }
            out.push(line);
          }
          return out;
        },
        getValue: () => {
          const src = rows[row - 1] || [];
          return src[col - 1] === undefined ? '' : src[col - 1];
        },
        setValue: function (v) {
          while (rows.length < row) rows.push(new Array(14).fill(''));
          rows[row - 1][col - 1] = v;
          return this;
        },
        setValues: function (v) {
          for (let r = 0; r < v.length; r++) {
            while (rows.length < row + r) rows.push(new Array(14).fill(''));
            for (let c = 0; c < v[r].length; c++) rows[row - 1 + r][col - 1 + c] = v[r][c];
          }
          return this;
        }
      };
    }
  };
}

function build(opts) {
  opts = opts || {};
  const counters = { miBulkReads: 0, mapBuilderCalls: 0, statusFlips: 0 };
  const pendingRows = opts.pendingRows || [pendingHeaderRow()];
  const miRows = [MI_HEADERS].concat(MI_ROWS);

  const pendingSheet = makeSheet('Pending Sales Orders', pendingRows, counters);
  const miSheet      = makeSheet('Master Inventory',     miRows,      counters);

  // In production this is getDataRange() on MI — so it counts as a bulk read.
  function mapBuilderStub() {
    counters.mapBuilderCalls++;
    counters.miBulkReads++;
    return { locationMap: new Map(), inventoryMap: new Map() };
  }

  const sandbox = {
    console: { log: () => {} },
    Logger:  { log: () => {} },
    SpreadsheetApp: {
      openById: () => ({
        getSheetByName: (n) => {
          if (n === 'Master Inventory') return opts.miAvailable === false ? null : miSheet;
          if (n === 'Pending Sales Orders') return pendingSheet;
          return null;
        }
      })
    },
    buildLocationAndInventoryMaps: mapBuilderStub,
    updateOrderStatus: function () { counters.statusFlips++; return { count: 1 }; },
    setupPendingSalesOrdersSheet: function () {},
    logActivity: function () {},
    logActivityBatch: function () {},
    getBoundaryRow: function () { return -1; }
  };
  sandbox.global = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'Config.js'), 'utf8'), sandbox, { filename: 'Config.js' });
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'Schema.js'), 'utf8'), sandbox, { filename: 'Schema.js' });
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'ZohoSalesOrders.js'), 'utf8'), sandbox,
                  { filename: 'ZohoSalesOrders.js' });

  /* ⚠ Re-apply AFTER loading. ZohoSalesOrders.js does not define this today, but
     the sibling harness lost eleven assertions to exactly this ordering when
     LiveSync.js did. Cheap insurance; suspect the harness first. */
  sandbox.buildLocationAndInventoryMaps = mapBuilderStub;

  return { sandbox, counters, pendingRows };
}

function so(over) {
  return Object.assign({
    salesorder_number: 'SO-99001',
    customer_name: 'Test Customer',
    sales_channel: 'direct_sales',
    status: 'open',
    paid_status: 'paid',
    shipped_status: 'pending',
    total_formatted: '$150.00',
    line_items: [
      { sku: '111111', quantity: 1, rate: 110, name: 'Part A' },  // eBay 100 → HIGH
      { sku: '222222', quantity: 1, rate: 50,  name: 'Part B' }   // eBay 50  → ok
    ]
  }, over || {});
}

const N = 13;                                     // PRICE_CHECK, 0-based
const LIVE_SUMMARY = '⚠ Zoho HIGH · 1/2 · +$10.00';

/* ============================================================ A · the gate */
section('A · _priceCheckIsMoot — when is the check worth paying for?', function () {
  const { sandbox } = build();
  const moot = sandbox._priceCheckIsMoot ||
               function () { return 'MISSING — function not in this build'; };
  t('open is live',                     moot(so({ status: 'open' })), false);
  t('draft is live',                    moot(so({ status: 'draft' })), false);
  t('closed is moot',                   moot(so({ status: 'closed' })), true);
  t('void is moot',                     moot(so({ status: 'void' })), true);
  t('shipped_status shipped is moot',   moot(so({ shipped_status: 'shipped' })), true);
  t('shipped_status fulfilled is moot', moot(so({ shipped_status: 'fulfilled' })), true);
  t('partially_shipped stays LIVE',     moot(so({ shipped_status: 'partially_shipped' })), false);
  t('case + whitespace tolerated',      moot(so({ status: '  Closed ' })), true);
  t('null payload is moot',             moot(null), true);
});

/* ====================================================== B · the dead read */
section('B · the dead Master Inventory read', function () {
  const { sandbox, counters } = build();
  const res = sandbox._computePriceCheck(so());
  t('buildLocationAndInventoryMaps NEVER called', counters.mapBuilderCalls, 0);
  t('MI read exactly ONCE, not twice',            counters.miBulkReads, 1);
  t('summary unchanged for a live SO  [net]',     res.summary, LIVE_SUMMARY);
  t('direction unchanged  [net]',                 res.direction, 'HIGH');
});

/* ================================================= C · terminal SOs skip */
section('C · a terminal SO must not touch Master Inventory at all', function () {
  ['closed', 'void'].forEach(function (st) {
    const { sandbox, counters } = build();
    sandbox.upsertPendingSalesOrder(so({ status: st }));
    t(st + ': zero MI bulk reads',     counters.miBulkReads, 0);
    t(st + ': zero map-builder calls', counters.mapBuilderCalls, 0);
  });
  const { sandbox, counters } = build();
  sandbox.upsertPendingSalesOrder(so({ shipped_status: 'shipped' }));
  t('shipped: zero MI bulk reads',            counters.miBulkReads, 0);
  t('shipped: status still propagates  [net]', counters.statusFlips, 1);
});

/* ================================== D · the flag must survive going terminal */
section('D · a terminal SO keeps the flag it already earned', function () {
  const existing = new Array(14).fill('');
  existing[0]  = 'SO-99001';
  existing[9]  = 'PULLED';
  existing[10] = 'yesterday';
  existing[12] = 'INV-000123';
  // ⚠ Deliberately a value a recompute could NEVER produce from this fixture.
  //   With the live summary here, HEAD would overwrite it with an identical
  //   string and the assertion would pass by coincidence, proving nothing.
  existing[13] = '⚠ Zoho LOW · 2/2 · -$40.00';
  const { sandbox, pendingRows, counters } =
    build({ pendingRows: [pendingHeaderRow(), existing] });

  sandbox.upsertPendingSalesOrder(so({ status: 'closed' }));
  t('PRICE_CHECK preserved, not recomputed', pendingRows[1][N], '⚠ Zoho LOW · 2/2 · -$40.00');
  t('and it cost no MI read',                counters.miBulkReads, 0);
  t('PULLED still preserved  [net]',         pendingRows[1][9],  'PULLED');
  t('PULLED_AT still preserved  [net]',      pendingRows[1][10], 'yesterday');
  t('INVOICE still preserved  [net]',        pendingRows[1][12], 'INV-000123');
});

/* ============================== E · a live SO still recomputes */
section('E · a live SO still recomputes the flag', function () {
  const existing = new Array(14).fill('');
  existing[0]  = 'SO-99001';
  existing[13] = 'STALE VALUE';
  const { sandbox, pendingRows, counters } =
    build({ pendingRows: [pendingHeaderRow(), existing] });

  sandbox.upsertPendingSalesOrder(so({ status: 'open' }));
  t('stale flag overwritten  [net]', pendingRows[1][N], LIVE_SUMMARY);
  t('one MI read, not two',          counters.miBulkReads, 1);
});

/* ====================== F · unreadable MI must SAY so */
section('F · MI unreadable is not the same as "not found"', function () {
  const { sandbox } = build({ miAvailable: false });
  const res = sandbox._computePriceCheck(so());
  // ⚠ HEAD reports "⚠ NOT FOUND · 2/2" here — a clean-sounding answer ABOUT THE
  //   SO for what is actually a broken MI. HEAD's own "— MI unavailable" branch
  //   was near-unreachable: it fired only if the map builder THREW, and the real
  //   builder returns empty maps for a missing sheet rather than throwing.
  t('says MI unavailable', res.summary, '— MI unavailable');
  t('not a clean answer about the SO', res.direction, 'EMPTY');
});

/* ========================= G · a new row for a terminal SO */
section('G · brand-new row for an already-terminal SO', function () {
  const { sandbox, pendingRows, counters } = build();
  sandbox.upsertPendingSalesOrder(so({ status: 'closed' }));
  t('row was inserted  [net]',               String(pendingRows[1][0]), 'SO-99001');
  t('flag left blank — we never checked it', pendingRows[1][N], '');
  t('and it cost nothing',                   counters.miBulkReads, 0);
});

console.log('\n' + (fail ? '✗ ' : '✓ ') + pass + ' passed, ' + fail + ' failed\n');
process.exit(fail ? 1 : 0);
