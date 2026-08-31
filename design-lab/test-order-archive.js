/**
 * test-order-archive.js — the PURE CORE of the order archive, loaded from the REAL
 * OrderArchive.js so the tests cannot drift from what ships.
 *
 * What is worth proving here:
 *   1. the two intervals are the RIGHT two (queue vs execution), incl. the awkward
 *      cases — printed before picked, canceled before anyone touched it
 *   2. the HONESTY rules: an order whose RECEIVED was purged is still written, with
 *      blanks rather than zeros; an unfinished order is not written at all
 *   3. the day key comes from the TERMINAL event, so an order that sat for weeks
 *      lands on the day it actually completed
 *
 * ⚠ EVERY SECTION FAILS SOFT. The choosePicker lesson: a harness that throws in
 *   section C tells you nothing about D–I, and a before/after proof is only useful
 *   if every section reports.
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');

const SRC  = process.env.SRC || path.join(__dirname, '..');
const CODE = fs.readFileSync(path.join(SRC, 'OrderArchive.js'), 'utf8');

const sandbox = {
  console,
  // ⚠⚠ HARNESS BUG, CAUGHT ON THE FIRST RUN — every section that rolled up events
  //    failed while the pure Date arithmetic passed. Total failure with a healthy
  //    control is this project's own signal to suspect the harness, and it was:
  //    vm.createContext creates a SEPARATE REALM with its OWN Date constructor, so a
  //    Date built out here is NOT `instanceof Date` in there. _oaRollUpOrders leans on
  //    that check everywhere, so it correctly saw zero usable timestamps and emitted
  //    nothing. Injecting the outer Date makes the two realms agree.
  //    (14th instance of a diagnostic accusing working code in this project.)
  Date,
  // Only the pure core is exercised; the sheet layer is never called.
  SpreadsheetApp: null, PropertiesService: null, Utilities: null,
  SPREADSHEET_ID: 'x', ACTIVITY_LOG: null
};
vm.createContext(sandbox);
vm.runInContext(CODE, sandbox, { filename: 'OrderArchive.js' });
const A = sandbox;

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label +
    (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};
const section = (name, fn) => {
  console.log('\n' + name);
  try { fn(); }
  catch (e) { fail++; console.log('  ✗ SECTION THREW (soft): ' + e.message); }
};

// ---- fixtures -------------------------------------------------------------
const dayOf = d => d.toISOString().slice(0, 10);   // deterministic, tz-free
const ev = (ts, event, orderId, extra) => Object.assign(
  { ts: new Date(ts), event, orderId, sku: '', qty: '', detail: '', note: '', picker: '' },
  extra || {});
const roll = evts => A._oaRollUpOrders(evts, dayOf);
const one  = evts => roll(evts)[0];
const find = (evts, id) => roll(evts).find(r => r.orderId === id);

// ===========================================================================
section('A · THE BASICS', () => {
  const r = one([
    ev('2026-08-20T14:00:00Z', 'RECEIVED',  '24-11111-22222', { qty: 2 }),
    ev('2026-08-20T15:00:00Z', 'PREPARING', '24-11111-22222'),
    ev('2026-08-20T16:00:00Z', 'SHIPPED',   '24-11111-22222')
  ]);
  t('order id survives',        r.orderId, '24-11111-22222');
  t('terminal status',          r.terminalStatus, 'SHIPPED');
  t('one RECEIVED = 1 line',    r.lines, 1);
  t('units from qty',           r.units, 2);
  t('queue = received→pick',    r.queueMin, 60);
  t('exec  = pick→shipped',     r.execMin, 60);
  t('total = received→shipped', r.totalMin, 120);

  const m = one([
    ev('2026-08-20T14:00:00Z', 'RECEIVED', 'SO-24609', { qty: 2 }),
    ev('2026-08-20T14:00:01Z', 'RECEIVED', 'SO-24609', { qty: 3 }),
    ev('2026-08-20T14:00:02Z', 'RECEIVED', 'SO-24609', { qty: 5 }),
    ev('2026-08-20T18:00:00Z', 'SHIPPED',  'SO-24609')
  ]);
  t('3 RECEIVED = 3 lines',     m.lines, 3);
  t('units SUMMED, not counted', m.units, 10);   // the falsy-zero / count-vs-sum trap
});

// ===========================================================================
section('B · THE TWO INTERVALS — the point of the whole thing', () => {
  // PRINTED before PREPARING: printing the list IS starting.
  const p = one([
    ev('2026-08-20T09:00:00Z', 'RECEIVED',  '24-33333-44444'),
    ev('2026-08-20T10:00:00Z', 'PRINTED',   '24-33333-44444'),
    ev('2026-08-20T11:00:00Z', 'PREPARING', '24-33333-44444'),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',   '24-33333-44444')
  ]);
  t('started = PRINTED when it came first', p.queueMin, 60);
  t('exec measured from PRINTED',           p.execMin, 120);

  // PREPARING before PRINTED — the other order.
  const q = one([
    ev('2026-08-20T09:00:00Z', 'RECEIVED',  'SO-100'),
    ev('2026-08-20T11:00:00Z', 'PRINTED',   'SO-100'),
    ev('2026-08-20T10:00:00Z', 'PREPARING', 'SO-100'),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',   'SO-100')
  ]);
  t('started = PREPARING when it came first', q.queueMin, 60);

  // ⚠ ADDED AFTER A MUTATION SURVIVED. Swapping _oaEarlier→_oaLater on firstPick
  //   passed all 54 assertions, because no fixture had TWO PREPARING events. Real
  //   orders always do: updateOrderStatus logs one transition PER ROW, so a 5-line
  //   order emits 5 PREPARING events. "Started" must be the FIRST of them — the last
  //   one is when picking FINISHED, which would silently collapse queue time and
  //   inflate execution time on exactly the multi-line orders worth studying.
  const multi = one([
    ev('2026-08-20T09:00:00Z', 'RECEIVED',  'SO-150'),
    ev('2026-08-20T09:00:01Z', 'RECEIVED',  'SO-150'),
    ev('2026-08-20T10:00:00Z', 'PREPARING', 'SO-150'),
    ev('2026-08-20T10:30:00Z', 'PREPARING', 'SO-150'),
    ev('2026-08-20T11:00:00Z', 'PREPARING', 'SO-150'),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',   'SO-150')
  ]);
  t('multi-line: started = FIRST pick', multi.queueMin, 60);
  t('multi-line: exec from FIRST pick', multi.execMin, 120);

  // Same rule for PRINTED — a reprint must not move the start.
  const rp = one([
    ev('2026-08-20T09:00:00Z', 'RECEIVED', 'SO-151'),
    ev('2026-08-20T10:00:00Z', 'PRINTED',  'SO-151'),
    ev('2026-08-20T11:30:00Z', 'PRINTED',  'SO-151'),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',  'SO-151')
  ]);
  t('a REPRINT does not move the start', rp.queueMin, 60);

  // Canceled without ever being touched: no execution phase existed.
  const c = one([
    ev('2026-08-20T09:00:00Z', 'RECEIVED', 'SO-200'),
    ev('2026-08-20T12:00:00Z', 'CANCELED', 'SO-200')
  ]);
  t('never started → queue spans to terminal', c.queueMin, 180);
  t('never started → exec is BLANK, not 0',    c.execMin, '');
  t('canceled is a terminal state',            c.terminalStatus, 'CANCELED');
});

// ===========================================================================
section('C · HONESTY — blanks, never reassuring zeros', () => {
  // RECEIVED already purged (order open >90 days). The row must still be written.
  const r = one([
    ev('2026-08-20T10:00:00Z', 'PREPARING', 'SO-300'),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',   'SO-300')
  ]);
  t('purged RECEIVED → still archived', r.orderId, 'SO-300');
  t('  received blank',                 r.received, '');
  t('  queue blank',                    r.queueMin, '');
  t('  total blank',                    r.totalMin, '');
  t('  lines blank, NOT 0',             r.lines, '');
  t('  units blank, NOT 0',             r.units, '');
  t('  exec still measurable',          r.execMin, 120);   // real net: what we DO know

  // Out-of-order events (backdated edit / revert / clock skew).
  const neg = one([
    ev('2026-08-20T12:00:00Z', 'RECEIVED', 'SO-400'),
    ev('2026-08-20T10:00:00Z', 'SHIPPED',  'SO-400')
  ]);
  t('negative gap → blank, not a negative', neg.totalMin, '');

  // An order still being worked has no cycle time yet.
  t('no terminal → NOT archived', roll([
    ev('2026-08-20T10:00:00Z', 'RECEIVED',  'SO-500'),
    ev('2026-08-20T11:00:00Z', 'PREPARING', 'SO-500')
  ]).length, 0);

  t('orderless rows ignored', roll([
    ev('2026-08-20T10:00:00Z', 'NOTE',    ''),
    ev('2026-08-20T11:00:00Z', 'SHIPPED', '')
  ]).length, 0);
});

// ===========================================================================
section('D · FLAGS', () => {
  const h = one([
    ev('2026-08-20T09:00:00Z', 'RECEIVED', 'SO-600'),
    ev('2026-08-20T10:00:00Z', 'NOTE',     'SO-600', { note: 'HOLD — buyer wants 2-Day' }),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',  'SO-600')
  ]);
  t('hold detected from NOTE', h.held, 'yes');

  const hh = one([
    ev('2026-08-20T09:00:00Z', 'RECEIVED', 'SO-601'),
    ev('2026-08-20T10:00:00Z', 'NOTE',     'SO-601', { note: 'household goods, withhold nothing' }),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',  'SO-601')
  ]);
  t('"household"/"withhold" do NOT fire', hh.held, '');   // word-boundary net

  const k = one([
    ev('2026-08-20T09:00:00Z', 'RECEIVED', 'SO-602', { detail: 'kit expansion from 158652' }),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',  'SO-602')
  ]);
  t('kit detected from RECEIVED detail', k.kit, 'yes');

  const nk = one([
    ev('2026-08-20T09:00:00Z', 'RECEIVED', 'SO-603', { detail: 'DIRECT manual' }),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',  'SO-603')
  ]);
  t('ordinary order is not flagged as a kit', nk.kit, '');
});

// ===========================================================================
section('E · CHANNEL — id shape only', () => {
  const ch = id => one([
    ev('2026-08-20T09:00:00Z', 'RECEIVED', id),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',  id)
  ]).channel;
  t('eBay digit-dash id',        ch('24-14979-87359'), 'eBay');
  t('Zoho SO-',                  ch('SO-24609'),       'DIRECT');
  t('Zoho INV-',                 ch('INV-022496'),     'DIRECT');
  t('manual replacement text',   ch('Replacement #: 19-14597-26309'), 'DIRECT');
});

// ===========================================================================
section('F · PICKER', () => {
  const r = one([
    ev('2026-08-20T09:00:00Z', 'RECEIVED',  'SO-700'),
    ev('2026-08-20T10:00:00Z', 'PREPARING', 'SO-700', { picker: 'Yassin · 1' }),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',   'SO-700', { picker: 'Hatem · 2' })
  ]);
  t('picker at TERMINAL wins', r.picker, 'Hatem · 2');

  const f = one([
    ev('2026-08-20T10:00:00Z', 'PREPARING', 'SO-701', { picker: 'Yassin · 1' }),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',   'SO-701')     // n8n sweep leaves it blank
  ]);
  t('falls back to any known picker', f.picker, 'Yassin · 1');

  const n = one([
    ev('2026-08-20T10:00:00Z', 'RECEIVED', 'SO-702'),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',  'SO-702')
  ]);
  t('unattributed stays blank', n.picker, '');
});

// ===========================================================================
section('G · DAY + TERMINAL SELECTION', () => {
  // Received in June, shipped in August — it belongs to the day it COMPLETED.
  const s = one([
    ev('2026-06-01T09:00:00Z', 'RECEIVED', 'SO-800'),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',  'SO-800')
  ]);
  t('day comes from the TERMINAL event', s.day, '2026-08-20');

  // Shipped, reverted by the verify sweep, shipped again → the LATEST terminal wins.
  const rv = one([
    ev('2026-08-20T09:00:00Z', 'RECEIVED', 'SO-801'),
    ev('2026-08-20T10:00:00Z', 'SHIPPED',  'SO-801'),
    ev('2026-08-21T11:00:00Z', 'SHIPPED',  'SO-801')
  ]);
  t('latest terminal wins',  rv.day, '2026-08-21');
  t('  and sets the total',  rv.totalMin, 26 * 60);

  const two = roll([
    ev('2026-08-20T09:00:00Z', 'RECEIVED', 'SO-900'),
    ev('2026-08-20T10:00:00Z', 'RECEIVED', 'SO-901'),
    ev('2026-08-20T11:00:00Z', 'SHIPPED',  'SO-900'),
    ev('2026-08-20T12:00:00Z', 'SHIPPED',  'SO-901')
  ]);
  t('two orders, two rows',        two.length, 2);
  t('deterministic first-seen order', two.map(r => r.orderId), ['SO-900', 'SO-901']);
});

// ===========================================================================
section('H · DAY ARITHMETIC (_oaDayAdd)', () => {
  t('plus one day',        A._oaDayAdd('2026-08-20',  1), '2026-08-21');
  t('minus one day',       A._oaDayAdd('2026-08-20', -1), '2026-08-19');
  t('month boundary',      A._oaDayAdd('2026-08-31',  1), '2026-09-01');
  t('year boundary',       A._oaDayAdd('2026-12-31',  1), '2027-01-01');
  t('back over a month',   A._oaDayAdd('2026-09-01', -1), '2026-08-31');
  t('leap day 2028',       A._oaDayAdd('2028-02-28',  1), '2028-02-29');
  // ⚠ US DST springs forward 2026-03-08. The noon-UTC anchor must not hop a day.
  t('across US spring DST', A._oaDayAdd('2026-03-07',  1), '2026-03-08');
  t('across US autumn DST', A._oaDayAdd('2026-11-01',  1), '2026-11-02');
  t('45-day chunk span',   A._oaDayAdd('2026-06-01', 44), '2026-07-15');
});

// ===========================================================================
section('I · THE REAL INCIDENT — a documented straggler beats an invented fixture', () => {
  // 24-14979-87359 sat PENDING from 9:01 AM and was only noticed at 17:06.
  const r = one([
    ev('2026-08-05T09:01:00Z', 'RECEIVED',  '24-14979-87359', { qty: 1 }),
    ev('2026-08-05T17:06:00Z', 'PREPARING', '24-14979-87359'),
    ev('2026-08-05T17:30:00Z', 'SHIPPED',   '24-14979-87359')
  ]);
  t('the 8-hour wait is QUEUE time', r.queueMin, 485);
  t('the work itself was 24 min',    r.execMin, 24);
  t('total 509 min',                 r.totalMin, 509);
  t('and it is an eBay order',       r.channel, 'eBay');
});

console.log('\n' + '='.repeat(58));
console.log(pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
