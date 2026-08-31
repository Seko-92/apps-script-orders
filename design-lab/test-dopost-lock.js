/**
 * test-dopost-lock.js — the doPost lock decision, from the REAL OrderService.js.
 *
 * WHY THIS EXISTS
 * ---------------
 * eBay is connected to Zoho, so every eBay order becomes a Zoho Sales Order and
 * fires the SO webhook. upsertPendingSalesOrder discards them — but the filter sits
 * INSIDE it, i.e. AFTER waitLock(30000), so each one took a place in the lock queue
 * before discovering it had nothing to do. Measured 2026-08-28: bursts of 4–12+ in
 * the same second, all different SO numbers, all `ebay_us`, all skipped.
 *
 * `_doPostNeedsLock` moves that decision before the lock. What is worth proving is
 * not the saving — it is the four things the shortcut must NOT break, because this
 * function guards the resource the floor's ✓ Pick contends for.
 *
 * ⚠ Sections fail soft (the choosePicker lesson): a throw in one must not hide the rest.
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');

const SRC  = process.env.SRC || path.join(__dirname, '..');
const CODE = fs.readFileSync(path.join(SRC, 'OrderService.js'), 'utf8');

// OrderService.js is large and references globals from sibling files. Only the
// top-level DECLARATIONS run on load, so a permissive sandbox is enough — and a
// Proxy is better than enumerating stubs, because an unstubbed name would
// otherwise throw and look like a product failure (13 prior instances of that).
const sandbox = new Proxy({ console, JSON, String, Number, Boolean, Object, Array, Date, Math, RegExp, parseInt, parseFloat, isNaN },
  { has: () => true, get: (t, k) => (k in t ? t[k] : undefined) });
vm.createContext(sandbox);
vm.runInContext(CODE, sandbox, { filename: 'OrderService.js' });
const needsLock = sandbox._doPostNeedsLock;

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = got === want;
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label + (ok ? '' : `  → got ${got}, want ${want}`));
};
const section = (name, fn) => {
  console.log('\n' + name);
  try { fn(); } catch (e) { fail++; console.log('  ✗ SECTION THREW (soft): ' + e.message); }
};

const so = (channel) => ({ action: 'zohoSalesOrder', salesorder: { salesorder_number: 'SO-25223', sales_channel: channel } });

section('A · THE SAVING — what this change is for', () => {
  // Verbatim from the live payloads captured 2026-08-28.
  t('ebay_us SO skips the lock',      needsLock('zohoSalesOrder', so('ebay_us')), false);
  t('any non-direct channel skips',   needsLock('zohoSalesOrder', so('amazon_us')), false);
  t('case and padding tolerated',     needsLock('zohoSalesOrder', so('  EBAY_US  ')), false);
});

section('B · WHAT IT MUST NOT BREAK — the four guards', () => {
  // 1. a real direct sale does the work, so it must still serialise
  t('direct_sales STILL locks',       needsLock('zohoSalesOrder', so('direct_sales')), true);
  t('  and tolerates case',           needsLock('zohoSalesOrder', so('Direct_Sales')), true);

  // 2. blank channel is FAIL-OPEN — mirrors `if (channel && ...)` inside
  //    upsertPendingSalesOrder. If Zoho stops sending the field, nothing vanishes.
  t('blank channel STILL locks',      needsLock('zohoSalesOrder', so('')), true);
  t('missing channel STILL locks',    needsLock('zohoSalesOrder', { action: 'zohoSalesOrder', salesorder: { salesorder_number: 'SO-1' } }), true);

  // 3. INVOICES arrive on the same action and DO write (the INVOICE column)
  t('an INVOICE payload STILL locks', needsLock('zohoSalesOrder', { action: 'zohoSalesOrder', invoice: { invoice_number: 'INV-022496' } }), true);

  // 4. an unparseable body leaves payload undefined → fail safe
  t('undefined payload STILL locks',  needsLock('zohoSalesOrder', undefined), true);
  t('null payload STILL locks',       needsLock('zohoSalesOrder', null), true);
  t('empty object STILL locks',       needsLock('zohoSalesOrder', {}), true);
});

section('C · SCOPE — nothing else changes behaviour', () => {
  // The backfill path reaches the same handler but is user-triggered and rare.
  t('backfill is untouched, still locks',
    needsLock('zohoBackfillSalesOrder', so('ebay_us')), true);
  // A non-direct channel on some OTHER action must not leak the shortcut.
  t('another action with a salesorder still locks',
    needsLock('insertOrders', so('ebay_us')), true);
});

section('D · THE PRE-EXISTING ALLOWLIST IS UNTOUCHED — regression net', () => {
  ['boardTick','boardRadio','boardPickers','boardPrint','boardPart','boardPartLite','boardOrder']
    .forEach(a => t(a + ' stays lock-free', needsLock(a, {}), false));

  // ⚠ Locked BY DEFAULT is the safety boundary. Everything that writes must stay so.
  ['insertOrders','updateMiRows','updateOrderStatus','writeZohoStock','boardStatus',
   'boardLeft','boardAdjust','zohoKitUpdated','telegramCommand','recomputeHand']
    .forEach(a => t(a + ' still locks', needsLock(a, {}), true));
  t('unknown action locks by default', needsLock('somethingNew', {}), true);
  t('empty action locks by default',   needsLock('', {}), true);
});

console.log('\n' + '='.repeat(58));
console.log(pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
