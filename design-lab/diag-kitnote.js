// ============================================================================
// CAN A NOTE REACH THE FLOOR? — every place a human can write one.
//
// Floor report 2026-08-14: "let's say I want to leave a comment for this sub
// item, should I care, or leave the comment in the main kit, or I can't because
// the main item is pinned?"  Two real workflows behind it:
//   · a HOLD — "wait, don't ship this yet"
//   · a PACKAGING note — the live "Miguel" note means a customer whose orders
//     need a particular packing. It is on all 14 rows of SO-24853 on the sheet
//     and was showing on only 9 of them: nine items packed right, five blind.
//
// ⚠ THE BUG: humanNote() did `if (note starts with "↳") return ''` — throwing
// away THE WHOLE NOTE, because it assumed a component's note is only the machine
// tag. Expansion appends the kit row's own note AFTER the tag, so every human
// note on every component row was deleted on the way to the floor. Silently:
// perfect on the sheet, absent on the board.
//
// THE FULL MAP THIS ASSERTS:
//   A · note on a component row                → shows (was swallowed)
//   B · note on the parent BEFORE expanding    → shows on every component
//   C · note on the parent AFTER expanding     → shows, carried down (was lost)
//   D · note on an UNEXPANDED / READY kit row  → shows (already worked — net)
//   E · machine annotations never reach the floor, and are not mistaken for
//       human text; an UNKNOWN segment is kept (fail toward showing)
//   F · a Zoho flag keeps its red warning AND the human note under it
//
// Usage: node diag-kitnote.js
// Before/after proof (A, B, C, F fail on HEAD; D and E are the nets):
//   git show HEAD:FloorBoard.html > /tmp/board-before.html
//   BOARD_FILE=/tmp/board-before.html node diag-kitnote.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK = require('./mock-tick.js');

const failures = [];
function check(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  if (!ok) failures.push(`${label}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
}

const KIT = '159012';
const row = (sku, loc, note, extra) => Object.assign({
  channel: 'DIRECT', orderId: 'SO-24853', sku, qty: 1, location: loc,
  status: 'PENDING', note: note || '', isKit: false, hand: 20
}, extra || {});

function tickWith(rows, kits) {
  const t = JSON.parse(JSON.stringify(MOCK));
  t.openOrders = rows;
  t.openOrdersTotal = rows.length;
  t.openOrdersBy = { EBAY: 0, DIRECT: rows.length };
  t.kits = kits || [];
  return t;
}

async function boot(browser, tick) {
  const html = fs.readFileSync(BOARD, 'utf8');
  const ctx = await browser.newContext({
    viewport: { width: 1280, height: 900 }, hasTouch: true, timezoneId: 'America/Chicago'
  });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));
  await page.route('http://hqlab.test/**', route => {
    if (route.request().url().includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      let res = { ok: false };
      if (body.action === 'boardTick')  res = Object.assign({ ok: true }, tick);
      if (body.action === 'boardRadio') res = { ok: true, nowPlaying: '' };
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());
  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.waitForFunction(
    () => !document.getElementById('board').classList.contains('booting'),
    null, { timeout: 20000 }
  ).catch(() => errs.push('board never left booting'));
  await page.waitForTimeout(1300);
  return { page, ctx, errs };
}

/** What the picker actually reads on a row: the note text and the red warning. */
const noteOf = (page, sku) => page.evaluate(s => {
  const li = [...document.querySelectorAll('.pick-row')]
    .find(r => (r.textContent || '').indexOf(s) !== -1);
  if (!li) return null;
  const n = li.querySelector('.sub-note');
  const w = li.querySelector('.sub-warn');
  return { note: n ? n.textContent.replace(/^📌\s*/, '').trim() : '',
           warn: w ? w.textContent.replace(/^⚠\s*/, '').trim() : '' };
}, sku);

(async () => {
  const browser = await chromium.launch();
  console.log(`\nBOARD: ${BOARD}\n`);
  const kitEntry = [{ key: 'SO-24853|' + KIT, parent: KIT, order: 'SO-24853',
                      total: 2, done: 0, hue: 0, dash: 0, left: ['E-12'] }];

  // ── A/B/E · component notes, machine stripping ───────────────────────────
  {
    console.log('='.repeat(70));
    console.log('  A/B/E · notes on component rows, and machine noise');
    console.log('='.repeat(70));
    const rows = [
      // B — parent note copied down by expansion (the live SO-24853 shape)
      row('167400', 'E-37', '↳ from KIT-' + KIT + ' · Miguel', { kit: KIT }),
      // A + E — a per-part note behind a swap annotation
      row('173772', 'E-12', '↳ from KIT-' + KIT + ' · swapped 173763 → 173772 · check the seal',
          { kit: KIT }),
      // E — machine-only note must stay silent, not print the tag
      row('172593', 'F-20', '↳ from KIT-' + KIT, { kit: KIT }),
      // E — a human note containing its own " · " must survive whole
      row('157311', 'I-53', '↳ from KIT-' + KIT + ' · call customer · urgent', { kit: KIT }),
      // E — an UNRECOGNISED segment is human until proven otherwise
      row('194199', 'L-226', '↳ from KIT-' + KIT + ' · brand new machine word', { kit: KIT }),
      // net — an ordinary row is untouched
      row('158346', 'F-23', 'Miguel')
    ];
    const { page, ctx, errs } = await boot(browser, tickWith(rows, kitEntry));

    check('B · parent note reaches the component', (await noteOf(page, '167400')).note, 'Miguel');
    // ⭐ THE HEADLINE — a per-part comment survives the machine prefix.
    check('A · per-part note survives', (await noteOf(page, '173772')).note, 'check the seal');
    check('E · swap annotation stripped',
          (await noteOf(page, '173772')).note.indexOf('swapped'), -1);
    check('E · machine-only note prints nothing', (await noteOf(page, '172593')).note, '');
    check('E · human " · " kept whole', (await noteOf(page, '157311')).note, 'call customer · urgent');
    check('E · unknown segment kept (fails toward showing)',
          (await noteOf(page, '194199')).note, 'brand new machine word');
    check('net · ordinary row unchanged', (await noteOf(page, '158346')).note, 'Miguel');
    check('no page errors', errs, []);
    await ctx.close();
  }

  // ── C · the HOLD written AFTER expansion ─────────────────────────────────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  C · hold added to the parent AFTER it was expanded');
    console.log('='.repeat(70));
    const rows = [
      // Components created earlier, so they never carried the hold...
      row('167400', 'E-37', '↳ from KIT-' + KIT + ' · Miguel',
          { kit: KIT, kitNote: 'HOLD — do not ship yet' }),
      // ...and one that already carries the parent note (must not print twice).
      row('173772', 'E-12', '↳ from KIT-' + KIT + ' · HOLD — do not ship yet',
          { kit: KIT, kitNote: 'HOLD — do not ship yet' }),
      row('158346', 'F-23', 'Miguel')
    ];
    const { page, ctx, errs } = await boot(browser, tickWith(rows, kitEntry));

    // ⭐ Without this the hold is invisible: the parent row is collapsed away.
    check('C · hold reaches the component', (await noteOf(page, '167400')).note,
          'HOLD — do not ship yet · Miguel');
    check('C · order-level note LEADS the per-part one',
          (await noteOf(page, '167400')).note.indexOf('HOLD'), 0);
    check('C · never printed twice', (await noteOf(page, '173772')).note,
          'HOLD — do not ship yet');
    check('C · untouched rows stay untouched', (await noteOf(page, '158346')).note, 'Miguel');
    check('no page errors', errs, []);
    await ctx.close();
  }

  // ── D · REGRESSION NET: the READY / unexpanded kit row ───────────────────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  D · unexpanded / READY kit row — already worked, must keep working');
    console.log('='.repeat(70));
    const rows = [
      row(KIT, 'K-56', 'HOLD — call the office', { isKit: true }),
      row('158346', 'F-23', 'Miguel')
    ];
    const { page, ctx, errs } = await boot(browser, tickWith(rows, []));
    check('D · note on a READY kit shows', (await noteOf(page, KIT)).note, 'HOLD — call the office');
    check('no page errors', errs, []);
    await ctx.close();
  }

  // ── F · a Zoho flag must not eat the note under it ───────────────────────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  F · Zoho flag + human note on the same row');
    console.log('='.repeat(70));
    const rows = [
      // _flagDirectRow PREPENDS the warning as its own first LINE.
      row('167400', 'E-37', '⚠️ REMOVED IN ZOHO 8/14\n↳ from KIT-' + KIT + ' · Miguel',
          { kit: KIT }),
      row('158346', 'F-23', '⚠️ ZOHO QTY: 4 → 2 8/14\nMiguel')
    ];
    const { page, ctx, errs } = await boot(browser, tickWith(rows, kitEntry));
    const a = await noteOf(page, '167400');
    const b = await noteOf(page, '158346');
    check('F · flag still shown as the warning', a.warn, 'REMOVED IN ZOHO 8/14');
    check('F · and the human note underneath survives', a.note, 'Miguel');
    check('F · same on a non-kit row', b.warn, 'ZOHO QTY: 4 → 2 8/14');
    check('F · human note kept there too', b.note, 'Miguel');
    check('no page errors', errs, []);
    await ctx.close();
  }

  await browser.close();
  console.log('\n' + '='.repeat(70));
  if (failures.length) {
    console.log(`  ${failures.length} FAILURE(S)`);
    failures.forEach(f => console.log('   ✗ ' + f));
    process.exit(1);
  }
  console.log('  ALL CLEAR');
})().catch(e => { console.error(e); process.exit(1); });
