// ============================================================================
// THE KIT THREAD — can a picker tell which scattered rows are one box?
//
// Floor report 2026-08-14, with screenshots: "the kit items are scattered
// around... the picker has no idea what item belongs to the kit."
//
// Live case SO-24853: 14 rows, of which 5 are the components of kit 159012,
// sorted by AISLE so they land at E-12, E-37, F-20, I-53 and L-226 — scattered
// through 9 unrelated lines by construction. Every row painted identically.
//
// ⚠ THE BUG WAS A FALSE PREMISE IN A GATE. The spine was drawn only when an
// order held 2+ kits, reasoning that "otherwise the band already says which
// rows ship together". The band says `159012 · 0/5` — that a kit EXISTS and how
// far along it is. It never says WHICH five of the fourteen. The premise only
// holds when the order is nothing but that kit's parts.
//
// The real test is whether the colour DISCRIMINATES:
//   2+ kits                     → yes, they must be told apart
//   1 kit + any non-kit line    → yes  ← the reported case, was broken
//   1 kit and nothing else      → no, every row would wear one colour (the
//                                 "GRAB on all eleven rows" redundancy)
//
// ⚠ NOT FIXED BY REGROUPING. Rejected 2026-08-07: the list is sorted by AISLE
// because that is the WALK, and a kit's parts live in different aisles by
// definition. Colour carries membership; the walk keeps its order.
//
// Usage: node diag-kitthread.js
// Before/after proof (A fails on HEAD; B and C are the regression nets):
//   git show HEAD:FloorBoard.html > /tmp/board-before.html
//   BOARD_FILE=/tmp/board-before.html node diag-kitthread.js
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

const row = (sku, loc, kit) => ({
  channel: 'DIRECT', orderId: 'SO-24853', sku, qty: 1, location: loc,
  status: 'PENDING', note: kit ? ('↳ from KIT-' + kit) : 'Miguel',
  isKit: false, kit: kit || undefined, hand: 20
});

/** SO-24853 as it really is: 5 components of ONE kit among 9 loose lines. */
function fixtureOneKit() {
  const t = JSON.parse(JSON.stringify(MOCK));
  const rows = [
    row('173772', 'E-12', '159012'), row('167400', 'E-37', '159012'),
    row('172593', 'F-20', '159012'), row('157311', 'I-53', '159012'),
    row('194199', 'L-226/A-56', '159012'),
    row('158346', 'F-23'), row('173421', 'G-22'), row('176859', 'G-51'),
    row('167220', 'I-29'), row('171576', 'J-52'), row('171585', 'J-52'),
    row('164943', 'L-53/K-17'), row('197934', 'L-90'), row('167202', 'M-26/J-31')
  ];
  t.openOrders = rows;
  t.openOrdersTotal = rows.length;
  t.openOrdersBy = { EBAY: 0, DIRECT: rows.length };
  t.kits = [{ key: 'SO-24853|159012', parent: '159012', order: 'SO-24853',
              total: 5, done: 0, hue: 0, dash: 0, left: ['E-12', 'E-37'] }];
  return t;
}

/** TWO kits in one order — the case that already worked. Must still work. */
function fixtureTwoKits() {
  const t = fixtureOneKit();
  t.openOrders = t.openOrders.slice(0, 5)
    .concat([row('111111', 'B-1', '158652'), row('222222', 'B-2', '158652')])
    .concat(t.openOrders.slice(5));
  t.openOrdersTotal = t.openOrders.length;
  t.openOrdersBy = { EBAY: 0, DIRECT: t.openOrders.length };
  t.kits.push({ key: 'SO-24853|158652', parent: '158652', order: 'SO-24853',
                total: 2, done: 0, hue: 1, dash: 0, left: ['B-1'] });
  return t;
}

/** An order that is NOTHING BUT one kit's parts — a spine here is redundant. */
function fixtureAllKit() {
  const t = fixtureOneKit();
  t.openOrders = t.openOrders.filter(r => r.kit);
  t.openOrdersTotal = t.openOrders.length;
  t.openOrdersBy = { EBAY: 0, DIRECT: t.openOrders.length };
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
    const url = route.request().url();
    if (url.includes('/api/board')) {
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
  await page.waitForTimeout(1400);
  return { page, ctx, errs };
}

/** Read every pick row's spine as the PICKER sees it — the painted colour. */
const readSpines = page => page.evaluate(() => {
  return [...document.querySelectorAll('.pick-row')].map(li => {
    const cs = getComputedStyle(li);
    const sku = (li.querySelector('.pick-sku') || {}).textContent
             || (li.textContent.match(/\d{6}/) || [''])[0];
    return {
      sku: String(sku).trim().slice(0, 6),
      cls: ['kit-c0', 'kit-c1', 'kit-dash'].filter(c => li.classList.contains(c)).join('+') || '',
      color: cs.borderLeftColor,
      style: cs.borderLeftStyle
    };
  });
});

const readChips = page => page.evaluate(() =>
  [...document.querySelectorAll('.kitchip')].map(c => {
    const dot = c.querySelector('i');
    return {
      text: c.textContent.trim(),
      cls: ['kit-h0', 'kit-h1', 'kit-hd'].filter(x => c.classList.contains(x)).join('+'),
      dot: dot ? getComputedStyle(dot).backgroundColor : null
    };
  }));

const KIT_SKUS = ['173772', '167400', '172593', '157311', '194199'];

(async () => {
  const browser = await chromium.launch();
  console.log(`\nBOARD: ${BOARD}\n`);

  // ── A · THE REPORTED CASE: one kit scattered among loose lines ────────────
  {
    console.log('='.repeat(70));
    console.log('  A · SO-24853 — 5 kit parts among 9 loose lines');
    console.log('='.repeat(70));
    const { page, ctx, errs } = await boot(browser, fixtureOneKit());
    const spines = await readSpines(page);
    const kitRows = spines.filter(r => KIT_SKUS.indexOf(r.sku) !== -1);
    const loose = spines.filter(r => KIT_SKUS.indexOf(r.sku) === -1 && r.sku);

    check('all 5 kit rows found', kitRows.length, 5);
    // ⭐ THE HEADLINE — this is what the picker could not see.
    check('every kit row wears the thread', kitRows.filter(r => r.cls === 'kit-c0').length, 5);
    check('and it is actually PAINTED (not transparent)',
          [...new Set(kitRows.map(r => r.color))], ['rgb(99, 255, 255)']);
    check('loose lines carry NO thread', loose.filter(r => r.cls !== '').length, 0);
    check('loose lines stay transparent',
          [...new Set(loose.map(r => r.color))], ['rgba(0, 0, 0, 0)']);

    const chips = await readChips(page);
    check('band shows one kit chip', chips.length, 1);
    check('chip names the kit and its progress', chips[0].text.replace(/\s+/g, ' '), '159012 · 0/5');
    // The whole point: chip dot and row spine are visibly ONE thing.
    check('chip dot matches the spine colour', chips[0].dot, kitRows[0].color);
    check('no page errors', errs, []);
    await ctx.close();
  }

  // ── B · REGRESSION NET: two kits must still be told apart ─────────────────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  B · two kits in one order — must stay distinguishable');
    console.log('='.repeat(70));
    const { page, ctx, errs } = await boot(browser, fixtureTwoKits());
    const spines = await readSpines(page);
    const a = spines.filter(r => KIT_SKUS.indexOf(r.sku) !== -1);
    const b = spines.filter(r => ['111111', '222222'].indexOf(r.sku) !== -1);

    check('kit A threaded', a.filter(r => r.cls === 'kit-c0').length, 5);
    check('kit B threaded', b.filter(r => r.cls === 'kit-c1').length, 2);
    // Two hues, measured ΔE 57.8 apart across normal + all three dichromacies.
    check('the two kits are DIFFERENT colours', a[0].color !== b[0].color, true);
    check('no page errors', errs, []);
    await ctx.close();
  }

  // ── C · REGRESSION NET: don't paint a colour that says nothing ────────────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  C · an order that is ONLY one kit — thread would be noise');
    console.log('='.repeat(70));
    const { page, ctx, errs } = await boot(browser, fixtureAllKit());
    const spines = await readSpines(page).then(s => s.filter(r => r.sku));
    check('all rows belong to the kit', spines.length, 5);
    // Every row wearing the same colour is the "GRAB on all eleven rows"
    // redundancy — the band chip already carries this fact once.
    check('no thread drawn', spines.filter(r => r.cls !== '').length, 0);
    const chips = await readChips(page);
    check('the band still says it', chips.length, 1);
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
