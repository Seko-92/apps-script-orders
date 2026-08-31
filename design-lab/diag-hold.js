// ============================================================================
// ⏸ HOLD — "pick it, prepare it, do NOT buy the label yet"
//
// Floor workflow (2026-08-15, 3–5×/week): mid-pick a buyer asks for expedited
// shipping and the payment conversation runs on, or an engine serial needs
// verifying. The shipping-responsible person writes a hold; the order still
// gets picked and prepared, then set aside until they say how to ship it.
//
// ⚠ WHAT IS HELD IS THE LABEL, NOT THE PICKING — so ✓ Pick must stay live.
// ⚠ AND THE BOARD CANNOT ENFORCE IT: the label is bought in eBay's Seller Hub,
//    not here. This makes a hold impossible to MISS, not impossible to ignore.
//    It has already been missed once in production.
//
// WHAT THIS ASSERTS:
//   A · a hold anywhere in a note marks the ORDER, and every row stays pickable
//   B · it is found wherever it is written — appended after an existing note,
//       leading, lower-case, or on a collapsed kit parent (carried via kitNote)
//   C · an unheld order gets NOTHING (no chip, no colour)
//   D · ⚠ THE COLOUR: amber, the same token as the unset-picker chip — NOT the
//       red used for deviances. Red means "something disagrees, act on it"; a
//       hold is a deliberate pause. The user rejected red explicitly.
//
// Usage: node diag-hold.js
// Before/after proof (A, B, D fail on HEAD; C passes — it is the net):
//   git show HEAD:FloorBoard.html > /tmp/board-before.html
//   BOARD_FILE=/tmp/board-before.html node diag-hold.js
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
  if (!ok) failures.push(label);
}

const row = (sku, loc, note, extra) => Object.assign({
  channel: 'DIRECT', orderId: 'SO-24853', sku, qty: 1, location: loc,
  status: 'PENDING', note: note || '', isKit: false, hand: 20
}, extra || {});

function tick(rows) {
  const t = JSON.parse(JSON.stringify(MOCK));
  t.openOrders = rows;
  t.openOrdersTotal = rows.length;
  t.openOrdersBy = { EBAY: 0, DIRECT: rows.length };
  t.kits = [];
  return t;
}

async function boot(browser, payload) {
  const html = fs.readFileSync(BOARD, 'utf8');
  const ctx = await browser.newContext({
    viewport: { width: 1280, height: 900 }, hasTouch: true, timezoneId: 'America/Chicago'
  });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));
  await page.route('http://hqlab.test/**', route => {
    if (route.request().url().includes('/api/board')) {
      const b = JSON.parse(route.request().postData() || '{}');
      let res = { ok: false };
      if (b.action === 'boardTick')  res = Object.assign({ ok: true }, payload);
      if (b.action === 'boardRadio') res = { ok: true, nowPlaying: '' };
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

/**
 * Is the hold VISIBLE ANYWHERE the floor would see it — band chip or row chip?
 *
 * ⚠ THIS USED TO CHECK ONLY `.holdchip`, i.e. only the band, and every fixture
 * used channel:'DIRECT' — which ALWAYS gets a band. So the suite passed while
 * the feature did nothing on a single-line eBay order, which is what the floor
 * hit within the hour. THE FIXTURE ONLY EXERCISED THE CASE THAT WORKS.
 */
const holdState = page => page.evaluate(() => {
  const band = document.querySelector('.holdchip');
  const rowc = document.querySelector('.pick-hold');
  const el = band || rowc;
  const cs = el ? getComputedStyle(el) : null;
  return {
    present: !!el,
    where: band ? 'band' : (rowc ? 'row' : 'nowhere'),
    text: el ? el.textContent.trim() : '',
    // The band chip colours its TEXT; the row chip colours its BACKGROUND.
    color: cs ? (band ? cs.color : cs.backgroundColor) : '',
    pickable: document.querySelectorAll('.pick-do').length
  };
});
const bandState = holdState;

(async () => {
  const browser = await chromium.launch();
  console.log(`\nBOARD: ${BOARD}\n`);

  const AMBER = 'rgb(255, 184, 77)';   // --money, the unset-picker chip's colour
  const RED   = 'rgb(255, 107, 107)';  // --alarm, deviances. Must NOT be used here.

  // ── A/D · a hold marks the order; picking stays live; colour is amber ─────
  {
    console.log('='.repeat(70));
    console.log('  A/D · hold appended to an existing note');
    console.log('='.repeat(70));
    const { page, ctx, errs } = await boot(browser, tick([
      row('158346', 'F-23', 'Miguel'),
      row('173421', 'G-22', 'Miguel · HOLD — verify engine SN'),   // appended, not leading
      row('176859', 'G-51', 'Miguel')
    ]));
    const b = await bandState(page);
    check('HOLD chip on the band', b.present, true);
    check('it says the word', b.text, 'HOLD');
    // ⭐ THE ONE THE USER ASKED FOR EXPLICITLY.
    check('chip is AMBER, not red', b.color, AMBER);
    check('and is definitely not the deviance red', b.color === RED, false);
    // ⚠ A hold must never stop the picking — that was the whole design ruling.
    check('every row still pickable', b.pickable, 3);
    check('no page errors', errs, []);
    await ctx.close();
  }

  // ── B · found wherever it is written ─────────────────────────────────────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  B · leading / lower-case / on a collapsed kit parent');
    console.log('='.repeat(70));

    const cases = [
      ['leading',              [row('158346', 'F-23', 'HOLD — waiting payment · Miguel')]],
      ['lower-case',           [row('158346', 'F-23', 'Miguel · hold for the SN check')]],
      ['mixed case',           [row('158346', 'F-23', 'Miguel · On Hold')]],
      ['carried from the kit parent (kitNote)',
                               [row('167400', 'E-37', '↳ from KIT-159012 · Miguel',
                                    { kit: '159012', kitNote: 'HOLD — do not ship yet' })]],
      ['on ANY line marks the whole order',
                               [row('158346', 'F-23', 'Miguel'),
                                row('173421', 'G-22', 'Miguel'),
                                row('176859', 'G-51', 'HOLD — check with the office')]]
    ];
    for (const [label, rows] of cases) {
      const { page, ctx } = await boot(browser, tick(rows));
      check(label, (await bandState(page)).present, true);
      await ctx.close();
    }
  }

  // ── E · ⭐ THE CASE THE FIRST FIXTURE MISSED — bandless eBay orders ───────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  E · eBay orders, which only get a band at 2+ lines');
    console.log('='.repeat(70));
    const ebay = (sku, loc, note, oid) => Object.assign(
      row(sku, loc, note), { channel: 'EBAY', orderId: oid || '14-15025-11695' });

    // ⭐ The exact live report: ONE eBay line, HOLD written on it, no band.
    {
      const { page, ctx, errs } = await boot(browser, tick([
        ebay('155682', 'L-200/E-49', 'HOLD')
      ]));
      const b = await holdState(page);
      check('single-line eBay order shows the hold', b.present, true);
      check('and it lands on the ROW (there is no band)', b.where, 'row');
      check('says the word', b.text, 'HOLD');
      check('amber, not red', b.color, AMBER);
      check('still pickable', b.pickable, 1);
      check('no page errors', errs, []);
      await ctx.close();
    }

    // A 2+ line eBay order DOES get a band — the chip belongs there instead,
    // and must not also duplicate onto every row.
    {
      const { page, ctx } = await boot(browser, tick([
        ebay('155682', 'L-200/E-49', 'HOLD — verify SN'),
        ebay('165411', 'J-20', 'Miguel')
      ]));
      const b = await holdState(page);
      check('multi-line eBay order shows the hold', b.present, true);
      check('and it lands on the BAND', b.where, 'band');
      check('not duplicated onto the rows too',
            await page.evaluate(() => document.querySelectorAll('.pick-hold').length), 0);
      await ctx.close();
    }

    // Two DIFFERENT eBay orders, only one held — the other must stay clean.
    {
      const { page, ctx } = await boot(browser, tick([
        ebay('155682', 'L-200/E-49', 'HOLD', '14-15025-11695'),
        ebay('165411', 'J-20', 'Miguel',     '22-99999-00000')
      ]));
      check('exactly one row marked',
            await page.evaluate(() => document.querySelectorAll('.pick-hold').length), 1);
      await ctx.close();
    }
  }

  // ── C · REGRESSION NET: no hold, no chip ─────────────────────────────────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  C · an ordinary order must gain nothing');
    console.log('='.repeat(70));
    const { page, ctx, errs } = await boot(browser, tick([
      row('158346', 'F-23', 'Miguel'),
      row('173421', 'G-22', 'Buyer Note: please pack carefully'),
      row('176859', 'G-51', '')
    ]));
    const b = await bandState(page);
    check('no HOLD chip', b.present, false);
    check('rows still pickable', b.pickable, 3);
    check('no page errors', errs, []);
    await ctx.close();
  }

  // ── C2 · the word must be a WORD, not a fragment ─────────────────────────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  C2 · "household" / "holder" must not trigger');
    console.log('='.repeat(70));
    for (const [label, note] of [
      ['household', 'Miguel · household items packing'],
      ['holder',    'Miguel · needs the holder bracket'],
      ['withhold',  'Miguel · do not withhold anything']
    ]) {
      const { page, ctx } = await boot(browser, tick([row('158346', 'F-23', note)]));
      check(`"${label}" does not fire`, (await bandState(page)).present, false);
      await ctx.close();
    }
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
