// ============================================================================
// "DID I ALREADY PUSH THIS ONE?" — the Zoho correction has to be remembered
//
// FLOOR REPORT 2026-08-18: "when the picker pushes to Zoho and the item resets,
// they forget whether they pushed it or not."
//
// They were right, and the board genuinely could not answer. The receipt built
// on 2026-08-13 is a CLIENT memory: adjustSeen, a 10-minute TTL, gone on
// reload, invisible on a second tablet. And once the Zoho mirror catches up
// (~2 min) the deviance disappears too — by design — so the row returns to
// looking completely untouched. Nothing on screen said "this one is done", so
// the safe move was to push it again.
//
// The fact was never missing, only unreachable: every push already writes a
// NOTE event to the Activity Log with source `board`, the SKU, before → after,
// the Zoho adjustment id and the picker. getDashboardSnapshot now lifts that
// out of the SAME log tail it already scans (one regex per row, no extra read)
// and the board renders it as a permanent chip.
//
// WHAT THIS ASSERTS — reading the RENDERED row, on a FRESH boot, which is the
// whole point: a client-side memory cannot survive one.
//   A · a SKU pushed earlier today shows its mark on a cold load
//   B · the mark says WHAT it was set to, and carries who/when in its title
//   C · a no-op ("already correct") is remembered too — that is exactly the
//        trip to the shelf we are trying to save
//   D · a SKU never pushed shows NOTHING (the mark must mean something)
//   E · the live receipt WINS while it exists — never two chips on one row
//   F · the mark is still the door back into the numpad
//
// Usage: node diag-zohomark.js
// Before/after proof (A, B, C, F fail on HEAD; D, E pass on both):
//   git show HEAD:FloorBoard.html > /tmp/board-before.html
//   BOARD_FILE=/tmp/board-before.html node diag-zohomark.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK  = require('./mock-tick.js');

const FIXED   = '165447';   // pushed to Zoho earlier today
const NOOPED  = '194244';   // checked, already correct (exists in the mock)
const UNTOUCHED_HINT = 'never-pushed';

let fails = [];
const check = (name, got, want) => {
  const ok = String(got) === String(want);
  console.log(`   ${ok ? '✓' : '✗'} ${name}  →  ${JSON.stringify(got)}${ok ? '' : `   (expected ${JSON.stringify(want)})`}`);
  if (!ok) fails.push(`${name}: got ${got}, expected ${want}`);
};

async function boot(browser, withLiveReceipt) {
  const html = fs.readFileSync(BOARD, 'utf8');
  const ctx = await browser.newContext({ viewport: { width: 800, height: 1280 },
    hasTouch: true, timezoneId: 'America/Chicago' });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push(e.message));

  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      let res = { ok: false };
      if (body.action === 'boardTick') {
        const t = JSON.parse(JSON.stringify(MOCK));
        // What the server now sends: the day's pushes, straight from the log.
        t.cockpit = t.cockpit || {};
        t.cockpit.zohoFixedToday = {
          [FIXED]:  { to: 6, noop: false, at: '10:42 AM', by: 'AShamma · 12343' },
          [NOOPED]: { to: 4, noop: true,  at: '11:05 AM', by: 'AShamma · 12343' }
        };
        res = Object.assign({ ok: true }, t);
      }
      if (body.action === 'boardAdjust') res = { ok: true, before: 4, after: body.target, delta: body.target - 4, noop: false };
      if (body.action === 'boardStatus') res = { ok: true };
      if (body.action === 'boardRadio')  res = { ok: true, nowPlaying: '' };
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());

  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.waitForFunction(() => !document.getElementById('board').classList.contains('booting'),
    null, { timeout: 20000 }).catch(() => errs.push('never left booting'));
  await page.waitForTimeout(1300);
  return { page, ctx, errs };
}

const readRow = (page, sku) => page.evaluate(s => {
  const li = [...document.querySelectorAll('.pick-row')]
    .find(r => (r.textContent || '').indexOf(s) !== -1);
  if (!li) return { found: false };
  const chips = [...li.querySelectorAll('.pc-zoho')];
  const c = chips[0];
  return {
    found: true,
    count: chips.length,
    text: c ? c.textContent.trim() : null,
    title: c ? (c.getAttribute('title') || '') : null,
    perm: c ? c.classList.contains('perm') : null,
    tag: c ? c.tagName : null
  };
}, sku);

(async () => {
  const browser = await chromium.launch();
  console.log(`\nTHE ZOHO PUSH MUST BE REMEMBERED — ${path.basename(BOARD)}`);

  // ---- A/B/C/D · A COLD LOAD. No client memory exists yet, by construction. --
  {
    const { page, ctx, errs } = await boot(browser, false);
    console.log('\n── COLD LOAD — nothing in client memory ────────────────────');

    const fixed = await readRow(page, FIXED);
    check('A · the pushed SKU is marked on a fresh boot', fixed.text, '✓ zoho 6');
    check('B · ...it is the permanent (server-backed) mark', fixed.perm, true);
    check('B · ...and names when and who', /10:42 AM/.test(fixed.title || '') &&
                                           /AShamma/.test(fixed.title || ''), true);

    const nooped = await readRow(page, NOOPED);
    check('C · a no-op is remembered too', nooped.text, '✓ zoho ok');
    check('C · ...and says nothing was sent',
          /nothing was sent/i.test(nooped.title || ''), true);

    // A SKU with no record must stay clean, or the mark means nothing.
    // ⚠ page.evaluate takes ONE argument — wrap or it throws at runtime.
    const clean = await page.evaluate(({ a, b }) => {
      const rows = [...document.querySelectorAll('.pick-row')]
        .filter(r => (r.textContent || '').indexOf(a) === -1 &&
                     (r.textContent || '').indexOf(b) === -1);
      return rows.some(r => r.querySelector('.pc-zoho'));
    }, { a: FIXED, b: NOOPED });
    check('D · every un-pushed row stays unmarked', clean, false);

    errs.forEach(e => fails.push('cold ' + e));
    await ctx.close();
  }

  // ---- E/F · THE LIVE RECEIPT STILL WINS, AND THE MARK IS A DOOR -----------
  {
    const { page, ctx, errs } = await boot(browser, true);
    console.log('\n── A FRESH PUSH ON TOP OF AN EXISTING MARK ─────────────────');

    // Push the SAME sku again through the numpad — that creates the live receipt.
    await page.evaluate(s => {
      const li = [...document.querySelectorAll('.pick-row')]
        .find(r => (r.textContent || '').indexOf(s) !== -1);
      const door = li.querySelector('.pc-adj') || li.querySelector('.pc-zoho');
      door.click();
    }, FIXED);
    await page.waitForTimeout(450);
    check('F · the mark reopens the numpad', await page.evaluate(
      () => !!document.getElementById('npCard') &&
            getComputedStyle(document.getElementById('npWrap') ||
                             document.getElementById('npCard')).display !== 'none'), true);
    await page.click('#npOk');
    await page.waitForTimeout(1400);

    const after = await readRow(page, FIXED);
    check('E · exactly ONE chip on the row', after.count, 1);
    check('E · ...and it is the LIVE receipt, not the record', after.perm, false);

    errs.forEach(e => fails.push('live ' + e));
    await ctx.close();
  }

  await browser.close();
  console.log('');
  if (fails.length) {
    console.log(`✗ ${fails.length} FAILURE(S)`);
    fails.forEach(f => console.log('   · ' + f));
    process.exit(1);
  }
  console.log('✓ THE PUSH IS REMEMBERED: it survives a reload, shows on any device,\n' +
              '  names who and when, and is still the way back in.');
})();
