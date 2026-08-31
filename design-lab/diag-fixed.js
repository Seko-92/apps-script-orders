// ============================================================================
// "IT HAS BEEN FIXED" — the row must stop shouting once Zoho is corrected.
//
// FLOOR FEEDBACK #3 (2026-08-13): after a count deviance is pushed to Zoho the
// row still showed the red "+N vs shelf" chip AND an amber button, which read
// as an outstanding problem. The receipt was there, but it was the third thing
// on a row where the other two said "broken". The picker had acted; only the
// board hadn't noticed.
//
// The design: once an adjustment exists for that SKU, the RECEIPT replaces both
// the red chip and the fix button — and becomes the door back in, because a
// first correction can itself be wrong (the reason "fix again" existed at all).
//
// THE ASSERTIONS, driven end to end against the real board:
//   before  → red chip, amber FIX ZOHO, row tinted .deviance
//   push    → numpad pre-filled with the count, confirm, boardAdjust fires
//   after   → NO red chip, NO fix button, receipt "⟳ zoho ← N", row NOT tinted
//   re-fix  → tapping the receipt reopens the numpad (capability retained)
//   heal    → when the 2-min mirror lands hand == target, receipt turns green
//             "✓ zoho at N" and the row is quiet by construction
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK  = require('./mock-tick.js');
const OUT   = path.join(__dirname, 'renders');

let fails = [];
const check = (name, got, want) => {
  const ok = String(got) === String(want);
  console.log(`   ${ok ? '✓' : '✗'} ${name}  →  ${JSON.stringify(got)}${ok ? '' : `   (expected ${JSON.stringify(want)})`}`);
  if (!ok) fails.push(`${name}: got ${got}, expected ${want}`);
};

// B-30 / 165447 — hand 4, qty 2, left 1 → a genuine −3.
const SKU = '165447';

(async () => {
  const html = fs.readFileSync(BOARD, 'utf8');
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 800, height: 1280 },
    hasTouch: true, timezoneId: 'America/Chicago' });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push(e.message));
  page.on('console', m => {
    if (m.type() !== 'error') return;
    const t = m.text();
    if (/Failed to load resource|ERR_FAILED|ERR_ABORTED/.test(t)) return;
    errs.push(t);
  });

  // `mirrorHand` lets us simulate the Zoho mirror catching up later.
  let mirrorHand = null;
  let adjustCalls = [];
  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      let res = { ok: false };
      if (body.action === 'boardTick') {
        const t = JSON.parse(JSON.stringify(MOCK));
        if (mirrorHand !== null) {
          t.openOrders.forEach(r => { if (String(r.sku) === SKU) r.hand = mirrorHand; });
        }
        res = Object.assign({ ok: true }, t);
      }
      if (body.action === 'boardAdjust') {
        adjustCalls.push(body);
        res = { ok: true, before: 4, after: body.target, delta: body.target - 4, noop: false };
      }
      if (body.action === 'boardStatus') res = { ok: true };
      if (body.action === 'boardRadio')  res = { ok: true, nowPlaying: '' };
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());

  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.waitForFunction(() => !document.getElementById('board').classList.contains('booting'),
    null, { timeout: 20000 });
  await page.waitForTimeout(1200);

  const read = () => page.evaluate(sku => {
    const li = [...document.querySelectorAll('.pick-row')]
      .find(r => (r.textContent || '').indexOf(sku) !== -1);
    if (!li) return null;
    const txt = s => { const e = li.querySelector(s); return e ? e.textContent.trim() : null; };
    const zoho = li.querySelector('.pc-zoho');
    return {
      dev: txt('.pc-dev'),
      fixBtn: txt('.pc-adj'),
      receipt: txt('.pc-zoho'),
      receiptTag: zoho ? zoho.tagName : null,
      receiptDone: zoho ? zoho.classList.contains('done') : null,
      rowRed: li.classList.contains('deviance'),
      counted: txt('.pc-btn')
    };
  }, SKU);

  console.log('\n── BEFORE — a genuine deviance, nothing done about it ───────');
  const before = await read();
  check('red deviance chip', before.dev, '-3 vs shelf');
  check('FIX ZOHO offered', before.fixBtn, 'fix zoho');
  check('row tinted red', before.rowRed, true);
  check('no receipt yet', before.receipt, null);

  console.log('\n── PUSH THE FIX (numpad → confirm) ─────────────────────────');
  await page.evaluate(sku => {
    const li = [...document.querySelectorAll('.pick-row')]
      .find(r => (r.textContent || '').indexOf(sku) !== -1);
    li.querySelector('.pc-adj').click();
  }, SKU);
  await page.waitForTimeout(450);
  const padVal = await page.evaluate(() => (document.getElementById('npValue') || {}).textContent);
  check('numpad pre-filled with the count', padVal, '1');
  await page.click('#npOk');

  /* ⭐ THE CELEBRATION (2026-08-18, picker's ask). Caught EARLY on purpose:
     it is a 1.1s one-shot that removes itself on animationend, so a check
     after the 1.5s wait below would always read false and look like a bug. */
  await page.waitForTimeout(400);
  const cheer = await page.evaluate(sku => {
    const li = [...document.querySelectorAll('.pick-row')]
      .find(r => (r.textContent || '').indexOf(sku) !== -1);
    return { glow: !!(li && li.classList.contains('fix-land')),
             pop:  (() => { const c = li && li.querySelector('.pc-zoho');
                            return c ? getComputedStyle(c).animationName : 'none'; })() };
  }, SKU);
  check('the row celebrates the landed push', cheer.glow, true);
  check('...and the receipt pops in with it', cheer.pop, 'zohoPop');

  await page.waitForTimeout(1500);
  check('boardAdjust fired once', adjustCalls.length, 1);
  // It must not linger: a permanent glow would read as a state, not an event.
  const lingering = await page.evaluate(sku => {
    const li = [...document.querySelectorAll('.pick-row')]
      .find(r => (r.textContent || '').indexOf(sku) !== -1);
    return !!(li && li.classList.contains('fix-land'));
  }, SKU);
  check('...and clears itself afterwards', lingering, false);
  check('pushed the counted number', adjustCalls[0] && adjustCalls[0].target, 1);

  console.log('\n── AFTER — the row must read DONE, not broken ──────────────');
  const after = await read();
  check('red chip GONE', after.dev, null);
  check('fix button GONE', after.fixBtn, null);
  check('receipt shown', after.receipt, '⟳ zoho ← 1');
  check('receipt is tappable (a button)', after.receiptTag, 'BUTTON');
  check('row NO LONGER tinted red', after.rowRed, false);
  check('count still visible', after.counted, 'counted 1');

  console.log('\n── RE-FIX — the capability must survive ────────────────────');
  await page.evaluate(sku => {
    const li = [...document.querySelectorAll('.pick-row')]
      .find(r => (r.textContent || '').indexOf(sku) !== -1);
    li.querySelector('.pc-zoho').click();
  }, SKU);
  await page.waitForTimeout(450);
  const reopened = await page.evaluate(() => ({
    open: !document.getElementById('numPad').classList.contains('hidden'),
    title: (document.getElementById('npTitle') || {}).textContent
  }));
  check('tapping the receipt reopens the numpad', reopened.open, true);
  check('and it is the adjust pad', /Correct Zoho stock/i.test(reopened.title || ''), true);
  await page.click('#npCancel');
  await page.waitForTimeout(300);

  console.log('\n── HEAL — the 2-min mirror lands, hand becomes 1 ───────────');
  mirrorHand = 1;                       // Zoho Stock sheet catches up
  await page.evaluate(() => { if (window.pollSoon) pollSoon(); });
  await page.waitForTimeout(2500);
  const healed = await read();
  check('receipt confirmed (green)', healed.receiptDone, true);
  check('receipt wording', healed.receipt, '✓ zoho at 1');
  check('still no red chip', healed.dev, null);
  check('still not tinted', healed.rowRed, false);

  await page.screenshot({ path: path.join(OUT, 'fixed-row.png') });
  if (errs.length) { console.log('\n⚠ JS: ' + errs.join(' | ')); fails.push('js errors'); }
  await browser.close();

  console.log('\n' + '='.repeat(62));
  if (fails.length) { console.log('✗ ' + fails.join('\n✗ ')); process.exit(1); }
  console.log('✓ A FIXED ROW READS FIXED: red and the call-to-action both retire,');
  console.log('  the receipt carries the state and the way back in, and it goes');
  console.log('  green on its own when the mirror catches up.');
})().catch(e => { console.error('CRASH', e); process.exit(1); });
