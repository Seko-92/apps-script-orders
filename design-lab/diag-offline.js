// ============================================================================
// THE PICKER WALKED OUT OF WIFI RANGE — dropouts, hangs, and coming back
//
// Asked 2026-08-15, while confirming the pick-TTL fix: "if the picker was moving
// around the warehouse and at some point he lost the wifi connection for a few
// seconds then reconnected back, what about this case then?"
//
// Reading the code for that question found THREE separate faults, and the first
// of them produces the SAME symptom the floor had already reported — so the
// original report probably had two causes and the TTL was only one:
//
//   1 · markPicked's .catch was SILENT. Row flips, reverts, no message. And
//       with no fetch timeout the hang ran for tens of seconds first, which is
//       where "after a while" came from. sendAdjust had always spoken up on the
//       identical failure; the busiest write on the board was the quiet one.
//   2 · NO TIMEOUT ON ANY REQUEST. A hang held pollBusy true, so nothing could
//       force a refresh — including a reconnect.
//   3 · A failed request fed pollLastMs, which drives the slow-server backoff.
//       So the board went quiet for up to 90s starting the moment the wifi came
//       back. The backoff was measuring "how long until we gave up".
//
// WHAT THIS ASSERTS — behaviour on a real offline transition:
//   A · tapping ✓ Pick with no line does NOT leave the row falsely picked
//        (regression net — this half already worked)
//   B · ...and the picker is TOLD, in words, that nothing was saved
//   C · a request that never answers is ABORTED on a bound, not left hanging
//   D · a failed poll does not poison the backoff (pollLastMs resets)
//   E · reconnecting refreshes AT ONCE instead of waiting for the schedule
//
// Usage: node diag-offline.js
// Before/after proof (B, C, D, E must FAIL on HEAD; A must PASS):
//   git show HEAD:FloorBoard.html > /tmp/board-before.html
//   BOARD_FILE=/tmp/board-before.html node diag-offline.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK = require('./mock-tick.js');
const ROW_SKU = '194244';

const failures = [];
function check(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  if (!ok) failures.push(`${label}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
}

async function boot(browser, state) {
  const html = fs.readFileSync(BOARD, 'utf8');
  const ctx = await browser.newContext({
    viewport: { width: 1280, height: 800 }, hasTouch: true, timezoneId: 'America/Chicago'
  });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));

  const calls = [];
  await page.route('http://hqlab.test/**', async route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      calls.push({ action: body.action, at: Date.now() });
      // ⚠ context.setOffline() does NOT reach a Playwright-INTERCEPTED route —
      // the handler runs before the network does, so the call sails through and
      // the whole "offline" premise evaporates. Cost three false failures on
      // the first run of this file. The line has to be killed HERE.
      if (state.dead) return route.abort('internetdisconnected');
      // A hanging pipe: accepted, never answered. That is what walking out of
      // range actually looks like to fetch — not a refusal, just silence.
      if (state.hang) return;                       // never fulfil, never abort
      let res = { ok: false, message: 'unknown action' };
      if (body.action === 'boardTick') {
        const t = JSON.parse(JSON.stringify(MOCK));
        t.picker = 'Shipping - Yassin 1';
        t.pickers = ['Shipping - Yassin 1'];
        res = Object.assign({ ok: true }, t);
      }
      if (body.action === 'boardPickers')   res = { ok: true, pickers: ['Shipping - Yassin 1'] };
      if (body.action === 'boardSetPicker') res = { ok: true };
      if (body.action === 'boardRadio')     res = { ok: true, nowPlaying: '' };
      if (body.action === 'boardStatus')    res = { ok: true, count: 1 };
      if (body.action === 'boardLeft')      res = { ok: true, count: 0 };
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
  return { page, ctx, calls, errs };
}

const rowState = page => page.evaluate(sku => {
  const li = [...document.querySelectorAll('.pick-row')]
    .find(r => (r.textContent || '').indexOf(sku) !== -1);
  if (!li) return { found: false };
  return { found: true, prep: !!li.querySelector('.pick-status.prep'), pickBtn: !!li.querySelector('.pick-do') };
}, ROW_SKU);

const tapPick = page => page.evaluate(sku => {
  const li = [...document.querySelectorAll('.pick-row')]
    .find(r => (r.textContent || '').indexOf(sku) !== -1);
  const b = li && li.querySelector('.pick-do');
  if (b) b.click();
  return !!b;
}, ROW_SKU);

const toastState = page => page.evaluate(() => {
  const t = document.getElementById('ctoast');
  return {
    shown: !!(t && t.classList.contains('show')),
    text: ((document.getElementById('ctoastMsg') || {}).textContent || '').trim()
  };
});

(async () => {
  const browser = await chromium.launch();
  console.log(`\nOFFLINE / RECONNECT — ${path.basename(BOARD)}\n`);

  // ---- A + B · TAPPING ✓ PICK WITH NO LINE ---------------------------------
  {
    const state = {};
    const { page, ctx, errs } = await boot(browser, state);
    console.log('A/B · ✓ Pick tapped while the line is down');

    state.dead = true;                       // the line goes down mid-aisle
    await tapPick(page);
    await page.waitForTimeout(2500);

    // A — the honest half, which already worked: never leave it falsely picked.
    const row = await rowState(page);
    check('  row is NOT left falsely picked', { prep: row.prep, btn: row.pickBtn }, { prep: false, btn: true });

    // B — the half that was silent. Reverting without a word is what made this
    // indistinguishable from the TTL bug on the floor.
    const t = await toastState(page);
    check('  the picker is TOLD it was not saved', t.shown, true);
    check('  ...and the message says so plainly',
          /not saved|nothing was sent|could not reach/i.test(t.text), true);

    errs.forEach(e => failures.push('A/B ' + e));
    await ctx.close();
  }

  // ---- C · A REQUEST THAT NEVER ANSWERS IS BOUNDED -------------------------
  // Shrink the bound from the page so the test does not sit for 15 real
  // seconds. That tests the MECHANISM; the shipped VALUE is asserted separately.
  {
    const state = {};
    const { page, ctx, errs } = await boot(browser, state);
    console.log('C · a hanging request is aborted, not left forever');

    const hasBound = await page.evaluate(() => typeof window.HQ_READ_TIMEOUT_MS === 'number');
    check('  a read timeout is defined at all', hasBound, true);
    if (hasBound) {
      check('  shipped read bound is 15s', await page.evaluate(() => window.HQ_READ_TIMEOUT_MS), 15000);
    }

    state.hang = true;                       // from here on, nothing is answered
    const outcome = await page.evaluate(() => {
      window.HQ_READ_TIMEOUT_MS = 1500;      // mechanism, not the value
      return new Promise(resolve => {
        let settled = false;
        window.hqCall('boardTick')
          .then(() => { settled = true; resolve('resolved'); })
          .catch(e => { settled = true; resolve(e && e.transport ? 'transport' : 'error'); });
        setTimeout(() => { if (!settled) resolve('STILL HANGING'); }, 6000);
      });
    });
    check('  it fails as a transport fault within the bound', outcome, 'transport');

    errs.forEach(e => failures.push('C ' + e));
    await ctx.close();
  }

  // ---- D · A FAILURE MUST NOT POISON THE BACKOFF ---------------------------
  {
    const state = {};
    const { page, ctx, errs } = await boot(browser, state);
    console.log('D · a failed poll does not slow the recovery');

    state.dead = true;
    await page.evaluate(() => window.pollSoon());
    await page.waitForTimeout(3000);

    const lastMs = await page.evaluate(() => window.pollLastMs);
    check('  pollLastMs reset rather than inflated by the failure', lastMs, 0);

    // The property that matters: the next poll is at the normal cadence, not
    // stretched toward POLL_MAX_MS by a timeout that measured nothing useful.
    const nextWait = await page.evaluate(() =>
      Math.min(Math.max(window.POLL_MS, window.pollLastMs || 0), window.POLL_MAX_MS));
    check('  next poll stays at the 20s cadence', nextWait, 20000);

    errs.forEach(e => failures.push('D ' + e));
    await ctx.close();
  }

  // ---- E · COMING BACK ------------------------------------------------------
  // The scenario as described: walks out of range, walks back. The board must
  // not sit on its old schedule showing the pre-outage picture.
  {
    const state = {};
    const { page, ctx, calls, errs } = await boot(browser, state);
    console.log('E · reconnecting refreshes at once');

    state.dead = true;                       // out of range
    await page.evaluate(() => window.pollSoon());
    await page.waitForTimeout(1500);

    const before = calls.filter(c => c.action === 'boardTick').length;
    state.dead = false;                      // back in range
    await page.evaluate(() => window.dispatchEvent(new Event('online')));
    await page.waitForTimeout(2000);         // well inside the 20s schedule
    const after = calls.filter(c => c.action === 'boardTick').length;

    check('  a tick is fetched within 2s of coming back', after > before, true);

    const banner = await page.evaluate(() =>
      document.getElementById('connLost').classList.contains('show'));
    check('  the "lost the line" banner clears', banner, false);

    errs.forEach(e => failures.push('E ' + e));
    await ctx.close();
  }

  await browser.close();

  console.log('');
  if (failures.length) {
    console.log(`✗ ${failures.length} failure(s)`);
    failures.forEach(f => console.log('   · ' + f));
    process.exit(1);
  }
  console.log('✓ all green');
})();
