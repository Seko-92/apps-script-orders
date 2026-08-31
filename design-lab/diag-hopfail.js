// ============================================================================
// THE PROXY ANSWERED, BUT NOT WITH AN ANSWER — malformed replies and stalls
//
// Floor report 2026-08-17, after the pick-TTL fix had shipped: "they pick the
// item, the pick button comes back on, and the board says 'could not reach the
// sheet - this pick was NOT saved. Tap it again.' For some items it took 3 or 4
// clicks. Same for the count." They swapped the warehouse wifi for a phone
// hotspot and it changed nothing — correctly, because none of this was the wifi.
//
// MEASURED THE SAME MORNING, on the live pipe, during the shift:
//   · read (boardTick)                   0.80s, served from the published cell
//   · Apps Script DIRECT, 12 calls       12/12 fine, 1.5–4.2s
//   · through the n8n proxy, good window 2.5–5.9s
//   · through the n8n proxy, BAD window  8.7s · 42.3s · 43.5s · 59.1s,
//                                        one `proxy: 404`, and TWICE a reply
//                                        that was the REQUEST echoed back
// Google was healthy and the box was idle (load 0.27, n8n at 0.13% CPU). The
// fault is the proxy hop, and it arrives in bursts — which is the whole reason
// it reads as "sometimes it works, sometimes I tap four times".
//
// TWO FAULTS FELL OUT OF THAT, BOTH PRODUCING THE ONE SYMPTOM:
//   1 · A MALFORMED REPLY READ AS A REFUSAL. n8n's error path passes the
//       incoming item straight through, so the board received
//       {"headers":{…},"body":{…}} — no `ok`, and no `error` either. hqCall
//       only treated a missing `ok` as a hop failure when `error` was ALSO
//       present, so this sailed past into "the server said no": the row
//       un-picked itself and the picker was told NOTHING. That is the ORIGINAL
//       report, and it survived the TTL fix because it was never the TTL.
//   2 · A TIMEOUT WAS REPORTED AS A FAILURE. Apps Script answers in ~2s; the
//       proxy stalls for 40–60s. A write that outran the 25s bound had most
//       likely LANDED, but the board cleared the row and said "NOT saved" —
//       so every retry was re-answering a question already answered right.
//
// WHAT THIS ASSERTS — reading the RENDERED row and the toast the picker sees:
//   A · a malformed reply to ✓ Pick is TOLD, never a silent revert   [FAILS on HEAD]
//   B · a malformed reply to a COUNT is told too                     [FAILS on HEAD]
//   C · a write that outruns the bound HOLDS the row and does not
//       claim it failed                                              [FAILS on HEAD]
//   D · a genuine refusal still opens the picker drawer              [regression net]
//   E · a healthy tick — which has NO `ok` key in production — still
//       renders the board                                            [regression net]
//   F · a real {error:"proxy: …"} is still reported as transport     [regression net]
//
// ⚠ E IS THE GUARD ON THE FIX ITSELF. boardTick returns the tick payload, which
// carries no `ok`. Widening the strict test from writes to every action would
// classify every healthy tick as a transport fault and black out the board.
// If E ever fails, the rule has been applied too broadly.
//
// Usage: node diag-hopfail.js
// Before/after proof (A, B, C must FAIL on HEAD; D, E, F must PASS on both):
//   git show HEAD:FloorBoard.html > /tmp/board-before.html
//   BOARD_FILE=/tmp/board-before.html node diag-hopfail.js
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

// The EXACT body captured from production, twice, on 2026-08-17. n8n echoing
// the request back. Kept verbatim so this test cannot drift from what the
// proxy really sent.
const ECHOED_REQUEST = {
  headers: { host: 'hq.yassinqurabi.com', 'user-agent': 'curl/8.21.0',
             'content-length': '38', accept: '*/*', 'content-type': 'application/json' },
  params: {}, query: {},
  body: { action: 'boardStatus', orderId: 'x', status: 'PREPARING' },
  webhookUrl: 'https://hq.yassinqurabi.com/api/board', executionMode: 'production'
};

async function boot(browser, state) {
  const html = fs.readFileSync(BOARD, 'utf8');
  const ctx = await browser.newContext({
    viewport: { width: 1280, height: 800 }, hasTouch: true, timezoneId: 'America/Chicago'
  });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));

  await page.route('http://hqlab.test/**', async route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      const isWrite = body.action === 'boardStatus' || body.action === 'boardLeft';

      // ⚠ THE HARNESS BUG WORTH REMEMBERING (carried from diag-offline):
      // context.setOffline() does NOT reach a Playwright-intercepted route —
      // the handler runs before the network does. A dead line has to be killed
      // HERE, and a stall has to be a route that never fulfils.
      if (isWrite && state.stall) return;                    // never answer
      if (isWrite && state.echo) {
        return route.fulfill({ contentType: 'application/json',
                               body: JSON.stringify(ECHOED_REQUEST) });
      }
      if (isWrite && state.proxyErr) {
        return route.fulfill({ contentType: 'application/json',
                               body: JSON.stringify({ error: 'proxy: Request failed with status code 404' }) });
      }
      if (isWrite && state.refuse) {
        return route.fulfill({ contentType: 'application/json',
                               body: JSON.stringify({ ok: false, error: 'Set the Pick ID before picking', needsPicker: true }) });
      }

      let res = { ok: false, message: 'unknown action' };
      if (body.action === 'boardTick') {
        // ⚠ NO `ok` KEY — this is what production actually returns for a tick,
        // and case E depends on it being faithful.
        const t = JSON.parse(JSON.stringify(MOCK));
        t.picker  = state.noPicker ? '' : 'Shipping - Yassin 1';
        t.pickers = ['Shipping - Yassin 1'];
        res = t;
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
  return { page, ctx, errs };
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

// The shelf-count path: open the numpad from the row, type a number, confirm.
const saveCount = page => page.evaluate(sku => {
  const li = [...document.querySelectorAll('.pick-row')]
    .find(r => (r.textContent || '').indexOf(sku) !== -1);
  const b = li && li.querySelector('.pc-btn');
  if (!b) return false;
  b.click();
  return true;
}, ROW_SKU);

(async () => {
  const browser = await chromium.launch();
  console.log(`\nMALFORMED PROXY REPLIES AND STALLS — ${path.basename(BOARD)}\n`);

  // ---- A · A MALFORMED REPLY TO ✓ PICK -------------------------------------
  // The headline. On HEAD this is a silent revert with no message at all.
  {
    const state = { echo: true };
    const { page, ctx, errs } = await boot(browser, state);
    console.log('A · the proxy echoes the request back on a ✓ Pick');

    await tapPick(page);
    await page.waitForTimeout(2500);

    const t = await toastState(page);
    check('  the picker is TOLD something went wrong', t.shown, true);
    check('  ...and it names the connection, not a refusal',
          /could not reach|not saved/i.test(t.text), true);
    const row = await rowState(page);
    check('  row is not left falsely picked', { prep: row.prep, btn: row.pickBtn },
                                              { prep: false, btn: true });

    errs.forEach(e => failures.push('A ' + e));
    await ctx.close();
  }

  // ---- B · THE SAME REPLY TO A SHELF COUNT ---------------------------------
  {
    const state = { echo: true };
    const { page, ctx, errs } = await boot(browser, state);
    console.log('B · the same malformed reply on a shelf count');

    const opened = await saveCount(page);
    if (!opened) {
      check('  count button present on the row', false, true);
    } else {
      await page.waitForTimeout(400);
      await page.evaluate(() => {
        const k = [...document.querySelectorAll('.np-key')].find(b => b.textContent.trim() === '5');
        if (k) k.click();
        const ok = document.getElementById('npOk');
        if (ok) ok.click();
      });
      await page.waitForTimeout(2500);
      const t = await toastState(page);
      check('  the picker is TOLD the count did not save', t.shown, true);
      check('  ...and it names the connection',
            /could not reach|not saved/i.test(t.text), true);
    }

    errs.forEach(e => failures.push('B ' + e));
    await ctx.close();
  }

  // ---- C · A WRITE THAT OUTRUNS THE BOUND ----------------------------------
  // The 42–59s stall. HEAD clears the row and says "NOT saved" — a claim it
  // cannot support, and the one that cost the extra taps.
  {
    const state = { stall: true };
    const { page, ctx, errs } = await boot(browser, state);
    console.log('C · the write stalls past the bound (the 42–59s case)');

    // Drive the clock rather than waiting 25 real seconds.
    await page.clock.install();
    await tapPick(page);
    await page.waitForTimeout(300);
    await page.clock.runFor(27000);
    await page.waitForTimeout(600);

    const row = await rowState(page);
    check('  the row STAYS picked while we find out',
          { prep: row.prep, btn: row.pickBtn }, { prep: true, btn: false });
    const t = await toastState(page);
    check('  ...and the message does not claim it failed',
          /not saved/i.test(t.text), false);
    check('  ...it says we are still checking',
          /still checking|slow connection/i.test(t.text), true);

    errs.forEach(e => failures.push('C ' + e));
    await ctx.close();
  }

  // ---- D · A GENUINE REFUSAL STILL BEHAVES ---------------------------------
  // Regression net: the fix must not turn our own refusals into transport faults.
  // That inversion has already happened once, on 2026-08-14.
  {
    const state = { refuse: true };
    const { page, ctx, errs } = await boot(browser, state);
    console.log('D · a genuine refusal (needsPicker) — regression net');

    await tapPick(page);
    await page.waitForTimeout(2500);

    const drawerOpen = await page.evaluate(() => {
      const d = document.getElementById('drw');
      return !!(d && d.classList.contains('on'));
    });
    check('  the picker drawer opens rather than an error', drawerOpen, true);
    const t = await toastState(page);
    check('  ...and it is NOT reported as a connection fault',
          /could not reach/i.test(t.text), false);

    errs.forEach(e => failures.push('D ' + e));
    await ctx.close();
  }

  // ---- E · A HEALTHY TICK HAS NO `ok` KEY ----------------------------------
  // ⚠ THE GUARD ON THE FIX. If the strict write test is ever widened to reads,
  // the board goes dark here and nowhere else.
  {
    const state = {};
    const { page, ctx, errs } = await boot(browser, state);
    console.log('E · a healthy tick carries no `ok` — the board must still paint');

    const rows = await page.evaluate(() => document.querySelectorAll('.pick-row').length);
    check('  the pick list rendered', rows > 0, true);
    const lost = await page.evaluate(() => {
      const c = document.getElementById('connlost');
      return !!(c && getComputedStyle(c).display !== 'none' && c.classList.contains('show'));
    });
    check('  ...and no "lost the line" banner', lost, false);

    errs.forEach(e => failures.push('E ' + e));
    await ctx.close();
  }

  // ---- F · A REAL PROXY ERROR IS STILL TRANSPORT ---------------------------
  {
    const state = { proxyErr: true };
    const { page, ctx, errs } = await boot(browser, state);
    console.log('F · a real {error:"proxy: 404"} — regression net');

    await tapPick(page);
    await page.waitForTimeout(2500);

    const t = await toastState(page);
    check('  reported as a connection fault', /could not reach/i.test(t.text), true);
    const row = await rowState(page);
    check('  and the row is put back', { prep: row.prep, btn: row.pickBtn },
                                       { prep: false, btn: true });

    errs.forEach(e => failures.push('F ' + e));
    await ctx.close();
  }

  await browser.close();
  console.log('');
  if (failures.length) {
    console.log(`✗ ${failures.length} FAILURE(S)`);
    failures.forEach(f => console.log('   · ' + f));
    process.exit(1);
  }
  console.log('✓ all green');
})();
