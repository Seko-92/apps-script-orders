// ============================================================================
// THE PICK THAT UN-PICKED ITSELF — the optimistic override vs the publish chain
//
// Floor report 2026-08-15: "they pick the item, after a while the item activates
// the pick button again for some reason, so they have to click pick again."
// A couple of times a day, never reproducible on demand.
//
// It was arithmetic, not a race in the code. PICK_OVERRIDE_MS was 45s, and the
// chain that carries a pick back to the board is:
//
//     tap → dirty flag → 1-min publish trigger → 15s n8n cache → 20s poll
//         = 45s TYPICAL, 95s WORST
//
// So the backstop expired exactly where the answer usually arrives and well
// before it sometimes arrives. Whether a pick stuck was decided by where the tap
// landed relative to the publish trigger — which is why it looked intermittent.
//
// TWO fixes, and they are complementary, not redundant:
//   · boardSetStatus now publishes INLINE (kills the 0–60s trigger wait)
//   · PICK_OVERRIDE_MS → 120s (the inline publish is DEBOUNCED at 15s, so on a
//     burst every tap after the first still rides the flag and pays the full 95s)
//
// WHAT THIS ASSERTS — what the picker's eye receives, on a driven clock:
//   A · THE REPORTED BUG. 60s after the tap, with the server STILL reporting
//        PENDING, the row must still read PREPARING and show no ✓ Pick button.
//   B · IT IS STILL A TTL, NOT FOREVER. At 150s with the server still
//        disagreeing, the override lets go. Proves A was not bought by making
//        the mask permanent, which would hide real state.
//   C · SERVER AGREES → the override is RETIRED, not merely masked.
//   D · SERVER REFUSES → the row snaps back on the spot, NOT after the TTL.
//        The regression net that keeps a longer TTL from slowing a refusal.
//
// Usage: node diag-pickttl.js
// Before/after proof (A must FAIL on HEAD; B, C, D must PASS on both):
//   git show HEAD:FloorBoard.html > /tmp/board-before.html
//   BOARD_FILE=/tmp/board-before.html node diag-pickttl.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK = require('./mock-tick.js');

// The clean case from mock-tick: A-14 / 194244, PENDING, qty 1.
const ROW_SKU = '194244';

// A Thursday, 10:00 Houston. Working hours ON PURPOSE — off-hours the board
// draws the rest veil and there is no pick list to assert against.
const T0 = '2026-08-13T15:00:00.000Z';

const failures = [];
function check(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  if (!ok) failures.push(`${label}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
}

/**
 * state.serverStatus  what the TICK reports for ROW_SKU ('PENDING'|'PREPARING')
 * state.statusRes     what boardStatus answers
 * Both are read per-request, so a test can change the server's mind mid-run.
 */
async function boot(browser, state) {
  const html = fs.readFileSync(BOARD, 'utf8');
  const ctx = await browser.newContext({
    viewport: { width: 1280, height: 800 }, hasTouch: true, timezoneId: 'America/Chicago'
  });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));

  // Fake clock BEFORE navigation so the first paint sees it (shoot-rest.js pattern).
  await page.clock.install({ time: new Date(T0) });

  const calls = [];
  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      calls.push(body);
      let res = { ok: false, message: 'unknown action' };
      if (body.action === 'boardTick') {
        const t = JSON.parse(JSON.stringify(MOCK));
        // The whole point: the tick can legitimately still say PENDING for up to
        // 95s after a successful write. Model exactly that.
        (t.openOrders || []).forEach(r => {
          if (String(r.sku) === ROW_SKU) r.status = state.serverStatus;
        });
        t.picker  = 'Shipping - Yassin 1';   // gate satisfied; not what we test
        t.pickers = ['Shipping - Yassin 1'];
        res = Object.assign({ ok: true }, t);
      }
      if (body.action === 'boardPickers')   res = { ok: true, pickers: ['Shipping - Yassin 1'] };
      if (body.action === 'boardSetPicker') res = { ok: true };
      if (body.action === 'boardRadio')     res = { ok: true, nowPlaying: '' };
      if (body.action === 'boardStatus')    res = state.statusRes || { ok: true, count: 1 };
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
  await page.clock.runFor(1500);
  await page.waitForTimeout(800);
  return { page, ctx, calls, errs };
}

const clickPick = page => page.evaluate(sku => {
  const li = [...document.querySelectorAll('.pick-row')]
    .find(r => (r.textContent || '').indexOf(sku) !== -1);
  if (!li) throw new Error('row not found: ' + sku);
  const b = li.querySelector('.pick-do');
  if (!b) throw new Error('no ✓ Pick button on row ' + sku);
  b.click();
}, sku => sku);

/**
 * What the PICKER'S EYE receives, not an internal value: does the row wear the
 * PREP marker, and is the ✓ Pick button offering itself again?
 */
const rowStateOf = (page, sku) => page.evaluate(s => {
  const li = [...document.querySelectorAll('.pick-row')]
    .find(r => (r.textContent || '').indexOf(s) !== -1);
  if (!li) return { found: false };
  return {
    found: true,
    prep: !!li.querySelector('.pick-status.prep'),
    pickBtn: !!li.querySelector('.pick-do')
  };
}, sku);
const rowState = page => rowStateOf(page, ROW_SKU);

const clickPickOf = (page, sku) => page.evaluate(s => {
  const li = [...document.querySelectorAll('.pick-row')]
    .find(r => (r.textContent || '').indexOf(s) !== -1);
  if (!li) return false;
  const b = li.querySelector('.pick-do');
  if (!b) return false;
  b.click();
  return true;
}, sku);

/** Advance the page clock and let the poll it triggers actually land. */
async function advance(page, ms) {
  await page.clock.fastForward(ms);
  await page.waitForTimeout(1200);        // real time — for the fetch to resolve
}

(async () => {
  const browser = await chromium.launch();
  console.log(`\nPICK TTL — ${path.basename(BOARD)}\n`);

  // ---- A · THE REPORTED BUG -------------------------------------------------
  // Tap, server accepts, but the tick still says PENDING 60s later — which is
  // ORDINARY, not an error: the publish chain is 45s typical / 95s worst.
  {
    const state = { serverStatus: 'PENDING', statusRes: { ok: true, count: 1 } };
    const { page, ctx, errs } = await boot(browser, state);
    console.log('A · server still reports PENDING 60s after a good pick');

    const before = await rowState(page);
    check('  row starts pickable', { prep: before.prep, btn: before.pickBtn }, { prep: false, btn: true });

    await page.evaluate(sku => {
      const li = [...document.querySelectorAll('.pick-row')]
        .find(r => (r.textContent || '').indexOf(sku) !== -1);
      li.querySelector('.pick-do').click();
    }, ROW_SKU);
    await page.waitForTimeout(600);

    const justAfter = await rowState(page);
    check('  flips optimistically on tap', { prep: justAfter.prep, btn: justAfter.pickBtn }, { prep: true, btn: false });

    await advance(page, 60000);
    const at60 = await rowState(page);
    check('  STILL picked at 60s (the reported bug)',
          { prep: at60.prep, btn: at60.pickBtn }, { prep: true, btn: false });

    errs.forEach(e => failures.push('A ' + e));
    await ctx.close();
  }

  // ---- B · IT IS STILL A TTL ------------------------------------------------
  // A was not bought by making the mask permanent. If the server never agrees,
  // the override must eventually let go rather than hide real state forever.
  {
    const state = { serverStatus: 'PENDING', statusRes: { ok: true, count: 1 } };
    const { page, ctx, errs } = await boot(browser, state);
    console.log('B · the backstop still expires (not a permanent mask)');

    await page.evaluate(sku => {
      const li = [...document.querySelectorAll('.pick-row')]
        .find(r => (r.textContent || '').indexOf(sku) !== -1);
      li.querySelector('.pick-do').click();
    }, ROW_SKU);
    await page.waitForTimeout(600);

    await advance(page, 150000);
    const at150 = await rowState(page);
    check('  override released at 150s', { prep: at150.prep, btn: at150.pickBtn }, { prep: false, btn: true });

    errs.forEach(e => failures.push('B ' + e));
    await ctx.close();
  }

  // ---- C · SERVER AGREES ----------------------------------------------------
  {
    const state = { serverStatus: 'PENDING', statusRes: { ok: true, count: 1 } };
    const { page, ctx, errs } = await boot(browser, state);
    console.log('C · server catches up → override retired, row stays picked');

    await page.evaluate(sku => {
      const li = [...document.querySelectorAll('.pick-row')]
        .find(r => (r.textContent || '').indexOf(sku) !== -1);
      li.querySelector('.pick-do').click();
    }, ROW_SKU);
    await page.waitForTimeout(600);

    state.serverStatus = 'PREPARING';      // the publish finally lands
    await advance(page, 30000);
    const agreed = await rowState(page);
    check('  picked once the tick agrees', { prep: agreed.prep, btn: agreed.pickBtn }, { prep: true, btn: false });

    // And it must SURVIVE past the TTL — because the override was retired, not
    // because it is still masking. This is the half that proves C is real.
    await advance(page, 150000);
    const later = await rowState(page);
    check('  still picked long past the TTL', { prep: later.prep, btn: later.pickBtn }, { prep: true, btn: false });

    errs.forEach(e => failures.push('C ' + e));
    await ctx.close();
  }

  // ---- D · SERVER REFUSES ---------------------------------------------------
  // THE REGRESSION NET for the longer TTL: a refusal must correct the row on the
  // spot. If refusals ever started relying on the TTL, doubling it would have
  // doubled how long the board lies to the floor.
  {
    const state = { serverStatus: 'PENDING', statusRes: { ok: false, error: 'Line not found — it may have shipped' } };
    const { page, ctx, errs } = await boot(browser, state);
    console.log('D · server refuses → row snaps back immediately, not on the TTL');

    await page.evaluate(sku => {
      const li = [...document.querySelectorAll('.pick-row')]
        .find(r => (r.textContent || '').indexOf(sku) !== -1);
      li.querySelector('.pick-do').click();
    }, ROW_SKU);
    await page.waitForTimeout(1500);       // no clock advance — this is "at once"

    const refused = await rowState(page);
    check('  back to pickable without advancing the clock',
          { prep: refused.prep, btn: refused.pickBtn }, { prep: false, btn: true });

    errs.forEach(e => failures.push('D ' + e));
    await ctx.close();
  }

  // ---- E · STRESS: A BURST ON A MULTI-LINE ORDER ---------------------------
  // The realistic shape, and the one the debounce actually bites on. A picker
  // walking order 08-15017-44806 taps four lines inside a minute. Only the FIRST
  // tap publishes inline — publishBoardTickInline is debounced at 15s — so taps
  // 2..4 fall back to the dirty flag and pay the FULL chain. If the TTL were
  // sized for the inline path alone, a burst would quietly un-pick behind them.
  {
    const state = { serverStatus: 'PENDING', statusRes: { ok: true, count: 1 } };
    const { page, ctx, errs } = await boot(browser, state);
    console.log('E · STRESS — four lines of one order, tapped in a burst');

    const BURST = ['167517', '155394', '176701', '190455'];
    for (const sku of BURST) {
      const hit = await clickPickOf(page, sku);
      if (!hit) failures.push(`E could not tap ${sku}`);
      await page.waitForTimeout(120);          // thumb speed, not machine speed
    }
    await page.waitForTimeout(700);

    const flipped = [];
    for (const sku of BURST) flipped.push((await rowStateOf(page, sku)).prep);
    check('  all four flip on tap', flipped, [true, true, true, true]);

    // Server has published NOTHING yet — the whole burst is still unconfirmed.
    await advance(page, 60000);
    const held = [];
    for (const sku of BURST) held.push((await rowStateOf(page, sku)).prep);
    check('  all four STILL picked at 60s', held, [true, true, true, true]);

    // And no cross-contamination: a row nobody touched stays pickable.
    const untouched = await rowStateOf(page, ROW_SKU);
    check('  an untouched row is unaffected',
          { prep: untouched.prep, btn: untouched.pickBtn }, { prep: false, btn: true });

    errs.forEach(e => failures.push('E ' + e));
    await ctx.close();
  }

  // ---- F · STRESS: A DEGRADED PIPE -----------------------------------------
  // THE TRAP UNDER THE TRAP. schedulePoll stretches the poll toward POLL_MAX_MS
  // when ticks are slow, so a FLAT TTL is sized against whatever pipe we had the
  // day we wrote it. Tier 2 failed silently for a week in August and ticks ran
  // 7.5–35s — exactly when the override matters most. pickOverrideUntil must
  // scale with the cadence actually in use.
  //
  // Tested against the REAL function with the REAL constants, not a re-derivation.
  {
    const state = { serverStatus: 'PENDING', statusRes: { ok: true, count: 1 } };
    const { page, ctx, errs } = await boot(browser, state);
    console.log('F · STRESS — the TTL scales when the pipe degrades');

    /* ⚠ FAILS SOFT ON PURPOSE — the diag-pickgate lesson. On an OLD board there
       is no pickOverrideUntil at all, and a throw here would abort the run
       before E and G got to report. A before/after proof is only useful if it
       shows every section's verdict. */
    const hasFn = await page.evaluate(() => typeof window.pickOverrideUntil === 'function');
    if (!hasFn) {
      console.log('  ✗ pickOverrideUntil() does not exist — the TTL is a flat constant');
      failures.push('pickOverrideUntil() missing: the TTL cannot scale with the poll cadence');
      errs.forEach(e => failures.push('F ' + e));
      await ctx.close();
    } else {

    const ttlAt = ms => page.evaluate(v => {
      window.pollLastMs = v;
      return window.pickOverrideUntil() - Date.now();
    }, ms);

    const healthy = await ttlAt(0);            // fresh boot / fast ticks
    check('  healthy pipe holds the 120s floor', Math.round(healthy / 1000), 120);

    const degraded = await ttlAt(35000);       // the observed tier-2-failure range
    check('  35s ticks → 145s', Math.round(degraded / 1000), 145);

    const crawling = await ttlAt(200000);      // clamped by POLL_MAX_MS (90s)
    check('  crawling pipe → 255s, clamped by POLL_MAX_MS', Math.round(crawling / 1000), 255);

    // The property that actually matters, stated as a property: the override
    // must always outlive the publish lag plus two polls at the live cadence.
    const survives = await page.evaluate(() => {
      const chain = 75000 + 2 * Math.min(Math.max(window.POLL_MS, window.pollLastMs || 0), window.POLL_MAX_MS);
      return (window.pickOverrideUntil() - Date.now()) >= chain;
    });
    check('  always outlasts publish lag + two polls', survives, true);

    errs.forEach(e => failures.push('F ' + e));
    await ctx.close();
    }
  }

  // ---- G · STRESS: THE UNDO RIDES THE SAME CHAIN ---------------------------
  // unpick() sets the reverse override. It was never separately considered, and
  // it faces the identical lag — a picker who taps by mistake and corrects it
  // must not watch the row re-pick itself 45s later.
  {
    const state = { serverStatus: 'PENDING', statusRes: { ok: true, count: 1 } };
    const { page, ctx, errs } = await boot(browser, state);
    console.log('G · STRESS — undo survives the same chain');

    await clickPickOf(page, ROW_SKU);
    await page.waitForTimeout(500);
    state.serverStatus = 'PREPARING';          // the pick lands for real
    await advance(page, 30000);
    const picked = await rowState(page);
    check('  picked and confirmed', { prep: picked.prep }, { prep: true });

    // Now undo. The server will keep saying PREPARING for the whole chain.
    const undone = await page.evaluate(() => {
      const u = document.getElementById('ctoastUndo');
      if (u && u.offsetParent !== null) { u.click(); return 'toast'; }
      return false;
    });
    if (!undone) {
      console.log('  · undo toast had already closed — driving unpick directly');
      await page.evaluate(sku => {
        const r = (window.lastTick.openOrders || []).find(x => String(x.sku) === sku);
        window.unpick(r.orderId, r.sku);
      }, ROW_SKU);
    }
    await page.waitForTimeout(700);

    const justUndone = await rowState(page);
    check('  row is pickable again straight away',
          { prep: justUndone.prep, btn: justUndone.pickBtn }, { prep: false, btn: true });

    await advance(page, 60000);                // server still reports PREPARING
    const stillUndone = await rowState(page);
    check('  and does NOT re-pick itself at 60s',
          { prep: stillUndone.prep, btn: stillUndone.pickBtn }, { prep: false, btn: true });

    errs.forEach(e => failures.push('G ' + e));
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
