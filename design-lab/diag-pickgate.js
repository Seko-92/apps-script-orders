// ============================================================================
// WHO PICKED IT — the attribution gate on ✓ Pick and the shelf count.
//
// Floor report 2026-08-14: "the workers started picking while no one picked the
// picker." Printing has refused without a Pick ID since 2026-05-01 and stock
// adjustment since 2026-08-11, but ✓ Pick — by far the most frequent thing the
// floor does — never asked. So a shift that happened not to print logged a
// whole day of picks against nobody, and the only warning was an amber chip
// that everyone had learned to ignore.
//
// ⚠ AND IT IS NOT AN OCCASIONAL STATE: resetDailyPickIds blanks F2 at 4am ON
// PURPOSE, so EVERY morning starts with no picker and the first action of the
// day is a pick, not a print.
//
// WHAT THIS ASSERTS — behaviour, not appearance:
//   A · no picker set  → ✓ Pick WRITES NOTHING and the picker list opens
//                        instead, and the row is NOT left optimistically
//                        flipped (no flip-then-snap-back)
//   B · choosing a name REPLAYS the exact pick — same order, same SKU, exactly
//        once — so the tap the picker already made is never lost, and no
//        window is claimed (with no print to follow that is an orphan tab)
//   C · picker already set → ZERO extra taps, ZERO extra calls. THE REGRESSION
//        NET: this gate must not slow down the floor's most frequent action.
//   D · server backstop — the client thought it had a picker, the server says
//        needsPicker → the list opens rather than a lie about the connection
//   E · a REAL transport failure must NOT open the picker list and must not
//        blame the picker. Protects the 2026-08-13 transport fix from the
//        2026-08-14 change that taught hqCall to tell the two apart.
//   F · the shelf count is gated the same way, and replays the typed number
//
// Usage: node diag-pickgate.js
// Before/after proof (A, B, D, F must FAIL on HEAD; C and E must PASS):
//   git show HEAD:FloorBoard.html > /tmp/board-before.html
//   BOARD_FILE=/tmp/board-before.html node diag-pickgate.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK = require('./mock-tick.js');
const PICKERS = ['Shipping - Yassin 1', 'Shipping - Miguel 2', 'Shipping - Ana 3'];

// The clean case from mock-tick: A-14 / 194244, PENDING, hand 9, qty 1.
const ROW_SKU = '194244';
const ROW_ORDER = '24-15021-77421';

const failures = [];
function check(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  if (!ok) failures.push(`${label}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
}
function checkTruthy(label, got) {
  const ok = !!got;
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → got ${JSON.stringify(got)}`));
  if (!ok) failures.push(`${label}: got ${JSON.stringify(got)}`);
}

/**
 * opts.picker      what the TICK says the picker is ('' = nobody)
 * opts.statusRes   what boardStatus answers (default success)
 * opts.leftRes     what boardLeft answers (default success)
 */
async function boot(browser, opts) {
  opts = opts || {};
  const html = fs.readFileSync(BOARD, 'utf8');
  const ctx = await browser.newContext({
    viewport: { width: 1280, height: 800 }, hasTouch: true, timezoneId: 'America/Chicago'
  });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));

  const calls = [];                       // every {action, body} the board fired
  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      calls.push(body);
      let res = { ok: false, message: 'unknown action' };
      if (body.action === 'boardTick') {
        const t = JSON.parse(JSON.stringify(MOCK));
        t.picker = opts.picker === undefined ? '' : opts.picker;
        t.pickers = PICKERS;              // rides on the tick since 2026-08-14
        res = Object.assign({ ok: true }, t);
      }
      if (body.action === 'boardPickers')   res = { ok: true, pickers: PICKERS };
      if (body.action === 'boardSetPicker') res = { ok: true };
      if (body.action === 'boardRadio')     res = { ok: true, nowPlaying: '' };
      if (body.action === 'boardStatus')    res = opts.statusRes || { ok: true };
      if (body.action === 'boardLeft')      res = opts.leftRes || { ok: true, count: 0 };
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());

  // Record window.open without ever letting a real tab appear.
  await page.addInitScript(() => {
    window.__opened = 0;
    window.open = function () { window.__opened++; return null; };
  });

  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.waitForFunction(
    () => !document.getElementById('board').classList.contains('booting'),
    null, { timeout: 20000 }
  ).catch(() => errs.push('board never left booting'));
  await page.waitForTimeout(1400);
  return { page, ctx, calls, errs };
}

const clickPick = page => page.evaluate(sku => {
  const li = [...document.querySelectorAll('.pick-row')]
    .find(r => (r.textContent || '').indexOf(sku) !== -1);
  if (!li) throw new Error('row not found: ' + sku);
  const b = li.querySelector('.pick-do');
  if (!b) throw new Error('no ✓ Pick button on row ' + sku);
  b.click();
}, ROW_SKU);

const drawerState = page => page.evaluate(() => ({
  open: document.getElementById('drw').classList.contains('on'),
  title: (document.getElementById('drwTitle').textContent || '').trim(),
  sub: (document.getElementById('drwSub').textContent || '').trim(),
  bodyText: (document.getElementById('drwBody').textContent || '').trim()
}));

const rowIsPrep = page => page.evaluate(sku => {
  const li = [...document.querySelectorAll('.pick-row')]
    .find(r => (r.textContent || '').indexOf(sku) !== -1);
  return !!(li && li.querySelector('.pick-status.prep'));
}, ROW_SKU);

/**
 * ⚠ FAILS SOFT ON PURPOSE. On an OLD board the drawer never opens, so there is
 * no name to tap — and a throw here would abort the run before the REGRESSION
 * NETS (C and E) got to prove they still pass. A before/after proof is only
 * useful if it shows which sections changed AND which did not.
 */
const choosePicker = async (page, who) => {
  const found = await page.evaluate(w => {
    const b = document.querySelector(`#drwBody [data-picker="${w}"]`);
    if (!b) return false;
    b.click();
    return true;
  }, who);
  if (!found) {
    console.log(`  ✗ picker "${who}" was never offered — the drawer did not open`);
    failures.push(`picker "${who}" was never offered`);
    return false;
  }
  await page.waitForTimeout(900);
  return true;
};

const only = (calls, action) => calls.filter(c => c.action === action);

(async () => {
  const browser = await chromium.launch();
  console.log(`\nBOARD: ${BOARD}\n`);

  // ── A · NO PICKER SET → ✓ Pick must write NOTHING ─────────────────────────
  {
    console.log('='.repeat(70));
    console.log('  A · no picker set — ✓ Pick must refuse and ask instead');
    console.log('='.repeat(70));
    const { page, ctx, calls, errs } = await boot(browser, { picker: '' });
    calls.length = 0;
    await clickPick(page);
    await page.waitForTimeout(900);

    // ⭐ THE HEADLINE. This is the exact bug the floor reported: the write went
    // through with nobody's name on it.
    check('boardStatus NOT called', only(calls, 'boardStatus').length, 0);
    const d = await drawerState(page);
    check('picker list opened instead', d.open, true);
    check('drawer asks who is picking', d.title, 'Who is picking?');
    check('and says why', d.sub, 'required before picking');
    checkTruthy('the list is populated', d.bodyText.indexOf('Yassin') !== -1);
    // No round trip: the Pick ID list rides on the tick.
    check('no boardPickers round trip', only(calls, 'boardPickers').length, 0);
    // The row must not flip and snap back — the gate fires BEFORE the
    // optimistic write, so the list never lies about what happened.
    check('row NOT left flipped to PREPARING', await rowIsPrep(page), false);
    check('no page errors', errs, []);

    // ── B · choosing a name REPLAYS the pick ────────────────────────────────
    console.log('\n' + '-'.repeat(70));
    console.log('  B · choose a picker → the refused pick replays, exactly once');
    console.log('-'.repeat(70));
    calls.length = 0;
    if (await choosePicker(page, PICKERS[0])) {
      check('boardSetPicker fired once', only(calls, 'boardSetPicker').length, 1);
      const st = only(calls, 'boardStatus');
      check('the pick REPLAYED exactly once', st.length, 1);
      if (st.length === 1) {
        check('replayed the right order', st[0].orderId, ROW_ORDER);
        check('replayed the right SKU', st[0].sku, ROW_SKU);
        check('replayed as PREPARING', st[0].status, 'PREPARING');
      }
      // With no print to follow, a claimed window is an orphan blank tab.
      check('no window claimed', await page.evaluate(() => window.__opened), 0);
      check('drawer closed', (await drawerState(page)).open, false);
      const chip = await page.evaluate(() => document.getElementById('footPicker').textContent.trim());
      checkTruthy('chip now names the picker', chip.indexOf('Yassin') !== -1);
      check('row is now PREPARING', await rowIsPrep(page), true);
    }
    await ctx.close();
  }

  // ── C · REGRESSION NET: picker set → zero friction ────────────────────────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  C · picker already set — the normal path must not change');
    console.log('='.repeat(70));
    const { page, ctx, calls, errs } = await boot(browser, { picker: 'Yassin · 1' });
    calls.length = 0;
    await clickPick(page);
    await page.waitForTimeout(900);

    check('pick went straight through', only(calls, 'boardStatus').length, 1);
    check('nobody was asked anything', only(calls, 'boardSetPicker').length, 0);
    check('no drawer in the way', (await drawerState(page)).open, false);
    check('row flipped to PREPARING', await rowIsPrep(page), true);
    check('no page errors', errs, []);
    await ctx.close();
  }

  // ── D · SERVER BACKSTOP — stale client view ───────────────────────────────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  D · client thought it had a picker, server refused');
    console.log('='.repeat(70));
    // F2 cleared mid-shift, or a tick old enough to still name someone.
    const { page, ctx, calls, errs } = await boot(browser, {
      picker: 'Yassin · 1',
      statusRes: { ok: false, needsPicker: true,
                   error: 'Set the Pick ID first — every pick is filed under a name.' }
    });
    calls.length = 0;
    await clickPick(page);
    await page.waitForTimeout(1100);

    check('the write was attempted', only(calls, 'boardStatus').length, 1);
    const d = await drawerState(page);
    // ⚠ THIS IS WHAT THE hqCall CHANGE BUYS. Before 2026-08-14 a refusal
    // carrying an `error` key was thrown as a TRANSPORT failure, so the board
    // told the floor the connection had dropped when the server had in fact
    // answered clearly. The drawer could never open from this path.
    check('picker list opened, not a connection lie', d.open, true);
    check('and says why', d.sub, 'required before picking');
    check('optimistic flip rolled back', await rowIsPrep(page), false);
    check('no page errors', errs, []);
    await ctx.close();
  }

  // ── E · REGRESSION NET: a real transport failure is still transport ───────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  E · a genuine proxy failure must NOT blame the picker');
    console.log('='.repeat(70));
    // The measured 2026-08-13 fault: HTTP 200 carrying {error:"proxy: …"} and
    // NO `ok` key. That absence is exactly what tells it apart from a refusal.
    const { page, ctx, calls, errs } = await boot(browser, {
      picker: 'Yassin · 1',
      statusRes: { error: 'proxy: Request failed with status code 404' }
    });
    calls.length = 0;
    await clickPick(page);
    await page.waitForTimeout(1100);

    check('picker list did NOT open', (await drawerState(page)).open, false);
    check('nobody was asked to identify themselves', only(calls, 'boardSetPicker').length, 0);
    check('WRITES ARE NEVER RETRIED', only(calls, 'boardStatus').length, 1);
    check('optimistic flip rolled back', await rowIsPrep(page), false);
    check('no page errors', errs, []);
    await ctx.close();
  }

  // ── F · the shelf count is gated too, and replays the number ──────────────
  {
    console.log('\n' + '='.repeat(70));
    console.log('  F · shelf count — gated, and the typed number survives');
    console.log('='.repeat(70));
    const { page, ctx, calls, errs } = await boot(browser, { picker: '' });
    calls.length = 0;

    await page.evaluate(sku => {
      const li = [...document.querySelectorAll('.pick-row')]
        .find(r => (r.textContent || '').indexOf(sku) !== -1);
      li.querySelector('.pc-btn').click();
    }, ROW_SKU);
    await page.waitForTimeout(500);
    check('numpad opened', await page.evaluate(
      () => !document.getElementById('numPad').classList.contains('hidden')), true);

    // Type 7, confirm.
    await page.evaluate(() => {
      const k = [...document.querySelectorAll('.np-key')].find(b => b.textContent.trim() === '7');
      k.click();
    });
    await page.click('#npOk');
    await page.waitForTimeout(900);

    check('boardLeft NOT called', only(calls, 'boardLeft').length, 0);
    const d = await drawerState(page);
    check('picker list opened instead', d.open, true);
    check('and says why', d.sub, 'required before saving a count');
    check('numpad closed first (no stacked modals)', await page.evaluate(
      () => document.getElementById('numPad').classList.contains('hidden')), true);

    calls.length = 0;
    if (await choosePicker(page, PICKERS[1])) {
      const left = only(calls, 'boardLeft');
      check('the count REPLAYED exactly once', left.length, 1);
      if (left.length === 1) {
        check('the typed number survived', String(left[0].count), '7');
        check('on the right SKU', left[0].sku, ROW_SKU);
      }
    }
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
