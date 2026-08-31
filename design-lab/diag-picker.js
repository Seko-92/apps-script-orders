// ============================================================================
// WHO IS PICKING — the footer Pick ID control.
//
// Floor report 2026-08-13: "the picker should be able to choose the Pick ID
// from the tablet." The capability (askForPicker → boardPickers →
// boardSetPicker) was already built, tested and live — but its ONLY door was
// tapping PRINT and being refused, so the floor believed it was PC-only.
//
// This asserts the door works AND that opening it repeatedly is safe:
//   · the footer chip is a real button with a real touch target
//   · unset is LOUD (it blocks printing)
//   · tapping it opens the list WITHOUT starting a print
//   · picking a name calls boardSetPicker EXACTLY ONCE and never claims a
//     window (with no print to follow, a claimed window is an orphan tab)
//   · ⚠ THE REGRESSION THAT MATTERS: opening the list N times then picking
//     once must still be ONE call. drwBody persists across innerHTML swaps,
//     so a per-call addEventListener stacks. Proven by opening 3× first.
//   · the PRINT path still chains into printing after choosing
//
// Usage: node diag-picker.js
// Before/after proof:
//   git show HEAD:FloorBoard.html > /tmp/board-before.html
//   BOARD_FILE=/tmp/board-before.html node diag-picker.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK  = require('./mock-tick.js');
const OUT   = path.join(__dirname, 'renders');
const PICKERS = ['Shipping - Yassin 1', 'Shipping - Miguel 2', 'Shipping - Ana 3'];

const SIZES = [['tablet-portrait', 800, 1280], ['tablet-landscape', 1280, 800]];
const failures = [];

async function boot(browser, w, h, opts) {
  opts = opts || {};
  const html = fs.readFileSync(BOARD, 'utf8');
  const ctx = await browser.newContext({
    viewport: { width: w, height: h }, hasTouch: true, timezoneId: 'America/Chicago'
  });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));
  page.on('console', m => {
    if (m.type() !== 'error') return;
    const t = m.text();
    if (/Failed to load resource|ERR_FAILED|ERR_ABORTED/.test(t)) return;
    errs.push(t);
  });

  // Count every action the page fires, so "exactly once" is measurable.
  const calls = [];
  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      calls.push(body.action);
      let res = { ok: false, message: 'unknown action' };
      if (body.action === 'boardTick') {
        const t = JSON.parse(JSON.stringify(MOCK));
        t.picker = opts.picker || '';
        // The tick carries the Pick ID list from 2026-08-14. Absent = a board
        // reading an older published payload, which must still work.
        if (opts.tickPickers) t.pickers = opts.tickPickers;
        res = Object.assign({ ok: true }, t);
      }
      // Failure injection — reproduces the measured 2026-08-13 fault: the n8n
      // proxy answers HTTP 200 carrying {error:"proxy: …404"} when the hop to
      // Apps Script drops. `opts.failPickers` fails that many boardPickers
      // calls before answering properly.
      if (body.action === 'boardPickers') {
        if (opts.failPickers && opts.failPickers > 0) {
          opts.failPickers--;
          res = { error: 'proxy: Request failed with status code 404' };
        } else {
          res = { ok: true, pickers: PICKERS };
        }
      }
      if (body.action === 'boardSetPicker') {
        if (opts.failSetPicker && opts.failSetPicker > 0) {
          opts.failSetPicker--;
          res = { error: 'proxy: Request failed with status code 404' };
        } else {
          res = { ok: true };
        }
      }
      if (body.action === 'boardPrint')     res = { ok: false, needsPicker: true };
      if (body.action === 'boardStatus')    res = { ok: true };
      if (body.action === 'boardRadio')     res = { ok: true, nowPlaying: '' };
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());

  // Record window.open without ever letting a real tab appear.
  await page.addInitScript(() => {
    window.__opened = 0;
    const real = window.open;
    window.open = function () { window.__opened++; return null; };
    window.__realOpen = real;
  });

  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.waitForFunction(() => !document.getElementById('board').classList.contains('booting'),
    null, { timeout: 20000 }).catch(() => errs.push('board never left booting'));
  await page.waitForTimeout(1500);
  return { page, ctx, calls, errs };
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch();

  for (const [name, w, h] of SIZES) {
    console.log(`\n${'='.repeat(68)}\n  ${name} (${w}×${h})\n${'='.repeat(68)}`);
    const p = [];

    // ── A. the chip itself ────────────────────────────────────────────────
    let { page, ctx, calls, errs } = await boot(browser, w, h, { picker: '' });
    const chip = await page.evaluate(() => {
      const el = document.getElementById('footPicker');
      if (!el) return null;
      const b = el.getBoundingClientRect();
      const cs = getComputedStyle(el);
      return { tag: el.tagName, unset: el.classList.contains('unset'),
               text: el.textContent.trim(), w: Math.round(b.width), h: Math.round(b.height),
               bottom: Math.round(b.bottom), cursor: cs.cursor, border: cs.borderTopColor };
    });
    if (!chip) p.push('#footPicker missing entirely');
    else {
      if (chip.tag !== 'BUTTON') p.push(`#footPicker is a <${chip.tag}>, not a button — not tappable`);
      if (chip.h < 40) p.push(`touch target only ${chip.h}px tall (want ≥40)`);
      if (!chip.unset) p.push('no picker set but the chip is not in its loud .unset state');
      if (chip.cursor !== 'pointer') p.push(`cursor is "${chip.cursor}" — does not read as tappable`);
    }
    console.log(`  chip: <${chip && chip.tag}> "${chip && chip.text}"  ${chip && chip.w}×${chip && chip.h}`
              + `  unset=${chip && chip.unset}  border=${chip && chip.border}`);

    // ── B. tap it: the list opens and NOTHING prints ──────────────────────
    calls.length = 0;
    await page.click('#footPicker');
    await page.waitForTimeout(900);
    const opened = await page.evaluate(() => ({
      drwOn: document.getElementById('drw').classList.contains('on'),
      title: document.getElementById('drwTitle').textContent.trim(),
      sub: document.getElementById('drwSub').textContent.trim(),
      names: [...document.querySelectorAll('[data-picker]')].length,
      windows: window.__opened
    }));
    if (!opened.drwOn) p.push('tapping the chip did not open the drawer');
    if (opened.title !== 'Who is picking?') p.push(`drawer title is "${opened.title}"`);
    if (opened.names !== PICKERS.length) p.push(`expected ${PICKERS.length} names, got ${opened.names}`);
    if (calls.includes('boardPrint')) p.push('tapping the chip STARTED A PRINT');
    if (opened.windows !== 0) p.push(`claimed ${opened.windows} window(s) with no print to follow (orphan tab)`);
    console.log(`  tap → "${opened.title}" · "${opened.sub}" · ${opened.names} names`
              + `   printed=${calls.includes('boardPrint')}  windows=${opened.windows}`);

    // ── C. ⚠ THE STACKING REGRESSION — reopen 3×, then pick ONCE ──────────
    for (let i = 0; i < 3; i++) {
      await page.evaluate(() => document.getElementById('drwX').click());
      await page.waitForTimeout(150);
      await page.click('#footPicker');
      await page.waitForTimeout(700);
    }
    calls.length = 0;
    await page.evaluate(() => window.__opened = 0);
    await page.click('[data-picker]');
    // ⚠ Wait long enough for at least one FULL tick to land. The mock keeps
    // returning picker:'' forever — i.e. a permanently stale server view — so
    // if the chip still shows the chosen name here, the override is genuinely
    // surviving stale ticks rather than just winning a repaint race.
    await page.waitForTimeout(4000);
    const after = await page.evaluate(() => ({
      sets: null, drwOn: document.getElementById('drw').classList.contains('on'),
      chipText: document.getElementById('footPicker').textContent.trim(),
      chipUnset: document.getElementById('footPicker').classList.contains('unset'),
      windows: window.__opened
    }));
    const setCalls = calls.filter(a => a === 'boardSetPicker').length;
    if (setCalls !== 1)
      p.push(`boardSetPicker fired ${setCalls}× after 4 opens — listeners are STACKING`);
    if (calls.includes('boardPrint')) p.push('choosing a name started a print (should only confirm)');
    if (after.windows !== 0) p.push(`claimed ${after.windows} orphan window(s) on a non-print pick`);
    if (after.drwOn) p.push('drawer stayed open after choosing');
    if (after.chipUnset) p.push('picker chosen but the chip is still in its unset state');
    if (!/Yassin|Miguel|Ana/.test(after.chipText)) p.push(`chip did not show the chosen name ("${after.chipText}")`);
    // Every Pick ID starts "Shipping - ", so showing it burns the chip's whole
    // width on a constant and truncates the only part that names anyone.
    if (/shipping\s*-/i.test(after.chipText))
      p.push(`chip still shows the constant "Shipping -" prefix ("${after.chipText}")`);
    console.log(`  4 opens → 1 pick: boardSetPicker ×${setCalls}   windows=${after.windows}`
              + `   drawer closed=${!after.drwOn}   chip="${after.chipText}"`);

    await page.screenshot({ path: path.join(OUT, `picker-${name}.png`) });
    if (errs.length) p.push('JS: ' + errs.join(' | '));
    await ctx.close();

    // ── D. the PRINT path must still chain into printing ──────────────────
    const second = await boot(browser, w, h, { picker: '' });
    second.calls.length = 0;
    await second.page.click('#printBtn');
    await second.page.waitForTimeout(1400);
    const askedTitle = await second.page.evaluate(() =>
      document.getElementById('drwTitle').textContent.trim());
    if (askedTitle !== 'Who is picking?') p.push(`print refusal did not offer the list (drawer said "${askedTitle}")`);
    second.calls.length = 0;
    await second.page.click('[data-picker]');
    await second.page.waitForTimeout(1600);
    const printCalls = second.calls.filter(a => a === 'boardPrint').length;
    const setCalls2  = second.calls.filter(a => a === 'boardSetPicker').length;
    if (setCalls2 !== 1) p.push(`print path: boardSetPicker fired ${setCalls2}×`);
    if (printCalls < 1)  p.push('print path: choosing a picker did NOT chain back into printing');
    console.log(`  print path → refused, offered list, chose: setPicker ×${setCalls2}, print ×${printCalls}`);
    if (second.errs.length) p.push('JS(print path): ' + second.errs.join(' | '));
    await second.ctx.close();

    // ── E. THE 404 BURST — a read must retry, a write must NOT ────────────
    // Measured live 2026-08-13: the proxy returns HTTP 200 with an error body
    // when the Apps Script hop 404s. The board used to read that as "the sheet
    // says there are no pickers" and told the floor to go fix a dropdown that
    // was already correct.
    const flaky = await boot(browser, w, h, { picker: '', failPickers: 1 });
    flaky.calls.length = 0;
    await flaky.page.click('#footPicker');
    await flaky.page.waitForTimeout(2600);
    const recovered = await flaky.page.evaluate(() => ({
      names: [...document.querySelectorAll('[data-picker]')].length,
      body: document.getElementById('drwBody').textContent
    }));
    const pickerTries = flaky.calls.filter(a => a === 'boardPickers').length;
    if (pickerTries !== 2) p.push(`one 404 should be retried once (saw ${pickerTries} boardPickers calls)`);
    if (recovered.names !== PICKERS.length) p.push('did not recover from a single 404');
    if (/No pickers configured/i.test(recovered.body)) p.push('blamed the sheet for a transport failure');
    console.log(`  one 404 → retried ${pickerTries}×, recovered with ${recovered.names} names`);
    await flaky.ctx.close();

    // Both attempts fail → the message must be honest about WHY.
    const dead = await boot(browser, w, h, { picker: '', failPickers: 9 });
    await dead.page.click('#footPicker');
    await dead.page.waitForTimeout(3000);
    const deadBody = await dead.page.evaluate(() =>
      document.getElementById('drwBody').textContent);
    if (/No pickers configured/i.test(deadBody))
      p.push('persistent transport failure still blamed on the sheet dropdown');
    if (!/could not reach|connection dropped/i.test(deadBody))
      p.push(`transport failure message unclear: "${deadBody.trim().slice(0, 80)}"`);
    console.log(`  persistent 404 → "${deadBody.trim().slice(0, 66)}…"`);
    await dead.ctx.close();

    // A WRITE must never be retried — a duplicate adjustment is worse than a
    // visible failure. boardSetPicker stands in for the whole write family.
    const wr = await boot(browser, w, h, { picker: '', failSetPicker: 1 });
    await wr.page.click('#footPicker');
    await wr.page.waitForTimeout(1200);
    wr.calls.length = 0;
    await wr.page.click('[data-picker]');
    await wr.page.waitForTimeout(2200);
    const setTries = wr.calls.filter(a => a === 'boardSetPicker').length;
    if (setTries !== 1) p.push(`WRITE was retried ${setTries}× — writes must never auto-retry`);
    console.log(`  write under 404 → boardSetPicker attempted ${setTries}× (must be 1)`);
    await wr.ctx.close();

    // ── F. THE INSTANT PATH — the tick carries the list, so ZERO calls ────
    // boardPickers cost 3-6s live, ~3s of which is just reaching Apps Script
    // through n8n, for four strings that change roughly never.
    const fast = await boot(browser, w, h, { picker: '', tickPickers: PICKERS });
    fast.calls.length = 0;
    const t0 = Date.now();
    await fast.page.click('#footPicker');
    await fast.page.waitForFunction(
      () => document.querySelectorAll('[data-picker]').length > 0,
      null, { timeout: 5000 }).catch(() => {});
    const openMs = Date.now() - t0;
    const fastState = await fast.page.evaluate(() => ({
      names: [...document.querySelectorAll('[data-picker]')].length,
      drwOn: document.getElementById('drw').classList.contains('on')
    }));
    const fetched = fast.calls.filter(a => a === 'boardPickers').length;
    if (fastState.names !== PICKERS.length) p.push('instant path did not render the list');
    if (!fastState.drwOn) p.push('instant path did not open the drawer');
    if (fetched !== 0) p.push(`instant path still called boardPickers ${fetched}×`);
    // Picking from the instant list must still write exactly once.
    fast.calls.length = 0;
    await fast.page.click('[data-picker]');
    await fast.page.waitForTimeout(1500);
    const fastSet = fast.calls.filter(a => a === 'boardSetPicker').length;
    if (fastSet !== 1) p.push(`instant path fired boardSetPicker ${fastSet}×`);
    console.log(`  instant path → ${fastState.names} names in ${openMs}ms, `
              + `boardPickers ×${fetched}, boardSetPicker ×${fastSet}`);
    await fast.ctx.close();

    if (p.length) { failures.push(`${name} — ${p.join('; ')}`); p.forEach(x => console.log(`      ✗ ${x}`)); }
    else console.log('  ✓ all good');
  }

  await browser.close();
  console.log('\n' + '='.repeat(68));
  if (failures.length) {
    console.log('✗ FAILURES:\n' + failures.map(f => '   ' + f).join('\n'));
    process.exit(1);
  }
  console.log('✓ PICK ID: reachable from the footer, loud when unset, opens without');
  console.log('  printing, fires exactly once however many times it is reopened,');
  console.log('  claims no orphan window, and the print path still chains through.');
})().catch(e => { console.error('CRASH', e); process.exit(1); });
