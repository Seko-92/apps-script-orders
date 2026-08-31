// ============================================================================
// THE FLOOR'S DOOR — /missing and /replacement on the Floor Board.
//
// 2026-08-28: a SALES_ORDER cell was overwritten by hand — `09-15094-35132`
// became `Missing #: 05-15052-93025` — on a row already picked and shelf-counted.
// A slip, not an intrusion: the right intent, the wrong cell. Once All Orders is
// column-locked an employee cannot type into col D at all, and they have no
// sidebar and no ⚙️ menu (measured in incognito), so THIS is the only door they
// have. It goes through doPost, which runs as the owner and writes past the lock.
//
// WHAT THIS ASSERTS — behaviour, not appearance:
//   A · the ⋯ menu carries the door, and it opens a real form
//   B · empty fields are refused ON THE CLIENT — no wasted round trip
//   C · a valid submit sends the right payload, exactly once
//   D · ⭐ a SERVER REFUSAL KEEPS THE TYPED VALUES. The engine's refusals teach
//        ("row 5 is already 212498 on …"), and they are useless if reading one
//        costs you the order id you just typed on a tablet.
//   E · needsPicker opens the picker list and REPLAYS the same payload
//   F · success clears the identity fields but KEEPS the kind — a second missing
//        line for the same shipment is the common follow-up
//   G · toggling Missing ⇄ Replacement preserves typing
//   H · a transport failure says nothing was added, and does NOT open the picker
//        list (protects the 2026-08-13 transport fix)
//
// Usage: node diag-missingline.js
// Before/after: every section must FAIL on HEAD — the feature does not exist there.
//   git show HEAD:FloorBoard.html > /tmp/board-before.html
//   BOARD_FILE=/tmp/board-before.html node diag-missingline.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK = require('./mock-tick.js');
const PICKERS = ['Shipping - Yassin 1', 'Shipping - Miguel 2'];

const failures = [];
function check(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  if (!ok) failures.push(label);
}
function has(label, hay, needle) {
  const ok = String(hay || '').indexOf(needle) !== -1;
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → ${JSON.stringify(String(hay).slice(0, 140))}`));
  if (!ok) failures.push(label);
}
async function section(name, fn) {
  console.log('\n' + name);
  // ⚠ FAIL SOFT — the choosePicker lesson. On an old board nothing here exists,
  // and a throw in section B would hide whether C-H changed at all.
  try { await fn(); }
  catch (e) { failures.push(name + ' THREW'); console.log('  ✗ SECTION THREW (soft): ' + e.message); }
}

/** opts.picker · opts.lineRes (what boardMissingLine answers) · opts.lineFail (abort it) */
async function boot(browser, opts) {
  opts = opts || {};
  const html = fs.readFileSync(BOARD, 'utf8');
  const ctx = await browser.newContext({
    viewport: { width: 1280, height: 800 }, hasTouch: true, timezoneId: 'America/Chicago'
  });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));

  const calls = [];
  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      calls.push(body);
      if (body.action === 'boardMissingLine' && opts.lineFail) {
        return route.abort('internetdisconnected');   // a real dead line
      }
      let res = { ok: false, message: 'unknown action' };
      if (body.action === 'boardTick') {
        const t = JSON.parse(JSON.stringify(MOCK));
        t.picker = opts.picker === undefined ? 'Yassin · 1' : opts.picker;
        t.pickers = PICKERS;
        res = Object.assign({ ok: true }, t);
      }
      if (body.action === 'boardPickers')     res = { ok: true, pickers: PICKERS };
      if (body.action === 'boardSetPicker')   res = { ok: true };
      if (body.action === 'boardRadio')       res = { ok: true, nowPlaying: '' };
      if (body.action === 'boardMissingLine') res = opts.lineRes || { ok: true, message: '✅ Added MISSING line · 212498 ×1 · E-57 · for 05-15052-93025', warnings: [] };
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());
  await page.addInitScript(() => { window.open = function () { return null; }; });

  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.waitForFunction(
    () => !document.getElementById('board').classList.contains('booting'),
    null, { timeout: 20000 }).catch(() => errs.push('board never left booting'));
  await page.waitForTimeout(1200);
  return { page, ctx, calls, errs };
}

// Open the door the way a picker does: ⋯ then the item.
const openDoor = async page => {
  await page.click('#menuBtn');
  await page.waitForTimeout(120);
  await page.click('#missingBtn');
  await page.waitForTimeout(200);
};

const formState = page => page.evaluate(() => {
  const v = id => { const e = document.getElementById(id); return e ? e.value : null; };
  const on = document.querySelector('.rl-seg button.on');
  return {
    open:  document.getElementById('drw').classList.contains('on'),
    title: (document.getElementById('drwTitle').textContent || '').trim(),
    kind:  on ? on.getAttribute('data-rlkind') : null,
    order: v('rlOrder'), sku: v('rlSku'), qty: v('rlQty'), note: v('rlNote'),
    msg:   (document.querySelector('.rl-msg') || {}).textContent || '',
    body:  (document.getElementById('drwBody').textContent || '').trim()
  };
});

const fill = async (page, o) => {
  if (o.order !== undefined) await page.fill('#rlOrder', o.order);
  if (o.sku   !== undefined) await page.fill('#rlSku', o.sku);
  if (o.qty   !== undefined) await page.fill('#rlQty', o.qty);
  if (o.note  !== undefined) await page.fill('#rlNote', o.note);
};
const submit = async page => { await page.click('#rlGo'); await page.waitForTimeout(450); };
const lineCalls = calls => calls.filter(c => c.action === 'boardMissingLine');

(async () => {
  console.log('\nTHE FLOOR\'S DOOR — missing / replacement line\n' + '='.repeat(62));
  console.log('board: ' + BOARD);
  const browser = await chromium.launch();
  browser.contexts;

  await section('A · the ⋯ menu carries the door', async () => {
    const { page, ctx, errs } = await boot(browser, {});
    await openDoor(page);
    const f = await formState(page);
    check('A1 the drawer opens', f.open, true);
    check('A2 titled for the job', f.title, 'Add a line');
    check('A3 defaults to Missing', f.kind, 'missing');
    check('A4 all four fields render', [f.order, f.sku, f.qty, f.note], ['', '', '1', '']);
    check('A5 no page errors', errs, []);
    await ctx.close();
  });

  await section('B · empty fields are refused ON THE CLIENT', async () => {
    const { page, ctx, calls } = await boot(browser, {});
    await openDoor(page);
    await submit(page);
    check('B1 nothing was sent', lineCalls(calls).length, 0);
    has('B2 and it says what is needed', (await formState(page)).msg, 'both required');

    // SKU alone is still incomplete — the order id is what validates.
    await fill(page, { sku: '212498' });
    await submit(page);
    check('B3 sku without an order is still refused', lineCalls(calls).length, 0);
    await ctx.close();
  });

  await section('C · a valid submit sends the right payload, once', async () => {
    const { page, ctx, calls } = await boot(browser, {});
    await openDoor(page);
    await fill(page, { order: '05-15052-93025', sku: '212498', qty: '2', note: 'gaskets only' });
    await submit(page);
    const c = lineCalls(calls);
    check('C1 exactly one call', c.length, 1);
    check('C2 the payload is complete and correct', c[0] && {
      kind: c[0].kind, originalOrder: c[0].originalOrder,
      sku: c[0].sku, qty: c[0].qty, note: c[0].note
    }, { kind: 'missing', originalOrder: '05-15052-93025', sku: '212498', qty: '2', note: 'gaskets only' });
    await ctx.close();
  });

  await section('D · ⭐ a SERVER REFUSAL KEEPS THE TYPED VALUES', async () => {
    const { page, ctx } = await boot(browser, {
      lineRes: { ok: false, message: "That exact line already exists — row 5 is 212498 on 'Missing #: 05-15052-93025'. If you need more units, raise the qty on that row instead of adding a second one." }
    });
    await openDoor(page);
    await fill(page, { order: '05-15052-93025', sku: '212498', qty: '3', note: 'keep me' });
    await submit(page);
    const f = await formState(page);
    has('D1 the refusal is shown in full', f.msg, 'row 5');
    has('D2 ...including what to do instead', f.msg, 'raise the qty');
    check('D3 the order id survived', f.order, '05-15052-93025');
    check('D4 the sku survived', f.sku, '212498');
    check('D5 the qty survived', f.qty, '3');
    check('D6 the note survived', f.note, 'keep me');
    await ctx.close();
  });

  await section('E · needsPicker opens the list and REPLAYS', async () => {
    let served = 0;
    const { page, ctx, calls } = await boot(browser, {
      picker: '',
      get lineRes() { served++; return served === 1
        ? { ok: false, needsPicker: true, error: 'Set the Pick ID first — every pick is filed under a name.' }
        : { ok: true, message: '✅ Added MISSING line · 212498 ×1', warnings: [] }; }
    });
    await openDoor(page);
    await fill(page, { order: '05-15052-93025', sku: '212498' });
    await submit(page);
    has('E1 the picker list opened', (await formState(page)).body, 'Shipping - Yassin 1');

    await page.click('[data-picker="Shipping - Yassin 1"]');
    await page.waitForTimeout(900);
    const c = lineCalls(calls);
    check('E2 the line was attempted twice — original + replay', c.length, 2);
    check('E3 the replay carried the SAME payload', c[1] && [c[1].originalOrder, c[1].sku],
          ['05-15052-93025', '212498']);
    await ctx.close();
  });

  await section('F · success clears identity, KEEPS the kind', async () => {
    const { page, ctx } = await boot(browser, {});
    await openDoor(page);
    await page.click('[data-rlkind="replacement"]');
    await page.waitForTimeout(120);
    await fill(page, { order: '19-14597-26309', sku: '171378', qty: '2', note: 'studs' });
    await submit(page);
    const f = await formState(page);
    has('F1 the success is shown', f.msg, 'Added');
    check('F2 order cleared', f.order, '');
    check('F3 sku cleared', f.sku, '');
    check('F4 qty back to 1', f.qty, '1');
    check('F5 ⭐ the KIND is kept for the next line', f.kind, 'replacement');
    await ctx.close();
  });

  await section('G · the kind toggle preserves typing', async () => {
    const { page, ctx } = await boot(browser, {});
    await openDoor(page);
    await fill(page, { order: '05-15052-93025', sku: '212498', note: 'do not lose me' });
    await page.click('[data-rlkind="replacement"]');
    await page.waitForTimeout(150);
    const f = await formState(page);
    check('G1 switched kind', f.kind, 'replacement');
    check('G2 the order survived the toggle', f.order, '05-15052-93025');
    check('G3 the sku survived', f.sku, '212498');
    check('G4 the note survived', f.note, 'do not lose me');
    await ctx.close();
  });

  await section('H · a transport failure is honest, and not the picker\'s fault', async () => {
    const { page, ctx } = await boot(browser, { lineFail: true });
    await openDoor(page);
    await fill(page, { order: '05-15052-93025', sku: '212498' });
    await page.click('#rlGo');
    await page.waitForTimeout(2500);
    const f = await formState(page);
    has('H1 it says nothing was added', f.msg, 'nothing was added');
    check('H2 it did NOT open the picker list', f.body.indexOf('Shipping - Yassin') === -1, true);
    check('H3 the typed values survived', [f.order, f.sku], ['05-15052-93025', '212498']);
    await ctx.close();
  });

  await browser.close();
  console.log('\n' + '='.repeat(62));
  if (failures.length) {
    console.log('❌ ' + failures.length + ' FAILED\n' + failures.map(f => '   · ' + f).join('\n') + '\n');
    process.exit(1);
  }
  console.log('✅ all assertions passed\n');
})();
