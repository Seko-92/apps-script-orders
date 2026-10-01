// ============================================================================
// PARTS FINDER ON THE FLOOR BOARD (diag-find.js, 2026-10-01)
// The tablet had no search. Find → the same one box as the sidebar (findParts on
// the server picks SKU / part numbers / words) → tap a result → the board's own
// three-stage part drawer, with "← results" back. Drives the REAL board over a
// routed /api/board, so every request the board makes is visible here.
//
// Usage: node diag-find.js
// Before/after:  git show HEAD:FloorBoard.html > /tmp/board-before.html
//                BOARD_FILE=/tmp/board-before.html node diag-find.js
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

const M = (sku, title, loc, qty, price, extra) => Object.assign({ sku, title, location: loc, available: qty, price,
  active: true, status: 'Active', isKit: false, url: '', image: '', via: 'main' }, extra || {});
const KW = { ok: true, mode: 'keywords', auto: true, ms: 3400, query: 'v2203 piston', total: 3, matches: [
  M('166527', 'Piston With Rings STD For Kubota, 16423-21110, V2203 IDI', 'E-84', 8, 79.99),
  M('173817', 'Piston With Rings 0.50 For Kubota, 16423-21910, V2203 IDI', 'E-54', 0, 74),
  M('215756', 'Engine Overhaul Kit 0.50 For Kubota V2203', 'K-12', 1, 440, { isKit: true, active: false, status: 'Completed' }) ] };
const MPN = { ok: true, mode: 'mpn', auto: true, ms: 3900, found: 1, missing: 1, results: [
  { query: '04270701', matches: [M('207069', 'Piston With Ring 0.50 For Deutz, 04270701, BF1011', 'E-11', 99, 99, { via: 'extra MPN' })] },
  { query: '04179234', matches: [] } ] };
const DOSSIER = sku => ({ ok: true, dossier: { sku, found: true, isKit: false, usedIn: [], unblock: [],
  part: { title: 'Piston With Rings STD For Kubota', available: 8, location: 'E-84' },
  sizes: [{ sku: '166527', size: 'STD', self: sku === '166527', num: '16423-21110', location: 'E-84', available: 8, active: true },
          { sku: '173817', size: '0.50', self: sku === '173817', num: '16423-21910', location: 'E-54', available: 0, active: true }],
  identity: { numbers: [{ num: '16423-21110', main: true, via: '' }, { num: '1A091-21110', main: false, via: '' }],
              engines: ['V2203', 'D1703'], brands: ['Kubota'], machines: [] } } });

const CALLS = [];
let MODE = 'normal';
let SLOW_LITE = 0;   // ms — make stage 2 land AFTER stage 3
function answer(body) {
  if (body.action === 'boardTick') return Object.assign({ ok: true }, JSON.parse(JSON.stringify(MOCK)));
  if (body.action === 'boardRadio') return { ok: true, nowPlaying: '' };
  if (body.action === 'boardPartLite') return { ok: true, basics: { sku: body.sku, found: true, title: 'Piston ' + body.sku,
    location: 'E-84', images: [], available: 8, ebayPrice: 79.99 } };
  if (body.action === 'boardPart') return DOSSIER(body.sku);
  if (body.action === 'boardFind') {
    CALLS.push({ text: body.text, force: body.force || '' });
    if (MODE === 'refuse') return { ok: false, reason: 'No part numbers found in what was pasted.' };
    if (body.force === 'mpn' || /^0\d{7}/.test(body.text)) return MPN;
    if (/^\d{6}$/.test(body.text)) return { ok: true, mode: 'sku', auto: true, sku: body.text, text: body.text };
    return KW;
  }
  return { ok: false };
}

async function boot(browser) {
  const html = fs.readFileSync(BOARD, 'utf8');
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 800 }, hasTouch: true, timezoneId: 'America/Chicago' });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));
  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      if (MODE === 'down' && body.action === 'boardFind') { CALLS.push({ text: body.text, force: body.force || '' }); return route.abort('internetdisconnected'); }
      const send = () => route.fulfill({ contentType: 'application/json', body: JSON.stringify(answer(body)) });
      if (SLOW_LITE && body.action === 'boardPartLite') return setTimeout(send, SLOW_LITE);
      return send();
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());
  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.waitForFunction(() => !document.getElementById('board').classList.contains('booting'), null, { timeout: 20000 })
    .catch(() => errs.push('board never left booting'));
  await page.waitForTimeout(1000);
  return { page, ctx, errs };
}

const tap = (page, sel) => page.evaluate(sel => { const e = document.querySelector(sel); if (e) e.click(); return !!e; }, sel);
const state = page => page.evaluate(() => ({
  open: typeof drwOpen !== 'undefined' && !!drwOpen,
  title: document.getElementById('drwTitle').textContent,
  back: !!document.getElementById('drwRes') && !document.getElementById('drwRes').classList.contains('hidden'),
  items: [...document.querySelectorAll('#drwBody .fnd-item .fnd-sku')].map(e => e.textContent),
  mode: (document.querySelector('#drwBody .fnd-mode') || {}).textContent || '',
  out: ((document.getElementById('fndOut') || {}).textContent || '').trim(),
  input: (document.getElementById('fndIn') || {}).value
}));
async function search(page, text) {
  await page.evaluate(t => { const i = document.getElementById('fndIn'); if (i) i.value = t; }, text);
  await page.press('#fndIn', 'Enter').catch(() => {});
  await page.waitForTimeout(500);
}

(async () => {
  const browser = await chromium.launch();
  console.log(`\nBOARD: ${BOARD}\n`);
  const { page, ctx, errs } = await boot(browser);

  // A · the door
  const btn = await page.evaluate(() => { const b = document.getElementById('findBtn');
    if (!b) return null; const r = b.getBoundingClientRect(); return { inFooter: !!b.closest('.ftr'), w: Math.round(r.width), h: Math.round(r.height), text: b.textContent }; });
  check('A1 Find button sits in the footer bar', btn && btn.inFooter, true);
  check('A2 it is a real tap target (≥ 40 px tall) and says the word', btn && btn.h >= 40 && btn.text, 'Find');
  await tap(page, '#findBtn'); await page.waitForTimeout(400);
  let s = await state(page);
  check('A3 tapping it opens the drawer on the search box', [s.open, s.title, s.input], [true, 'Find a part', '']);
  check('A4 nothing is searched just by opening it', CALLS.length, 0);

  // B · words
  await search(page, 'v2203 piston');
  s = await state(page);
  check('B1 one request, the text as typed, no force', CALLS, [{ text: 'v2203 piston', force: '' }]);
  check('B2 the results list in server order', s.items, ['166527', '173817', '215756']);
  check('B3 it says how it searched, with the switch', /Searched as words · try as SKU or part numbers/.test(s.mode), true);
  const row2 = await page.evaluate(() => { const it = document.querySelectorAll('#drwBody .fnd-item')[1]; if (!it) return { qty: '', txt: '' };
    return { qty: it.querySelector('.q').className, txt: it.querySelector('.fnd-meta').textContent }; });
  check('B4 zero on hand is marked red', /zero/.test(row2.qty), true);
  const ended = await page.evaluate(() => { const it = document.querySelectorAll('#drwBody .fnd-item')[2]; if (!it) return null;
    return [it.classList.contains('ended'), [...it.querySelectorAll('.fnd-tag')].map(t => t.textContent)]; });
  check('B5 an ended kit is dimmed and tagged', ended, [true, ['kit', 'Completed']]);

  // C · tap a result → the part drawer, with ← results
  await page.evaluate(() => { const it = document.querySelector('#drwBody .fnd-item'); if (it) it.click(); });
  await page.waitForTimeout(700);
  s = await state(page);
  check('C1 the part drawer opens for that SKU', s.title, '166527');
  check('C2 "← results" is offered', s.back, true);
  const sizes = await page.evaluate(() => [...document.querySelectorAll('#drwBody .drw-size')].map(e =>
    (e.classList.contains('self') ? '*' : '') + e.querySelector('b').textContent));
  check('C3 the dossier paints Other sizes, this one marked', sizes, ['*STD', '0.50']);
  const chips = await page.evaluate(() => [...document.querySelectorAll('#drwBody .drw-chip')].map(e =>
    (e.classList.contains('main') ? '*' : '') + e.textContent));
  check('C4 ...and the part numbers, the main one marked', chips, ['*16423-21110', '1A091-21110']);
  check('C5 ...and the engines', await page.evaluate(() => (document.querySelector('#drwBody .drw-fit') || {}).textContent), 'Engines V2203, D1703');

  // D · a size card opens that part and keeps the way back
  await page.evaluate(() => { const c = document.querySelector('#drwBody .drw-size[data-goto]'); if (c) c.click(); });
  await page.waitForTimeout(700);
  s = await state(page);
  check('D1 tapping the 0.50 card opens 173817', s.title, '173817');
  check('D2 "← results" still offered', s.back, true);
  check('D3 now the 0.50 card is the marked one', await page.evaluate(() =>
    (document.querySelector('#drwBody .drw-size.self b') || {}).textContent), '0.50');

  // E · back to the results — from memory, no new search
  const before = CALLS.length;
  await tap(page, '#drwRes'); await page.waitForTimeout(400);
  s = await state(page);
  check('E1 ← results restores the list', s.items, ['166527', '173817', '215756']);
  check('E2 ...and the query', s.input, 'v2203 piston');
  check('E3 ...without searching again', CALLS.length, before);
  check('E4 the back button hides on the search view', s.back, false);

  // F · the mode switch
  await page.evaluate(() => { const a = [...document.querySelectorAll('#drwBody .fnd-mode a')].find(x => x.dataset.force === 'mpn'); if (a) a.click(); });
  await page.waitForTimeout(500);
  check('F1 "part numbers" re-runs with force=mpn', CALLS[CALLS.length - 1], { text: 'v2203 piston', force: 'mpn' });
  s = await state(page);
  check('F2 part-number view: each number with its listings', s.items, ['207069']);
  check('F3 a number on no listing is listed as such', /Not on any listing \(1\)\s*04179234/.test(s.out), true);
  check('F4 an extra-MPN hit says so', await page.evaluate(() => [...document.querySelectorAll('#drwBody .fnd-tag')].map(t => t.textContent)), ['extra MPN']);

  // G · a SKU goes straight to the part, no back button (there is no list)
  await search(page, '157554');
  s = await state(page);
  check('G1 a SKU opens the part drawer directly', s.title, '157554');
  check('G2 ...with no "← results"', s.back, false);

  // H · close, reopen: the last list is still there (a SKU search leaves no list → the hint)
  await tap(page, '#drwX'); await page.waitForTimeout(300);
  check('H1 closing closes', (await state(page)).open, false);
  await tap(page, '#findBtn'); await page.waitForTimeout(400);
  s = await state(page);
  check('H2 reopening keeps the last text', s.input, '157554');
  check('H3 ...and shows the hint (a SKU search has no list)', /Type a SKU, part numbers/.test(s.out), true);

  // I · a refusal is shown, not swallowed
  MODE = 'refuse';
  await search(page, 'xx');
  check('I1 the server\'s reason is shown', (await state(page)).out, 'No part numbers found in what was pasted.');

  // J · a dropped connection says so
  MODE = 'down';
  await search(page, 'v2203 piston');
  check('J1 a transport failure says "could not reach"', /Could not reach the sheet/.test((await state(page)).out), true);
  MODE = 'normal';

  // K · Escape closes; the pick list is untouched behind it
  await page.keyboard.press('Escape'); await page.waitForTimeout(300);
  check('K1 Escape closes the drawer', (await state(page)).open, false);
  check('K2 the pick list is still there', await page.evaluate(() => document.querySelectorAll('.pick-row').length > 0), true);

  // L · the live race: stage 3 (dossier) lands BEFORE stage 2 (basics) — must survive
  SLOW_LITE = 900;
  await tap(page, '#findBtn'); await page.waitForTimeout(300);
  await search(page, 'v2203 piston');
  await page.evaluate(() => { const it = document.querySelector('#drwBody .fnd-item'); if (it) it.click(); });
  await page.waitForTimeout(1600);
  check('L1 dossier first, basics second: Other sizes still shown', await page.evaluate(() =>
    document.querySelectorAll('#drwBody .drw-size').length), 2);
  check('L2 ...and the basics painted too (title)', await page.evaluate(() =>
    ((document.querySelector('#drwBody .drw-name') || {}).textContent || '')), 'Piston 166527');
  SLOW_LITE = 0;

  check('no page errors', errs, []);
  await ctx.close(); await browser.close();
  console.log('\n' + (failures.length ? `✗ ${failures.length} FAILED` : '✓ ALL PASSED'));
  process.exit(failures.length ? 1 : 0);
})();
