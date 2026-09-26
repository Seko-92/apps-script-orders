/**
 * diag-amazon-board.js — the Floor Board and the wall with a THIRD channel (2026-09-26).
 *
 * Drives the REAL FloorBoard.html and wall.html with the standard mock tick plus two
 * Amazon orders, and asserts what a picker would see: three tabs with honest counts, the
 * Amazon tab showing only Amazon rows under an order band that carries the Seller Central
 * reminder, eBay and Direct untouched, and the arrival beacon naming Amazon. Screenshots
 * go to renders/amazon-*.png.
 *
 * Run:  node design-lab/diag-amazon-board.js     (BOARD_FILE / WALL_FILE to point elsewhere)
 */
'use strict';
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const WALL  = process.env.WALL_FILE  || path.join(__dirname, '..', 'wall.html');
const OUT = path.join(__dirname, 'renders');
const BASE = JSON.parse(JSON.stringify(require('./mock-tick.js')));

let pass = 0, fail = 0;
const ok = (label, cond, detail) => {
  cond ? pass++ : fail++;
  console.log((cond ? '  ✓ ' : '  ✗ ') + label + (cond ? '' : '  → ' + JSON.stringify(detail)));
};

function tickWithAmazon(extra) {
  const t = JSON.parse(JSON.stringify(BASE));
  t.openOrders = t.openOrders.concat([
    { channel: 'AMAZON', orderId: 'AMZ-114-3941689-8772232', sku: '167517', qty: 2, location: 'K-7',
      status: 'PENDING', note: 'AMAZON · ship by 9/29', isKit: false, hand: 12 },
    { channel: 'AMAZON', orderId: 'AMZ-114-3941689-8772232', sku: '171378', qty: 1, location: 'L-3',
      status: 'PENDING', note: 'AMAZON · ship by 9/29', isKit: false, hand: 1 },
    { channel: 'AMAZON', orderId: 'AMZ-111-1111111-1111111', sku: '194244', qty: 1, location: 'A-14',
      status: 'PREPARING', note: 'AMAZON · ship by 9/30', isKit: false, hand: 9 }
  ]).concat(extra || []);
  t.openOrdersTotal = (t.openOrdersTotal || 0) + 3;
  t.openOrdersBy = Object.assign({}, t.openOrdersBy, { AMAZON: 3 });
  t.cockpit.receivedAmazon = 2;
  return t;
}

async function openPage(browser, file, tickFn, vp) {
  const html = fs.readFileSync(file, 'utf8');
  const ctx = await browser.newContext({ viewport: vp || { width: 1280, height: 800 }, hasTouch: true,
                                         timezoneId: 'America/Chicago' });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push('pageerror: ' + e.message));
  page.on('console', m => { if (m.type() === 'error' && !/Failed to load resource/.test(m.text())) errors.push(m.text()); });
  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      let res = { ok: false, message: 'unknown' };
      if (body.action === 'boardTick')  res = Object.assign({ ok: true }, tickFn());
      if (body.action === 'boardRadio') res = { ok: true, nowPlaying: '' };
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com|fonts\.g/, r => r.abort());
  await page.goto('http://hqlab.test/' + (file === WALL ? 'wall' : ''), { waitUntil: 'load' });
  await page.waitForTimeout(2500);
  return { page, ctx, errors };
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch();

  // ---------------- THE FLOOR BOARD ----------------
  console.log('\nFloor Board');
  {
    const { page, ctx, errors } = await openPage(browser, BOARD, () => tickWithAmazon());
    const tabs = await page.$$eval('.ct', b => b.map(x => [x.getAttribute('data-ch'), x.querySelector('b').textContent]));
    ok('B1 three tabs: eBay · Direct · Amazon', JSON.stringify(tabs.map(t => t[0])) === '["EBAY","DIRECT","AMAZON"]', tabs);
    ok('B2 the Amazon tab counts what is open (3)', (tabs.find(t => t[0] === 'AMAZON') || [])[1] === '3', tabs);
    ok('B3 eBay and Direct counts unchanged', (tabs.find(t => t[0] === 'EBAY') || [])[1] === '13' &&
       (tabs.find(t => t[0] === 'DIRECT') || [])[1] === '14', tabs);

    await page.click('.ct[data-ch="AMAZON"]');
    await page.waitForTimeout(400);
    const amz = await page.evaluate(() => {
      const ul = document.getElementById('pickListAmazon');
      const vis = ul && getComputedStyle(ul).display !== 'none';
      return {
        visible: vis,
        otherVisible: ['pickListEbay', 'pickListDirect'].some(id => getComputedStyle(document.getElementById(id)).display !== 'none'),
        rows: Array.from(ul.querySelectorAll('.pick-row')).map(li => li.textContent),
        heads: Array.from(ul.querySelectorAll('.ph-id')).map(x => x.textContent),
        chips: Array.from(ul.querySelectorAll('.ph-amz')).map(x => x.textContent)
      };
    });
    ok('B4 tapping Amazon shows the Amazon list', amz.visible, amz);
    ok('B5 …and only the Amazon list', !amz.otherVisible);
    ok('B6 three Amazon rows', amz.rows.length === 3, amz.rows.length);
    ok('B7 every row is an Amazon part (no eBay/Direct rows leaked in)',
       amz.rows.every(t => /167517|171378|194244/.test(t)), amz.rows);
    ok('B8 each Amazon order gets a band, even a one-line one (it is a box, like Direct)',
       JSON.stringify(amz.heads.slice().sort()) === JSON.stringify(['AMZ-111-1111111-1111111', 'AMZ-114-3941689-8772232']), amz.heads);
    ok('B9 every Amazon band carries the Seller Central reminder',
       amz.chips.length === 2 && amz.chips.every(c => /Seller Central/.test(c)), amz.chips);
    await page.screenshot({ path: path.join(OUT, 'amazon-board-tab.png') });

    await page.click('.ct[data-ch="DIRECT"]');
    await page.waitForTimeout(300);
    const dir = await page.$$eval('#pickListDirect .pick-row', r => r.map(x => x.textContent).join('|'));
    ok('B10 the Direct list holds no Amazon parts', !/AMZ-/.test(dir) && dir.length > 0);
    const heads = await page.$$eval('#pickListDirect .ph-amz', x => x.length);
    ok('B11 no Seller Central chip on Direct bands', heads === 0, heads);
    ok('B12 no page errors', errors.length === 0, errors);
    await ctx.close();
  }

  // an Amazon arrival: the beacon names the channel
  {
    let n = 0;
    const { page, ctx } = await openPage(browser, BOARD, () => (n++ < 1 ? tickWithAmazon() : tickWithAmazon([
      { channel: 'AMAZON', orderId: 'AMZ-222-2222222-2222222', sku: '155394', qty: 1, location: 'B-12',
        status: 'PENDING', note: 'AMAZON', isKit: false, hand: 5 }])));
    await page.evaluate(() => { if (typeof pollSoon === 'function') pollSoon(); });
    await page.waitForTimeout(2500);
    const beacon = await page.evaluate(() => document.getElementById('beaconBody').textContent);
    ok('B13 a new Amazon order lights the beacon as "Amazon · …"', /^Amazon · B-12/.test(beacon), beacon);
    await ctx.close();
  }

  // no Amazon rows → the board behaves as before
  {
    const { page, ctx, errors } = await openPage(browser, BOARD, () => JSON.parse(JSON.stringify(BASE)));
    const tabs = await page.$$eval('.ct', b => b.map(x => [x.getAttribute('data-ch'), x.querySelector('b').textContent,
                                                           x.classList.contains('none')]));
    ok('B14 without Amazon orders the Amazon tab reads 0 and is dimmed', JSON.stringify(tabs[2]) === '["AMAZON","0",true]', tabs);
    const active = await page.evaluate(() => document.body.className.match(/ch-\w+/)[0]);
    ok('B15 the board still opens on eBay', active === 'ch-ebay', active);
    ok('B16 no page errors', errors.length === 0, errors);
    await ctx.close();
  }

  // ---------------- THE WALL ----------------
  console.log('\nWall');
  {
    const { page, ctx, errors } = await openPage(browser, WALL, () => tickWithAmazon(), { width: 1920, height: 1080 });
    const w = await page.evaluate(() => ({
      heads: Array.from(document.querySelectorAll('.chan-head h2')).map(h => h.textContent),
      amzRows: document.querySelectorAll('#listA .row, #listA li:not(.none)').length,
      amzText: document.getElementById('listA').textContent,
      dText: document.getElementById('listD').textContent,
      cnt: document.getElementById('cntA').textContent
    }));
    ok('W1 three columns: eBay · Direct · Amazon', JSON.stringify(w.heads) === '["eBay","Direct","Amazon"]', w.heads);
    ok('W2 the Amazon column lists the Amazon parts', /167517/.test(w.amzText) && /194244/.test(w.amzText), w.amzText.slice(0, 200));
    ok('W3 the Direct column holds no Amazon order', !/AMZ-/.test(w.dText));
    ok('W4 no page errors', errors.length === 0, errors);
    await page.screenshot({ path: path.join(OUT, 'amazon-wall.png') });
    await ctx.close();
  }

  await browser.close();
  console.log('\n' + (fail ? '✗ ' + fail + ' FAILED · ' : '✓ all · ') + pass + ' passed');
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error('CRASH', e); process.exit(1); });
