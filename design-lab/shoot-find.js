// shoot-find.js — render the board's Parts Finder drawer (results, then a part with sizes).
//   node shoot-find.js [outPrefix]   → <prefix>-land-list.png, -land-part.png, -port-list.png
const fs = require('fs'), path = require('path'); const { chromium } = require('playwright');
const MOCK = require('./mock-tick.js');
const img = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="80" height="80"><rect width="80" height="80" fill="#eee"/><circle cx="40" cy="40" r="22" fill="#999"/></svg>');
const M = (sku, title, loc, qty, price, extra) => Object.assign({ sku, title, location: loc, available: qty, price, active: true, status: 'Active', image: img, via: 'main' }, extra || {});
const KW = { ok: true, mode: 'keywords', auto: true, ms: 3400, total: 16, matches: [
  M('155430', 'Piston rings STD For Kubota, 1G790-21050, V2203, V2403, V2203-M-DI', 'J-29', 85, 22),
  M('166527', 'Piston With Rings STD For Kubota, 16423-21110, V2203 IDI, D1703, F2803, 87mm.', 'E-84', 8, 79.99),
  M('173817', 'Piston With Rings 0.50 For Kubota, 16423-21910, V2203 IDI, D1703, F2803, 87mm.', 'E-54', 0, 74),
  M('215756', 'Engine Overhaul, Rebuild Kit 0.50 For Kubota V2203 IDI, 16423-21910, Metal.', 'K-12', 1, 440, { isKit: true }),
  M('163341', 'Piston With Ring 0.50 For Kubota, 1G790-21900, V2203-E', 'D-24', 0, 99, { active: false, status: 'Completed' }) ] };
const DOS = { ok: true, dossier: { sku: '166527', found: true, usedIn: [{ kitSku: '217205', kitName: 'Engine Overhaul Kit STD Kubota V2203', qtyPer: 4, buildable: 3 }], unblock: [],
  part: { available: 8, zohoAvailable: 8, miAvailable: 8, committed: 2 },
  sizes: [{ sku: '166527', size: 'STD', self: true, num: '16423-21110', location: 'E-84', available: 8, active: true },
          { sku: '173817', size: '0.50', self: false, num: '16423-21910', location: 'E-54', available: 0, active: true }],
  identity: { numbers: [{ num: '16423-21110', main: true }, { num: '1A091-21110' }, { num: 'H1900-21100' }], engines: ['V2203', 'D1703', 'V2403', 'F2803'], brands: ['Kubota', 'Bobcat'], machines: [] } } };
const html = fs.readFileSync(path.join(__dirname, '..', 'FloorBoard.html'), 'utf8');
async function shot(browser, vp, name, part) {
  const ctx = await browser.newContext({ viewport: vp, hasTouch: true, timezoneId: 'America/Chicago' }); const p = await ctx.newPage();
  const errs = []; p.on('pageerror', e => errs.push(e.message));
  await p.route('http://hqlab.test/**', r => { const u = r.request().url();
    if (u.includes('/api/board')) { const b = JSON.parse(r.request().postData() || '{}'); let res = { ok: false };
      if (b.action === 'boardTick') res = Object.assign({ ok: true }, MOCK);
      if (b.action === 'boardFind') res = KW;
      if (b.action === 'boardPartLite') res = { ok: true, basics: { sku: b.sku, found: true, title: 'Piston With Rings STD For Kubota, 16423-21110, V2203 IDI, D1703, F2803, 87mm.', location: 'E-84', images: [img, img], available: 8, ebayPrice: 79.99, listingStatus: 'Active' } };
      if (b.action === 'boardPart') res = DOS;
      return r.fulfill({ contentType: 'application/json', body: JSON.stringify(res) }); }
    return r.fulfill({ contentType: 'text/html; charset=utf-8', body: html }); });
  await p.route(/aladhan\.com|open-meteo\.com|fonts\./, r => r.abort());
  await p.goto('http://hqlab.test/'); await p.waitForTimeout(1500);
  await p.click('#findBtn'); await p.waitForTimeout(350);
  await p.fill('#fndIn', 'v2203 piston'); await p.press('#fndIn', 'Enter'); await p.waitForTimeout(700);
  if (part) { await p.evaluate(() => document.querySelectorAll('.fnd-item')[1].click()); await p.waitForTimeout(900); }
  await p.screenshot({ path: name }); console.log(name, errs.length ? errs : 'no errors'); await ctx.close();
}
(async () => { const b = await chromium.launch(); const pre = process.argv[2] || 'renders/find';
  await shot(b, { width: 1280, height: 800 }, pre + '-land-list.png');
  await shot(b, { width: 1280, height: 800 }, pre + '-land-part.png', true);
  await shot(b, { width: 800, height: 1280 }, pre + '-port-list.png');
  await b.close(); })();
