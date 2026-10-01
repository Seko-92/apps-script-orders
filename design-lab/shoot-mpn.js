// shoot-mpn.js — render the Parts Finder (PartConsoleModal.html) headless.
//   node shoot-mpn.js out.png            keyword result, first row selected, full dossier arrived
//   MODE=mpn node shoot-mpn.js out.png   SERPIC paste result
//   MODE=loading ...                     the instant before the full dossier arrives
const fs = require('fs'), path = require('path'); const { chromium } = require('playwright');
const img = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="80" height="80"><rect width="80" height="80" fill="#ddd"/><circle cx="40" cy="40" r="22" fill="#999"/></svg>');
const M = (sku, title, loc, qty, price, extra) => Object.assign({ sku, title, location: loc, available: qty, price, active: true, url: 'x', image: img, mpn: '' }, extra || {});
const kw = { ok: true, mode: 'keywords', auto: true, ms: 3600, query: 'v2203 piston', total: 16, matches: [
  M('155430', 'Piston rings STD For Kubota, 1G790-21050, 1G790-21090, V2203, V2003, D1503', 'J-29', 85, 22),
  M('163332', 'Piston With Ring STD For Kubota, 1G790-21110, V2203-M, V2003-T engines', 'E-84', 8, 79.99),
  M('173817', 'Piston With Rings 0.50 For Kubota, 16423-21910, V2203 Indirect Injection', 'E-54', 7, 74),
  M('215756', 'Engine Overhaul Kit 0.50, 16423-21910 For Kubota V2203', 'K-12', 1, 440, { isKit: true }),
  M('163341', 'Piston With Ring 0.50 For Kubota, 1G790-21900, V2203-E', 'D-24', 0, 99, { active: false, status: 'Completed' }) ] };
const mpn = { ok: true, mode: 'mpn', auto: true, ms: 4000, found: 2, missing: 1, columns: 22, results: [
  { query: '04270701', matches: [ M('163485', 'Piston With Ring STD For Deutz, 04179921, 04270701, BF4M1011F', 'E-11', 203, 119.99),
                                  M('157554', 'Engine Overhaul, Rebuild Kit, For Deutz BF4M1011F STD', 'D-1', 1, 494, { isKit: true }) ] },
  { query: '04179921', matches: [ M('207069', 'Piston With Ring 0.50 For Deutz, 04270702, BF4M1011F', 'E-11', 99, 99, { via: 'extra MPN' }) ] },
  { query: '04179234', matches: [] } ] };
const dossier = { sku: '155430', found: true, isKit: false, ebayUrl: 'x', unblock: [],
  part: { title: 'Piston rings STD For Kubota, 1G790-21050', images: [img, img, img], location: 'J-29', available: 85,
          zohoAvailable: 85, miAvailable: 86, sold: 412, ebayPrice: 22, zohoPrice: 22, listingStatus: 'Active', committed: 2 },
  usedIn: [{ kitSku: '217205', kitName: 'Engine Overhaul Kit STD Kubota V2203', qtyPer: 4, buildable: 3, priceStatus: 'IN LINE' }] };
const res = process.env.MODE === 'mpn' ? mpn : kw;
const text = process.env.MODE === 'mpn' ? '0427 0701\n0417 9234\n0417 9921' : 'v2203 piston';
const stub = `<script>window.google={script:{run:(function(){var h={};var api={withSuccessHandler:function(f){h.ok=f;return api;},
 withFailureHandler:function(f){return api;}, getPartData:function(q){ ${process.env.MODE === 'loading' ? '' : 'setTimeout(function(){h.ok({ok:true,dossier:' + JSON.stringify(dossier) + '});},50);'} },
 findParts:function(){}};return api;})()}};</script>`;
const html = fs.readFileSync(path.join(__dirname, '..', 'PartConsoleModal.html'), 'utf8')
  .replace('<?!= initJson ?>', JSON.stringify({ text, res }))
  .replace('<head>', '<head><meta charset="utf-8">' + stub);
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1180, height: 760 } });
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  await p.route('http://hq.test/', r => r.fulfill({ body: html, contentType: 'text/html; charset=utf-8' }));
  await p.goto('http://hq.test/'); await p.waitForTimeout(400);
  await p.screenshot({ path: process.argv[2] || 'renders/parts-finder.png' });
  console.log('errors:', errs.length ? errs : 'none'); await b.close();
})();
