// shoot-mpn.js — render MpnSearchModal.html headless with a realistic result.
const fs = require('fs'), path = require('path'); const { chromium } = require('playwright');
const res = { ok: true, found: 2, missing: 2, columns: 21, ms: 2140, results: [
  { query: '02102238', matches: [
    { sku: '157554', isKit: true, title: 'Engine Overhaul, Rebuild Kit, For Deutz BF4M1011F', location: 'E-17', available: 6, price: 85, via: 'main', active: true, url: 'https://ebay.com/itm/1' },
    { sku: '171729', title: 'Main Bearing Set STD Deutz BF4M1012', location: 'NOT FOUND', available: 0, price: 92.5, via: 'extra MPN', active: false, status: 'Completed', url: '' } ] },
  { query: '0415 7075', matches: [
    { sku: '163872', title: 'Cylinder Head Gasket for Deutz F3L912', location: 'A-9', available: 3, price: 39.99, via: 'Interchange Part Number', active: true, url: 'https://ebay.com/itm/2' } ] },
  { query: '04292547', matches: [] }, { query: '1C010-74110', matches: [] } ] };
let html = fs.readFileSync(path.join(__dirname, '..', 'MpnSearchModal.html'), 'utf8')
  .replace('<?!= initJson ?>', JSON.stringify({ text: '1  0429 2547  Seal ring  1\n2  0415 7075  Gasket  2', res }));
html = html.replace('<head>', '<head><meta charset="utf-8"><script>window.google={script:{run:{}}};</script>');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1080, height: 720 } });
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  await p.route('http://hq.test/', r => r.fulfill({ body: html, contentType: 'text/html; charset=utf-8' }));
  await p.goto('http://hq.test/'); await p.waitForTimeout(300);
  await p.screenshot({ path: process.argv[2] || 'renders/mpn-finder.png' });
  console.log('errors:', errs.length ? errs : 'none'); await b.close();
})();
