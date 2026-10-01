// ============================================================================
// QUOTE BASKET in the Parts Finder window (diag-quote.js, 2026-10-01)
// Several searches for one customer → one message. Drives the REAL
// PartConsoleModal.html with google.script.run stubbed, from a real origin so
// localStorage works (the basket must survive a new search AND a reload).
//
// Usage: node diag-quote.js
// Before/after:  git show HEAD:PartConsoleModal.html > /tmp/pc-before.html
//                MODAL_FILE=/tmp/pc-before.html node diag-quote.js
// ============================================================================
'use strict';
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');
const MODAL = process.env.MODAL_FILE || path.join(__dirname, '..', 'PartConsoleModal.html');

const failures = [];
function check(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  if (!ok) failures.push(label);
}

const M = (sku, title, loc, qty, price, extra) => Object.assign({ sku, title, location: loc, available: qty, price,
  active: true, status: 'Active', url: 'https://www.ebay.com/itm/' + sku, image: '', mpn: '', via: 'main', isKit: false }, extra || {});
const KW = { ok: true, mode: 'keywords', auto: true, ms: 3000, query: 'v2203 piston', total: 2, matches: [
  M('166527', 'Piston With Rings STD For Kubota, 16423-21110', 'E-84', 8, 79.99, { mpn: '16423-21110' }),
  M('173817', 'Piston With Rings 0.50 For Kubota, 16423-21910', 'E-54', 0, 74, { mpn: '16423-21910' }) ] };
const MPN = { ok: true, mode: 'mpn', auto: true, ms: 3000, found: 2, missing: 1, results: [
  { query: '04103655', matches: [ M('195990', 'Connecting Rod Bearing STD For Deutz 04103655', 'F-12', 30, 18.5, { mpn: '04103655' }) ] },
  { query: '04270701', matches: [ M('157554', 'Engine Overhaul Kit STD Deutz', 'D-1', 1, 494, { isKit: true, active: false, status: 'Completed' }),
                                  M('163485', 'Piston With Ring STD For Deutz, 04179921', 'E-11', 5, 119.99, { mpn: '04179921' }) ] },
  { query: '04179234', matches: [] } ] };
const DOSSIER = { sku: '158805', found: true, isKit: false, usedIn: [], unblock: [], ebayUrl: 'https://www.ebay.com/itm/158805',
  part: { title: 'Main Bearing Set STD For Deutz, 02234014, F2L 511', ebayPrice: null, available: 3, location: 'F-32' },
  identity: { numbers: [{ num: '02234014', main: true, via: '' }], engines: [], brands: [], machines: [] } };

function page(html, init) {
  const stub = `<script>window.google={script:{run:(function(){var h={};var api={withSuccessHandler:function(f){h.ok=f;return api;},
   withFailureHandler:function(f){return api;}, getPartData:function(q){ setTimeout(function(){h.ok({ok:true,dossier:${JSON.stringify(DOSSIER)}});},30); },
   findParts:function(){}};return api;})()}};window.confirm=function(){return true;};
   window.__copied=null;document.execCommand=function(c){ if(c==='copy'){ var a=document.activeElement; window.__copied=a&&a.value; return true;} return false; };</script>`;
  return html.replace('<?!= initJson ?>', JSON.stringify(init)).replace('<head>', '<head><meta charset="utf-8">' + stub);
}

(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1180, height: 820 } });
  const p = await ctx.newPage(); const errs = []; p.on('pageerror', e => errs.push(e.message));
  let init = { text: 'v2203 piston', res: KW };
  const html = fs.readFileSync(MODAL, 'utf8');
  await p.route('http://hq.test/', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: page(html, init) }));
  await p.goto('http://hq.test/'); await p.waitForTimeout(300);
  await p.evaluate(() => { try { localStorage.clear(); } catch (e) {} }); await p.reload(); await p.waitForTimeout(400);

  const bar = () => p.evaluate(() => ({ sum: (document.getElementById('qbSum') || {}).textContent,
    copyOff: !!(document.getElementById('qbCopy') || {}).disabled }));
  const ev = (fn, arg) => p.evaluate(fn, arg);

  // A · empty
  check('A1 the bar starts empty, Copy disabled', await bar(), { sum: 'Quote is empty — use + on any result', copyOff: true });
  check('A2 every row has a + button', await ev(() => document.querySelectorAll('.pf-item .pf-add').length), 2);

  // B · add from a row (and the row click itself still just selects)
  await ev(() => document.querySelector('#pfi1 .pf-add').click()); await p.waitForTimeout(100);
  check('B1 + adds it: 1 part, its price', (await bar()).sum, 'Quote · 1 part · $74.00');
  check('B2 the button turns to ✓', await ev(() => document.querySelector('#pfi1 .pf-add').textContent), '✓');
  check('B3 + did not also open the row', await ev(() => document.getElementById('pfi1').classList.contains('sel')), false);
  await ev(() => document.querySelector('#pfi1 .pf-add').click()); await p.waitForTimeout(100);
  check('B4 + again = one more of the same, not a second line', (await bar()).sum, 'Quote · 1 part (2 pcs) · $148.00');

  // C · a NEW search keeps the basket (the whole point)
  await ev(r => showResult(r), MPN); await p.waitForTimeout(200);
  check('C1 after a new search the quote is still there', (await bar()).sum, 'Quote · 1 part (2 pcs) · $148.00');
  await ev(() => qbAddFirsts()); await p.waitForTimeout(150);
  check('C2 "+ Add all to quote" takes the first ACTIVE match per number (skips the ended kit)', await ev(() => qb.items.map(x => x.sku)), ['173817', '195990', '163485']);
  await ev(() => qbAddMissing()); await p.waitForTimeout(100);
  check('C3 missing numbers go in as "not available"', await ev(() => qb.missing), ['04179234']);
  check('C4 the summary counts them', (await bar()).sum, 'Quote · 3 parts (4 pcs) · $286.49 · 1 not available');

  // D · add from the right pane (a SKU dossier — no price → "on request")
  await ev(d => renderDossier(d), DOSSIER); await p.waitForTimeout(100);
  await ev(() => document.getElementById('pcAddQ').click()); await p.waitForTimeout(100);
  check('D1 "+ Add to quote" in the header adds the open part, number from its identity',
    await ev(() => qb.items.slice(-1).map(x => [x.sku, x.mpn, x.price])), [['158805', '02234014', null]]);
  check('D2 the header button now says it is in', await ev(() => document.getElementById('pcAddQ').textContent), '✓ In quote · add one');
  check('D3 an unpriced part is counted apart, never as $0', (await bar()).sum, 'Quote · 4 parts (5 pcs) · $286.49 + 1 on request · 1 not available');

  // E · the panel: qty, remove, note, links
  await ev(() => qbOpen(true)); await p.waitForTimeout(100);
  check('E1 the panel lists every part', await ev(() => document.querySelectorAll('.qb-row').length), 4);
  await ev(() => document.querySelectorAll('.qb-row')[1].querySelectorAll('.qb-qty button')[1].click());   // bearing +1
  await ev(() => document.querySelectorAll('.qb-row')[1].querySelectorAll('.qb-qty button')[1].click());
  await ev(() => document.querySelectorAll('.qb-row')[1].querySelectorAll('.qb-qty button')[1].click());
  check('E2 − / + changes the quantity (bearing ×4)', await ev(() => qb.items[1].qty), 4);
  await ev(() => document.querySelectorAll('.qb-row')[0].querySelectorAll('.qb-qty button')[0].click());
  await ev(() => document.querySelectorAll('.qb-row')[0].querySelectorAll('.qb-qty button')[0].click());
  check('E3 never below 1', await ev(() => qb.items[0].qty), 1);
  await ev(() => { const n = document.getElementById('qbNote'); n.value = 'Deutz F2L511'; n.dispatchEvent(new Event('input')); });
  check('E4 the note is kept', await ev(() => qb.note), 'Deutz F2L511');

  // F · the message, exactly
  await ev(() => qbCopy()); await p.waitForTimeout(100);
  const withLinks = await ev(() => window.__copied);
  check('F1 the copied message, links on', withLinks, [
    'HQ Motor Service — quote for Deutz F2L511', '',
    '1 × Piston With Rings 0.50 For Kubota, 16423-21910',
    '    Part # 16423-21910 · SKU 173817 · $74.00 · currently out of stock',
    '    https://www.ebay.com/itm/173817', '',
    '4 × Connecting Rod Bearing STD For Deutz 04103655',
    '    Part # 04103655 · SKU 195990 · $18.50 each · $74.00',
    '    https://www.ebay.com/itm/195990', '',
    '1 × Piston With Ring STD For Deutz, 04179921',
    '    Part # 04179921 · SKU 163485 · $119.99',
    '    https://www.ebay.com/itm/163485', '',
    '1 × Main Bearing Set STD For Deutz, 02234014, F2L 511',
    '    Part # 02234014 · SKU 158805 · price on request',
    '    https://www.ebay.com/itm/158805', '',
    'Not available: 04179234', '',
    'Total: $267.99 (+ 1 item priced on request)',
    'Prices exclude shipping.'].join('\n'));
  await ev(() => { const c = document.getElementById('qbLinks'); c.checked = false; c.dispatchEvent(new Event('change')); });
  await ev(() => qbCopy());
  const noLinks = await ev(() => window.__copied);
  check('F2 links off: no URL lines at all', /https?:/.test(noLinks), false);
  check('F3 links off: everything else identical', noLinks, withLinks.split('\n').filter(l => !/https?:/.test(l)).join('\n').replace(/\n{3,}/g, '\n\n'));
  check('F4 the button confirms', await ev(() => document.getElementById('qbCopy').textContent), 'Copied ✓');

  // G · remove
  await ev(() => document.querySelectorAll('.qb-row')[0].querySelector('.qb-x').click()); await p.waitForTimeout(100);
  check('G1 ✕ removes the line', await ev(() => qb.items.map(x => x.sku)), ['195990', '163485', '158805']);
  await ev(() => document.querySelector('.qb-na .qb-x').click());
  check('G2 ✕ removes a not-available number', await ev(() => qb.missing), []);

  // H · survives a reload (closing the window)
  init = { text: '', res: null };
  await p.reload(); await p.waitForTimeout(400);
  check('H1 after a reload the quote is still there', (await bar()).sum, 'Quote · 3 parts (6 pcs) · $193.99 + 1 on request');
  check('H2 ...with the note and the links choice', await ev(() => [qb.note, qb.links]), ['Deutz F2L511', false]);

  // I · clear
  await ev(() => { qbOpen(true); qbClear(); }); await p.waitForTimeout(100);
  check('I1 Clear empties it', await bar(), { sum: 'Quote is empty — use + on any result', copyOff: true });

  check('no page errors', errs, []);
  await browser.close();
  console.log('\n' + (failures.length ? `✗ ${failures.length} FAILED` : '✓ ALL PASSED'));
  process.exit(failures.length ? 1 : 0);
})();
