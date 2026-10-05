// shoot-kit-batch.js — the 2026-10-05 "review first, write once" kit window,
// driven end to end against the real SO-25980 queue (13 kits, 7 READY).
//
//   node shoot-kit-batch.js                 → renders/kitbatch-*.png + assertions
//   MODAL_FILE=/tmp/old.html node ...       → run against another revision
//
// google.script.run is a recording stub: commitKitBatchFromModal answers per item,
// finishKitBatchFromModal returns results built from what was committed. Every
// call is kept in window.__calls so the test can assert WHAT was sent — the
// payload is the contract, the pixels are only the picture.
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');
const ROOT = path.join(__dirname, '..');
const OUT = path.join(__dirname, 'renders');
const FILE = process.env.MODAL_FILE || path.join(ROOT, 'KitExpansionModal.html');
const QUEUE = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixture-kit-batch.json'), 'utf8'));

let pass = 0, fail = 0;
const t = (name, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + name + (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};

function page_html(queue) {
  let html = fs.readFileSync(FILE, 'utf8');
  html = html.replace(/<\?!=\s*sessionId\s*\?>/g, JSON.stringify('sess-1'))
             .replace(/<\?!=\s*queueJson\s*\?>/g, JSON.stringify(queue))
             .replace(/<\?=\s*queueLength\s*\?>/g, String(queue.length));
  const left = html.match(/<\?[^>]{0,40}\?>/g);
  if (left) throw new Error('unresolved scriptlet ' + left.join(', '));
  return html;
}

async function open(browser, queue, failIndex) {
  const page = await browser.newPage({ viewport: { width: 1180, height: 820 }, deviceScaleFactor: 1 });
  const errs = []; page.on('pageerror', e => errs.push(String(e)));
  await page.addInitScript((failIndex) => {
    window.__calls = [];
    const committed = [], skipped = [], failed = [];
    const impl = {
      commitKitBatchFromModal(sid, items) {
        return { ok: true, results: items.map(it => {
          if (it.index === failIndex) { failed.push({ kitSku: 'x', reason: 'Kit row not found' }); return { index: it.index, ok: false, action: 'expand', reason: 'Kit row not found' }; }
          if (it.action !== 'expand') { skipped.push({ kitSku: 'k' + it.index, kitType: it.action === 'box' ? 'READY' : 'MANUAL' }); return { index: it.index, ok: true, action: it.action }; }
          const c = { kitSku: 'k' + it.index, componentsAdded: 5, excludedSkus: it.excludedSkus, extras: it.extras, totalKits: 1, rowQty: 1, forced: it.force };
          committed.push(c); return { index: it.index, ok: true, action: 'expand', committed: c };
        }) };
      },
      finishKitBatchFromModal() { return { ok: true, results: { committed, skipped, failed } }; },
      lookupSkuForKitAlter(sku) { return { found: true, sku, name: 'Swapped part', location: 'Z-1', available: 4 }; },
      closeKitExpansionSession() { return { ok: true }; }
    };
    const mk = (su, fa) => new Proxy({}, { get(_, k) {
      if (k === 'withSuccessHandler') return f => mk(f, fa);
      if (k === 'withFailureHandler') return f => mk(su, f);
      return (...args) => { window.__calls.push({ fn: k, args: JSON.parse(JSON.stringify(args)) });
        const r = impl[k] ? impl[k](...args) : null; setTimeout(() => su && su(r), 120); };
    }});
    window.google = { script: { run: mk(null, null), host: { close() { window.__closed = true; } } } };
    window.confirm = () => true;
  }, failIndex == null ? -1 : failIndex);
  const html = page_html(queue);
  await page.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: html }));
  await page.goto('http://hq.test/kit', { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(500);
  return { page, errs };
}
const shot = (page, n) => page.screenshot({ path: path.join(OUT, 'kitbatch-' + n + '.png') });
const txt = (page, sel) => page.$eval(sel, e => e.textContent.trim()).catch(() => null);
const calls = (page, fn) => page.evaluate(fn => window.__calls.filter(c => c.fn === fn), fn);

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch();

  console.log('\nA · overview');
  let { page, errs } = await open(browser, QUEUE);
  t('A1 opens on the overview, not kit 1', await page.$('.ov') != null, true);
  t('A2 counts 7 READY', await page.$$eval('.ov-group.ready .ov-chip', e => e.length), 7);
  t('A3 158670 (registry said MANUAL) is shown READY at K-55',
    await page.$$eval('.ov-group.ready .ov-chip', e => e.map(x => x.innerText).some(s => s.includes('158670') && s.includes('K-55'))), true);
  t('A4 nothing written yet', (await page.evaluate(() => window.__calls.length)), 0);
  await shot(page, '1-overview');

  console.log('\nB · "Ship as boxes" → READY kits skip the review');
  await page.click('text=Start review →');
  t('B1 counter says 6 to review of 13', await txt(page, '#counter'), 'Kit 1 of 6 · 13 total');
  t('B2 first page is a MANUAL kit (158949)', (await txt(page, '.kit-sku')), '158949');
  await page.waitForTimeout(600); await shot(page, '2-kit');

  console.log('\nC · choices survive Back');
  await page.uncheck('.comp-cb >> nth=0');
  await page.fill('#deployInput', '2'); await page.dispatchEvent('#deployInput', 'input');
  await page.click('#btnNext');
  t('C1 Next moves to kit 2', await txt(page, '#counter'), 'Kit 2 of 6 · 13 total');
  t('C2 still nothing written', (await page.evaluate(() => window.__calls.filter(c => /Batch/.test(c.fn)).length)), 0);
  await page.click('#btnBack');
  t('C3 Back restores the unticked part', await page.$eval('.comp-cb', e => e.checked), false);
  t('C4 Back restores spares = 2', await page.$eval('#deployInput', e => e.value), '2');

  console.log('\nD · walk to the summary');
  for (let i = 0; i < 6; i++) { if (await page.$('.sum')) break; await page.click('#btnNext'); await page.waitForTimeout(50); }
  t('D1 summary lists all 13 kits', await page.$$eval('.sum-line', e => e.length), 13);
  t('D2 the hand-unchecked kit is flagged amber', await page.$eval('#sum-3', e => e.classList.contains('flag')), true);
  t('D3 boxes are quiet lines', await page.$eval('#sum-0', e => e.classList.contains('quiet')), true);
  t('D4 button says Expand all 6', await txt(page, '#btnWrite'), '✓ Expand all 6');
  await shot(page, '3-summary');

  console.log('\nE · edit from the summary returns to the summary');
  await page.click('#sum-0 .sum-edit');
  t('E1 opens 158670', await txt(page, '.kit-sku'), '158670');
  await page.click('.act-btn:has-text("Expand into parts")');
  t('E2 primary says Back to summary', await txt(page, '#btnNext'), 'Back to summary →');
  await page.click('#btnNext');
  t('E3 back on summary, now 7 to expand', await txt(page, '#btnWrite'), '✓ Expand all 7');
  t('E4 158670 now marked READY · EXPANDED', await page.$eval('#sum-0', e => e.innerText.includes('READY · EXPANDED')), true);

  console.log('\nF · the one write');
  await page.click('#btnWrite');
  await page.waitForTimeout(150);
  t('F1 progress shown on the lines, no veil', await page.$eval('#kxOverlay', e => e.classList.contains('shown')), false);
  await shot(page, '4-writing');
  await page.waitForSelector('.done-screen', { timeout: 8000 });
  const batches = await calls(page, 'commitKitBatchFromModal');
  const sent = batches.flatMap(b => b.args[1]);
  t('F2 every kit sent exactly once', sent.map(s => s.index).sort((a, b) => a - b), QUEUE.map((_, i) => i));
  t('F3 no chunk carries more than 3 expansions', batches.every(b => b.args[1].filter(x => x.action === 'expand').length <= 3), true);
  t('F4 158670 expanded with force', sent.find(s => s.index === 0), sent.find(s => s.index === 0) && Object.assign({}, sent.find(s => s.index === 0), { action: 'expand', force: true }));
  t('F5 158679 still ships as a box', sent.find(s => s.index === 1).action, 'box');
  const k3 = sent.find(s => s.index === 3);
  t('F6 kit 158949 carries spares=2 and the unticked part', [k3.extras, k3.excludedSkus.includes(QUEUE[3].components[0].sku)], [2, true]);
  t('F7 finish called once, after the chunks', (await calls(page, 'finishKitBatchFromModal')).length, 1);
  await shot(page, '5-done');
  t('F8 no page errors', errs, []);
  await page.close();

  console.log('\nG · a single kit skips overview and summary');
  ({ page, errs } = await open(browser, [QUEUE[3]]));
  t('G1 opens straight on the kit', await txt(page, '.kit-sku'), '158949');
  t('G2 button writes directly', await txt(page, '#btnNext'), 'Expand →');
  await page.click('#btnNext');
  await page.waitForSelector('.done-screen', { timeout: 8000 });
  t('G3 one batch call', (await calls(page, 'commitKitBatchFromModal')).length, 1);
  t('G4 no page errors', errs, []);
  await page.close();

  console.log('\nH · a failed kit is named, the rest still land');
  ({ page, errs } = await open(browser, QUEUE, 5));
  await page.click('.act-btn:has-text("Expand all into parts")');
  await page.click('text=Start review →');
  for (let i = 0; i < 14; i++) { if (await page.$('.sum')) break; await page.click('#btnNext'); await page.waitForTimeout(40); }
  await page.click('#btnWrite');
  await page.waitForSelector('.done-screen', { timeout: 8000 });
  const doneTxt = await txt(page, '.done-screen');
  t('H1 done screen shows the failure', /Kit row not found/.test(doneTxt), true);
  t('H2 the other 12 committed', (await page.$$eval('.done-result-line.committed', e => e.length)), 12);
  await shot(page, '6-failure');
  t('H3 no page errors', errs, []);

  await browser.close();
  console.log('\n' + pass + ' passed · ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
})();
