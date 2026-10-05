// shoot-loading.js — the 2026-10-05 in-place loading look, in every window.
//
//   node shoot-loading.js        → renders/loading-<window>.png + assertions
//
// For each window: load it, switch its OLD overlay on the way its own code does
// (add 'shown' / 'on'), and assert what the eye gets — the old veil is invisible,
// the top bar shows, the message line sits inside the result area, the area fades,
// the header and footer do NOT fade — then switch it off and assert it all clears.
// The Kit Price Calculator is also driven through its REAL button, because a
// class toggle cannot prove the window's own code path reaches the new look.
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');
const ROOT = path.join(__dirname, '..'), OUT = path.join(__dirname, 'renders');

const KIT = { kitSku: '217475', sourceSalesOrder: 'SO-24853', sourceQty: 2, sourceNote: '', originalRow: 47, table: 'DIRECT',
  kitName: 'Engine Overhaul Kit 0.50', kitType: 'MANUAL', kitLocation: 'D-1', kitEngine: '', salesDescription: 'Full Gasket Set + Head Gasket',
  components: [{ sku: '167517', name: 'Piston Kit STD', qty: 4, location: 'E-37', available: 12 }], unparsedLines: [], alreadyExpanded: { count: 0, skus: [] } };
const DIFF = { ok: true, soNumber: 'SO-24853', customerName: 'Triple M', totalFormatted: '$5,347.79', isFirstPull: true, pulled: false,
  zohoStatus: 'confirmed', zohoShippedStatus: 'pending', lines: [
    { sku: '168138', name: 'Full Gasket Set', status: 'new', zohoQty: 2, directQty: 0, delta: 2, location: 'L-226', available: 3, directRows: [] }],
  summary: { totalLines: 1, unchanged: 0, new: 1, qtyChanged: 0, removed: 0, anyChanges: true } };
const PUSH = { sessionId: 's', candidates: [{ sku: '170154', name: 'Piston Ring Set', direction: 'ZOHO LOW', zohoBefore: 13.94, ebayTarget: 16, delta: 2.06, itemId: '1', pushable: true }],
  meta: { pushable: 1, skipped: 0, cap: 30, hasPassphrase: true, zohoSyncedAt: Date.now() } };
const DOSSIER = { found: true, query: 'SO-24853', rows: [], events: [], notes: [], summary: { statuses: ['PENDING'], skus: [] }, links: {} };
const PART = { query: '', mode: 'sku', results: [] };

const W = [
  { file: 'KitCalculatorModal.html', ov: 'kcOverlay', msg: 'kcOverlayMsg', reg: '.kc-body', cls: 'shown', keep: '.kc-head',
    subs: [['kitListJson', [{ sku: '217475', name: 'Engine Overhaul Kit' }]]] },
  { file: 'ZohoPullModal.html', ov: 'zpOverlay', msg: 'zpOverlayMsg', reg: '#bodyEl', cls: 'shown', keep: '#ftrEl', subs: [['diffJson', DIFF]] },
  { file: 'PricePushModal.html', ov: 'pxOverlay', msg: 'pxOverlayMsg', reg: '#bodyEl', cls: 'shown', keep: '#ftrEl', subs: [['dataJson', PUSH]] },
  { file: 'OrderCaseModal.html', ov: 'ocOverlay', msg: 'ocOverlayTxt', reg: '.oc-body', cls: 'on', keep: '.oc-head', subs: [['dossierJson', DOSSIER]] },
  { file: 'PartConsoleModal.html', ov: 'pcOverlay', msg: 'pcOverlayTxt', reg: '.pc-main', cls: 'on', keep: '.pc-head', subs: [['initJson', PART]] },
  { file: 'KitExpansionModal.html', ov: 'kxOverlay', msg: 'kxOverlayMsg', reg: '#bodyEl', cls: 'shown', keep: '#ftrEl',
    subs: [['sessionId', 's'], ['queueJson', [KIT]], ['queueLength', 1], ['kitJson', KIT], ['kitIndex', 0]] }
];

let pass = 0, fail = 0;
const t = (n, got, want) => { const ok = JSON.stringify(got) === JSON.stringify(want); ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + n + (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want))); };

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch();
  for (const w of W) {
    console.log('\n═══ ' + w.file);
    let html = fs.readFileSync(path.join(ROOT, w.file), 'utf8');
    for (const [k, v] of w.subs) html = html.replace(new RegExp('<\\?!?=\\s*' + k + '\\s*\\?>', 'g'), JSON.stringify(v));
    const left = html.match(/<\?[^>]{0,40}\?>/g); if (left) throw new Error('unresolved ' + left);
    const page = await browser.newPage({ viewport: { width: 1000, height: 700 } });
    const errs = []; page.on('pageerror', e => errs.push(String(e)));
    await page.addInitScript(() => {
      // Server calls never answer, so a loader the page starts stays up for the shot.
      const mk = () => new Proxy({}, { get(_, k) { return (k === 'withSuccessHandler' || k === 'withFailureHandler' || k === 'withUserObject') ? () => mk() : () => {}; } });
      window.google = { script: { run: mk(), host: { close() {}, setHeight() {}, setWidth() {} } } };
      window.confirm = () => true;
    });
    await page.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: html }));
    await page.goto('http://hq.test/w', { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(700);

    await page.evaluate(w => { document.getElementById(w.msg).textContent = 'Fetching component prices'; document.getElementById(w.ov).classList.add(w.cls); }, w);
    await page.waitForTimeout(450);
    const s = await page.evaluate(w => {
      const reg = document.querySelector(w.reg), keep = document.querySelector(w.keep);
      const line = reg && reg.querySelector(':scope > .wk-line');
      const sib = reg && [...reg.children].find(c => !c.classList.contains('wk-line'));
      return {
        veil: getComputedStyle(document.getElementById(w.ov)).display,
        bar: getComputedStyle(document.querySelector('.wk-bar')).opacity,
        line: line ? line.textContent : null,
        fades: sib ? /opacity\(0\.4/.test(getComputedStyle(sib).filter) : null,
        keepOpaque: keep ? getComputedStyle(keep).opacity : null
      };
    }, w);
    t('old veil hidden', s.veil, 'none');
    t('top bar visible', s.bar, '1');
    t('message line in the result area', s.line && s.line.startsWith('Fetching component prices'), true);
    t('result area fades', s.fades, true);
    t('header/footer stay crisp', s.keepOpaque, '1');
    await page.screenshot({ path: path.join(OUT, 'loading-' + w.file.replace('.html', '') + '.png') });

    await page.evaluate(w => document.getElementById(w.ov).classList.remove(w.cls), w);
    await page.waitForTimeout(300);
    const c = await page.evaluate(() => ({ on: document.body.classList.contains('wk-on'), lines: document.querySelectorAll('.wk-line').length, btns: document.querySelectorAll('.wk-btn').length }));
    t('clears completely', c, { on: false, lines: 0, btns: 0 });

    if (w.file === 'KitCalculatorModal.html') {
      const inp = await page.$('input[type=text], #kitSku');
      if (inp) await inp.fill('217475');
      await page.click('#goKit');
      await page.waitForTimeout(400);
      const r = await page.evaluate(() => ({ on: document.body.classList.contains('wk-on'), btn: !!document.querySelector('#goKit.wk-btn') }));
      t('REAL Compute click → in-place look + its button spins', r, { on: true, btn: true });
      await page.screenshot({ path: path.join(OUT, 'loading-KitCalculator-real.png') });
    }
    t('no page errors', errs, []);
    await page.close();
  }
  await browser.close();
  console.log('\n' + pass + ' passed · ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
})();
