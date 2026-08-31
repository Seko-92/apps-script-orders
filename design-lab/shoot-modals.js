// ============================================================================
// THE FOUR COMMIT MODALS — render the REAL files headless and inventory them.
//
// These are the surfaces you sit in and work: Kit Expansion runs a per-kit
// queue with swaps and qty edits, Zoho Pull is a per-line diff review, the
// Calculator has editable per-line prices, Price Push is review-then-confirm.
// They were never renderable here — each is an HtmlService template opened with
// showModalDialog, so nothing could see them. This resolves their scriptlets
// against realistic fixtures and serves them like the sidebar harness does.
//
// ⚠ EACH MODAL IS THE REAL FILE, UNTOUCHED apart from resolving its injections.
// If it renders wrong here, it renders wrong in the sheet.
//
// Usage: node shoot-modals.js            (renders + inventories)
//        TAG=before node shoot-modals.js (label the output)
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const ROOT = path.join(__dirname, '..');
const OUT = path.join(__dirname, 'renders');
const TAG = process.env.TAG || 'modal';

// ── fixtures ────────────────────────────────────────────────────────────────
// Shaped from the real producers: KitExpansion.js queue.push, ZohoPull.js
// computeZohoSoDiff, PriceWriteback.js candidates.push, KitPricing.js kits[].
const KIT = {
  kitSku: '217475', sourceSalesOrder: 'SO-24853', sourceQty: 2,
  sourceNote: 'Miguel · pack with care', originalRow: 47, table: 2,
  userMultiplier: 1, kitName: 'Engine Overhaul Kit 0.50, 16423-21910',
  kitType: 'MANUAL', kitLocation: 'K-55', kitEngine: 'D1105',
  salesDescription: 'Full Gasket Set 19077-03310 + Head Gasket\nPiston Kit STD 167517\nMain Bearing Set 159579',
  components: [
    { sku: '167517', name: 'Piston Kit STD',            qty: 4, location: 'E-37',  available: 12,   missing: false },
    { sku: '168138', name: 'Full Gasket Set',           qty: 1, location: 'L-226', available: 3,    missing: false },
    { sku: '162198', name: 'Head Gasket',               qty: 1, location: 'I-53',  available: 8,    missing: false,
      bundled: true, bundledIn: 'Full Gasket Set 19077-03310' },
    { sku: '159579', name: 'Main Bearing Set STD',      qty: 1, location: 'E-12',  available: 0,    missing: false },
    { sku: '195306', name: 'Thrust Washer Set',         qty: 2, location: 'NOT FOUND', available: null, missing: true }
  ],
  unparsedLines: ['-1 Head Gasket (161262)'],
  alreadyExpanded: { count: 5, skus: ['167517', '168138', '162198', '159579', '195306'] }
};

const DIFF = {
  ok: true, reason: '', soNumber: 'SO-24853', matchedVia: 'so', pendingRowFound: true,
  isFirstPull: false, pulled: true, pulledAt: '8/19/26 2:38 PM',
  customerName: 'Triple M Equipment', totalFormatted: '$5,347.79',
  zohoStatus: 'confirmed', zohoShippedStatus: 'pending',
  invoiceNumber: 'INV-022759', pendingLastUpdated: '8/19/26 6:02 PM',
  lines: [
    { sku: '167517', name: 'Piston Kit STD', status: 'unchanged', zohoQty: 4, directQty: 4,
      delta: 0, location: 'E-37', available: 12, missing: false, directRows: [{ row: 48, qty: 4, status: 'PENDING' }] },
    { sku: '168138', name: 'Full Gasket Set', status: 'new', zohoQty: 2, directQty: 0,
      delta: 2, location: 'L-226', available: 3, missing: false, directRows: [] },
    { sku: '159579', name: 'Main Bearing Set STD', status: 'qty_changed', zohoQty: 6, directQty: 3,
      delta: 3, location: 'E-12', available: 0, missing: false, directRows: [{ row: 51, qty: 3, status: 'PREPARING' }] },
    { sku: '195306', name: 'Thrust Washer Set', status: 'removed', zohoQty: 0, directQty: 2,
      delta: -2, location: 'NOT FOUND', available: null, missing: true, directRows: [{ row: 53, qty: 2, status: 'PENDING' }] }
  ],
  summary: { totalLines: 4, unchanged: 1, new: 1, qtyChanged: 1, removed: 1, anyChanges: true }
};

const PUSH = {
  sessionId: '11111111-2222-3333-4444-555555555555',
  candidates: [
    { sku: '170154', name: 'Piston Ring Set STD', direction: 'ZOHO LOW',  zohoBefore: 13.94,
      ebayTarget: 16.00, delta: 2.06,  bigSwing: false, itemId: '9001', pushable: true,  skipReason: '' },
    { sku: '170145', name: 'Valve Guide Set',     direction: 'ZOHO LOW',  zohoBefore: 13.00,
      ebayTarget: 14.00, delta: 1.00,  bigSwing: false, itemId: '9002', pushable: true,  skipReason: '' },
    { sku: '155457', name: 'Full Gasket Set',     direction: 'ZOHO HIGH', zohoBefore: 79.99,
      ebayTarget: 59.99, delta: -20.00, bigSwing: true,  itemId: '9003', pushable: true,  skipReason: '' },
    { sku: '218204', name: 'Crankshaft Assembly', direction: 'OOS / NO REF', zohoBefore: 340.00,
      ebayTarget: null, delta: null, bigSwing: false, itemId: '', pushable: false, skipReason: 'no eBay price' }
  ],
  meta: { pushable: 3, skipped: 1, cap: 30, hasPassphrase: true, zohoSyncedAt: Date.now() - 90 * 1000 }
};

const KITLIST = [
  { sku: '159093', name: 'Engine Overhaul Kit 0.50, 16423-21910' },
  { sku: '215756', name: 'Engine Overhaul Kit 0.50, 16423-21910' },
  { sku: '217475', name: 'Engine Overhaul, Rebuild Kit STD, 16423-21110' },
  { sku: '217853', name: 'Engine Repair Kit 0.25' }
];

const MODALS = [
  { file: 'KitExpansionModal.html', tag: 'kit-expand', w: 1180, h: 820,
    subs: [[/<\?!?=\s*sessionId\s*\?>/g, JSON.stringify('sess-abc')],
           [/<\?!?=\s*kitJson\s*\?>/g,   JSON.stringify(KIT)],
           [/<\?!?=\s*queueLength\s*\?>/g, '3'],
           [/<\?!?=\s*kitIndex\s*\?>/g,  '0']] },
  { file: 'ZohoPullModal.html', tag: 'zoho-pull', w: 920, h: 620,
    subs: [[/<\?!?=\s*diffJson\s*\?>/g, JSON.stringify(DIFF)]] },
  { file: 'PricePushModal.html', tag: 'price-push', w: 880, h: 620,
    subs: [[/<\?!?=\s*dataJson\s*\?>/g, JSON.stringify(PUSH)]] },
  { file: 'KitCalculatorModal.html', tag: 'kit-calc', w: 1000, h: 700,
    subs: [[/<\?!?=\s*kitListJson\s*\?>/g, JSON.stringify(KITLIST)],
           [/<\?!?=\s*defaultDiscount\s*\?>/g, '0.10']] }
];

// The glyphs this audit is actually about — the same MEANING spelled different
// ways across four files is the "one soul" break, not emoji density.
const VOCAB = { '✓': 'ok', '✅': 'ok', '✔': 'ok',
                '✗': 'fail', '❌': 'fail', '✕': 'close', '✖': 'close',
                '⚠': 'warn', '⚠️': 'warn' };

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch();
  const seen = {};

  for (const m of MODALS) {
    let html = fs.readFileSync(path.join(ROOT, m.file), 'utf8');
    for (const [re, val] of m.subs) html = html.replace(re, val);
    const leftover = html.match(/<\?[^>]{0,40}\?>/g);
    if (leftover) throw new Error(m.file + ': unresolved scriptlet ' + leftover.join(', '));

    const page = await browser.newPage({ viewport: { width: m.w, height: m.h }, deviceScaleFactor: 2 });
    const errs = []; page.on('pageerror', e => errs.push(String(e)));
    page.on('console', c => { if (c.type() === 'error') errs.push('console: ' + c.text()); });

    // ⚠ Same origin lesson as the sidebar harness: setContent leaves an opaque
    // origin where localStorage throws, and no charset renders '·' as mojibake.
    await page.addInitScript(() => {
      const mk = (su, fa) => new Proxy({}, { get(_, k) {
        if (k === 'withSuccessHandler') return f => mk(f, fa);
        if (k === 'withFailureHandler') return f => mk(su, f);
        return () => { if (su) setTimeout(() => su(null), 0); };
      }});
      window.google = { script: { run: mk(null, null),
        host: { close() {}, setHeight() {}, setWidth() {} },
        url: { getLocation(f) { f({ parameter: {} }); } } } };
    });
    await page.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: html }));
    await page.goto('http://hq.test/' + m.tag, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(900);

    const info = await page.evaluate((vocab) => {
      const txt = document.body.innerText || '';
      const found = {};
      Object.keys(vocab).forEach(g => { const n = (txt.split(g).length - 1); if (n) found[g] = n; });
      const btns = [...document.querySelectorAll('button')];
      return {
        glyphs: found,
        buttons: btns.length,
        btnLabels: btns.map(b => (b.innerText || '').trim().replace(/\s+/g, ' ')).filter(Boolean).slice(0, 12),
        hasVeil: !!document.querySelector('[class*="overlay"], [class*="veil"], [class*="loading"]'),
        bodyBg: getComputedStyle(document.body).backgroundColor,
        fonts: [...new Set([...document.querySelectorAll('h1,h2,h3,button,.hdr-title')]
          .map(e => getComputedStyle(e).fontFamily.split(',')[0].replace(/["']/g, '')))].slice(0, 4),
        height: document.body.scrollHeight
      };
    }, VOCAB);

    seen[m.tag] = info;
    console.log('\n═══ ' + m.file + ' ═══');
    console.log('  buttons     :', info.buttons, '|', info.btnLabels.slice(0, 6).join(' · '));
    console.log('  vocabulary  :', Object.keys(info.glyphs).length
      ? Object.entries(info.glyphs).map(([g, n]) => g + '×' + n).join('  ') : '(none rendered)');
    console.log('  loading veil:', info.hasVeil ? 'yes' : 'NO');
    console.log('  body bg     :', info.bodyBg, ' fonts:', info.fonts.join(', '));
    if (errs.length) console.log('  ⚠ errors    :', errs.slice(0, 3));

    await page.screenshot({ path: path.join(OUT, `${TAG}-${m.tag}.png`) });

    // ── the loading veil, PROVEN rather than grepped ─────────────────────────
    // ⚠ The class-name detector above is loose (it matches anything with
    // "loading" in a class). This drives the real code path and reads the
    // COMPUTED backdrop-filter, which is what actually makes the veil frosted
    // rather than a black box — the 2026-07-31 finding.
    const veil = await page.evaluate(() => {
      const fn = (typeof setLoading === 'function') ? setLoading
               : (typeof showOverlay === 'function') ? showOverlay : null;
      if (!fn) return { drove: false };
      try { fn(true); } catch (e) { return { drove: false, err: String(e) }; }
      const on = [...document.querySelectorAll('*')].find(
        e => /overlay/i.test(e.className || '') && e.classList.contains('shown'));
      if (!on) return { drove: true, shown: false };
      const cs = getComputedStyle(on);
      return { drove: true, shown: true,
               blur: cs.backdropFilter || cs.webkitBackdropFilter || 'none',
               bg: cs.backgroundColor };
    });
    if (veil.drove) {
      const okv = veil.shown && /blur/.test(veil.blur || '');
      console.log('  veil (driven):', okv ? '✓ frosted — ' + veil.blur
                                          : '✗ ' + JSON.stringify(veil));
      if (okv) await page.screenshot({ path: path.join(OUT, `${TAG}-${m.tag}-veil.png`) });
    } else {
      console.log('  veil (driven): no setLoading/showOverlay to drive');
    }
    await page.close();
  }

  // ── the finding this harness exists for ───────────────────────────────────
  console.log('\n═══ ONE MEANING, HOW MANY GLYPHS? ═══');
  const byMeaning = {};
  Object.entries(seen).forEach(([tag, i]) => {
    Object.keys(i.glyphs).forEach(g => {
      const mean = VOCAB[g];
      (byMeaning[mean] = byMeaning[mean] || {})[g] = (byMeaning[mean][g] || []).concat
        ? (byMeaning[mean][g] || []).concat(tag) : [tag];
    });
  });
  Object.entries(byMeaning).forEach(([mean, gl]) => {
    const spellings = Object.keys(gl);
    const flag = spellings.length > 1 ? '  ← SPLIT' : '';
    console.log('  %s: %s%s', mean.padEnd(6),
      spellings.map(g => g + ' (' + gl[g].join(',') + ')').join('   '), flag);
  });
  console.log('\nrenders → ' + OUT);
  await browser.close();
})();
