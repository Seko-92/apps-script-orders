/**
 * shoot-kitbuild.js — render the Kit Build modal, screen AND print.
 *
 * ⚠ Serves over a REAL ORIGIN (page.route + http://hq.test/), never setContent —
 *   setContent leaves an opaque origin where localStorage throws, and one throw at
 *   the top of the script kills every statement after it.
 * ⚠ charset=utf-8 on the fulfilled response, or every · and → is mojibake that
 *   looks exactly like a string bug.
 *
 *   node shoot-kitbuild.js
 */
const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');

const SRC = path.join(__dirname, '..', 'KitBuildModal.html');
const OUT = path.join(__dirname, 'renders');

// The fixture the server would send — two kits SHARING a piston, one bundled part.
const KITS = [
  { sku:'158679', name:'Engine Overhaul Kit STD', type:'READY',  location:'K-55' },
  { sku:'217205', name:'Repair Kit 0.50',         type:'MANUAL', location:'NOT FOUND' },
  { sku:'171018', name:'Main Bearing Set STD',    type:'MANUAL', location:'E-17' }
];

const PLAN = {
  ok: true, picker: 'Yassin · 1', builtAt: '9/3/26 10:14 AM',
  warnings: [],
  totals: { kits:2, kitUnits:5, lines:5, pieces:24, shelves:5, notFound:0, shared:1 },
  kits: [
    { found:true, kitSku:'158679', kitName:'Engine Overhaul Kit STD', kitType:'READY',
      kitLocation:'K-55', kitEngine:'V2203', qty:2, unparsedLines:[], components:[
        { sku:'155394', name:'Full Gasket Set', qty:2, location:'A-50', available:12, excluded:false, bundled:false, bundledInto:'', missing:false, short:false },
        { sku:'167517', name:'Piston Kit STD',  qty:4, location:'A-9',  available:18, excluded:false, bundled:false, bundledInto:'', missing:false, short:false },
        { sku:'162198', name:'Head Gasket',     qty:2, location:'L-226',available:7,  excluded:true,  bundled:true,  bundledInto:'155394', missing:false, short:false },
        { sku:'171018', name:'Main Bearing Set',qty:2, location:'E-17', available:3,  excluded:false, bundled:false, bundledInto:'', missing:false, short:false }
      ]},
    { found:true, kitSku:'217205', kitName:'Repair Kit 0.50', kitType:'MANUAL',
      kitLocation:'NOT FOUND', kitEngine:'', qty:3, unparsedLines:[], components:[
        { sku:'167517', name:'Piston Kit STD', qty:9, location:'A-9', available:18, excluded:false, bundled:false, bundledInto:'', missing:false, short:false },
        { sku:'173763', name:'Thrust Washer',  qty:6, location:'B-4', available:2,  excluded:false, bundled:false, bundledInto:'', missing:false, short:true }
      ]}
  ],
  gather: [
    { sku:'167517', name:'Piston Kit STD',   location:'A-9',   available:18, totalQty:13, missing:false, usedBy:[{kitSku:'158679',kits:2,subtotal:4},{kitSku:'217205',kits:3,subtotal:9}] },
    { sku:'155394', name:'Full Gasket Set',  location:'A-50',  available:12, totalQty:2,  missing:false, usedBy:[{kitSku:'158679',kits:2,subtotal:2}] },
    { sku:'173763', name:'Thrust Washer',    location:'B-4',   available:2,  totalQty:6,  missing:false, usedBy:[{kitSku:'217205',kits:3,subtotal:6}] },
    { sku:'171018', name:'Main Bearing Set', location:'E-17',  available:3,  totalQty:2,  missing:false, usedBy:[{kitSku:'158679',kits:2,subtotal:2}] },
    { sku:'900999', name:'Oil Seal',         location:'NOT FOUND', available:null, totalQty:1, missing:true, usedBy:[{kitSku:'217205',kits:3,subtotal:1}] }
  ]
};

const html = fs.readFileSync(SRC, 'utf8')
  .replace('<?!= kitListJson ?>',    JSON.stringify(KITS))
  .replace('<?!= pickerNameJson ?>', JSON.stringify('Yassin · 1'))
  .replace('<?!= maxKitsJson ?>',    '12');

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1180, height: 820 }, deviceScaleFactor: 2 });
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  p.on('console', m => { if (m.type() === 'error') errs.push(m.text()); });

  await p.route('**/*', r => r.fulfill({
    status: 200, contentType: 'text/html; charset=utf-8', body: html }));

  // Stub google.script.run BEFORE the page script runs.
  await p.addInitScript(plan => {
    window.__PLAN = plan;
    const chain = { _s: null, _f: null,
      withSuccessHandler(f) { this._s = f; return this; },
      withFailureHandler(f) { this._f = f; return this; },
      getKitBuildPlan() { const s = this._s; setTimeout(() => s && s(window.__PLAN), 5); }
    };
    window.google = { script: { run: chain, host: { close(){} } } };
  }, PLAN);

  await p.goto('http://hq.test/', { waitUntil: 'domcontentloaded' });

  // Drive it the way a picker would.
  await p.evaluate(() => {
    selected = [{ sku:'158679', qty:2 }, { sku:'217205', qty:3 }];
    refresh();
  });
  await p.waitForTimeout(400);

  await p.screenshot({ path: path.join(OUT, 'kitbuild-screen.png'), fullPage: true });

  await p.emulateMedia({ media: 'print' });
  await p.waitForTimeout(150);
  await p.screenshot({ path: path.join(OUT, 'kitbuild-print.png'), fullPage: true });

  const summary = await p.evaluate(() => ({
    summary: document.getElementById('summary').textContent.trim(),
    printDisabled: document.getElementById('printBtn').disabled,
    gatherRows: document.querySelectorAll('#gatherTable tbody tr:not(.empty-row)').length,
    emptyRows: document.querySelectorAll('#gatherTable tbody tr.empty-row').length,
    asmBlocks: document.querySelectorAll('.asm-block').length,
    firstLoc: document.querySelector('#gatherTable tbody .c-loc').textContent
  }));
  console.log(JSON.stringify(summary, null, 2));
  if (errs.length) { console.log('PAGE ERRORS:'); errs.forEach(e => console.log('  ' + e)); }
  else console.log('no page errors');
  await b.close();
})();
