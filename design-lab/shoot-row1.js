const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1200, height: 560 }, deviceScaleFactor: 2 });
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  await p.goto('file://' + path.join(__dirname, 'mockup-row1.html'), { waitUntil: 'networkidle' });
  await p.waitForSelector('html[data-ready="1"]');
  await p.evaluate(() => document.fonts.ready);
  const out = path.join(__dirname, 'renders', 'mast', '_row1.png');
  await p.screenshot({ path: out, fullPage: true });
  await b.close();
  if (errs.length) console.log('⚠ ' + errs.join('\n⚠ '));
  console.log('row1 mockup → ' + out);
})();
