const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1210, height: 1800 }, deviceScaleFactor: 2 });
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  await p.goto('file://' + path.join(__dirname, 'mockup-banner.html'), { waitUntil: 'networkidle' });
  await p.waitForSelector('html[data-ready="1"]');
  await p.evaluate(() => document.fonts.ready);
  await p.waitForTimeout(600);
  const out = path.join(__dirname, 'renders', 'mast', '_banner.png');
  await p.screenshot({ path: out, fullPage: true });
  await b.close();
  if (errs.length) console.log('⚠ ' + errs.join('\n⚠ '));
  console.log('banner mockup → ' + out);
})();
