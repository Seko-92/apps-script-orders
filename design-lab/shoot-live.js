// END-TO-END: render the LIVE production pages against REAL data.
// No route interception — this proves Caddy → n8n → Apps Script → Sheets.
'use strict';
const { chromium } = require('playwright');
(async () => {
  const browser = await chromium.launch();
  for (const [name, url, vp, touch] of [
    ['live-tool-portrait',  'https://hq.yassinqurabi.com/',     [800, 1280], true],
    ['live-tool-landscape', 'https://hq.yassinqurabi.com/',     [1280, 800], true],
    ['live-wall',           'https://hq.yassinqurabi.com/wall', [1920, 1080], false]
  ]) {
    const ctx = await browser.newContext({ viewport: { width: vp[0], height: vp[1] },
      hasTouch: touch, timezoneId: 'America/Chicago' });
    const page = await ctx.newPage();
    const errs = [];
    page.on('pageerror', e => errs.push(e.message));
    await page.goto(url, { waitUntil: 'load', timeout: 45000 });
    // wait for a REAL tick to land
    await page.waitForFunction(() => {
      const b = document.getElementById('board');
      const g = document.getElementById('grabNum');
      return (b ? !b.classList.contains('booting') : true) && g && g.textContent !== '—';
    }, null, { timeout: 60000 }).catch(() => errs.push('no live tick within 60s'));
    await page.waitForTimeout(2500);
    const st = await page.evaluate(() => ({
      grab: (document.getElementById('grabNum') || {}).textContent,
      rows: document.querySelectorAll('.pick-row, .wrow').length,
      bands: document.querySelectorAll('.pick-head, .wband').length,
      live: (document.getElementById('liveText') || {}).textContent,
      lost: !!document.querySelector('.connlost.show')
    }));
    await page.screenshot({ path: 'renders/' + name + '.png' });
    console.log('✓', name.padEnd(20), JSON.stringify(st), errs.length ? '⚠ ' + errs.join(' | ') : '');
    await ctx.close();
  }
  await browser.close();
})().catch(e => { console.error('CRASH', e); process.exit(1); });
