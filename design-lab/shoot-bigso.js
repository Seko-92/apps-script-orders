// The live board's Direct tab with the big SO open, at the warehouse tablet sizes.
const { chromium } = require('playwright');
(async () => {
  const out = process.argv[2] || '/tmp';
  const browser = await chromium.launch();
  for (const [name, vp] of [['tablet-land', [1280, 800]], ['tablet-port', [800, 1280]]]) {
    const ctx = await browser.newContext({ viewport: { width: vp[0], height: vp[1] }, hasTouch: true, isMobile: false });
    const page = await ctx.newPage();
    await page.goto('https://hq.yassinqurabi.com/', { waitUntil: 'load', timeout: 45000 });
    await page.waitForFunction(() => window.lastTick && window.lastTick.openOrders, null, { timeout: 45000 }).catch(() => {});
    await page.waitForTimeout(2500);
    const info = await page.evaluate(() => {
      const t = window.lastTick || {};
      const oo = t.openOrders || [];
      const big = oo.filter(r => r.orderId === 'SO-25980');
      return { open: oo.length, total: t.openOrdersTotal, perChannel: t.openOrdersByChannel || null,
               big: big.length, bigStatuses: [...new Set(big.map(r => r.status))],
               tabs: [...document.querySelectorAll('[data-chan], .chan-tab, .tab')].map(e => e.textContent.trim()).slice(0, 6) };
    });
    console.log(name, JSON.stringify(info));
    // open the Direct tab, then the big order's card
    const tab = await page.$('text=/^\\s*Direct/i'); if (tab) { await tab.click().catch(() => {}); await page.waitForTimeout(800); }
    await page.screenshot({ path: `${out}/bigso-${name}-1.png` });
    const card = await page.$('text=SO-25980'); if (card) { await card.click().catch(() => {}); await page.waitForTimeout(1200); }
    await page.screenshot({ path: `${out}/bigso-${name}-2.png` });
    const more = await page.evaluate(() => [...document.querySelectorAll('*')].map(e => e.childElementCount === 0 ? e.textContent : '').filter(t => /more|not shown|\+\d+/.test(t)).slice(0, 8));
    console.log('  notes:', JSON.stringify(more));
    await ctx.close();
  }
  await browser.close();
})();
