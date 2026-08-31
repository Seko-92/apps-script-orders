/*
 * probe-video.js — drives the REAL video.html in a real browser.
 * ---------------------------------------------------------------------------
 * ⚠ SERVED FROM A REAL ORIGIN, never page.setContent — setContent leaves an
 *   opaque origin where localStorage THROWS, and this page reads it at boot,
 *   which would kill every statement after it. 30 phantom failures were spent
 *   learning that on the sidebar harness.
 * ⚠ charset=utf-8 on the fulfilled response is LOAD-BEARING — without it
 *   Chromium falls back to windows-1252 and every · and ⚠ is mojibake that
 *   looks exactly like a string bug.
 *
 *   node probe-video.js            # UI + refusals, youtube intercepted
 *   LIVE=1 node probe-video.js     # ALSO plays a real video (real network)
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ROOT = process.env.SRC || path.join(__dirname, '..');
const LIVE = !!process.env.LIVE;
const SHOT = path.join(__dirname, 'renders');
const HTML = fs.readFileSync(path.join(ROOT, 'video.html'), 'utf8');
const PROVEN_ID = 'Rs7St51oDDc';   // the board's Sunna · Madinah — proven embeddable

let pass = 0, fail = 0;
const t = (n, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? (pass++, console.log('  ✓ ' + n))
     : (fail++, console.log('  ✗ ' + n + '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};

(async () => {
  if (!fs.existsSync(SHOT)) fs.mkdirSync(SHOT, { recursive: true });
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  ctx.setDefaultTimeout(4000);
  const page = await ctx.newPage();

  const errors = [];
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('pageerror', e => errors.push('PAGEERROR ' + e.message));

  let ytScriptRequested = false;
  await page.route('**/*', route => {
    const u = route.request().url();
    if (u === 'http://hq.test/' || u === 'http://hq.test/video') {
      return route.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: HTML });
    }
    if (/youtube\.com|ytimg|googlevideo/.test(u)) {
      ytScriptRequested = true;
      if (LIVE) return route.continue();
      return route.abort();
    }
    return route.continue();
  });

  await page.goto('http://hq.test/');
  await page.waitForTimeout(400);

  console.log('THE PAGE STANDS UP');
  t('no console errors on boot', errors, []);
  t('the idle face is showing', await page.locator('.hqv-mark').isVisible(), true);
  t('  … and says what to do', (await page.locator('.hqv-sub').textContent()).trim(), 'nothing playing');
  t('the dock is present', await page.locator('#box').isVisible(), true);
  t('nothing is playing yet', await page.locator('body.hqv-playing').count(), 0);

  console.log('\n⚠ LAZY — an unopened page costs NOTHING');
  t('the YouTube API was never fetched', ytScriptRequested, false);
  t('  … and no iframe exists', await page.locator('iframe').count(), 0);

  console.log('\n⚠ EVERY REFUSAL SAYS WHY (silence reads as broken)');
  const refuse = async (text, must) => {
    await page.fill('#box', text);
    await page.click('.hqv-go:not(.hqv-stop)');
    await page.waitForTimeout(120);
    const vis = await page.locator('#msg.on.err').count();
    const msg = (await page.locator('#msg').textContent()).trim();
    t('refused: ' + (text || '(empty)').slice(0, 34), vis > 0 && must.test(msg), true);
    return msg;
  };
  await refuse('http://stream.example.com/a.mp3', /https/i);
  await refuse('https://www.youtube.com/@chan',   /no video id/i);
  await refuse('hello there',                     /not a link/i);
  await refuse('',                                /paste/i);
  t('  … and none of them started a player', await page.locator('body.hqv-playing').count(), 0);
  t('  … nor saved anything', await page.locator('.hqv-chip').count(), 0);

  console.log('\nA GOOD LINK PLAYS, IS SAVED, AND NAMES ITSELF');
  await page.fill('#box', 'https://youtu.be/' + PROVEN_ID);
  await page.click('.hqv-go:not(.hqv-stop)');
  await page.waitForTimeout(LIVE ? 4500 : 600);
  t('the page went into playing state', await page.locator('body.hqv-playing').count(), 1);
  t('the box was cleared for the next paste', await page.inputValue('#box'), '');
  t('it was saved as a chip', await page.locator('.hqv-chip').count(), 1);
  t('  … marked as the live one', await page.locator('.hqv-chip.live').count(), 1);
  t('the YouTube API was fetched only NOW', ytScriptRequested, true);

  if (LIVE) {
    const state = await page.evaluate(() => {
      const f = document.querySelector('#ytHost');
      return { tag: f ? f.tagName.toLowerCase() : null, iframes: document.querySelectorAll('iframe').length };
    });
    // ⚠ YT.Player SWAPS the div for an iframe that KEEPS THE SAME ID — two
    //   harness bugs were written against the opposite assumption.
    t('LIVE · the host became an iframe (same id)', state.tag, 'iframe');
    t('LIVE · exactly one player exists', state.iframes, 1);
    const err = (await page.locator('#msg.on.err').count()) > 0
      ? (await page.locator('#msg').textContent()).trim() : '';
    console.log('  ℹ LIVE player message: ' + (err || '(none — it loaded clean)'));
    await page.screenshot({ path: path.join(SHOT, 'video-live.png') });
  }

  await page.screenshot({ path: path.join(SHOT, 'video-playing.png') });

  console.log('\n⚠ REMOVE MUST NOT PLAY THE THING YOU ARE DELETING');
  await page.fill('#box', 'https://youtu.be/dQw4w9WgXcQ');
  await page.click('.hqv-go:not(.hqv-stop)');
  await page.waitForTimeout(300);
  t('two saved now', await page.locator('.hqv-chip').count(), 2);
  const firstName = (await page.locator('.hqv-chip').first().locator('.nm').textContent()).trim();
  await page.locator('.hqv-chip').first().hover();
  await page.locator('.hqv-chip').first().locator('.hqv-x').click();
  await page.waitForTimeout(200);
  t('one left after remove', await page.locator('.hqv-chip').count(), 1);
  t('  … and it is the OTHER one',
    (await page.locator('.hqv-chip').first().locator('.nm').textContent()).trim() !== firstName, true);

  console.log('\nSAVED LIST SURVIVES A RELOAD (it is per-device, in localStorage)');
  await page.reload();
  await page.waitForTimeout(400);
  t('the chip came back', await page.locator('.hqv-chip').count(), 1);
  t('  … but nothing auto-played', await page.locator('body.hqv-playing').count(), 0);
  t('  … and the idle face is back', await page.locator('.hqv-mark').isVisible(), true);

  console.log('\nSTOP RETURNS IT TO THE LOUNGE');
  await page.locator('.hqv-chip').first().click();
  await page.waitForTimeout(300);
  t('playing again', await page.locator('body.hqv-playing').count(), 1);
  await page.click('.hqv-stop');
  await page.waitForTimeout(250);
  t('stopped', await page.locator('body.hqv-playing').count(), 0);
  t('  … idle face restored', await page.locator('.hqv-mark').isVisible(), true);
  t('  … and no chip is marked live', await page.locator('.hqv-chip.live').count(), 0);

  await page.screenshot({ path: path.join(SHOT, 'video-idle.png') });

  console.log('\nPHONE / TABLET VIEWPORTS');
  for (const [w, h, tag] of [[800, 1280, 'tablet-portrait'], [390, 844, 'phone']]) {
    await page.setViewportSize({ width: w, height: h });
    await page.waitForTimeout(250);
    const box = await page.locator('#box').boundingBox();
    const dock = await page.locator('#dock').boundingBox();
    t(tag + ' · the paste box is on screen', box.x >= 0 && box.x + box.width <= w + 1, true);
    t(tag + ' · the dock sits inside the viewport', Math.round(dock.y + dock.height) <= h + 1, true);
    await page.screenshot({ path: path.join(SHOT, 'video-' + tag + '.png') });
  }

  /* ⚠ THE HARNESS CAUSES ITS OWN NOISE. In offline mode we deliberately
     route.abort() every youtube.com request, and Chromium logs each one as
     net::ERR_FAILED. Those are MINE, not the page's — filtering them is the
     honest thing; asserting on them would be the harness accusing working
     code, which is the most frequent failure mode in this project. */
  const real = errors.filter(e => !/ERR_FAILED|ERR_BLOCKED|net::/.test(e));
  t('still no console errors at the end', real, []);

  await browser.close();
  console.log('\n' + (fail ? '❌ ' : '✅ ') + pass + ' passed · ' + fail + ' failed');
  console.log('renders → design-lab/renders/video-*.png');
  process.exit(fail ? 1 : 0);
})();
