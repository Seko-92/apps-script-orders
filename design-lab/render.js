// ============================================================================
// FLOOR BOARD RENDER LOOP — see what you ship BEFORE the user does.
//
// Serves the REAL FloorBoard.html (read fresh from the repo on every run, so a
// stale snapshot cannot pass) via Playwright route interception, answers
// /api/board with mock-tick.js, and screenshots the four layouts the board
// must survive:
//     wall      1920×1080  fine pointer  → monitor layout
//     tablet-l  1280×800   coarse touch  → tool layout (pointer:coarse branch)
//     tablet-p   800×1280  coarse touch  → tool layout (width branch)
//     phone      390×844   coarse touch  → phone layout
//
// Usage:  node render.js [outdir]     (default ./renders)
// Emits:  <outdir>/board-<name>.png + render-log.json (mode class, pointer
//         media results, console errors) — the honesty file: if pointer:coarse
//         did not emulate, the log SAYS so instead of a silently-wrong render.
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const BOARD = path.join(__dirname, '..', 'FloorBoard.html');
const OUT = process.argv[2] || path.join(__dirname, 'renders');
const MOCK = require('./mock-tick.js');

const VIEWPORTS = [
  { name: 'wall',            width: 1920, height: 1080, touch: false },
  { name: 'tablet-landscape', width: 1280, height: 800,  touch: true  },
  { name: 'tablet-portrait',  width: 800,  height: 1280, touch: true  },
  { name: 'phone',            width: 390,  height: 844,  touch: true  }
];

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const html = fs.readFileSync(BOARD, 'utf8');   // fresh read — never a snapshot
  const browser = await chromium.launch();
  const log = { renderedAt: new Date().toISOString(), boardBytes: html.length, shots: [] };

  for (const vp of VIEWPORTS) {
    const ctx = await browser.newContext({
      viewport: { width: vp.width, height: vp.height },
      hasTouch: vp.touch,
      isMobile: vp.touch && vp.width < 500,
      timezoneId: 'America/Chicago'   // the board lives on Houston time
    });
    const page = await ctx.newPage();

    const errors = [];
    page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
    page.on('pageerror', e => errors.push('pageerror: ' + e.message));

    // The page itself + the board API, both fulfilled locally.
    await page.route('http://hqlab.test/**', route => {
      const url = route.request().url();
      if (url.includes('/api/board')) {
        const body = JSON.parse(route.request().postData() || '{}');
        let res = { ok: false, message: 'unknown action ' + body.action };
        if (body.action === 'boardTick')   res = Object.assign({ ok: true }, MOCK);
        if (body.action === 'boardStatus') res = { ok: true };
        if (body.action === 'boardRadio')  res = { ok: true, nowPlaying: '' };
        return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
      }
      return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
    });
    // Third-party niceties (prayer/weather) — let them fail fast instead of
    // hanging the settle window; the board degrades gracefully by design.
    await page.route(/aladhan\.com|open-meteo\.com/, route => route.abort());

    await page.goto('http://hqlab.test/', { waitUntil: 'load' });
    await page.waitForFunction(
      () => !document.getElementById('board').classList.contains('booting'),
      null, { timeout: 20000 }
    ).catch(() => errors.push('board never left booting state'));
    await page.waitForTimeout(2500);   // fonts + enter animations settle

    const state = await page.evaluate(() => ({
      pointerCoarse: matchMedia('(pointer: coarse)').matches,
      toolLayout:   matchMedia('(max-width: 1200px), (pointer: coarse) and (max-width: 1600px)').matches,
      modeClass:    (document.getElementById('board').className.match(/mode-\w+/) || [''])[0],
      pickRows:     document.querySelectorAll('.pick-row').length,
      fontsReady:   document.fonts ? document.fonts.status : 'unknown'
    }));

    const file = path.join(OUT, 'board-' + vp.name + '.png');
    await page.screenshot({ path: file, fullPage: false });
    log.shots.push(Object.assign({ name: vp.name, file, viewport: vp, errors }, state));
    console.log(`✓ ${vp.name.padEnd(17)} ${vp.width}×${vp.height}  coarse=${state.pointerCoarse}  tool=${state.toolLayout}  ${state.modeClass}  rows=${state.pickRows}  errs=${errors.length}`);
    await ctx.close();
  }

  await browser.close();
  fs.writeFileSync(path.join(OUT, 'render-log.json'), JSON.stringify(log, null, 2));
  console.log('\nRenders in ' + OUT);
})().catch(e => { console.error('RENDER CRASH:', e); process.exit(1); });
