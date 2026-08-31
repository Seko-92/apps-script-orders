// ============================================================================
// SAFE-AREA / FOOTER DIAGNOSTIC
//
// The 2026-08-12 tablet report: the footer (PRINT · picker · ⋯) sat under
// Android's gesture-nav bar in BOTH orientations. Headless has no system bars,
// so a plain render CANNOT reproduce it — which is exactly why the bug shipped.
//
// So this runs every viewport TWICE:
//   inset 0px   → the desktop/wall/PWA case. Nothing may move. (regression net)
//   inset 48px  → a device WITH a nav bar, simulated by overriding the
//                 --safe-* tokens that :root feeds from env(). The whole
//                 chain responds exactly as it does on glass.
//
// THE ASSERTION THAT MATTERS: every interactive footer control's bottom edge
// must sit ABOVE the system-bar line (innerHeight - inset). "Visible" is not
// the bar — "reachable" is: a half-covered 44px target is a mis-tap.
//
// Usage: node diag-footer.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

// BOARD_FILE lets this run against an older revision for before/after proof:
//   git show HEAD:FloorBoard.html > /tmp/before.html
//   BOARD_FILE=/tmp/before.html TAG=before node diag-footer.js
const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const TAG   = process.env.TAG ? process.env.TAG + '-' : '';
const MOCK  = require('./mock-tick.js');
const OUT   = path.join(__dirname, 'renders');

const SIZES = [
  ['tablet-portrait',   800, 1280, true],
  ['tablet-landscape', 1280,  800, true],
  ['tablet-land-1600', 1600,  795, true],   // what the real landscape shot implies
  ['phone',             390,  844, true],
  ['wall',             1920, 1080, false]
];
const INSETS = [0, 48];

let failures = [];

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const html = fs.readFileSync(BOARD, 'utf8');
  const browser = await chromium.launch();

  for (const inset of INSETS) {
    console.log(`\n${'='.repeat(64)}\n  SYSTEM-BAR INSET = ${inset}px  ${inset === 0
      ? '(desktop / wall / installed PWA — nothing may move)'
      : '(Android gesture-nav bar — the reported case)'}\n${'='.repeat(64)}`);

    for (const [name, w, h, touch] of SIZES) {
      const ctx = await browser.newContext({
        viewport: { width: w, height: h }, hasTouch: touch,
        isMobile: touch && w < 500, timezoneId: 'America/Chicago'
      });
      const page = await ctx.newPage();
      const errs = [];
      page.on('pageerror', e => errs.push('pageerror: ' + e.message));
      // The prayer/weather fetches are aborted by this harness on purpose;
      // their load failures are OUR noise, not the board's.
      page.on('console', m => {
        if (m.type() !== 'error') return;
        const t = m.text();
        if (/Failed to load resource|ERR_FAILED|ERR_ABORTED/.test(t)) return;
        errs.push(t);
      });

      await page.route('http://hqlab.test/**', route => {
        const url = route.request().url();
        if (url.includes('/api/board')) {
          const body = JSON.parse(route.request().postData() || '{}');
          let res = { ok: false, message: 'unknown action' };
          if (body.action === 'boardTick')   res = Object.assign({ ok: true }, MOCK);
          if (body.action === 'boardStatus') res = { ok: true };
          if (body.action === 'boardRadio')  res = { ok: true, nowPlaying: '' };
          return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
        }
        return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
      });
      await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());

      // Simulate the device's system bars. :root defines --safe-* from env();
      // env() is always 0 in headless, so we override the tokens themselves —
      // every consumer downstream behaves exactly as it would on glass.
      if (inset) {
        await page.addInitScript(px => {
          document.addEventListener('DOMContentLoaded', () => {
            const s = document.createElement('style');
            s.textContent = `:root{--safe-b:${px}px;--safe-t:0px;--safe-l:0px;--safe-r:0px;}`;
            document.head.appendChild(s);
          });
        }, inset);
      }

      await page.goto('http://hqlab.test/', { waitUntil: 'load' });
      await page.waitForFunction(
        () => !document.getElementById('board').classList.contains('booting'),
        null, { timeout: 20000 }
      ).catch(() => errs.push('board never left booting'));
      await page.waitForTimeout(1800);

      const m = await page.evaluate(inset => {
        // ⚠ SIMULATE A STATION PLAYING A LONG TRACK. Reported 2026-08-14: with
        // the radio on, the now-playing text grew the footer ~188px past an
        // 800px portrait screen and pushed PRINT and the ⋯ menu off the right
        // edge. Headless starts with the radio idle, so the bar looks fine —
        // the bug only exists once something is playing. Force it.
        const rm = document.getElementById('radioMeta');
        if (rm) rm.textContent =
          'SomaFM Drone Zone — Steve Roach · Structures from Silence (Remastered 2024)';

        const safeBottom = window.innerHeight - inset;   // the system-bar line
        const r = sel => {
          const e = document.querySelector(sel);
          if (!e) return null;
          const b = e.getBoundingClientRect();
          return { top: Math.round(b.top), bottom: Math.round(b.bottom), h: Math.round(b.height) };
        };
        // Every control a picker must actually TAP in the footer.
        const controls = [...document.querySelectorAll(
          '.ftr .ftr-btn, .ftr .today-bar, .ftr .ftr-picker'
        )].map(e => {
          const b = e.getBoundingClientRect();
          return {
            label: (e.textContent || '').trim().slice(0, 18).replace(/\s+/g, ' '),
            bottom: Math.round(b.bottom), top: Math.round(b.top),
            right: Math.round(b.right), left: Math.round(b.left)
          };
        }).filter(c => c.bottom > 0);
        // A CLOSED drawer must be entirely off-screen — not merely mostly.
        // Measured, because a shadow check alone let a 49px peek pass.
        const sheet = document.querySelector('.today-sheet');
        const sheetRect = sheet ? sheet.getBoundingClientRect() : null;
        // ⚠ FULLSCREEN PLACEMENT. Floor report 2026-08-13: nobody could find
        // it inside the ⋯ menu. It must live in the BAR — visible without
        // opening anything — and still be a real target.
        const fs = document.getElementById('fsToggle');
        const fsRect = fs ? fs.getBoundingClientRect() : null;
        return {
          fsExists: !!fs,
          fsInMenu: fs ? !!fs.closest('#menuPop') : null,
          fsInBar: fs ? !!fs.closest('.ftr') : null,
          fsH: fsRect ? Math.round(fsRect.height) : null,
          fsVisible: fs ? (getComputedStyle(fs).display !== 'none' && fsRect.width > 0) : null,
          fsText: fs ? fs.textContent.trim() : null,
          sheetClosedTop: sheetRect ? Math.round(sheetRect.top) : null,
          sheetClosedH: sheetRect ? Math.round(sheetRect.height) : null,
          sheetOpen: sheet ? sheet.classList.contains('open') : null,
          innerHeight: window.innerHeight, innerWidth: window.innerWidth, safeBottom,
          docScrollW: document.documentElement.scrollWidth,
          docScrollH: document.documentElement.scrollHeight,
          boardH: Math.round(document.getElementById('board').getBoundingClientRect().height),
          bar: r('.bar'), ftr: r('.ftr'),
          controls,
          worstControl: controls.reduce((a, c) => Math.max(a, c.bottom), 0),
          sheetClosedShadow: sheet ? getComputedStyle(sheet).boxShadow : null,
          rows: document.querySelectorAll('.pick-row').length
        };
      }, inset);

      // ── assertions ────────────────────────────────────────────────────
      const problems = [];
      if (m.docScrollH > m.innerHeight + 1)
        problems.push(`page scrolls (${m.docScrollH} > ${m.innerHeight})`);
      if (m.ftr.bottom > m.innerHeight + 1)
        problems.push(`footer ground ${m.ftr.bottom - m.innerHeight}px past screen`);
      for (const c of m.controls) {
        if (c.bottom > m.safeBottom + 1)
          problems.push(`"${c.label}" ${c.bottom - m.safeBottom}px under the system bar`);
        // ⚠ HORIZONTAL TOO. A control pushed off the right edge by the radio's
        // now-playing text is just as unreachable as one under the nav bar.
        if (c.right > m.innerWidth + 1)
          problems.push(`"${c.label}" ${c.right - m.innerWidth}px PAST THE RIGHT EDGE (radio overflow)`);
        if (c.left < -1)
          problems.push(`"${c.label}" pushed ${-c.left}px off the LEFT edge`);
      }
      if (m.sheetClosedShadow && m.sheetClosedShadow !== 'none')
        problems.push(`closed drawer still casts a shadow (${m.sheetClosedShadow})`);
      if (!m.sheetOpen && m.sheetClosedTop !== null && m.sheetClosedTop < m.innerHeight)
        problems.push(`closed drawer PEEKS ${m.innerHeight - m.sheetClosedTop}px over the footer`
                      + ` (top ${m.sheetClosedTop}, h ${m.sheetClosedH})`);
      if (!m.fsExists) problems.push('fullscreen control missing entirely');
      else {
        if (m.fsInMenu) problems.push('fullscreen is still buried in the ⋯ menu');
        if (!m.fsInBar) problems.push('fullscreen is not in the footer bar');
        // The phone deliberately sheds ornaments, so only hold the bar to the
        // touch-target rule where it is actually shown.
        if (m.fsVisible && m.fsH < 40) problems.push(`fullscreen target only ${m.fsH}px tall`);
      }
      if (errs.length) problems.push('JS: ' + errs.join(' | '));

      const ok = problems.length === 0;
      if (!ok) failures.push(`${name}@${inset}px — ${problems.join('; ')}`);

      console.log(`\n  ${ok ? '✓' : '✗'} ${name} (${w}×${h})   rows=${m.rows}`);
      console.log(`      screen ${m.innerHeight}  ·  system-bar line ${m.safeBottom}  ·  board ${m.boardH}`);
      console.log(`      bar  ${JSON.stringify(m.bar)}`);
      console.log(`      ftr  ${JSON.stringify(m.ftr)}  → lowest control bottom ${m.worstControl}`
                  + `  (headroom ${m.safeBottom - m.worstControl}px)`);
      if (!ok) problems.forEach(p => console.log(`      ✗ ${p}`));

      await page.screenshot({ path: path.join(OUT, `${TAG}safe-${inset}-${name}.png`) });
      await ctx.close();
    }
  }

  await browser.close();
  console.log('\n' + '='.repeat(64));
  if (failures.length) {
    console.log('✗ FAILURES:\n' + failures.map(f => '   ' + f).join('\n'));
    process.exit(1);
  }
  console.log('✓ ALL VIEWPORTS × BOTH INSETS: footer reachable, page does not scroll,');
  console.log('  closed drawer casts nothing.');
})().catch(e => { console.error('CRASH', e); process.exit(1); });
