// ============================================================================
// WALL SAFE-AREA DIAGNOSTIC
//
// wall.html draws edge-to-edge. It has NO write surface, so unlike the tool
// nothing here can be mis-tapped — but a monitor's entire job is being READ,
// and a row hidden under Android's gesture-nav bar is a row that isn't doing
// that job. `grep safe-area|viewport-fit|env(` on wall.html returned ZERO hits
// until 2026-08-13; this proves the fix and nets the regression.
//
// Headless has no system bars, so a plain render CANNOT reproduce any of this.
// :root feeds --safe-* from env(), which is always 0 here — so we override the
// TOKENS and every consumer downstream behaves exactly as it does on glass.
//
// THE ASSERTIONS (geometry + behaviour, never "it looks fine"):
//   · page never scrolls — the wall is one fixed screen, by design
//   · header content sits BELOW the top inset (its ground may paint under it)
//   · the last VISIBLE pick row in each channel clears the bottom bar line
//   · the last metric card in the rail clears it too
//   · the conn-lost toast and the rest-veil text stay inside the safe box
//   · left/right insets (landscape nav bar) never clip a column's content
//
// Usage:  node diag-wall-safe.js
// Before/after proof:
//   git show HEAD:wall.html > /tmp/wall-before.html
//   WALL_FILE=/tmp/wall-before.html TAG=before node diag-wall-safe.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const WALL = process.env.WALL_FILE || path.join(__dirname, '..', 'wall.html');
const TAG  = process.env.TAG ? process.env.TAG + '-' : '';
const MOCK = require('./mock-tick.js');
const OUT  = path.join(__dirname, 'renders');

const SIZES = [
  ['wall',             1920, 1080, false],
  ['tablet-landscape', 1280,  800, true],
  ['tablet-portrait',   800, 1280, true]
];

// Each case is a real device shape, not an arbitrary number.
const CASES = [
  { name: 'none',       t: 0,  b: 0,  l: 0,  r: 0,  why: 'wall display / desktop — NOTHING may move (regression net)' },
  { name: 'android-nav', t: 0, b: 48, l: 0,  r: 0,  why: 'Android gesture-nav bar, portrait — the reported class of bug' },
  { name: 'ios-land',   t: 0,  b: 21, l: 44, r: 44, why: 'iOS landscape — home indicator + both notch rails' }
];

const failures = [];

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const html = fs.readFileSync(WALL, 'utf8');
  const browser = await chromium.launch();

  for (const c of CASES) {
    console.log(`\n${'='.repeat(70)}\n  INSETS  top ${c.t}  bottom ${c.b}  left ${c.l}  right ${c.r}`
              + `\n  ${c.why}\n${'='.repeat(70)}`);

    for (const [name, w, h, touch] of SIZES) {
      const ctx = await browser.newContext({
        viewport: { width: w, height: h }, hasTouch: touch,
        isMobile: touch && w < 500, timezoneId: 'America/Chicago'
      });
      const page = await ctx.newPage();
      const errs = [];
      page.on('pageerror', e => errs.push('pageerror: ' + e.message));
      page.on('console', m => {
        if (m.type() !== 'error') return;
        const t = m.text();
        // The prayer fetch is aborted by THIS harness on purpose — our noise.
        if (/Failed to load resource|ERR_FAILED|ERR_ABORTED/.test(t)) return;
        errs.push(t);
      });

      await page.route('http://hqlab.test/**', route => {
        if (route.request().url().includes('/api/board')) {
          return route.fulfill({ contentType: 'application/json',
            body: JSON.stringify(Object.assign({ ok: true }, MOCK)) });
        }
        return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
      });
      await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());

      if (c.t || c.b || c.l || c.r) {
        await page.addInitScript(ins => {
          document.addEventListener('DOMContentLoaded', () => {
            const s = document.createElement('style');
            s.textContent = `:root{--safe-t:${ins.t}px;--safe-b:${ins.b}px;`
                          + `--safe-l:${ins.l}px;--safe-r:${ins.r}px;}`;
            document.head.appendChild(s);
          });
        }, c);
      }

      await page.goto('http://hqlab.test/wall', { waitUntil: 'load' });
      await page.waitForTimeout(2200);

      const m = await page.evaluate(ins => {
        // Force the two transient overlays visible so they can be MEASURED.
        // A toast you cannot read is a toast that failed.
        document.getElementById('connLost').classList.add('show');
        document.getElementById('restveil').classList.add('show');

        const H = window.innerHeight, W = window.innerWidth;
        const safeBottom = H - ins.b, safeTop = ins.t;
        const safeLeft = ins.l, safeRight = W - ins.r;
        const box = sel => {
          const e = document.querySelector(sel);
          if (!e) return null;
          const b = e.getBoundingClientRect();
          return { top: Math.round(b.top), bottom: Math.round(b.bottom),
                   left: Math.round(b.left), right: Math.round(b.right) };
        };

        // The LAST row actually shown in a channel: rows past the ul's own
        // bottom are clipped by overflow:hidden (the wall's deliberate cap),
        // so they are not "hidden by the bar" — they were never on screen.
        const lastVisibleRow = ulSel => {
          const ul = document.querySelector(ulSel);
          if (!ul) return null;
          const ulBottom = ul.getBoundingClientRect().bottom;
          let last = null;
          ul.querySelectorAll('.wrow, .wband').forEach(el => {
            const b = el.getBoundingClientRect();
            if (b.bottom <= ulBottom + 1 && b.height > 0) {
              if (!last || b.bottom > last.bottom)
                last = { bottom: Math.round(b.bottom), right: Math.round(b.right),
                         left: Math.round(b.left),
                         text: (el.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 26) };
            }
          });
          return last;
        };

        const cards = [...document.querySelectorAll('.rail .m')].map(e => {
          const b = e.getBoundingClientRect();
          return { bottom: Math.round(b.bottom), left: Math.round(b.left) };
        }).filter(c => c.bottom > 0);

        return {
          H, W, safeBottom, safeTop, safeLeft, safeRight,
          docScrollH: document.documentElement.scrollHeight,
          docScrollW: document.documentElement.scrollWidth,
          wall: box('.wall'),
          hdrGround: box('.hdr'),
          hdrChip: box('.hdr-hq'),            // first content in the header lane
          hdrClock: box('.hdr-clock'),
          rail: box('.rail'),
          lastCard: cards.length ? cards[cards.length - 1] : null,
          firstCardLeft: cards.length ? cards[0].left : null,
          lastE: lastVisibleRow('#listE'),
          lastD: lastVisibleRow('#listD'),
          conn: box('.connlost'),
          rvClock: box('.rv-clock'),
          rvSub: box('.rv-sub'),
          rows: document.querySelectorAll('.wrow').length
        };
      }, c);

      // ── assertions ──────────────────────────────────────────────────────
      const p = [];
      const under = (what, bottom) => {
        if (bottom > m.safeBottom + 1) p.push(`${what} ${bottom - m.safeBottom}px UNDER the bottom bar`);
      };
      const pastRight = (what, right) => {
        if (right > m.safeRight + 1) p.push(`${what} ${right - m.safeRight}px past the right inset`);
      };

      if (m.docScrollH > m.H + 1) p.push(`page scrolls vertically (${m.docScrollH} > ${m.H})`);
      if (m.docScrollW > m.W + 1) p.push(`page scrolls horizontally (${m.docScrollW} > ${m.W})`);

      // Header ground may paint under the status bar; its CONTENT may not.
      if (m.hdrChip && m.hdrChip.top < m.safeTop - 1)
        p.push(`header chip ${m.safeTop - m.hdrChip.top}px under the top bar`);
      if (m.hdrChip && m.hdrChip.left < m.safeLeft - 1)
        p.push(`header chip ${m.safeLeft - m.hdrChip.left}px past the left inset`);
      if (m.hdrClock) pastRight('header clock', m.hdrClock.right);

      if (m.lastCard) under('last rail card', m.lastCard.bottom);
      if (m.firstCardLeft !== null && m.firstCardLeft < m.safeLeft - 1)
        p.push(`rail card ${m.safeLeft - m.firstCardLeft}px past the left inset`);

      if (m.lastE) { under(`last eBay row ("${m.lastE.text}")`, m.lastE.bottom); }
      if (m.lastD) { under(`last DIRECT row ("${m.lastD.text}")`, m.lastD.bottom);
                     pastRight('last DIRECT row', m.lastD.right); }

      if (m.conn) under('conn-lost toast', m.conn.bottom);
      if (m.rvSub) under('rest-veil sub-line', m.rvSub.bottom);
      if (m.rvClock && m.rvClock.top < m.safeTop - 1)
        p.push(`rest-veil clock ${m.safeTop - m.rvClock.top}px under the top bar`);

      if (errs.length) p.push('JS: ' + errs.join(' | '));

      const ok = p.length === 0;
      if (!ok) failures.push(`${name} @ ${c.name} — ${p.join('; ')}`);

      console.log(`\n  ${ok ? '✓' : '✗'} ${name} (${w}×${h})   rows=${m.rows}`);
      console.log(`      screen ${m.W}×${m.H}   safe box: top ${m.safeTop} · bottom ${m.safeBottom}`
                + ` · left ${m.safeLeft} · right ${m.safeRight}`);
      console.log(`      hdr ground ${JSON.stringify(m.hdrGround)}`);
      console.log(`      hdr chip   ${JSON.stringify(m.hdrChip)}`);
      console.log(`      last eBay row bottom ${m.lastE ? m.lastE.bottom : '—'}`
                + `   headroom ${m.lastE ? m.safeBottom - m.lastE.bottom : '—'}px`);
      console.log(`      last rail card bottom ${m.lastCard ? m.lastCard.bottom : '—'}`
                + `   headroom ${m.lastCard ? m.safeBottom - m.lastCard.bottom : '—'}px`);
      console.log(`      conn toast bottom ${m.conn ? m.conn.bottom : '—'}`
                + `   headroom ${m.conn ? m.safeBottom - m.conn.bottom : '—'}px`);
      if (!ok) p.forEach(x => console.log(`      ✗ ${x}`));

      await page.screenshot({ path: path.join(OUT, `${TAG}wallsafe-${c.name}-${name}.png`) });
      await ctx.close();
    }
  }

  await browser.close();
  console.log('\n' + '='.repeat(70));
  if (failures.length) {
    console.log('✗ FAILURES:\n' + failures.map(f => '   ' + f).join('\n'));
    process.exit(1);
  }
  console.log('✓ ALL VIEWPORTS × ALL INSET SHAPES: every painted element inside the safe box,');
  console.log('  page does not scroll in either axis.');
})().catch(e => { console.error('CRASH', e); process.exit(1); });
