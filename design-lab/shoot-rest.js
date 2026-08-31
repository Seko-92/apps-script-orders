// ============================================================================
// THE NIGHT DIAL — render + assert the /wall rest scene.
//
// The standard mock is a busy 9:48 AM with 19 orders to grab, so the rest scene
// can NEVER appear under it. This drives a fake clock to real Houston moments
// and feeds a quiet tick, then asserts the dial's MECHANICS — not just that it
// drew something:
//   · the night arc's fill matches the hour (empty 5PM → full 9AM)
//   · the glowing head sits on the arc's front, at the right angle
//   · --dawn is 0 deep night and warm near open
//   · "opens in" skips the weekend (Friday night must not read "17h")
//   · the seconds jewel is phase-aligned to the wall clock, not merely spinning
//
// Usage: node shoot-rest.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const WALL = process.env.WALL_FILE || path.join(__dirname, '..', 'wall.html');
const MOCK = require('./mock-tick.js');
const OUT  = path.join(__dirname, 'renders');

// A quiet tick: shop closed, nothing left to grab. `straggler` leaves exactly
// one open order so the footnote case can be seen.
function quietTick(straggler) {
  const t = JSON.parse(JSON.stringify(MOCK));
  t.cockpit.ebayGrab = 0;
  t.cockpit.directGrab = 0;
  t.openOrders = straggler ? [MOCK.openOrders[0]] : [];
  t.openOrdersBy = {};
  t.openOrdersTotal = straggler ? 1 : 0;
  return t;
}

// Houston is CDT (UTC−5) in August.
const MOMENTS = [
  { name: 'deep-night', iso: '2026-08-13T22:15:00-05:00', straggler: false,
    why: 'Thu 10:15 PM — 5h15m into a 16h night ≈ 33% arc, dawn 0',
    expectProg: 5.25 / 16, expectDawn: 0 },
  { name: 'straggler',  iso: '2026-08-13T23:40:00-05:00', straggler: true,
    why: 'Thu 11:40 PM — one order still on the floor, footnote shows',
    expectProg: 6.667 / 16, expectDawn: 0 },
  { name: 'near-dawn',  iso: '2026-08-14T07:30:00-05:00', straggler: false,
    why: 'Fri 7:30 AM — 14h30m in ≈ 91% arc, dawn warm, line flips to "about to wake"',
    expectProg: 14.5 / 16, expectDawn: 0.55 },
  { name: 'friday-night', iso: '2026-08-14T21:00:00-05:00', straggler: false,
    why: 'FRI 9 PM — "opens in" must SKIP the weekend (≈ 2d 12h, never 12h)',
    expectProg: 4 / 16, expectDawn: 0, expectOpensDays: true }
];

const SIZES = [['wall', 1920, 1080], ['tablet-landscape', 1280, 800]];
const failures = [];

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const html = fs.readFileSync(WALL, 'utf8');
  const browser = await chromium.launch();

  for (const mo of MOMENTS) {
    console.log(`\n${'='.repeat(70)}\n  ${mo.name.toUpperCase()}  —  ${mo.why}\n${'='.repeat(70)}`);

    for (const [sizeName, w, h] of SIZES) {
      const ctx = await browser.newContext({
        viewport: { width: w, height: h }, timezoneId: 'America/Chicago'
      });
      const page = await ctx.newPage();
      const errs = [];
      page.on('pageerror', e => errs.push('pageerror: ' + e.message));
      page.on('console', m => {
        if (m.type() !== 'error') return;
        const t = m.text();
        if (/Failed to load resource|ERR_FAILED|ERR_ABORTED/.test(t)) return;
        errs.push(t);
      });

      // Fake clock BEFORE navigation so the page's very first paint sees it.
      await page.clock.install({ time: new Date(mo.iso) });

      const body = quietTick(mo.straggler);
      await page.route('http://hqlab.test/**', route => {
        if (route.request().url().includes('/api/board')) {
          return route.fulfill({ contentType: 'application/json',
            body: JSON.stringify(Object.assign({ ok: true }, body)) });
        }
        return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
      });
      await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());

      await page.goto('http://hqlab.test/wall', { waitUntil: 'load' });
      // Let the poll land and the 1s clock interval fire a few times.
      await page.clock.runFor(4000);
      await page.waitForTimeout(1200);

      const m = await page.evaluate(() => {
        const g = id => document.getElementById(id);
        const txt = id => (g(id) ? g(id).textContent.trim() : null);
        const nf = g('nightFill'), nh = g('nightHead');
        const CIRC = 295.31;
        const dash = nf ? nf.getAttribute('stroke-dasharray') : '';
        const filled = parseFloat((dash || '0').split(' ')[0]) || 0;
        const sweep = g('dialSweep');
        return {
          showing: g('restveil').classList.contains('show'),
          greet: txt('restGreet'), clock: txt('rvClock'), line: txt('restLine'),
          shipped: txt('restShipped'), received: txt('restReceived'),
          opens: txt('restOpens'), prayer: txt('restPrayer'),
          straggler: txt('restStraggler'),
          dawnLine: g('restLine') ? g('restLine').classList.contains('dawn') : null,
          prog: filled / CIRC,
          headCx: nh ? parseFloat(nh.getAttribute('cx')) : null,
          headCy: nh ? parseFloat(nh.getAttribute('cy')) : null,
          dawn: parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--dawn')) || 0,
          sweepDelay: sweep ? sweep.style.animationDelay : null,
          embers: document.querySelectorAll('.ember').length,
          // the dial must not overflow the screen at any size
          dialBox: (() => { const b = document.querySelector('.dial').getBoundingClientRect();
                            return { w: Math.round(b.width), h: Math.round(b.height),
                                     top: Math.round(b.top), bottom: Math.round(b.bottom),
                                     left: Math.round(b.left), right: Math.round(b.right) }; })(),
          screenW: window.innerWidth, screenH: window.innerHeight,
          scrollW: document.documentElement.scrollWidth,
          scrollH: document.documentElement.scrollHeight
        };
      });

      // ── assertions ──────────────────────────────────────────────────────
      const p = [];
      if (!m.showing) p.push('rest scene did NOT show on a quiet off-hours tick');
      if (m.embers !== 7) p.push(`expected 7 embers, got ${m.embers}`);
      if (Math.abs(m.prog - mo.expectProg) > 0.02)
        p.push(`night arc ${(m.prog * 100).toFixed(1)}% but the hour implies ${(mo.expectProg * 100).toFixed(1)}%`);
      // head must sit ON the r=47 circle, at the arc's leading angle
      if (m.headCx !== null) {
        const ang = (m.prog * 360 - 90) * Math.PI / 180;
        const wantX = 50 + 47 * Math.cos(ang), wantY = 50 + 47 * Math.sin(ang);
        if (Math.abs(m.headCx - wantX) > 0.6 || Math.abs(m.headCy - wantY) > 0.6)
          p.push(`arc head at (${m.headCx},${m.headCy}) but the fill ends at (${wantX.toFixed(2)},${wantY.toFixed(2)})`);
      }
      if (mo.expectDawn === 0 && m.dawn > 0.01) p.push(`--dawn ${m.dawn} should be 0 in deep night`);
      if (mo.expectDawn > 0 && m.dawn < mo.expectDawn) p.push(`--dawn ${m.dawn} too cold near open (want ≥ ${mo.expectDawn})`);
      if (mo.expectDawn > 0.5 && !m.dawnLine) p.push('near dawn but the line never flipped to "about to wake"');
      if (mo.expectOpensDays && !/d\s/.test(m.opens || ''))
        p.push(`Friday night "opens in" reads "${m.opens}" — the weekend was NOT skipped`);
      if (mo.straggler && !(m.straggler || '').includes('still on the floor'))
        p.push('one open order but no straggler footnote');
      if (!mo.straggler && m.straggler) p.push(`straggler footnote shown with 0 open orders: "${m.straggler}"`);
      if (!m.sweepDelay || !m.sweepDelay.startsWith('-'))
        p.push(`seconds jewel not phase-aligned (animation-delay "${m.sweepDelay}")`);
      if (m.dialBox.right > m.screenW + 1 || m.dialBox.left < -1)
        p.push(`dial overflows horizontally (${m.dialBox.left}..${m.dialBox.right} in ${m.screenW})`);
      if (m.dialBox.bottom > m.screenH + 1 || m.dialBox.top < -1)
        p.push(`dial overflows vertically (${m.dialBox.top}..${m.dialBox.bottom} in ${m.screenH})`);
      if (m.scrollW > m.screenW + 1 || m.scrollH > m.screenH + 1)
        p.push(`page scrolls (${m.scrollW}×${m.scrollH} vs ${m.screenW}×${m.screenH})`);
      if (errs.length) p.push('JS: ' + errs.join(' | '));

      const ok = p.length === 0;
      if (!ok) failures.push(`${mo.name} @ ${sizeName} — ${p.join('; ')}`);

      console.log(`\n  ${ok ? '✓' : '✗'} ${sizeName} (${w}×${h})`);
      console.log(`      ${m.greet} · ${m.clock} · "${m.line}"`);
      console.log(`      shipped ${m.shipped} · received ${m.received}  |  opens in ${m.opens} · prayer ${m.prayer}`);
      console.log(`      night arc ${(m.prog * 100).toFixed(1)}%   head (${m.headCx}, ${m.headCy})   --dawn ${m.dawn}`);
      console.log(`      dial ${m.dialBox.w}×${m.dialBox.h}   jewel delay ${m.sweepDelay}   embers ${m.embers}`);
      if (m.straggler) console.log(`      footnote: ${m.straggler}`);
      if (!ok) p.forEach(x => console.log(`      ✗ ${x}`));

      await page.screenshot({ path: path.join(OUT, `rest-${mo.name}-${sizeName}.png`) });
      await ctx.close();
    }
  }

  await browser.close();
  console.log('\n' + '='.repeat(70));
  if (failures.length) {
    console.log('✗ FAILURES:\n' + failures.map(f => '   ' + f).join('\n'));
    process.exit(1);
  }
  console.log('✓ THE DIAL: arc tracks the hour, head rides the fill, dawn warms toward open,');
  console.log('  the weekend is skipped, the jewel is phase-aligned, nothing overflows.');
})().catch(e => { console.error('CRASH', e); process.exit(1); });
