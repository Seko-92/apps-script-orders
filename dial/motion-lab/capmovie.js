/* capmovie.js — the whole row-1 movie, frame by frame, into movie/NNNN.png + manifest.json */
const { chromium } = require('/home/yassin/Desktop/Projects/Projects/Excel Code/design-lab/node_modules/playwright');
const fs = require('fs'), path = require('path');
const FPS = 24, STEP = Math.round(1000 / FPS), HOLD = 2400;
const A = 'HIGH QUALITY MOTOR SERVICE', H = '★ HOUSTON, TEXAS', E = 'ENGINE PARTS & OVERHAUL KITS';
// [scene, seconds, second line, film blur, eBay light pass in the pause before it]
const CUT = [
  ['scanplus',   3.6, H, true,  false],
  ['typewriter', 3.4, E, false, false],
  ['firing',     3.6, E, true,  true ],
  ['odometer',   2.2, H, true,  false],
  ['loanplus',   5.8, E, false, false],
  ['decode',     2.6, E, false, true ],
  ['gear',       3.2, H, true,  false],
  ['assembly',   2.8, H, true,  false],
  ['conveyor',   2.6, E, true,  true ],
  ['torque',     4.2, E, true,  false],
  ['title',      3.0, H, true,  false],
  ['kit',        9.5, A, true,  true ],   // the finale: A → A, no hold, no return
];
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 820, height: 160 }, deviceScaleFactor: 1 });
  const errs = []; p.on('pageerror', e => errs.push(e.message));
  await p.goto('file://' + __dirname + '/movie.html');
  await p.waitForSelector('body[data-ready="1"]', { timeout: 20000 });
  const dir = path.join(__dirname, 'movie'); fs.rmSync(dir, { recursive: true, force: true }); fs.mkdirSync(dir);
  const out = []; let i = 0;
  const put = async (url, ms, tag) => { const f = path.join(dir, String(i++).padStart(4, '0') + '.png'); fs.writeFileSync(f, Buffer.from(url.split(',')[1], 'base64')); out.push({ f, ms, tag }); };
  const still = t => p.evaluate(a => still(a), t);
  const frame = (name, x, y, u, t, blur, n) => blur
    ? p.evaluate(([m, x, y, u, t, sh]) => { window.NOGHOST = true; return moveBlur(m, x, y, u, t, sh, 8); }, [name, x, y, u, t, 0.5 / n])
    : p.evaluate(([m, x, y, u, t]) => move(m, x, y, u, t), [name, x, y, u, t]);
  for (const [name, D, B, blur, sheen] of CUT) {
    const n = Math.round(D * FPS);
    if (sheen) {                                  // the pause, split around a light pass on the eBay logo
      await put(await still(A), 20000, 'rest-a');
      const m = Math.round(1.4 * FPS);
      for (let k = 1; k <= m; k++) await put(await p.evaluate(u => move('sheen', 'HIGH QUALITY MOTOR SERVICE', 'HIGH QUALITY MOTOR SERVICE', u, 0), k / m), STEP, 'sheen');
      await put(await still(A), 45000 - 20000 - m * STEP, 'rest-b');
    } else await put(await still(A), 45000, 'rest');
    for (let k = 1; k <= n; k++) await put(await frame(name, A, B, k / n, k / FPS, blur, n), STEP, name);
    if (name === 'kit') continue;
    await put(await still(B), HOLD, 'hold');
    for (let k = 1; k <= n; k++) await put(await frame(name, B, A, k / n, k / FPS, blur, n), STEP, name + '-back');
    console.log(name, 'done ·', out.length, 'frames so far');
  }
  fs.writeFileSync(path.join(dir, 'manifest.json'), JSON.stringify(out));
  console.log('frames', out.length, '· movie length', (out.reduce((s, x) => s + x.ms, 0) / 60000).toFixed(2), 'min · errors:', errs.length ? errs : 'none');
  await b.close();
})();
