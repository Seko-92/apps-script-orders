const { chromium } = require('/home/yassin/Desktop/Projects/Projects/Excel Code/design-lab/node_modules/playwright');
const fs = require('fs'), path = require('path');
const FPS = 24, STEP = Math.round(1000 / FPS), A = 'HIGH QUALITY MOTOR SERVICE';
const PLAN = [
  ['typewriter', 3.4, 'ENGINE PARTS & OVERHAUL KITS'],
  ['odometer',   2.2, '★ HOUSTON, TEXAS'],
  ['decode',     2.6, 'ENGINE PARTS & OVERHAUL KITS'],
  ['conveyor',   2.6, 'ENGINE PARTS & OVERHAUL KITS'],
  ['assembly',   2.8, '★ HOUSTON, TEXAS'],
  ['title',      3.0, '★ HOUSTON, TEXAS'],
];
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 820, height: 160 }, deviceScaleFactor: 1 });
  const errs = []; p.on('pageerror', e => errs.push(e.message));
  await p.goto('file://' + __dirname + '/moves.html');
  await p.waitForSelector('body[data-ready="1"]', { timeout: 20000 });
  const man = {};
  const save = (dir, i, url) => { const f = path.join(dir, String(i).padStart(3, '0') + '.png'); fs.writeFileSync(f, Buffer.from(url.split(',')[1], 'base64')); return f; };
  for (const [name, D, B] of PLAN.filter(([n]) => !process.env.ONLY || process.env.ONLY === n)) {
    const dir = path.join(__dirname, 'moves', name); fs.mkdirSync(dir, { recursive: true });
    const n = Math.round(D * FPS), out = []; let i = 0;
    out.push({ f: save(dir, i++, await p.evaluate(a => still(a), A)), ms: 'REST' });
    for (let k = 1; k <= n; k++) out.push({ f: save(dir, i++, await p.evaluate(([m, a, bb, u, t]) => move(m, a, bb, u, t), [name, A, B, k / n, k / FPS])), ms: STEP });
    out.push({ f: save(dir, i++, await p.evaluate(a => still(a), B)), ms: 2400 });
    for (let k = 1; k <= n; k++) out.push({ f: save(dir, i++, await p.evaluate(([m, a, bb, u, t]) => move(m, a, bb, u, t), [name, B, A, k / n, k / FPS])), ms: STEP });
    man[name] = { D, B, frames: out };
    console.log(name, out.length, 'frames');
  }
  const mf = path.join(__dirname, 'moves', 'manifest.json'); const old = fs.existsSync(mf) ? JSON.parse(fs.readFileSync(mf)) : {}; fs.writeFileSync(mf, JSON.stringify(Object.assign(old, man)));
  console.log('errors:', errs.length ? errs : 'none'); await b.close();
})();
