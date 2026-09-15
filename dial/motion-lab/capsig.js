const { chromium } = require('/home/yassin/Desktop/Projects/Projects/Excel Code/design-lab/node_modules/playwright');
const fs = require('fs'), path = require('path');
const FPS = 24, STEP = Math.round(1000 / FPS), A = 'HIGH QUALITY MOTOR SERVICE';
const PLAN = [
  ['firing', 3.6, 'ENGINE PARTS & OVERHAUL KITS'],
  ['gear',   3.2, '★ HOUSTON, TEXAS'],
  ['loan',   5.2, 'ENGINE PARTS & OVERHAUL KITS'],
  ['scan',   3.6, '★ HOUSTON, TEXAS'],
  ['torque', 4.2, 'ENGINE PARTS & OVERHAUL KITS'],
];
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 820, height: 160 }, deviceScaleFactor: 1 });
  const errs = []; p.on('pageerror', e => errs.push(e.message));
  await p.goto('file://' + __dirname + '/sig.html');
  await p.waitForSelector('body[data-ready="1"]', { timeout: 20000 });
  const man = {};
  const save = (dir, i, url) => { const f = path.join(dir, String(i).padStart(3, '0') + '.png'); fs.writeFileSync(f, Buffer.from(url.split(',')[1], 'base64')); return f; };
  for (const [name, D, B] of PLAN.filter(([n]) => !process.env.ONLY || process.env.ONLY === n)) {
    const dir = path.join(__dirname, 'sig', name); fs.mkdirSync(dir, { recursive: true });
    const n = Math.round(D * FPS), out = []; let i = 0;
    out.push({ f: save(dir, i++, await p.evaluate(a => still(a), A)), ms: 'REST' });
    for (let k = 1; k <= n; k++) out.push({ f: save(dir, i++, await p.evaluate(([m, a, bb, u, t]) => move(m, a, bb, u, t), [name, A, B, k / n, k / FPS])), ms: STEP });
    out.push({ f: save(dir, i++, await p.evaluate(a => still(a), B)), ms: 2400 });
    for (let k = 1; k <= n; k++) out.push({ f: save(dir, i++, await p.evaluate(([m, a, bb, u, t]) => move(m, a, bb, u, t), [name, B, A, k / n, k / FPS])), ms: STEP });
    man[name] = { D, B, frames: out };
    console.log(name, out.length, 'frames');
  }
  const mf = path.join(__dirname, 'sig', 'manifest.json'); const old = fs.existsSync(mf) ? JSON.parse(fs.readFileSync(mf)) : {}; fs.writeFileSync(mf, JSON.stringify(Object.assign(old, man)));
  console.log('errors:', errs.length ? errs : 'none'); await b.close();
})();
