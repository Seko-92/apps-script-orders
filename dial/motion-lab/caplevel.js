const { chromium } = require('/home/yassin/Desktop/Projects/Projects/Excel Code/design-lab/node_modules/playwright');
const fs = require('fs'), path = require('path');
const FPS = 24, STEP = Math.round(1000 / FPS), A = 'HIGH QUALITY MOTOR SERVICE', H_ = '★ HOUSTON, TEXAS', E_ = 'ENGINE PARTS & OVERHAUL KITS';
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 820, height: 160 }, deviceScaleFactor: 1 });
  const errs = []; p.on('pageerror', e => errs.push(e.message));
  await p.goto('file://' + __dirname + '/level.html');
  await p.waitForSelector('body[data-ready="1"]', { timeout: 20000 });
  const man = {};
  const job = async (name, parts) => {
    const dir = path.join(__dirname, 'level', name); fs.mkdirSync(dir, { recursive: true });
    const out = []; let i = 0;
    for (const [fn, ms] of parts) { const url = await p.evaluate(fn.src, fn.args); const f = path.join(dir, String(i++).padStart(3, '0') + '.png'); fs.writeFileSync(f, Buffer.from(url.split(',')[1], 'base64')); out.push({ f, ms }); }
    man[name] = out; console.log(name, out.length, 'frames');
  };
  const still = (t) => [{ src: a => still(a), args: t }, 'REST'];
  const holdB = (t) => [{ src: a => still(a), args: t }, 2400];
  const seq = (name, a, bb, D, blur) => {
    const n = Math.round(D * FPS), res = [];
    const f = (x, y, k) => blur
      ? { src: ([m, x, y, u, t, sh]) => { window.NOGHOST = true; return moveBlur(m, x, y, u, t, sh, 8); }, args: [name, x, y, k / n, k / FPS, 0.5 / n] }
      : { src: ([m, x, y, u, t]) => move(m, x, y, u, t), args: [name, x, y, k / n, k / FPS] };
    return { n, fwd: k => [f(a, bb, k), STEP], back: k => [f(bb, a, k), STEP] };
  };
  const round = (name, jobName, bb, D, blur) => { const s = seq(name, A, bb, D, blur); const parts = [still(A)];
    for (let k = 1; k <= s.n; k++) parts.push(s.fwd(k)); parts.push(holdB(bb)); for (let k = 1; k <= s.n; k++) parts.push(s.back(k)); return job(jobName, parts); };
  // 1 · the logo is a kit (A → A)
  { const n = Math.round(9.5 * FPS), parts = [still(A)]; for (let k = 1; k <= n; k++) parts.push([{ src: ([u, t]) => move('kit', 'HIGH QUALITY MOTOR SERVICE', 'HIGH QUALITY MOTOR SERVICE', u, t), args: [k / n, k / FPS] }, STEP]); await job('kit', parts); }
  // 2 · film blur
  await round('gear', 'gear-blur', H_, 3.2, true);
  await round('conveyor', 'conveyor-blur', E_, 2.6, true);
  // 3 · the stage joins in
  await round('scanplus', 'scanplus', H_, 3.6, false);
  await round('loanplus', 'loanplus', E_, 5.8, false);
  { const n = Math.round(1.4 * FPS), parts = [still(A)]; for (let k = 1; k <= n; k++) parts.push([{ src: ([u]) => move('sheen', 'HIGH QUALITY MOTOR SERVICE', 'HIGH QUALITY MOTOR SERVICE', u, 0), args: [k / n] }, STEP]); await job('sheen', parts); }
  fs.writeFileSync(path.join(__dirname, 'level', 'manifest.json'), JSON.stringify(man));
  console.log('errors:', errs.length ? errs : 'none'); await b.close();
})();
