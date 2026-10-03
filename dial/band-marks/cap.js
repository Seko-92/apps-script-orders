// cap.js — render the ACTION window of each mark at 24 fps → frames/<which>/NNN.png + manifest.json
const { chromium } = require('../../design-lab/node_modules/playwright');
const fs = require('fs'), path = require('path');
const FPS = 24, OUT = path.join(__dirname, 'frames');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage();
  await p.route('http://bm.test/**', r => {   // http, not file://: a file image taints the canvas
    const f = path.join(__dirname, new URL(r.request().url()).pathname);
    r.fulfill({ body: fs.readFileSync(f), contentType: f.endsWith('.png') ? 'image/png' : 'text/html; charset=utf-8' });
  });
  await p.goto('http://bm.test/marks.html');
  const cyc = await p.evaluate(() => window.CYC), man = { cyc };
  for (const which of ['direct', 'amazon']) {
    const [t0, t1] = await p.evaluate(w => window.ACTION[w], which);
    const dir = path.join(OUT, which); fs.mkdirSync(dir, { recursive: true });
    const n = Math.ceil((t1 - t0) / 1000 * FPS), list = [];
    for (let i = 0; i <= n; i++) {
      const t = Math.min(t1, t0 + i * 1000 / FPS);
      const url = await p.evaluate(([w, t]) => window.render(w, t), [which, t]);
      const f = path.join(dir, String(i).padStart(3, '0') + '.png');
      fs.writeFileSync(f, Buffer.from(url.split(',')[1], 'base64')); list.push(f);
    }
    man[which] = { t0, t1, frames: list }; console.log(which, list.length, 'frames');
  }
  fs.writeFileSync(path.join(__dirname, 'manifest.json'), JSON.stringify(man, null, 1));
  await b.close();
})();
