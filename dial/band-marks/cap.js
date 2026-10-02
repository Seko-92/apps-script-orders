// cap.js — render both band marks frame by frame → frames/<which>/NNN.png + manifest.json
//   node cap.js        (24 fps action; rest frames held by enc.py)
const { chromium } = require('../../design-lab/node_modules/playwright');
const fs = require('fs'), path = require('path');
const FPS = 24, OUT = path.join(__dirname, 'frames');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage();
  // served over http: a file:// image taints the canvas and blocks toDataURL
  await p.route('http://bm.test/**', r => {
    const f = path.join(__dirname, new URL(r.request().url()).pathname);
    r.fulfill({ body: fs.readFileSync(f), contentType: f.endsWith('.png') ? 'image/png' : 'text/html; charset=utf-8' });
  });
  await p.goto('http://bm.test/marks.html');
  const man = {};
  for (const which of ['direct', 'parcel']) {
    const dir = path.join(OUT, which); fs.mkdirSync(dir, { recursive: true });
    const action = await p.evaluate(w => window.ACTION_MS[w], which);
    const n = Math.ceil(action / 1000 * FPS);
    const list = [];
    for (let i = 0; i <= n; i++) {
      const t = Math.min(action, i * 1000 / FPS);
      const url = await p.evaluate(([w, t]) => window.render(w, t), [which, t]);
      const f = path.join(dir, String(i).padStart(3, '0') + '.png');
      fs.writeFileSync(f, Buffer.from(url.split(',')[1], 'base64'));
      list.push(f);
    }
    man[which] = list; console.log(which, list.length, 'frames');
  }
  fs.writeFileSync(path.join(__dirname, 'manifest.json'), JSON.stringify(man, null, 1));
  await b.close();
})();
