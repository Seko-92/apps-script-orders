const { chromium } = require('playwright');
const path = require('path');
const SB = '/tmp/claude-1000/-home-yassin-Desktop-Projects-Projects-Excel-Code/4b8083d7-c1fd-4c06-afd2-fd819e2a5e0e/scratchpad';

(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport:{width:1300,height:1100}, deviceScaleFactor:2 });
  const errs=[]; p.on('pageerror',e=>errs.push(String(e)));
  p.on('console',m=>{if(m.type()==='error')errs.push(m.text());});
  await p.goto('file://'+path.join(SB,'idle-panel.html'));
  await p.waitForTimeout(400);

  let fail=0;
  const ok=(n,c,x)=>{ if(c) console.log('  ok  '+n);
    else {fail++; console.log('  XX  '+n+(x!==undefined?'  -> '+JSON.stringify(x):''));} };

  for (const mode of ['day','night']) {
    await p.click(`[data-go="${mode}"]`);
    await p.waitForTimeout(2600);
    console.log(`\n-- ${mode.toUpperCase()} ------------------`);
    const r = await p.evaluate((m)=>{
      const panel=document.getElementById('live'), pr=panel.getBoundingClientRect();
      const rest=panel.querySelector('.rest.'+m);
      const kids=[...rest.children].filter(e=>e.offsetHeight>0 && getComputedStyle(e).position!=='absolute');
      const boxes=kids.map(e=>{const q=e.getBoundingClientRect();
        return {cls:e.className,top:+(q.top-pr.top).toFixed(1),bot:+(q.bottom-pr.top).toFixed(1)};});
      const lap=[]; for(let i=1;i<boxes.length;i++)
        if(boxes[i].top<boxes[i-1].bot-0.5) lap.push([boxes[i-1].cls,boxes[i].cls]);
      return {h:pr.height,w:pr.width,boxes,
        over:boxes.filter(x=>x.bot>pr.height+0.5), lap,
        opacity:getComputedStyle(rest).opacity};
    }, mode);
    r.boxes.forEach(x=>console.log('    '+String(x.top).padStart(6)+' -> '+String(x.bot).padStart(6)+'  .'+x.cls));
    ok('rest layer visible', r.opacity==='1', r.opacity);
    ok('nothing overflows', r.over.length===0, r.over);
    ok('no sibling overlaps', r.lap.length===0, r.lap);
    ok('panel is 310px', Math.round(r.w)===310, r.w);
  }

  await p.click('[data-go="day"]'); await p.waitForTimeout(2600);
  const t = await p.evaluate(()=>{
    const cs=n=>getComputedStyle(n); const d=document.querySelector('#live .rest.day');
    const n=document.querySelector('#live .rest.night');
    return {pips:d.querySelectorAll('.pip').length,
      pipsVis:[...d.querySelectorAll('.pip')].filter(e=>+cs(e).opacity>0.9).length,
      odos:[...d.querySelectorAll('[data-odo]')].map(e=>e.textContent.trim()),
      closes:d.querySelector('[data-closes]').textContent.trim(),
      rows:d.querySelectorAll('.nrow').length,
      nightRows:n.querySelectorAll('.nrow').length,
      dayBg:cs(d).backgroundImage.slice(0,30), nightBg:cs(n).backgroundImage.slice(0,30),
      dayHasWidgets:d.querySelectorAll('.lamp,.ob,.ft,.tach,.xrow').length,
      dayFlank:[...d.querySelectorAll('.nbox .k')].map(e=>e.textContent.trim()).join('|'),
      nightFlank:[...n.querySelectorAll('.nbox .k')].map(e=>e.textContent.trim()).join('|'),
      noneRows:[...document.querySelectorAll('#live .nrow b')].filter(e=>/^(none|0)$/i.test(e.textContent.trim())).length,
      restZ:+cs(d).zIndex, hdrZ:+cs(document.querySelector('#live .hdr')).zIndex};
  });
  console.log('\n-- FACE ------------------\n  ',JSON.stringify(t));
  ok('20 shipment pips on the face', t.pips===20 && t.pipsVis===20, {n:t.pips,vis:t.pipsVis});
  ok('flank rolled to 20 / 24', t.odos[0]==='20' && t.odos[1]==='24', t.odos);
  ok('closes-in is live', /^[0-9]+h [0-9][0-9]m$/.test(t.closes), t.closes);
  ok('day has 4 rows', t.rows===4, t.rows);
  ok('night drops the healthy row (3 not 4)', t.nightRows===3, t.nightRows);
  ok('flank is identical in both states', t.dayFlank===t.nightFlank, {d:t.dayFlank,n:t.nightFlank});
  ok('no row anywhere says "none"', t.noneRows===0, t.noneRows);
  ok('ZERO widgets left in the day', t.dayHasWidgets===0, t.dayHasWidgets);
  ok('day and night are different LIGHTS', t.dayBg!==t.nightBg, {d:t.dayBg,n:t.nightBg});
  ok('rest above the console header', t.restZ>t.hdrZ, {rest:t.restZ,hdr:t.hdrZ});
  ok('no console/page errors', errs.length===0, errs.slice(0,3));

  await p.click('[data-go="day"]');  await p.waitForTimeout(2600);
  await p.locator('#live').screenshot({path:path.join(SB,'idle-day.png')});
  await p.click('[data-go="night"]'); await p.waitForTimeout(1800);
  await p.locator('#live').screenshot({path:path.join(SB,'idle-night.png')});

  console.log('\n'+(fail?('XX '+fail+' FAILED'):'OK all checks passed'));
  await b.close(); process.exit(fail?1:0);
})();
