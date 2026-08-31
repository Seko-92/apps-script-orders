const { chromium } = require('playwright');
const path=require('path');
const SB='/tmp/claude-1000/-home-yassin-Desktop-Projects-Projects-Excel-Code/4b8083d7-c1fd-4c06-afd2-fd819e2a5e0e/scratchpad';
(async()=>{
  const b=await chromium.launch();
  const p=await b.newPage({viewport:{width:1240,height:1100},deviceScaleFactor:2});
  const errs=[]; p.on('pageerror',e=>errs.push(String(e)));
  p.on('console',m=>{if(m.type()==='error')errs.push(m.text());});
  await p.goto('file://'+path.join(SB,'marks.html'));
  await p.waitForTimeout(350);
  let fail=0; const ok=(n,c,x)=>{c?console.log('  ok  '+n):(fail++,console.log('  XX  '+n+(x!==undefined?'  -> '+JSON.stringify(x):'')));};

  const r=await p.evaluate(()=>{
    const uses=[...document.querySelectorAll('use')];
    const ids=new Set([...document.querySelectorAll('symbol')].map(s=>s.id));
    const bad=uses.map(u=>u.getAttribute('href')).filter(h=>!ids.has(h.slice(1)));
    // every rendered mark must actually paint pixels
    const empty=[...document.querySelectorAll('.chip svg,.mk svg,.c .ci svg,.brow svg,.wallrow svg')]
      .filter(s=>{const r=s.getBoundingClientRect();return r.width<6||r.height<6;}).length;
    const body=document.body.getBoundingClientRect();
    return {symbols:ids.size, uses:uses.length, brokenRefs:bad, empty,
      hScroll:document.documentElement.scrollWidth>document.documentElement.clientWidth+1,
      sidebars:[...document.querySelectorAll('.side')].map(e=>Math.round(e.getBoundingClientRect().width)),
      emojiLeftInProposed:[...document.querySelectorAll('.side')][1].querySelectorAll('.ci.emo').length,
      marksInProposed:[...document.querySelectorAll('.side')][1].querySelectorAll('.ci svg').length};
  });
  console.log('  ',JSON.stringify(r));
  ok('18 symbols defined', r.symbols===18, r.symbols);
  ok('every <use> resolves to a symbol', r.brokenRefs.length===0, r.brokenRefs);
  ok('no mark renders empty', r.empty===0, r.empty);
  ok('sidebars are 310px', r.sidebars.every(w=>w===310), r.sidebars);
  ok('no emoji in the proposed card stack', r.emojiLeftInProposed===0, r.emojiLeftInProposed);
  ok('7 drawn marks in the proposed stack', r.marksInProposed===7, r.marksInProposed);
  ok('page does not scroll sideways', !r.hScroll, r.hScroll);
  ok('no console errors', errs.length===0, errs.slice(0,3));

  await p.screenshot({path:path.join(SB,'marks-top.png'),clip:{x:0,y:0,width:1240,height:1080}});
  await p.locator('.stage').first().screenshot({path:path.join(SB,'marks-cards.png')});
  // dark theme too
  await p.emulateMedia({colorScheme:'dark'}); await p.waitForTimeout(250);
  await p.screenshot({path:path.join(SB,'marks-dark.png'),clip:{x:0,y:0,width:1240,height:900}});
  console.log('\n'+(fail?('XX '+fail+' FAILED'):'OK all checks passed'));
  await b.close(); process.exit(fail?1:0);
})();
