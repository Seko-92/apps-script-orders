// Does the third sort option push anything off the right edge?
// The 2026-08-14 radio bug was exactly this shape, and portrait only.
'use strict';
const fs=require('fs'),path=require('path');
const {chromium}=require('playwright');
const html=fs.readFileSync(process.env.BOARD_FILE||path.join(__dirname,'..','FloorBoard.html'),'utf8');
const MOCK=require('./mock-tick.js');
const VPS=[['tablet-portrait',800,1280],['tablet-landscape',1280,800],['phone',390,844],['wall',1920,1080]];
(async()=>{
  const b=await chromium.launch(); let bad=0;
  for(const [name,w,h] of VPS){
    const c=await b.newContext({viewport:{width:w,height:h},hasTouch:w<1600,timezoneId:'America/Chicago'});
    const p=await c.newPage();
    await p.route('http://hqlab.test/**',r=>{
      const u=r.request().url();
      if(u.includes('/api/board')){
        const body=JSON.parse(r.request().postData()||'{}');
        let res={ok:false};
        if(body.action==='boardTick') res=Object.assign({ok:true},MOCK);
        if(body.action==='boardStatus') res={ok:true};
        if(body.action==='boardRadio') res={ok:true,nowPlaying:''};
        return r.fulfill({contentType:'application/json',body:JSON.stringify(res)});
      }
      return r.fulfill({contentType:'text/html; charset=utf-8',body:html});
    });
    await p.route(/aladhan|open-meteo/,r=>r.abort());
    await p.goto('http://hqlab.test/',{waitUntil:'load'});
    await p.waitForTimeout(2400);
    const r=await p.evaluate(()=>{
      const vw=window.innerWidth, out={vw,items:[],overflow:document.documentElement.scrollWidth>window.innerWidth};
      document.querySelectorAll('.sortsw-opt').forEach(e=>{
        const b=e.getBoundingClientRect();
        out.items.push({t:e.textContent.trim(),l:Math.round(b.left),r:Math.round(b.right),w:Math.round(b.width),h:Math.round(b.height),vis:b.width>0});
      });
      const sw=document.getElementById('sortSw');
      if(sw){const b=sw.getBoundingClientRect(); out.sw={l:Math.round(b.left),r:Math.round(b.right)};}
      const pc=document.getElementById('pickCount');
      if(pc){const b=pc.getBoundingClientRect(); out.count={l:Math.round(b.left),r:Math.round(b.right),txt:pc.textContent.trim()};}
      return out;
    });
    const off=r.items.filter(i=>i.r>r.vw || i.l<0);
    const countOff=r.count && (r.count.r>r.vw);
    const tooSmall=r.items.filter(i=>i.vis && i.h<24 && r.vw<1600);
    const ok=off.length===0 && !countOff && !r.overflow;
    if(!ok) bad++;
    console.log((ok?'  ✓ ':'  ✗ ')+name.padEnd(17)+w+'×'+h+
      '  sort switch '+(r.sw?r.sw.l+'→'+r.sw.r:'?')+'  of '+r.vw+
      '  opts='+r.items.map(i=>i.t).join('/')+
      (r.count?('  count@'+r.count.r):'')+
      (r.overflow?'  ⚠ PAGE SCROLLS SIDEWAYS':'')+
      (off.length?'  ⚠ OFF-EDGE: '+off.map(i=>i.t).join(','):'')+
      (countOff?'  ⚠ line-count pushed off':'')+
      (tooSmall.length?'  ⚠ tap target <24px: '+tooSmall.map(i=>i.t+' '+i.h+'px').join(','):''));
    await c.close();
  }
  await b.close();
  console.log(bad?'\n❌ '+bad+' viewport(s) broken':'\n✅ the third option fits every viewport');
  process.exit(bad?1:0);
})();
