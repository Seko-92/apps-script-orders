// Does a HOLD still reach the picker when the sort removes every band?
// The 2026-08-15 bug shipped because the fixture only had banded orders.
'use strict';
const fs=require('fs'),path=require('path');
const {chromium}=require('playwright');
const html=fs.readFileSync(path.join(__dirname,'..','FloorBoard.html'),'utf8');
const MOCK=JSON.parse(JSON.stringify(require('./mock-tick.js')));
// put a HOLD on a MULTI-line eBay order (banded in aisle mode, bandless in walk)
// and on a single-line one (bandless in BOTH) — both must show, in every mode.
let multi=0, solo=0;
MOCK.openOrders.forEach(r=>{
  if(r.orderId==='08-15017-44806'){ r.note='HOLD — customer confirming shipping'; multi++; }
  if(r.orderId==='24-15004-11290'){ r.note='HOLD — verify serial'; solo++; }
});
console.log('seeded HOLD on '+multi+' rows of a 5-line order and '+solo+' single-line order');
(async()=>{
  const b=await chromium.launch();
  const c=await b.newContext({viewport:{width:1280,height:800},hasTouch:true,timezoneId:'America/Chicago'});
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
  await p.waitForTimeout(2600);
  let bad=0;
  for(const m of ['aisle','age','walk']){
    await p.evaluate(x=>setPickMode(x),m);
    await p.waitForTimeout(600);
    const r=await p.evaluate(()=>{
      const rowChips=document.querySelectorAll('.pick-row .pick-hold').length;
      const bandChips=document.querySelectorAll('.pick-head .hold, .pick-head .pick-hold, .pick-head [class*="hold"]').length;
      const bandTextHold=Array.from(document.querySelectorAll('.pick-head')).filter(e=>/HOLD/i.test(e.textContent)).length;
      return {rowChips,bandChips:bandChips||bandTextHold};
    });
    const total=r.rowChips+r.bandChips;
    const ok=total>0;
    if(!ok) bad++;
    console.log((ok?'  ✓ ':'  ✗ ')+m.padEnd(6)+' HOLD visible: '+total+
                '  (row chips '+r.rowChips+', band '+r.bandChips+')');
  }
  console.log(bad? '\n❌ a HOLD went invisible in '+bad+' mode(s)' : '\n✅ the HOLD reaches the picker in every mode');
  await b.close();
  process.exit(bad?1:0);
})();
