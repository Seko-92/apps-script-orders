// A TRUTHFUL picture of the hold surfaces — no fake clock, so CSS transitions
// are real and the shots are settled rather than mid-fade.
'use strict';
const fs=require('fs'),path=require('path');
const {chromium}=require('playwright');
const BOARD=process.env.BOARD_FILE||path.join(__dirname,'..','FloorBoard.html');
const html=fs.readFileSync(BOARD,'utf8');
const BASE=JSON.parse(JSON.stringify(require('./mock-tick.js')));
const OUT=path.join(__dirname,'renders');
const held=(o,x)=>Object.assign({orderId:o,channel:'EBAY',
  note:'HOLD — buyer wants expedited, change from Ground to 2-Day',
  acked:false,ackText:'',shipped:true,urgent:true,lines:3},x||{});
(async()=>{
  const b=await chromium.launch();
  const c=await b.newContext({viewport:{width:1280,height:800},hasTouch:true,timezoneId:'America/Chicago'});
  const p=await c.newPage();
  let TICK=Object.assign({},BASE,{held:[]});
  await p.route('http://hqlab.test/**',r=>{
    const u=r.request().url();
    if(u.includes('/api/board')){
      const body=JSON.parse(r.request().postData()||'{}');
      let res={ok:false};
      if(body.action==='boardTick')res=Object.assign({ok:true},TICK);
      if(body.action==='boardStatus')res={ok:true};
      if(body.action==='boardRadio')res={ok:true,nowPlaying:''};
      if(body.action==='boardAckHold')res={ok:true,rows:3,tag:'✓ SEEN 2:31 PM by Yassin · 1'};
      return r.fulfill({contentType:'application/json',body:JSON.stringify(res)});
    }
    return r.fulfill({contentType:'text/html; charset=utf-8',body:html});
  });
  await p.route(/aladhan|open-meteo|youtube|ytimg|somafm|walmradio|isekoi/,r=>r.abort());
  await p.goto('http://hqlab.test/',{waitUntil:'load'});
  await p.waitForTimeout(2500);

  const push=async h=>{TICK=Object.assign({},BASE,{held:h});
    await p.evaluate(()=>pollSoon());await p.waitForTimeout(1400);};

  await push([held('24-14979-87359')]);
  await p.screenshot({path:path.join(OUT,'hold-takeover.png')});
  await p.evaluate(()=>holdCloseTakeover());
  await p.waitForTimeout(500);
  await p.screenshot({path:path.join(OUT,'hold-strip-red.png')});

  await push([held('24-14979-87359',{acked:true,ackText:'2:31 PM by Yassin · 1',
    note:'HOLD — buyer wants expedited, change from Ground to 2-Day · ✓ SEEN 2:31 PM by Yassin · 1'})]);
  await p.screenshot({path:path.join(OUT,'hold-strip-calm.png')});

  // tablet portrait — the real floor device
  await p.setViewportSize({width:800,height:1280});
  await p.evaluate(()=>{holdSeen={};});
  await push([held('24-14979-87359'),held('SO-24609',{channel:'DIRECT',shipped:false,
    note:'HOLD — customer verifying the engine serial'})]);
  await p.screenshot({path:path.join(OUT,'hold-tablet.png')});
  console.log('shot 4 → design-lab/renders/hold-*.png');
  await b.close();
})();
