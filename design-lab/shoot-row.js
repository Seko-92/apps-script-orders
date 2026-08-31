'use strict';
const fs=require('fs'),path=require('path');const{chromium}=require('playwright');
const html=fs.readFileSync(path.join(__dirname,'..','FloorBoard.html'),'utf8');
const MOCK=require('./mock-tick.js');const SKU='165447';
(async()=>{const br=await chromium.launch();
const ctx=await br.newContext({viewport:{width:1000,height:1280},hasTouch:true,timezoneId:'America/Chicago'});
const p=await ctx.newPage();let mirror=null;
await p.route('http://hqlab.test/**',r=>{const u=r.request().url();
 if(u.includes('/api/board')){const b=JSON.parse(r.request().postData()||'{}');let res={ok:false};
  if(b.action==='boardTick'){const t=JSON.parse(JSON.stringify(MOCK));
    if(mirror!==null)t.openOrders.forEach(x=>{if(String(x.sku)===SKU)x.hand=mirror;});
    res=Object.assign({ok:true},t);}
  if(b.action==='boardAdjust')res={ok:true,before:4,after:b.target,delta:b.target-4,noop:false};
  if(b.action==='boardRadio')res={ok:true,nowPlaying:''};
  return r.fulfill({contentType:'application/json',body:JSON.stringify(res)});}
 return r.fulfill({contentType:'text/html; charset=utf-8',body:html});});
await p.route(/aladhan|open-meteo/,r=>r.abort());
await p.goto('http://hqlab.test/',{waitUntil:'load'});
await p.waitForTimeout(2200);
const row=()=>p.locator('.pick-row').filter({hasText:SKU}).first();
await row().screenshot({path:'renders/row-1-deviance.png'});
await p.evaluate(s=>{const li=[...document.querySelectorAll('.pick-row')].find(r=>r.textContent.includes(s));li.querySelector('.pc-adj').click();},SKU);
await p.waitForTimeout(400);await p.click('#npOk');await p.waitForTimeout(1400);
await row().screenshot({path:'renders/row-2-fixed.png'});
mirror=1;await p.evaluate(()=>{if(window.pollSoon)pollSoon();});await p.waitForTimeout(2400);
await row().screenshot({path:'renders/row-3-healed.png'});
console.log('3 states shot');await br.close();})();
