'use strict';
const fs=require('fs'),path=require('path');const{chromium}=require('playwright');
const html=fs.readFileSync(path.join(__dirname,'..','FloorBoard.html'),'utf8');
const MOCK=require('./mock-tick.js');
(async()=>{const b=await chromium.launch();
for(const [tag,pk] of [['unset',''],['set','Shipping - Yassin 1']]){
  const ctx=await b.newContext({viewport:{width:800,height:1280},hasTouch:true,timezoneId:'America/Chicago'});
  const p=await ctx.newPage();
  await p.route('http://hqlab.test/**',r=>{const u=r.request().url();
    if(u.includes('/api/board')){const bd=JSON.parse(r.request().postData()||'{}');
      let res={ok:false};
      if(bd.action==='boardTick'){const t=JSON.parse(JSON.stringify(MOCK));t.picker=pk;res=Object.assign({ok:true},t);}
      if(bd.action==='boardRadio')res={ok:true,nowPlaying:''};
      return r.fulfill({contentType:'application/json',body:JSON.stringify(res)});}
    return r.fulfill({contentType:'text/html; charset=utf-8',body:html});});
  await p.route(/aladhan|open-meteo/,r=>r.abort());
  await p.goto('http://hqlab.test/',{waitUntil:'load'});
  await p.waitForTimeout(2500);
  await p.locator('.ftr').screenshot({path:'renders/ftr-'+tag+'.png'});
  console.log(tag+' ok');
  await ctx.close();}
await b.close();})();
