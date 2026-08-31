// ============================================================================
// shoot-walk.js — the A–Z / WALK sort, seen next to the two modes it joins.
// Drives the REAL board (fresh read) on the picker's own device size, taps each
// mode, and reports what the picker would actually SEE: the shelf sequence, and
// whether any order band survived (in walk mode none may).
// ============================================================================
'use strict';
const fs=require('fs'), path=require('path');
const { chromium } = require('playwright');
const BOARD=path.join(__dirname,'..','FloorBoard.html');
const MOCK=require('./mock-tick.js');
const OUT=path.join(__dirname,'renders'); fs.mkdirSync(OUT,{recursive:true});

(async()=>{
  const html=fs.readFileSync(BOARD,'utf8');
  const browser=await chromium.launch();
  const ctx=await browser.newContext({viewport:{width:1280,height:800},hasTouch:true,timezoneId:'America/Chicago'});
  const page=await ctx.newPage();
  const errors=[];
  page.on('console',m=>{ if(m.type()==='error') errors.push(m.text()); });
  page.on('pageerror',e=>errors.push('pageerror: '+e.message));
  await page.route('http://hqlab.test/**', route=>{
    const url=route.request().url();
    if(url.includes('/api/board')){
      const body=JSON.parse(route.request().postData()||'{}');
      let res={ok:false,message:'unknown action '+body.action};
      if(body.action==='boardTick')   res=Object.assign({ok:true},MOCK);
      if(body.action==='boardStatus') res={ok:true};
      if(body.action==='boardRadio')  res={ok:true,nowPlaying:''};
      return route.fulfill({contentType:'application/json',body:JSON.stringify(res)});
    }
    return route.fulfill({contentType:'text/html; charset=utf-8',body:html});
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r=>r.abort());
  await page.goto('http://hqlab.test/',{waitUntil:'load'});
  await page.waitForFunction(()=>!document.getElementById('board').classList.contains('booting'),null,{timeout:20000})
    .catch(()=>errors.push('board never left booting'));
  await page.waitForTimeout(2200);

  const read=()=>page.evaluate(()=>{
    const out={ebayShelves:[],directShelves:[],bands:0,rowsWithOwnId:0,rows:0,spines:0,sections:0,holds:0};
    let ch=null;
    [['pickListEbay','E'],['pickListDirect','D']].forEach(([id,which])=>{
      const list=document.getElementById(id); if(!list) return;
      list.querySelectorAll('li').forEach(el=>{
        if(el.classList.contains('pick-sec')){ out.sections++; return; }
        if(el.classList.contains('pick-head')){ out.bands++; return; }
        if(!el.classList.contains('pick-row')) return;
        out.rows++;
        const loc=el.querySelector('.shelf');
        if(loc){ (which==='D'?out.directShelves:out.ebayShelves).push(loc.textContent.trim().replace(/\s+/g,' ')); }
        const meta=el.querySelector('.sub-id');
        if(meta && meta.textContent.trim()) out.rowsWithOwnId++;
        if(el.classList.contains('kitted')) out.spines++;
        if(el.querySelector('.pick-hold')) out.holds=(out.holds||0)+1;
      });
    });
    return out;
  });

  const modes=[['aisle','Aisle'],['age','Age'],['walk','A–Z']];
  const report={};
  for(const [mode,label] of modes){
    await page.evaluate(m=>setPickMode(m), mode);
    await page.waitForTimeout(700);
    const st=await read();
    report[mode]=st;
    const f=path.join(OUT,'walk-'+mode+'.png');
    await page.screenshot({path:f});
    console.log('✓ '+label.padEnd(5)+' rows='+String(st.rows).padEnd(3)+
                ' bands='+String(st.bands).padEnd(2)+
                ' kit-threads='+String(st.spines).padEnd(2)+
                ' holds='+String(st.holds).padEnd(2)+
                ' rows carrying their own order id='+String(st.rowsWithOwnId).padEnd(3)+
                ' → '+f.split('/').pop());
    console.log('    eBay   : '+st.ebayShelves.join('  '));
    console.log('    DIRECT : '+st.directShelves.join('  '));
  }
  console.log('\nconsole errors: '+(errors.length?errors.join(' | '):'none'));
  fs.writeFileSync(path.join(OUT,'walk-report.json'),JSON.stringify(report,null,1));
  await browser.close();
})().catch(e=>{console.error('CRASH:',e);process.exit(1);});
