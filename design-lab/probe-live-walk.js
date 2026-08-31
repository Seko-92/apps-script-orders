// Tap the new A–Z option on the REAL board, against REAL data.
'use strict';
const path=require('path');
const {chromium}=require('playwright');
(async()=>{
  const b=await chromium.launch();
  const c=await b.newContext({viewport:{width:1280,height:800},hasTouch:true,timezoneId:'America/Chicago'});
  const p=await c.newPage();
  const errs=[];
  p.on('console',m=>{ if(m.type()==='error') errs.push(m.text()); });
  p.on('pageerror',e=>errs.push('pageerror: '+e.message));
  await p.goto('https://hq.yassinqurabi.com/',{waitUntil:'load',timeout:45000});
  await p.waitForTimeout(9000);   // let a real tick land
  const opts=await p.evaluate(()=>Array.from(document.querySelectorAll('.sortsw-opt')).map(e=>e.textContent.trim()));
  console.log('  sort options on the live board: '+JSON.stringify(opts));
  for(const m of ['aisle','walk','age','aisle']){
    const r=await p.evaluate(mode=>{
      const el=document.querySelector('.sortsw-opt[data-mode="'+mode+'"]');
      if(!el) return {err:'option not found: '+mode};
      el.click();
      return null;
    }, m);
    if(r&&r.err){ console.log('  ✗ '+r.err); continue; }
    await p.waitForTimeout(900);
    const st=await p.evaluate(()=>{
      const shelves=[];
      ['pickListEbay','pickListDirect'].forEach(id=>{
        const l=document.getElementById(id); if(!l) return;
        l.querySelectorAll('li.pick-row .shelf').forEach(s=>shelves.push(s.textContent.trim()));
      });
      return {
        on:(document.querySelector('.sortsw-opt.on')||{}).textContent,
        word:(document.getElementById('pickModeWord')||{}).textContent,
        rows:document.querySelectorAll('.pick-row').length,
        bands:document.querySelectorAll('.pick-head').length,
        shelves
      };
    });
    console.log('  ✓ tapped '+m.padEnd(6)+' → lit "'+String(st.on).trim()+'"  word="'+st.word+
                '"  rows='+st.rows+' bands='+st.bands+'  shelves: '+st.shelves.join(' '));
  }
  await p.screenshot({path:path.join(__dirname,'renders','live-walk.png')});
  console.log('\n  console errors: '+(errs.length?errs.join(' | '):'none'));
  await b.close();
})().catch(e=>{console.error('CRASH:',e.message);process.exit(1);});
