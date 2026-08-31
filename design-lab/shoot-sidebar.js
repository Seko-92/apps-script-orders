const { chromium } = require('playwright');
const path=require('path');
const fs=require('fs');
// ⚠ SELF-SUFFICIENT: builds its own resolved copy from the REAL Sidebar.html
// instead of a pre-made file in a session scratchpad that no longer exists.
// Served over http:// because setContent/file:// give an origin where
// localStorage THROWS, and the sidebar reads it at the top of its script —
// one throw there kills every statement after it.
const SRC=path.join(__dirname,'..','Sidebar.html');
const HTML=fs.readFileSync(SRC,'utf8').replace("'<?!= boardApiUrl ?>'","''");

(async()=>{
  const b=await chromium.launch();
  const p=await b.newPage({viewport:{width:310,height:1500},deviceScaleFactor:2});
  const errs=[]; p.on('pageerror',e=>errs.push(String(e)));
  await p.route('http://hq.test/**',r=>r.fulfill({contentType:'text/html; charset=utf-8',body:HTML}));
  await p.goto('http://hq.test/sidebar');
  await p.waitForTimeout(2200);
  const info=await p.evaluate(()=>{
    const cs=n=>getComputedStyle(n);
    const cards=[...document.querySelectorAll('.card')].map(c=>({
      id:c.dataset.id, zone:c.dataset.zone,
      icon:(c.querySelector('.card-icon')||{}).textContent,
      accent:cs(c).borderLeftColor}));
    return {cards, zones:[...document.querySelectorAll('.zone-header,.zone-title')].map(e=>e.textContent.trim()),
      bodyBg:cs(document.body).backgroundColor,
      fonts:[...new Set([...document.querySelectorAll('.card-title,.cockpit-stat-value,.status-bar')]
        .map(e=>cs(e).fontFamily.split(',')[0].replace(/["']/g,'')))],
      h:document.body.scrollHeight};
  });
  console.log('cards:',info.cards.length,' zones:',JSON.stringify(info.zones));
  console.log('bodyBg:',info.bodyBg,' fonts:',JSON.stringify(info.fonts),' height:',info.h);
  console.log('accents:',JSON.stringify([...new Set(info.cards.map(c=>c.accent))]));
  console.log('errors:',errs.slice(0,4));
  await p.screenshot({path:path.join(__dirname,'renders','sidebar-shipped.png'),clip:{x:0,y:0,width:310,height:900}});
  await p.screenshot({path:path.join(__dirname,'renders','sidebar-shipped-full.png'),fullPage:true});
  await b.close();
})();
