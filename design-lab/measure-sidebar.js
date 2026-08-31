const { chromium } = require('playwright');
const path=require('path');
const SB='/tmp/claude-1000/-home-yassin-Desktop-Projects-Projects-Excel-Code/4b8083d7-c1fd-4c06-afd2-fd819e2a5e0e/scratchpad';
(async()=>{
  const b=await chromium.launch();
  const out={};
  for(const f of ['sidebar-live.html','sidebar-after.html']){
    const p=await b.newPage({viewport:{width:310,height:4200}});
    const errs=[]; p.on('pageerror',e=>errs.push(String(e)));
    await p.goto('file://'+path.join(SB,f)); await p.waitForTimeout(2200);
    out[f]=await p.evaluate(()=>{
      const q=s=>document.querySelector(s);
      const alerts=q('.card[data-id="alerts"]');
      const emoji=[...document.querySelectorAll('.card-icon,.ctrl-btn')]
        .filter(e=>/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/u.test(e.textContent)).length;
      const titles=[...document.querySelectorAll('.card-title h3')];
      const upper=titles.filter(t=>getComputedStyle(t).textTransform==='uppercase').length;
      return {alertsH:Math.round(alerts.getBoundingClientRect().height),
        docH:document.body.scrollHeight, emojiIcons:emoji,
        titles:titles.length, uppercaseTitles:upper,
        boxedAlertRows:[...document.querySelectorAll('.alert-row')]
          .filter(r=>getComputedStyle(r).backgroundColor!=='rgba(0, 0, 0, 0)').length,
        alertRowH:Math.round(document.querySelector('.alert-row').getBoundingClientRect().height),
        borderedThings:[...document.querySelectorAll('.card-body *')]
          .filter(e=>parseFloat(getComputedStyle(e).borderTopWidth)>0
                  && parseFloat(getComputedStyle(e).borderLeftWidth)>0).length,
        titleAlign:getComputedStyle(document.querySelector('.card-title')).justifyContent,
        marks:document.querySelectorAll('.card-icon svg, .ctrl-btn svg').length};
    });
    out[f].errors=errs.length;
    await p.close();
  }
  const A=out['sidebar-live.html'], B=out['sidebar-after.html'];
  const row=(k,a,b,unit='')=>console.log('  '+k.padEnd(26)+String(a).padStart(7)+unit+'   ->'+String(b).padStart(7)+unit);
  console.log('\n                              BEFORE       AFTER');
  row('Alerts card height',A.alertsH,B.alertsH,'px');
  row('Whole panel height',A.docH,B.docH,'px');
  row('Emoji icons',A.emojiIcons,B.emojiIcons);
  row('Drawn marks',A.marks,B.marks);
  row('Uppercase card titles',A.uppercaseTitles+'/'+A.titles,B.uppercaseTitles+'/'+B.titles);
  row('Boxed alert rows',A.boxedAlertRows,B.boxedAlertRows);
  row('One alert row height',A.alertRowH,B.alertRowH,'px');
  row('Fully-boxed elements',A.borderedThings,B.borderedThings);
  row('Card title alignment',A.titleAlign,B.titleAlign);
  row('JS errors',A.errors,B.errors);
  console.log('\n  Alerts card saves '+(A.alertsH-B.alertsH)+'px ('+
    Math.round((1-B.alertsH/A.alertsH)*100)+'%)');
  await b.close();
})();
