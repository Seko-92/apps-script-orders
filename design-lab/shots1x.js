const { chromium } = require('playwright');
const path=require('path');
const SB='/tmp/claude-1000/-home-yassin-Desktop-Projects-Projects-Excel-Code/4b8083d7-c1fd-4c06-afd2-fd819e2a5e0e/scratchpad';
(async()=>{
  const b=await chromium.launch();
  for(const [f,out] of [['sidebar-live.html','shot-before.png'],['sidebar-after.html','shot-after.png']]){
    const p=await b.newPage({viewport:{width:310,height:1180},deviceScaleFactor:1});
    await p.goto('file://'+path.join(SB,f)); await p.waitForTimeout(2200);
    await p.screenshot({path:path.join(SB,out),clip:{x:0,y:0,width:310,height:1180}});
    await p.close();
  }
  await b.close(); console.log('shot');
})();
