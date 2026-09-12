/* diag-palette.js — the command palette's WAYS OUT.
 *
 * Reported from use 2026-09-12: "the esc button not working, and when I click
 * out of the box sometimes it doesn't get out."  Both are the same root cause
 * seen twice — there was no working close affordance:
 *   · the ESC chip is styled exactly like a button and was INERT,
 *   · and the only real exit was a ~12px strip of backdrop beside the box,
 *     which on a TABLET (no Escape key) is the whole exit surface.
 *   · plus: click the SHEET behind the sidebar and focus leaves the iframe, so
 *     our document keydown never sees Escape again.
 *
 * PROVE IT BITES:
 *   git show HEAD:Sidebar.html > /tmp/old.html && SIDEBAR_SRC=/tmp/old.html node diag-palette.js
 */
const { chromium } = require('playwright');
const path=require('path'), fs=require('fs');
const SRC=process.env.SIDEBAR_SRC||path.join(__dirname,'..','Sidebar.html');
const HTML=fs.readFileSync(SRC,'utf8').replace("'<?!= boardApiUrl ?>'","''");

(async()=>{
  const b=await chromium.launch();
  const ctx=await b.newContext({viewport:{width:310,height:900},deviceScaleFactor:1});
  const p=await ctx.newPage();
  await p.addInitScript(()=>{
    const D={getDisplayUrls:{board:'',wall:'',hosted:false},getSidebarTick:null,
             getCurrentPicker:'',getActionableAlerts:null};
    const mk=(su,fa)=>new Proxy({},{get(_,k){
      if(k==='withSuccessHandler')return f=>mk(f,fa);
      if(k==='withFailureHandler')return f=>mk(su,f);
      return ()=>{const v=Object.prototype.hasOwnProperty.call(D,k)?D[k]:null; if(su)setTimeout(()=>su(v),0);};}});
    window.google={script:{run:mk(null,null),host:{close(){},setHeight(){}},url:{getLocation(f){f({parameter:{}});}}}};
  });
  await p.route('http://hq.test/**',r=>r.fulfill({contentType:'text/html; charset=utf-8',body:HTML}));
  await p.goto('http://hq.test/sidebar');
  await p.waitForTimeout(2200);

  let fail=0;
  const ok=(n,c,x)=>{c?console.log('  ok  '+n):(fail++,console.log('  XX  '+n+(x!==undefined?'  -> '+JSON.stringify(x):'')));};
  const open  = ()=>p.evaluate(()=>{ openPalette(); });
  const isOpen= ()=>p.evaluate(()=>document.getElementById('palette').classList.contains('active'));

  console.log('\nA · control — it opens and renders');
  await open(); await p.waitForTimeout(200);
  ok('palette is open', await isOpen());
  ok('commands rendered', await p.evaluate(()=>document.querySelectorAll('#cmdList .cmd').length)>0);

  console.log('\nB · THE HEADLINE — tapping the ESC chip gets you out');
  const chip=await p.$('.palette-hint');
  ok('the ESC chip exists', !!chip);
  if(chip){
    const box=await chip.boundingBox();
    // a REAL mouse click at real coordinates — the browser decides whether a
    // click happened, not us (the 2026-08-19 tapsteal lesson).
    await p.mouse.click(box.x+box.width/2, box.y+box.height/2);
    await p.waitForTimeout(250);
    ok('tapping ESC closed the palette', (await isOpen())===false, {stillOpen:await isOpen()});
    ok('the chip LOOKS tappable (cursor:pointer)',
       await p.evaluate(()=>getComputedStyle(document.querySelector('.palette-hint')).cursor)==='pointer',
       await p.evaluate(()=>getComputedStyle(document.querySelector('.palette-hint')).cursor));
  }

  console.log('\nC · focus leaves the sidebar (you clicked the SHEET) — it dismisses');
  await open(); await p.waitForTimeout(150);
  const p2=await ctx.newPage(); await p2.goto('about:blank'); await p2.bringToFront();
  await p.waitForTimeout(300);
  let realBlur=(await isOpen())===false;
  if(!realBlur){ // headless may not deliver a real window blur; drive the contract directly
    await p.bringToFront();
    await p.evaluate(()=>window.dispatchEvent(new Event('blur')));
    await p.waitForTimeout(200);
  }
  await p.bringToFront(); await p2.close();
  ok('leaving the panel closed it'+(realBlur?' (real tab blur)':' (blur event)'),
     (await isOpen())===false, {stillOpen:await isOpen()});

  console.log('\nD · regression nets — the exits that already worked');
  await open(); await p.waitForTimeout(150);
  await p.keyboard.press('Escape'); await p.waitForTimeout(200);
  ok('Escape still closes', (await isOpen())===false);

  await open(); await p.waitForTimeout(150);
  await p.mouse.click(155, 860);                       // backdrop, well below the box
  await p.waitForTimeout(200);
  ok('backdrop click still closes', (await isOpen())===false);

  console.log('\nE · it must not over-close');
  await open(); await p.waitForTimeout(150);
  const hdr=await (await p.$('.palette-header')).boundingBox();
  await p.mouse.click(hdr.x+6, hdr.y+hdr.height/2);     // inside the BOX chrome
  await p.waitForTimeout(200);
  ok('clicking the box itself keeps it open', (await isOpen())===true);
  await p.keyboard.press('Escape'); await p.waitForTimeout(150);

  console.log('\nF · a command still runs (close must not eat the action)');
  await p.evaluate(()=>{ window.__ran=null; window.run=function(fn){ window.__ran=fn; }; });
  await open(); await p.waitForTimeout(200);
  const row=await (await p.$('#cmdList .cmd')).boundingBox();
  await p.mouse.click(row.x+row.width/2, row.y+row.height/2);
  await p.waitForTimeout(250);
  ok('the command fired', await p.evaluate(()=>window.__ran)!==null, await p.evaluate(()=>window.__ran));
  ok('and the palette closed', (await isOpen())===false);

  console.log(fail? '\n'+fail+' FAILED\n' : '\nOK all checks passed\n');
  await b.close(); process.exit(fail?1:0);
})();
