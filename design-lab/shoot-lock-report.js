/**
 * shoot-lock-report.js — the lock's multi-line checklist, rendered in the card.
 *
 * ⚠ Nothing else populates #lockAcctResult with a REPORT, so the 260px scroll cap and the
 *   success-vs-failure colouring are invisible to every other harness. A successful install
 *   contains "⚠⚠ NOT in an incognito window", which an over-eager failure regex painted red.
 */
const { chromium } = require('playwright');
const path=require('path'), fs=require('fs');
const HTML=fs.readFileSync(path.join(__dirname,'..','Sidebar.html'),'utf8').replace("'<?!= boardApiUrl ?>'","''");
const REPORT = ["══ INSTALL THE ALL ORDERS LOCK ══","",
"1 · bridge      ✓ answering, allowlist holding",
"2 · n8n acct    ✓ none",
"3 · protection  ✅ All Orders LOCKED",
"   editable: E4:E · F4:F · H4:H · F2:G2 · H2",
"   n8n exception: none","",
"── NOW TEST IT THE WAY STAFF ACTUALLY WORK ──",
"  ⚠⚠ NOT in an incognito window. That is the mistake 2026-08-29 made: an",
"     anonymous user has no sidebar and no ⚙️ menu at all.","",
"  1. signed in as staff: a sort, Update Locations, a kit commit, a Zoho pull",
"  2. typing into SKU / QTY / LOCATION / SALES ORDER is refused",
"  3. NOTE, STATUS and LEFT still accept, and the F2/H2 dropdowns still work",
"  4. ⚠ let a real n8n sync land, then confirm the ~1 AM sweep still deletes",
"     shipped rows — the only failure here that is silent","",
"  Rollback at any point: unprotectAllOrdersSheet()"].join("\n");
(async()=>{
  const b=await chromium.launch();
  const p=await b.newPage({viewport:{width:310,height:2600},deviceScaleFactor:2});
  await p.addInitScript(rep=>{
    const D={getDisplayUrls:{board:'x',wall:'y',hosted:true},getSidebarTick:null,getCurrentPicker:'',
             getActionableAlerts:null,
             getN8nSheetsAccountState:{ok:true,owner:true,isSet:true,isNone:true,value:'none',key:'N8N_SHEETS_ACCOUNT'},
             installAllOrdersLock:rep};
    const mk=(su,fa)=>new Proxy({},{get(_,k){if(k==='withSuccessHandler')return f=>mk(f,fa);if(k==='withFailureHandler')return f=>mk(su,f);
      return()=>{const v=Object.prototype.hasOwnProperty.call(D,k)?D[k]:null;if(su)setTimeout(()=>su(v),0);};}});
    window.google={script:{run:mk(null,null),host:{close(){},setHeight(){}},url:{getLocation(f){f({parameter:{}});}}}};
    window.confirm=()=>true;
  }, REPORT);
  await p.route('http://hq.test/**',r=>r.fulfill({contentType:'text/html; charset=utf-8',body:HTML}));
  await p.goto('http://hq.test/sidebar'); await p.waitForTimeout(1800);
  await p.evaluate(()=>{document.querySelector('.card[data-id="sheet-protect"] .card-header').click();
    document.querySelectorAll('.card-body').forEach(b=>b.style.transition='none');});
  await p.waitForTimeout(700);
  await p.evaluate(()=>lockAllOrdersNow());
  await p.waitForTimeout(600);
  const st=await p.evaluate(()=>{const c=document.querySelector('.card[data-id="sheet-protect"]');
    const bd=c.querySelector('.card-body'); const r=document.getElementById('lockAcctResult');
    return {card:Math.round(c.getBoundingClientRect().height), clipped:bd.scrollHeight>bd.clientHeight+2,
            boxH:Math.round(r.getBoundingClientRect().height), scrolls:r.scrollHeight>r.clientHeight+2,
            cls:r.className};});
  console.log('  card '+st.card+'px  '+(st.clipped?'❌ CLIPPED':'✅ not clipped'));
  console.log('  report box '+st.boxH+'px  '+(st.scrolls?'(scrolls — capped)':'(fits)')+'  class="'+st.cls+'"');
  await p.locator('.card[data-id="sheet-protect"]').screenshot({path:path.join(__dirname,'renders','lock-report.png')});
  await b.close();
})();
