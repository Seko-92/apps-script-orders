const { chromium } = require('playwright');
const fs=require('fs'), path=require('path');
const ROOT = '/home/yassin/Desktop/Projects/Projects/Excel Code';
const D = { found:true, sku:'167517', name:'Piston Kit STD 0.50', location:'A-9',
  listingStatus:'Active', available:18, quantity:40, quantitySold:22,
  ebayPrice:34.5, zohoPrice:34.5, priceStatus:'ok', committed:6, images:[],
  isKit:true, viewItemURL:'', kit:{components:[{sku:'155394',name:'Full Gasket Set',qty:1,available:12,location:'A-50'}],buildable:4,limitedBy:''},
  usedIn:[{kitSku:'158679',kitName:'Engine Overhaul Kit STD',buildable:4,blocking:false},
          {kitSku:'217205',kitName:'Repair Kit 0.50',buildable:0,blocking:true}],
  unblock:['217205'] };
const html = fs.readFileSync(path.join(ROOT,'PartConsoleModal.html'),'utf8')
  .replace('<?!= dossierJson ?>', JSON.stringify(D))
  .replace(/<\?!?=\s*\w+\s*\?>/g, '{}');
(async()=>{
  const b=await chromium.launch();
  const p=await b.newPage({viewport:{width:1000,height:760},deviceScaleFactor:2});
  const errs=[]; p.on('pageerror',e=>errs.push(String(e)));
  await p.route('**/*', r=>r.fulfill({status:200,contentType:'text/html; charset=utf-8',body:html}));
  await p.goto('http://hq.test/',{waitUntil:'domcontentloaded'});
  await p.waitForTimeout(500);
  await p.screenshot({path:path.join(__dirname,'renders','partconsole-marks.png'),fullPage:false});
  const r = await p.evaluate(()=>({
    marks: document.querySelectorAll('svg.mk').length,
    broken: [...document.querySelectorAll('svg.mk use')].filter(u=>!document.querySelector(u.getAttribute('href'))).length,
    emoji: (document.body.innerText.match(/[\u{1F000}-\u{1FAFF}]/gu)||[]).length
  }));
  console.log(JSON.stringify(r));
  if(errs.length) console.log('ERRORS: '+errs.join(' | ')); else console.log('no page errors');
  await b.close();
})();
