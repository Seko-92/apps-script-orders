// Renders ZohoPullModal.html with a fixture diff and asserts the per-line note
// boxes: present on new + qty lines only, the Zoho note copies in on click, and
// Apply sends each line's note in its selection.
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'ZohoPullModal.html'), 'utf8');
const diff = { soNumber:'SO-99999', customerName:'Test Co', totalFormatted:'$1,234.00', invoiceNumber:'',
  summary:{ totalLines:4, new:2, qtyChanged:1, removed:1, unchanged:0 },
  lines:[
    { sku:'166500', name:'Piston Kit STD', zohoNote:'customer wants the 0.25 oversize — double check', status:'new', zohoQty:2, directQty:0, delta:2, location:'E-12', available:4, missing:false, directRows:[] },
    { sku:'173817', name:'Gasket Set', zohoNote:'', status:'new', zohoQty:1, directQty:0, delta:1, location:'F-3', available:9, missing:false, directRows:[] },
    { sku:'158373', name:'Bearing Set', zohoNote:'', status:'qty_changed', zohoQty:5, directQty:3, delta:2, location:'C-55', available:20, missing:false, directRows:[{row:30,status:'PENDING',qty:3}] },
    { sku:'199095', name:'Old line', zohoNote:'', status:'removed', zohoQty:0, directQty:1, delta:-1, location:'E-84', available:1, missing:false, directRows:[{row:31,status:'PENDING',qty:1}] } ] };
const html = src.replace('<?!= diffJson ?>', JSON.stringify(diff));
let ok = 0, bad = 0; const t = (n, c) => { c ? ok++ : bad++; console.log((c?'✓ ':'✗ ')+n); };
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport:{ width:920, height:900 } });
  await p.addInitScript(() => { window.__sent=null; window.google={script:{run:new Proxy({},{get:(o,k)=>{
    if(k==='withSuccessHandler'||k==='withFailureHandler') return ()=>window.google.script.run;
    return (...a)=>{ window.__sent={fn:k,args:a}; }; }})}}; });
  await p.route('http://hq.test/', r => r.fulfill({ contentType:'text/html; charset=utf-8', body: html }));
  await p.goto('http://hq.test/');
  t('note box on new line', await p.$('#note-166500') !== null);
  t('note box on qty-changed line', await p.$('#note-158373') !== null);
  t('no note box on removed line', await p.$('#note-199095') === null);
  t('Zoho note shown', (await p.textContent('#line-166500')).includes('0.25 oversize'));
  await p.click('#line-166500 .zoho-note');
  t('click copies Zoho note', (await p.inputValue('#note-166500')).includes('0.25 oversize'));
  await p.fill('#note-173817', 'pack separately');
  await p.screenshot({ path: path.join(__dirname, 'renders', 'pull-notes.png'), fullPage:true });
  await p.click('#btnApply');
  const s = await p.evaluate(() => window.__sent);
  const sels = s && s.args[1] || [];
  const by = k => sels.find(x => x.sku === k) || {};
  t('apply sends line note', by('173817').note === 'pack separately');
  t('apply sends copied Zoho note', /oversize/.test(by('166500').note || ''));
  t('blank note not sent', !('note' in by('158373')));
  await b.close(); console.log(ok + ' passed, ' + bad + ' failed'); process.exit(bad ? 1 : 0);
})();
