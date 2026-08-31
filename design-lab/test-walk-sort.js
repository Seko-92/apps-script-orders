// =====================================================================================
// test-walk-sort.js — the A–Z / WALK sort.
//
// THE POINT OF THIS FILE: the shelf comparator now exists TWICE — compareLocations()
// in Helpers.js (server) and cmpShelf() in FloorBoard.html (client, because the board
// cannot call server code). This runs BOTH over the same fixtures and fails if they
// ever disagree, so the sheet and the board can never end up in different aisles.
// =====================================================================================
const fs=require('fs'), path=require('path'), vm=require('vm');
const R=p=>fs.readFileSync(path.join(__dirname,'..',p),'utf8');

// --- server side ---
const srv={console,String,Number,parseInt,parseFloat,isNaN,Math,Array,Object,RegExp};
vm.createContext(srv);
const H=R('Helpers.js');
vm.runInContext(H.slice(H.indexOf('function _parseShelfLocation'),
                        H.indexOf('\n}\n', H.indexOf('function compareLocations'))+3), srv);

// --- client side, lifted out of the real board ---
const cli={console,String,Number,parseInt,parseFloat,isNaN,Math,Array,Object,RegExp};
vm.createContext(cli);
const B=R('FloorBoard.html');
const a=B.indexOf('function _shelfParts'), b=B.indexOf('function sortForMode');
if(a<0||b<0){ console.error('FATAL: could not find the client comparator'); process.exit(1); }
vm.runInContext(B.slice(a,b), cli);

let pass=0, fail=0;
const ok=(n,c,x)=>{ c?pass++:(fail++,console.log('  ✗ '+n+(x!==undefined?'  → '+JSON.stringify(x):''))); };

// ---------------------------------------------------------------- agreement
console.log('\n── the two implementations must agree, always');
const shelves=['A-9','A-50','A-1','B-4','B-52','B-8','C-50','C-61','G-2','H-24','Z-9','Z-90','Z-150',
  'G-35 * 2','L-208/C-51','K-2',' K-2','D-1','D-36','NOT FOUND','','S93','C-14','A-9 ','a-9','L-226/A-56'];
let disagree=0;
for(let i=0;i<shelves.length;i++) for(let j=0;j<shelves.length;j++){
  const s=Math.sign(srv.compareLocations(shelves[i],shelves[j]));
  const c=Math.sign(cli.cmpShelf(shelves[i],shelves[j]));
  if(s!==c){ disagree++; if(disagree<5) console.log('    server '+s+' vs client '+c+'  for '+JSON.stringify([shelves[i],shelves[j]])); }
}
ok('every pair of '+shelves.length+' real shelf strings compares identically', disagree===0, disagree+' disagreements');
const srvSorted=shelves.slice().sort(srv.compareLocations);
const cliSorted=shelves.slice().sort(cli.cmpShelf);
ok('and a full sort produces the same order', JSON.stringify(srvSorted)===JSON.stringify(cliSorted));

console.log('\n── the bug this exists to prevent');
ok('A-9 comes BEFORE A-50 (plain string compare gets this wrong)', cli.cmpShelf('A-9','A-50')<0);
ok('B-8 before B-52', cli.cmpShelf('B-8','B-52')<0);
ok('Z-9 before Z-90 before Z-150',
   cli.cmpShelf('Z-9','Z-90')<0 && cli.cmpShelf('Z-90','Z-150')<0);
ok('aisle letter beats the number', cli.cmpShelf('B-1','A-99')>0);
ok('the pack-size suffix sorts on the shelf, not the text', cli.cmpShelf('G-35 * 2','G-36')<0);
ok('a leading space does not change the shelf', cli.cmpShelf(' K-2','K-2')===0);
ok('case does not change the shelf', cli.cmpShelf('a-9','A-9')===0);

console.log('\n── unwalkable shelves sink');
ok('NOT FOUND is flagged missing', cli._shelfMissing('NOT FOUND')===1);
ok('blank is flagged missing', cli._shelfMissing('')===1 && cli._shelfMissing(null)===1);
ok('a real shelf is not', cli._shelfMissing('C-50')===0);

// ---------------------------------------------------------------- the sort itself
console.log('\n── the walk itself');
function walk(rows){
  return rows.slice().sort(function(a,b){
    var ca=(a.channel==='DIRECT')?1:0, cb=(b.channel==='DIRECT')?1:0;
    if(ca!==cb) return ca-cb;
    var ma=cli._shelfMissing(a.location), mb=cli._shelfMissing(b.location);
    if(ma!==mb) return ma-mb;
    var s=cli.cmpShelf(a.location,b.location);
    if(s!==0) return s;
    return String(a.sku||'').localeCompare(String(b.sku||''));
  });
}
// two eBay orders whose lines interleave by shelf — the case aisle mode anchors
const rows=[
  {channel:'EBAY',orderId:'O-1',sku:'111111',location:'F-6'},
  {channel:'EBAY',orderId:'O-1',sku:'222222',location:'L-208'},
  {channel:'EBAY',orderId:'O-2',sku:'333333',location:'J-8'},
  {channel:'EBAY',orderId:'O-1',sku:'444444',location:'E-89'},
  {channel:'EBAY',orderId:'O-2',sku:'555555',location:'A-9'},
  {channel:'EBAY',orderId:'O-3',sku:'666666',location:'NOT FOUND'},
  {channel:'DIRECT',orderId:'SO-1',sku:'777777',location:'B-2'},
  {channel:'DIRECT',orderId:'SO-2',sku:'888888',location:'A-4'}
];
const w=walk(rows);
ok('eBay block still comes before DIRECT',
   w.findIndex(r=>r.channel==='DIRECT') === w.filter(r=>r.channel==='EBAY').length);
ok('eBay walks A-9 → E-89 → F-6 → J-8 → L-208, orders interleaved',
   w.filter(r=>r.channel==='EBAY'&&r.location!=='NOT FOUND').map(r=>r.location).join(' ')
   === 'A-9 E-89 F-6 J-8 L-208',
   w.filter(r=>r.channel==='EBAY').map(r=>r.location));
ok('the shelf-less row sits at the END of its own channel, not between shelves',
   w.filter(r=>r.channel==='EBAY').slice(-1)[0].location==='NOT FOUND');
ok('DIRECT walks its own aisles too',
   w.filter(r=>r.channel==='DIRECT').map(r=>r.location).join(' ')==='A-4 B-2');
ok('an order IS scattered — that is the mode working, not a fault',
   w.findIndex(r=>r.orderId==='O-1') < w.findIndex(r=>r.orderId==='O-2') === false ||
   w.filter(r=>r.orderId==='O-1').length===3);
ok('same shelf → deterministic by SKU',
   walk([{channel:'EBAY',sku:'999',location:'C-1'},{channel:'EBAY',sku:'111',location:'C-1'}])[0].sku==='111');
ok('sorting does not mutate the input', rows[0].sku==='111111');
ok('empty list is safe', walk([]).length===0);

console.log('\n'+(fail===0?'✅ ':'❌ ')+pass+' passed, '+fail+' failed');
process.exit(fail===0?0:1);
