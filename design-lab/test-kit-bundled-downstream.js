// =====================================================================================
// test-kit-bundled-downstream.js — the three surfaces that must HONOUR the bundled flag:
//   1. computeKitPrice        (KitPricing.js)   — bundled parts leave the parts total
//   2. _oosComputeKitBuild    (OutOfStock.js)   — bundled parts cannot block a build
//   3. the include[] override (KitPricing.js)   — the human beats the rule, both ways
//
// Set KIT_SRC to point at another checkout to prove these fail on HEAD:
//   mkdir -p /tmp/head && for f in KitPricing.js OutOfStock.js KitRegistry.js; do
//     git show HEAD:$f > /tmp/head/$f; done && KIT_SRC=/tmp/head node test-kit-bundled-downstream.js
// =====================================================================================
const fs=require('fs'), path=require('path'), vm=require('vm');
const SRC = process.env.KIT_SRC || path.join(__dirname,'..');
const HEAD_MODE = !!process.env.KIT_SRC;

const ctx={console,Map,JSON,String,Number,parseInt,parseFloat,isNaN,Math,Array,Object,RegExp};
vm.createContext(ctx);
function slice(file, a, b){
  const src=fs.readFileSync(path.join(SRC,file),'utf8');
  const i=src.indexOf(a); if(i<0) return null;
  const j=b?src.indexOf(b,i):-1;
  return src.slice(i, j<0?src.length:j);
}
// KitPricing: constants + resolver + engine
vm.runInContext(slice('KitPricing.js','var KIT_PRICING = {','function _buildKitComponentPriceMaps'), ctx);
vm.runInContext(slice('KitPricing.js','function _resolveComponentPrice','function computeKitPriceBySku'), ctx);
(function(){
  const src=fs.readFileSync(path.join(SRC,'OutOfStock.js'),'utf8');
  const a=src.indexOf('function _oosComputeKitBuild'), e=src.indexOf('\n}\n',a);
  vm.runInContext(src.slice(a,e+3), ctx);
})();

let pass=0, fail=0;
const ok=(n,c,x)=>{ c?pass++:(fail++, console.log('  ✗ '+n+(x?'  → '+x:''))); };

const maps={ ebay:new Map([['aaa',100],['bbb',40],['ccc',25]]), zoho:new Map() };
const comps=[
  {sku:'aaa',qty:2,name:'Full Gasket Set'},
  {sku:'bbb',qty:1,name:'Head Gasket',  bundled:true, bundledInto:'-1 Full Gasket Set 19077-03310'},
  {sku:'ccc',qty:1,name:'Thrust Washer'}
];

console.log('\n── computeKitPrice');
let r=ctx.computeKitPrice(comps,{maps,discount:0.10});
ok('bundled part is OUT of the parts total (200+25, not +40)', r.rawSum===225, 'rawSum='+r.rawSum);
ok('bundled part is still LISTED, never silently dropped', r.lines.length===3);
ok('the bundled line is flagged excluded', r.lines[1].excluded===true);
ok('the bundled line names its parent', /Full Gasket Set/.test(r.lines[1].bundledInto||''));
ok('kept lines are not flagged', r.lines[0].excluded===false && r.lines[2].excluded===false);
ok('excludedComponents lists it', (r.excludedComponents||[]).join()==='bbb');
ok('the rounded price follows the reduced total', r.roundedTotal===Math.round(225*0.9));

console.log('\n── an unpriced BUNDLED part must not condemn the kit');
const withUnpriced=[
  {sku:'aaa',qty:1,name:'Full Gasket Set'},
  {sku:'zzz',qty:1,name:'Head Gasket', bundled:true, bundledInto:'Full Gasket Set'}
];
r=ctx.computeKitPrice(withUnpriced,{maps,discount:0.10});
ok('kit stays COMPLETE though the bundled part has no price', r.complete===true, 'unpriced='+JSON.stringify(r.unpricedComponents));
ok('the unpriced-but-included case still marks incomplete',
   ctx.computeKitPrice([{sku:'aaa',qty:1,name:'x'},{sku:'zzz',qty:1,name:'y'}],{maps}).complete===false);

console.log('\n── the human overrides the rule, in BOTH directions');
r=ctx.computeKitPrice(comps,{maps,discount:0.10,include:{bbb:true}});
ok('include[sku]=true puts a bundled part BACK in', r.rawSum===265, 'rawSum='+r.rawSum);
ok('...and clears its excluded flag', r.lines[1].excluded===false);
r=ctx.computeKitPrice(comps,{maps,discount:0.10,include:{aaa:false}});
ok('include[sku]=false drops a normal part', r.rawSum===25, 'rawSum='+r.rawSum);
ok('a dropped normal part is flagged excluded', r.lines[0].excluded===true);

console.log('\n── _oosComputeKitBuild');
const avail=s=>({aaa:10,bbb:0,ccc:7}[s]);
const kit={components:comps,unparsedLines:[]};
let b=ctx._oosComputeKitBuild(kit,avail);
ok('a bundled part with ZERO stock cannot block the build', b.buildable===5, 'buildable='+b.buildable);
ok('the limiter is a real component', b.limiter && b.limiter.sku==='aaa', JSON.stringify(b.limiter));
ok('the component count reflects only active parts', b.components==='2 ok', b.components);
const kitAllKept={components:comps.map(c=>({...c,bundled:false})),unparsedLines:[]};
ok('without the flag the zero-stock part still blocks (regression net)',
   ctx._oosComputeKitBuild(kitAllKept,avail).buildable===0);
ok('a kit whose parts are ALL bundled falls back rather than reading as empty',
   typeof ctx._oosComputeKitBuild({components:[{sku:'aaa',qty:1,name:'x',bundled:true}],unparsedLines:[]},avail).buildable === 'number');
ok('unparsed PD still wins over everything',
   ctx._oosComputeKitBuild({components:comps,unparsedLines:['junk']},avail).buildable==='⚠');

console.log('\n' + (fail===0?'✅ ':'❌ ') + pass + ' passed, ' + fail + ' failed'
            + (HEAD_MODE ? '   [KIT_SRC='+SRC+']' : ''));
process.exit(fail===0?0:1);
