// Sizes the "advertising more than we can build" population, WITH the bundled rule live.
const fs=require('fs'),path=require('path'),vm=require('vm');
const ctx={console,Map,JSON,String,Number,parseInt,parseFloat,isNaN,Math,Array,Object,RegExp};
vm.createContext(ctx);
const reg=fs.readFileSync(path.join(__dirname,'..','KitRegistry.js'),'utf8');
vm.runInContext(reg.slice(reg.indexOf('var KIT_BUNDLE = {'), reg.indexOf('// =======================================================================================\n// READ API')), ctx);
(function(){const s=fs.readFileSync(path.join(__dirname,'..','OutOfStock.js'),'utf8');
  const a=s.indexOf('function _oosComputeKitBuild'), e=s.indexOf('\n}\n',a); vm.runInContext(s.slice(a,e+3),ctx);})();

const kits=JSON.parse(fs.readFileSync(path.join(__dirname,'fixture-kit-registry.json'),'utf8'));
const px=JSON.parse(fs.readFileSync(path.join(__dirname,'fixture-prices.json'),'utf8'));
const mi=JSON.parse(fs.readFileSync(path.join(__dirname,'fixture-mi.json'),'utf8'));
const availOf=s=>{const z=px.zoho[s]; return z&&z.available!=null?z.available:null;};

let rows=[];
kits.forEach(k=>{
  const kit=JSON.parse(JSON.stringify(k));
  ctx._kbAnnotateKit(kit);
  const b=ctx._oosComputeKitBuild(kit,availOf);
  if(typeof b.buildable!=='number') return;
  const s=kit.sku.toLowerCase();
  const zo=(px.zoho[s]&&px.zoho[s].available)||0;
  const eb=(mi[s]&&mi[s].avail!=null)?mi[s].avail:null;
  const advertised=Math.max(zo, eb==null?0:eb);
  const gap=advertised-b.buildable;
  if(gap>0) rows.push({sku:kit.sku,type:kit.type,name:kit.name,adv:advertised,zo,eb,build:b.buildable,gap,lim:b.limitedBy});
});
const man=rows.filter(r=>r.type!=='READY'), rdy=rows.filter(r=>r.type==='READY');
console.log('kits advertising more than buildable : '+rows.length+'   (MANUAL '+man.length+' · READY '+rdy.length+')');
const crit=man.filter(r=>r.build===0), part=man.filter(r=>r.build>0);
console.log('  MANUAL, buildable 0  (next sale fails) : '+crit.length);
console.log('  MANUAL, partly covered                 : '+part.length);
console.log('  total units advertised but unmakeable  : '+man.reduce((s,r)=>s+r.gap,0));
console.log('  kits where eBay and Zoho disagree      : '+man.filter(r=>r.eb!=null&&r.eb!==r.zo).length);
console.log('\nworst 12 MANUAL:');
man.sort((a,b)=>b.gap-a.gap||a.build-b.build).slice(0,12).forEach(r=>
  console.log('   '+r.sku+'  adv '+r.adv+' (eBay '+r.eb+' / Zoho '+r.zo+')  build '+r.build+'  gap '+r.gap+'   '+String(r.lim).slice(0,40)));
