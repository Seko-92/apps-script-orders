// =====================================================================================
// sim-kit-health.js — replays the REAL engines (KitRegistry.js bundled rule,
// KitPricing.computeKitPrice, OutOfStock._oosComputeKitBuild) over the LIVE
// 223-kit registry + live prices/stock, WITH and WITHOUT the bundled rule.
// Proves the change against production data before it is deployed.
// =====================================================================================
const fs=require('fs'), path=require('path'), vm=require('vm');
function loadPure(file, startMark, endMark){
  const src=fs.readFileSync(path.join(__dirname,'..',file),'utf8');
  const a=src.indexOf(startMark), b=endMark?src.indexOf(endMark):src.length;
  if(a<0) throw new Error('anchor not found in '+file+': '+startMark);
  return src.slice(a, b<0?src.length:b);
}
const ctx={console,Map,JSON,String,Number,parseInt,parseFloat,isNaN,Math,Array,Object,RegExp,
           SpreadsheetApp:undefined};
vm.createContext(ctx);
vm.runInContext(loadPure('KitRegistry.js','var KIT_BUNDLE = {','// =======================================================================================\n// READ API'), ctx);
vm.runInContext(loadPure('KitPricing.js','var KIT_PRICING = {','function _buildKitComponentPriceMaps'), ctx);
vm.runInContext(loadPure('KitPricing.js','function _resolveComponentPrice','function computeKitPriceBySku'), ctx);
(function(){                       // extract exactly the _oosComputeKitBuild body
  const src=fs.readFileSync(path.join(__dirname,'..','OutOfStock.js'),'utf8');
  const a=src.indexOf('function _oosComputeKitBuild');
  const end=src.indexOf('\n}\n', a);
  vm.runInContext(src.slice(a, end+3), ctx);
})();

const kits=JSON.parse(fs.readFileSync(path.join(__dirname,'fixture-kit-registry.json'),'utf8'));
const px=JSON.parse(fs.readFileSync(path.join(__dirname,'fixture-prices.json'),'utf8'));
const maps={ebay:new Map(Object.entries(px.ebay)), zoho:new Map(Object.entries(px.zoho))};
const resolveAvail=s=>{const z=px.zoho[s]; if(z&&z.available!=null) return z.available; return null;};

function median(a){const s=a.slice().sort((x,y)=>x-y);const m=s.length>>1;return s.length%2?s[m]:(s[m-1]+s[m])/2;}

function run(applyBundled){
  const rows=[];
  const implied=[];
  kits.forEach(k=>{
    const kit=JSON.parse(JSON.stringify(k));
    if(applyBundled) ctx._kbAnnotateKit(kit); else kit.components.forEach(c=>c.bundled=false);
    const p0=ctx.computeKitPrice(kit.components,{maps});
    const listed=ctx._resolveComponentPrice(kit.sku,maps).price;
    const complete=(kit.unparsedLines||[]).length===0 && kit.components.length>0 && p0.unpricedComponents.length===0;
    if(complete && listed!=null && p0.rawSum>0) implied.push(1-listed/p0.rawSum);
    rows.push({kit,listed,complete});
  });
  const disc=Math.max(0,Math.min(0.6, implied.length?median(implied):0.10));
  let under=0,over=0,inline=0,noListed=0,incomplete=0,underBy=0,blocked=0,buildableNow=0;
  const per={};
  rows.forEach(({kit,listed,complete})=>{
    const p=ctx.computeKitPrice(kit.components,{maps,discount:disc});
    const b=ctx._oosComputeKitBuild(kit,resolveAvail);
    if(typeof b.buildable==='number'&&b.buildable>0) buildableNow++;
    if(typeof b.buildable!=='number') blocked+=0;
    let status;
    if(!complete){incomplete++;status='INCOMPLETE';}
    else if(listed==null){noListed++;status='NO LISTED';}
    else{
      const d=listed-p.roundedTotal, tol=Math.max(2,0.03*p.roundedTotal);
      if(Math.abs(d)<=tol){inline++;status='IN LINE';}
      else if(d<0){under++;underBy+=d;status='UNDERPRICED';}
      else{over++;status='OVERPRICED';}
    }
    per[kit.sku]={status,parts:p.rawSum,computed:p.roundedTotal,listed,
                  build:b.buildable,limited:b.limitedBy,excluded:p.excludedComponents.length};
  });
  return {disc,under,over,inline,noListed,incomplete,underBy,buildableNow,per};
}

const before=run(false), after=run(true);
const pc=x=>(x*100).toFixed(1)+'%';
console.log('                        BEFORE      AFTER');
console.log('calibrated discount     '+pc(before.disc).padEnd(11)+pc(after.disc));
console.log('UNDERPRICED             '+String(before.under).padEnd(11)+after.under);
console.log('OVERPRICED              '+String(before.over).padEnd(11)+after.over);
console.log('IN LINE                 '+String(before.inline).padEnd(11)+after.inline);
console.log('NO LISTED $             '+String(before.noListed).padEnd(11)+after.noListed);
console.log('⚠ INCOMPLETE            '+String(before.incomplete).padEnd(11)+after.incomplete);
console.log('$ left on the table     '+('$'+Math.abs(before.underBy).toFixed(0)).padEnd(11)+'$'+Math.abs(after.underBy).toFixed(0));
console.log('kits buildable > 0      '+String(before.buildableNow).padEnd(11)+after.buildableNow);

const partsDrop=Object.keys(before.per).reduce((s,k)=>s+(before.per[k].parts-after.per[k].parts),0);
console.log('\nphantom parts value removed: $'+partsDrop.toFixed(2));

const flips=Object.keys(before.per).filter(k=>before.per[k].status!==after.per[k].status);
console.log('kits whose PRICE verdict changed: '+flips.length);
flips.slice(0,12).forEach(k=>{
  const b=before.per[k],a=after.per[k];
  console.log('   '+k+'  '+b.status.padEnd(12)+'-> '+a.status.padEnd(12)+
              ' parts '+b.parts.toFixed(0)+'->'+a.parts.toFixed(0)+
              '  computed '+b.computed+'->'+a.computed+'  listed '+(b.listed??'-'));
});
const bflips=Object.keys(before.per).filter(k=>String(before.per[k].build)!==String(after.per[k].build));
console.log('\nkits whose BUILDABLE changed: '+bflips.length);
bflips.slice(0,10).forEach(k=>{
  console.log('   '+k+'  build '+before.per[k].build+' -> '+after.per[k].build+
              '   was blocked by: '+String(before.per[k].limited).slice(0,44));
});
