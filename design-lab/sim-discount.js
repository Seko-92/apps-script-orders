// Is the spread in DISC% explained by "10% then round to a nice number", or is it drift?
const fs=require('fs'),path=require('path'),vm=require('vm');
const ctx={console,Map,JSON,String,Number,parseInt,parseFloat,isNaN,Math,Array,Object,RegExp};
vm.createContext(ctx);
const reg=fs.readFileSync('../KitRegistry.js','utf8');
vm.runInContext(reg.slice(reg.indexOf('var KIT_BUNDLE = {'), reg.indexOf('// =======================================================================================\n// READ API')), ctx);
const kp=fs.readFileSync('../KitPricing.js','utf8');
vm.runInContext(kp.slice(kp.indexOf('var KIT_PRICING = {'), kp.indexOf('function _buildKitComponentPriceMaps')), ctx);
vm.runInContext(kp.slice(kp.indexOf('function _resolveComponentPrice'), kp.indexOf('function computeKitPriceBySku')), ctx);

const kits=JSON.parse(fs.readFileSync('fixture-kit-registry.json','utf8'));
const px=JSON.parse(fs.readFileSync('fixture-prices.json','utf8'));
const maps={ebay:new Map(Object.entries(px.ebay)),zoho:new Map(Object.entries(px.zoho))};

const rows=[];
kits.forEach(k=>{
  const kit=JSON.parse(JSON.stringify(k)); ctx._kbAnnotateKit(kit);
  const p=ctx.computeKitPrice(kit.components,{maps});
  const listed=ctx._resolveComponentPrice(kit.sku,maps).price;
  const complete=(kit.unparsedLines||[]).length===0 && kit.components.length>0 && p.unpricedComponents.length===0;
  if(!complete||listed==null||p.rawSum<=0) return;
  rows.push({sku:kit.sku,parts:p.rawSum,listed,disc:1-listed/p.rawSum});
});
rows.sort((a,b)=>a.disc-b.disc);
const pct=x=>(x*100).toFixed(1)+'%';
const q=f=>rows[Math.min(rows.length-1,Math.floor(f*rows.length))].disc;
console.log('priced + complete kits: '+rows.length);
console.log('DISC%  min '+pct(rows[0].disc)+'  p10 '+pct(q(.10))+'  p25 '+pct(q(.25))+
            '  median '+pct(q(.50))+'  p75 '+pct(q(.75))+'  p90 '+pct(q(.90))+'  max '+pct(rows[rows.length-1].disc));

console.log('\n── is the discount ever LESS than 10%? (rounding down can only ever increase it)');
const below=rows.filter(r=>r.disc<0.10), at=rows.filter(r=>r.disc>=0.10&&r.disc<=0.16), above=rows.filter(r=>r.disc>0.16);
console.log('   below 10%      : '+below.length+'  ('+(100*below.length/rows.length).toFixed(0)+'%)');
console.log('   10–16%         : '+at.length+'  ('+(100*at.length/rows.length).toFixed(0)+'%)');
console.log('   above 16%      : '+above.length+'  ('+(100*above.length/rows.length).toFixed(0)+'%)');

console.log('\n── how ROUND are the listed prices? (a human picking a nice number)');
const endings={};
rows.forEach(r=>{ const c=Math.round(r.listed)%10; endings[c]=(endings[c]||0)+1; });
console.log('   last digit: '+Object.keys(endings).sort((a,b)=>endings[b]-endings[a])
  .map(d=>d+'→'+endings[d]).join('  '));
const nice=rows.filter(r=>{const v=Math.round(r.listed); return v%5===0 || v%10===9 || v%100===99;});
console.log('   ends in 0/5/9 : '+nice.length+' of '+rows.length+'  ('+(100*nice.length/rows.length).toFixed(0)+'%)');

console.log('\n── if the rule really is 10%, how far off is each kit in DOLLARS?');
const gaps=rows.map(r=>({sku:r.sku,gap:Math.round(r.parts*0.9)-r.listed,parts:r.parts,listed:r.listed}));
gaps.sort((a,b)=>a.gap-b.gap);
const money=gaps.reduce((s,g)=>s+Math.max(0,g.gap),0);
console.log('   kits listed BELOW parts−10% : '+gaps.filter(g=>g.gap>0).length);
console.log('   kits listed ABOVE parts−10% : '+gaps.filter(g=>g.gap<0).length);
console.log('   total $ under the 10% line  : $'+money.toFixed(0));
console.log('   median gap                  : $'+gaps[Math.floor(gaps.length/2)].gap.toFixed(0));

console.log('\n── what tolerance around a FLAT 10% would call each share of the catalog "in line"?');
[0.03,0.05,0.07,0.10].forEach(tol=>{
  const inline=rows.filter(r=>{
    const computed=Math.round(r.parts*0.9);
    return Math.abs(r.listed-computed) <= Math.max(2, tol*computed);
  }).length;
  console.log('   ±'+(tol*100).toFixed(0)+'%  → '+inline+' in line ('+(100*inline/rows.length).toFixed(0)+
              '%),  '+(rows.length-inline)+' flagged');
});

console.log('\n════ WHERE THE MONEY ACTUALLY IS, at a flat 10% rule ════');
const g2=rows.map(r=>({sku:r.sku,parts:r.parts,listed:r.listed,disc:r.disc,
                       gap:Math.round(r.parts*0.9)-r.listed}))
             .filter(x=>x.gap>0).sort((a,b)=>b.gap-a.gap);
const tot=g2.reduce((s,x)=>s+x.gap,0);
let run=0, n80=0;
for(const x of g2){ run+=x.gap; n80++; if(run>=0.8*tot) break; }
console.log('   '+g2.length+' kits sit under the line, $'+tot.toFixed(0)+' total');
console.log('   80% of that money ($'+(0.8*tot).toFixed(0)+') is in just '+n80+' kits');
console.log('\n   worst 10:');
g2.slice(0,10).forEach(x=>console.log('     '+x.sku+'  parts $'+x.parts.toFixed(0).padStart(5)+
  '  listed $'+x.listed.toFixed(0).padStart(5)+'  should be $'+Math.round(x.parts*0.9).toString().padStart(5)+
  '  short $'+x.gap.toFixed(0).padStart(4)+'   disc '+(x.disc*100).toFixed(1)+'%'));

console.log('\n════ TOLERANCE SHAPES — flag count vs money captured ════');
[['±3% (today)',(c)=>Math.max(2,0.03*c)],
 ['±5%',(c)=>Math.max(2,0.05*c)],
 ['$25 or 5%',(c)=>Math.max(25,0.05*c)],
 ['$40 or 6%',(c)=>Math.max(40,0.06*c)],
 ['$50 or 8%',(c)=>Math.max(50,0.08*c)]].forEach(([label,tolFn])=>{
  let flagged=0, money=0;
  rows.forEach(r=>{
    const c=Math.round(r.parts*0.9), d=r.listed-c;
    if(Math.abs(d)>tolFn(c)){ flagged++; if(d<0) money+=-d; }
  });
  console.log('   '+label.padEnd(12)+' flags '+String(flagged).padStart(3)+' kits  ·  captures $'+money.toFixed(0)+' of the $'+tot.toFixed(0));
});
