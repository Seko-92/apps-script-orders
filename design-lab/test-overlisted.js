// =====================================================================================
// test-overlisted.js — the oversell verdicts: a MANUAL kit advertised beyond what its
// components can assemble. Pure functions from the REAL KitHealth.js.
//
// Prove against HEAD:  mkdir -p /tmp/head2 && git show HEAD:KitHealth.js > /tmp/head2/KitHealth.js
//                      KIT_SRC=/tmp/head2 node test-overlisted.js
// =====================================================================================
const fs=require('fs'), path=require('path'), vm=require('vm');
const SRC=process.env.KIT_SRC||path.join(__dirname,'..');
const ctx={console,Map,JSON,String,Number,parseInt,parseFloat,isNaN,Math,Array,Object,RegExp};
vm.createContext(ctx);
const src=fs.readFileSync(path.join(SRC,'KitHealth.js'),'utf8');
const a=src.indexOf('var KIT_HEALTH = {'), b=src.indexOf('// =======================================================================================\n// SETUP');
vm.runInContext(src.slice(a,b), ctx);
const S=ctx.KIT_HEALTH.stock;
// ⚠ FAIL SOFT. On an older KitHealth.js these simply don't exist, and a throw
// here would abort the run before the later sections could report — the point of
// a before/after proof is seeing WHICH parts changed and which did not.
const st   = ctx._kitStockStatus || function(){ return '(no _kitStockStatus)'; };
const risk = ctx._kitAtRisk      || function(){ return '(no _kitAtRisk)'; };

let pass=0, fail=0;
const ok=(n,c,x)=>{ c?pass++:(fail++,console.log('  ✗ '+n+(x?'  → '+String(x):''))); };

console.log('\n── a MANUAL kit is a promise, not a box');
ok("listed 1, can build 0 → CAN'T BUILD", st(1,0,'MANUAL',1)===S.CANT_BUILD, st(1,0,'MANUAL',1));
ok('listed 5, can build 1 → OVER-LISTED', st(5,1,'MANUAL',5)===S.OVER_LISTED, st(5,1,'MANUAL',5));
ok('listed 2, can build 2 → not flagged',  st(2,2,'MANUAL',2)===S.STOCK_BUILD);
ok('listed 2, can build 9 → not flagged',  st(2,9,'MANUAL',2)===S.STOCK_BUILD);
ok('listed 0, can build 0 → plain OOS',    st(0,0,'MANUAL',0)===S.OOS);
ok('listed 0, can build 4 → BUILD-ONLY',   st(0,4,'MANUAL',0)===S.BUILD_ONLY);

console.log('\n── a READY kit is a real box on a shelf — never an oversell');
ok('READY listed 8, can build 0 → IN STOCK', st(8,0,'READY',8)===S.IN_STOCK, st(8,0,'READY',8));
ok('READY listed 9, can build 4 → STOCK+BUILD', st(9,4,'READY',9)===S.STOCK_BUILD);
ok('READY is never CAN\'T BUILD', st(5,0,'READY',5)!==S.CANT_BUILD);

console.log('\n── untrustable buildability stays quiet');
ok('⚠ buildable → ⚠, never an oversell claim', st(3,'⚠','MANUAL',3)===S.UNKNOWN);
ok('...and AT RISK is blank', risk(3,'⚠','MANUAL')==='');

console.log('\n── advertised is what a buyer can order, not what Zoho happens to say');
ok('eBay 0 / Zoho 1 → advertised 1 still flags', st(1,0,'MANUAL',1)===S.CANT_BUILD);
ok('advertised beats a lower kitQty', st(0,0,'MANUAL',3)===S.CANT_BUILD, st(0,0,'MANUAL',3));
ok('missing advertised falls back to kitQty', st(2,0,'MANUAL')===S.CANT_BUILD);

console.log('\n── AT RISK is the size of the exposure');
ok('listed 5, build 1 → 4', risk(5,1,'MANUAL')===4, risk(5,1,'MANUAL'));
ok('listed 1, build 0 → 1', risk(1,0,'MANUAL')===1);
ok('covered → blank, so any ink is a finding', risk(2,5,'MANUAL')==='');
ok('exactly covered → blank', risk(3,3,'MANUAL')==='');
ok('READY → blank whatever the gap', risk(8,0,'READY')==='');

console.log('\n── the real kits from the floor');
ok('159093 (listed 1, piston short) → CAN\'T BUILD', st(1,0,'MANUAL',1)===S.CANT_BUILD);
ok('214631 (listed 5, build 0) → 5 units at risk', risk(5,0,'MANUAL')===5);
ok('210329 (listed 8, build 4) → 4 units at risk', risk(8,4,'MANUAL')===4);

console.log('\n── schema');
ok('AT RISK column exists', ctx.KIT_HEALTH.cols.AT_RISK===17, ctx.KIT_HEALTH.cols.AT_RISK);
ok('sheet is 18 wide', ctx.KIT_HEALTH.dataWidth===18, ctx.KIT_HEALTH.dataWidth);
ok('headers match the width', ctx.KIT_HEALTH.headers.length===ctx.KIT_HEALTH.dataWidth);
ok('LAST_CHECKED moved to the end', ctx.KIT_HEALTH.cols.LAST_CHECKED===18);

console.log('\n'+(fail===0?'✅ ':'❌ ')+pass+' passed, '+fail+' failed'+(process.env.KIT_SRC?'   [HEAD]':''));
process.exit(fail===0?0:1);
