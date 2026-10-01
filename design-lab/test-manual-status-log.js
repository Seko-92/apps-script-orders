const fs=require('fs'),vm=require('vm');
const R=process.argv[2];
let logged=[], cells={};
function mk(){ // rows 4..6: SKU,QTY,LOC,SO,NOTE,STATUS
  cells={4:['165834',1,'A-2','23-15215-72784','','PREPARING'],5:['1',1,'B','X-1','','PENDING'],6:['2',1,'C','X-2','','SHIPPED']};}
const sheet={getName:()=> 'All orders',getLastRow:()=>6,getMaxRows:()=>100,
 getRange(r,c,n,w){n=n||1;w=w||1;return{getValues:()=>{let o=[];for(let i=0;i<n;i++){let row=cells[r+i]||[];o.push(row.slice(c-1,c-1+w));}return o;},
  setValue(v){cells[r][c-1]=v;},setValues(v){v.forEach((x,i)=>cells[r+i][c-1]=x[0]);},getRow:()=>r,getHeight:()=>n,getWidth:()=>w,getSheet:()=>sheet,getColumn:()=>c};}};
const ctx={console,Object,String,parseInt,JSON,Math,
 LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
 SpreadsheetApp:{openById:()=>({getSheetByName:()=>sheet}),flush(){}},
 SPREADSHEET_ID:'x',MAIN_SHEET_NAME:'All orders',
 logActivityBatch:(e)=>logged.push(...e.map(x=>x[0]+'|'+x[4]+'|'+x[5])),
 _dashBustTickCache(){},syncStatusToTelegram(){},sortTableByStatusAndLocation(){},updateOrderStatsInSheet(){},
 _kitParentFollowUp(){},holdNoteHasHold:()=>false};
vm.createContext(ctx);
for(const f of ['Schema.js','StatusService.js','OrderService.js']) vm.runInContext(fs.readFileSync(R+'/'+f,'utf8'),ctx,{filename:f});
let pass=0,fail=0;const ok=(n,c,g)=>{c?pass++:(fail++,console.log('FAIL',n,g))};
mk(); ctx.handleManualStatusChange({range:sheet.getRange(4,6),oldValue:'PENDING'});
ok('manual PENDING->PREPARING logs',logged.length===1&&logged[0]==='PREPARING|manual-edit|from PENDING',logged);
logged=[];mk(); ctx.handleManualStatusChange({range:sheet.getRange(4,6),oldValue:'PREPARING'});
ok('same value = no log',logged.length===0,logged);
logged=[];mk(); cells[5][5]='PREPARING'; ctx.handleManualStatusChange({range:sheet.getRange(4,6,2,1)});
ok('paste of 2 rows logs both',logged.length===2&&logged.every(x=>x.endsWith('manual edit (paste)')),logged);
logged=[];mk(); ctx.updateOrderStatus('X-1','PREPARING',{source:'board'});
ok('board path unchanged',logged.length===1&&logged[0]==='PREPARING|board|from PENDING',logged);
logged=[];mk(); ctx.updateOrderStatus('23-15215-72784','PREPARING',{source:'board'});
ok('board no-op still silent',logged.length===0,logged);
console.log(pass+' pass, '+fail+' fail');
