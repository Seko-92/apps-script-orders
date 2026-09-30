const fs=require('fs'),vm=require('vm');
const SRC=process.env.SRC||require('path').join(__dirname,'..');
// sheet model: rows 1..20, colA and colD; DIRECT divider at 10, header 11
const colA={}, colD={}, fmt={};
for(let r=4;r<=9;r++){colA[r]='1'+r; colD[r]= r<6?'24-1':'24-'+r; fmt[r]='@';}
colA[10]='DIRECT'; colA[11]='◈ SKU'; colD[11]='SALES ORDER'; fmt[11]='"1️⃣ "@'; fmt[10]='@';
for(let r=12;r<=16;r++){colA[r]='2'+r; colD[r]= r<15?'SO-1':'SO-'+r; fmt[r]='@';}
const last=16,max=20;
function range(r,c,n){ return {
  getValues:()=>Array.from({length:n},(_,i)=>[c===1?(colA[r+i]||''):(colD[r+i]||'')]),
  getNumberFormats:()=>Array.from({length:n},(_,i)=>[fmt[r+i]||'@']),
  setNumberFormats:(a)=>a.forEach((x,i)=>fmt[r+i]=x[0]),
  getFontSizes:()=>Array.from({length:n},()=>[10]), setFontSizes:()=>{},
  setBorder:()=>{}, };}
const sheet={getLastRow:()=>last,getMaxRows:()=>max,getRange:(r,c,n)=>range(r,c,n||1),
  getConditionalFormatRules:()=>[],setConditionalFormatRules:()=>{}};
const ctx={console,SpreadsheetApp:{openById:()=>({getSheetByName:()=>sheet}),flush:()=>{}},SPREADSHEET_ID:'x',MAIN_SHEET_NAME:'All orders'};
vm.createContext(ctx);
for(const f of ['Config.js','Schema.js','Helpers.js','RowManagement.js']) vm.runInContext(fs.readFileSync(SRC+'/'+f,'utf8'),ctx,{filename:f});
ctx._paintDirectOrderDividers=()=>{};
ctx.setupDuplicateSalesOrderHighlighting();
const ok=(n,c,g)=>console.log((c?'PASS ':'FAIL ')+n+(c?'':' → got '+JSON.stringify(g)));
ok('DIRECT header badge cleared',fmt[11]==='@',fmt[11]);
ok('divider row untouched',fmt[10]==='@',fmt[10]);
ok('eBay group still badged',/^"1️⃣ "@$/.test(fmt[4]),fmt[4]);
ok('DIRECT group still badged',/"@$/.test(fmt[12])&&fmt[12]!=='@',fmt[12]);
ok('DIRECT single row plain',fmt[15]==='@',fmt[15]);
