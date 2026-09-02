/** strip-patterns.js — do the five patterns survive 260x121 (2:1) -> 876x56 (16:1)? */
'use strict';
const fs=require('fs'), path=require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const A=require('./src/ambient'), P=require('./src/patterns'), M=require('./src/board-matrix');
registerFonts();
const S=+(process.env.SCALE||1), BW=260, BH=121, SW=876, SH=56;

function field(w,h,name,ph){
  const {cols,rows}=A.gridSize(w,h);
  const mark=A.markField(cols,rows,[{text:'HQ',size:Math.min(19,rows*0.62),weight:'600'}]);
  const T=P.tickerStrip(createCanvas,cols,rows,'HQ MOTOR SERVICE · HOUSTON · ');
  const fd=(fn)=>{const o=createCanvas(cols,rows),c=o.getContext('2d');
    c.clearRect(0,0,cols,rows);fn(c);const px=c.getImageData(0,0,cols,rows).data,f=[];
    for(let y=0;y<rows;y++){const r=[];for(let x=0;x<cols;x++)r.push(px[(y*cols+x)*4+3]>110);f.push(r);}return f;};
  return {piston:()=>fd(c=>P.piston(c,cols,rows,ph)),ticker:()=>fd(c=>P.ticker(c,cols,rows,ph,null,T)),
          mark:()=>A.composeAmbient('mark',ph,cols,rows,mark),belt:()=>fd(c=>P.belt(c,cols,rows,ph)),
          night:()=>fd(c=>P.night(c,cols,rows,ph)),refresh:()=>P.refresh(cols,rows,ph,mark),
          inlinefour:()=>fd(c=>P.inlineFour(c,cols,rows,ph)),wave:()=>fd(c=>P.wave(c,cols,rows,ph)),
          aisle:()=>fd(c=>P.aisle(c,cols,rows,ph)),sweep:()=>fd(c=>P.sweep(c,cols,rows,ph))}[name]();
}
function tile(ctx,x,y,w,h,name,ph){
  const p=createCanvas(w*S,h*S), pc=p.getContext('2d');
  const g=pc.createLinearGradient(0,0,0,h*S);
  g.addColorStop(0,'#26221c');g.addColorStop(0.14,'#141210');g.addColorStop(1,'#100e0c');
  pc.fillStyle=g;pc.fillRect(0,0,p.width,p.height);
  M.drawMatrix(pc,{scale:S,w,h,field:field(w,h,name,ph)});
  ctx.drawImage(p,x*S,y*S);
}
const NAMES=(process.env.NAMES||'inlinefour,aisle,sweep,wave,refresh,ticker,belt,night').split(',');
const PAD=20,LAB=17,GAP=16;
const W=(PAD*2+BW+12+SW)*S, H=(PAD*2+NAMES.length*(LAB+BH+GAP))*S;
const cv=createCanvas(W,H),ctx=cv.getContext('2d');
ctx.fillStyle='#0b0b0b';ctx.fillRect(0,0,W,H);
let y=PAD;
NAMES.forEach(n=>{
  ctx.fillStyle='#ffd400';ctx.font=`600 ${10*S}px Oswald`;ctx.letterSpacing=`${1.5*S}px`;
  ctx.fillText(n.toUpperCase()+'   —   A1:C2 260x121 (2:1)          vs          D1:H1 876x56 (16:1)',PAD*S,(y+9)*S);
  ctx.letterSpacing='0px'; y+=LAB;
  tile(ctx,PAD,y,BW,BH,n,+(process.env.PH||0.35));
  tile(ctx,PAD+BW+12,y+(BH-SH)/2,SW,SH,n,+(process.env.PH||0.35));
  y+=BH+GAP;
});
fs.writeFileSync(path.join(__dirname,'renders',`${process.env.OUT||'strip-patterns'}${S===1?'':'@'+S+'x'}.png`),cv.toBuffer('image/png'));
console.log('strip-patterns  block 260x121 vs strip 876x56');
