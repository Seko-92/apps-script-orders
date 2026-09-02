/** motion-sheet.js — the new abstract patterns across FOUR phases, so motion is visible. */
'use strict';
const fs=require('fs'), path=require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const A=require('./src/ambient'), P=require('./src/patterns'), M=require('./src/board-matrix');
registerFonts();
const S=+(process.env.SCALE||1);
const W=+(process.env.W||539), H=+(process.env.H||56);
const NAMES=(process.env.NAMES||'pendulum,moire,liquid,ripple').split(',');
const PHASES=[0.05,0.30,0.55,0.80];

function field(w,h,name,ph){
  const {cols,rows}=A.gridSize(w,h);
  const mark=A.markField(cols,rows,[{text:'HQ',size:Math.min(19,rows*0.62),weight:'600'}]);
  const T=P.tickerStrip(createCanvas,cols,rows,'HQ MOTOR SERVICE · HOUSTON · ');
  const fd=fn=>{const o=createCanvas(cols,rows),c=o.getContext('2d');
    c.clearRect(0,0,cols,rows);fn(c);const px=c.getImageData(0,0,cols,rows).data,f=[];
    for(let y=0;y<rows;y++){const r=[];for(let x=0;x<cols;x++)r.push(px[(y*cols+x)*4+3]>110);f.push(r);}return f;};
  return {
    pendulum:()=>fd(c=>P.pendulum(c,cols,rows,ph)),
    moire:   ()=>P.moire(cols,rows,ph),
    liquid:  ()=>P.liquid(cols,rows,ph),
    ripple:  ()=>P.ripple(cols,rows,ph),
    inlinefour:()=>fd(c=>P.inlineFour(c,cols,rows,ph)),
    wave:()=>fd(c=>P.wave(c,cols,rows,ph)), aisle:()=>fd(c=>P.aisle(c,cols,rows,ph)),
    sweep:()=>fd(c=>P.sweep(c,cols,rows,ph)), belt:()=>fd(c=>P.belt(c,cols,rows,ph)),
    night:()=>fd(c=>P.night(c,cols,rows,ph)), ticker:()=>fd(c=>P.ticker(c,cols,rows,ph,null,T)),
    mark:()=>A.composeAmbient('mark',ph,cols,rows,mark), refresh:()=>P.refresh(cols,rows,ph,mark)
  }[name]();
}
function tile(ctx,x,y,name,ph){
  const p=createCanvas(W*S,H*S), pc=p.getContext('2d');
  const g=pc.createLinearGradient(0,0,0,H*S);
  g.addColorStop(0,'#26221c');g.addColorStop(0.14,'#141210');g.addColorStop(1,'#100e0c');
  pc.fillStyle=g;pc.fillRect(0,0,p.width,p.height);
  M.drawMatrix(pc,{scale:S,w:W,h:H,field:field(W,H,name,ph)});
  ctx.drawImage(p,x*S,y*S);
}
const PAD=20,LAB=17,GAP=8,BLK=22;
const cv=createCanvas((PAD*2+W)*S,(PAD*2+NAMES.length*(LAB+PHASES.length*(H+GAP)+BLK))*S);
const ctx=cv.getContext('2d');
ctx.fillStyle='#0b0b0b';ctx.fillRect(0,0,cv.width,cv.height);
let y=PAD;
NAMES.forEach(n=>{
  ctx.fillStyle='#ffd400';ctx.font=`600 ${10*S}px Oswald`;ctx.letterSpacing=`${1.5*S}px`;
  ctx.fillText(n.toUpperCase()+'   —   '+W+'x'+H+'   phases .05 / .30 / .55 / .80',PAD*S,(y+9)*S);
  ctx.letterSpacing='0px'; y+=LAB;
  PHASES.forEach(ph=>{ tile(ctx,PAD,y,n,ph); y+=H+GAP; });
  y+=BLK;
});
fs.writeFileSync(path.join(__dirname,'renders',(process.env.OUT||'motion-sheet')+'.png'),cv.toBuffer('image/png'));
console.log('  '+(process.env.OUT||'motion-sheet')+'.png  '+W+'x'+H+' x'+PHASES.length+' phases');
