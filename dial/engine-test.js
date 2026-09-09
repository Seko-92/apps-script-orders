/** engine-test.js — an engine is TALL. Does it read on the block (51 disc rows) where it
 *  could not on the strip (26)? And how fast does it need to turn to read as running? */
'use strict';
const fs=require('fs'),path=require('path');
const {createCanvas}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const P=require('./src/patterns'); const D=require('./discs.js');
registerFonts();
const BAND='#1a1a1a';
const ST={...D.BASE,pitch:3,lattice:'hex',ground:[BAND],offHi:BAND,offLo:BAND,
          onHi:'#ddd6c6',onLo:'#a49d8c'};
function field(name,phase,W,H){
  const {pts,cols,rows}=D.lattice(W,H,ST);
  const off=createCanvas(cols,rows),c=off.getContext('2d');
  c.clearRect(0,0,cols,rows); P[name](c,cols,rows,phase);
  const px=c.getImageData(0,0,cols,rows).data;
  const f=pts.map(p=>({...p,v:px[((p.r*cols)+p.c)*4+3]>110?1:0}));
  f.push({cols,rows,meta:true}); return f;
}
const CASES=[
  ['inlineFour on the BLOCK  260x133  (86x51 discs)','inlineFour',260,133],
  ['piston     on the BLOCK  260x133','piston',260,133],
  ['inlineFour on the STRIP  539x68   (179x26 discs) — for comparison','inlineFour',539,68],
];
const PH=[0,0.12,0.25,0.37,0.5,0.62];
const PAD=14,LAB=13,G=8;
let Wm=0,Ht=PAD*2;
for(const [,,w,h] of CASES){ Wm=Math.max(Wm,PH.length*(w+G)-G); Ht+=LAB+h+G+8; }
const cv=createCanvas(Wm+PAD*2,Ht),ctx=cv.getContext('2d');
ctx.fillStyle='#0d0d0d';ctx.fillRect(0,0,cv.width,cv.height);
let y=PAD;
for(const [label,name,W,H] of CASES){
  ctx.fillStyle='#ffd400';ctx.font='600 9px Oswald';ctx.letterSpacing='1.1px';
  ctx.fillText(label,PAD,y+9);ctx.letterSpacing='0px';y+=LAB;
  let x=PAD;
  for(const ph of PH){ D.paint(ctx,W,H,field(name,ph,W,H),ST,x,y); x+=W+G; }
  y+=H+G+8;
}
fs.writeFileSync(path.join(__dirname,'renders','engine-test.png'),cv.toBuffer('image/png'));
console.log('  engine-test.png '+cv.width+'x'+cv.height+' @1:1  — phases 0 · .12 · .25 · .37 · .5 · .62');
