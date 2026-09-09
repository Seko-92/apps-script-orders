/** flip-test.js — does a per-disc rotation actually read at 1:1? Frames sampled across one
 *  transition, on the band, at the real strip size. */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render'); const D=require('./discs.js');
registerFonts();
const W=539,H=68,BAND='#1a1a1a';
const ST={...D.BASE,pitch:3,lattice:'hex',ground:[BAND],offHi:BAND,offLo:BAND,
          onHi:'#ddd6c6',onLo:'#a49d8c'};
const TMP=path.join(__dirname,'renders','_ft.png');
execFileSync('rsvg-convert',['-w','1200','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);
(async()=>{
const LOGO=await loadImage(TMP);
const word=(t,s)=>(c,w,h)=>{
  let z=Math.round(h*(s||0.72));
  while(z>7){c.font='600 '+z+'px Oswald'; if(c.measureText(t).width<=w*0.88)break; z--;}
  c.fillStyle='#fff';c.textAlign='center';c.textBaseline='middle';c.fillText(t,w/2,h/2+1);
};
const A=D.sampleField(W,H,word('HQ MOTOR SERVICE'),ST);
const B=D.sampleField(W,H,word('HOUSTON, TEXAS'),ST);
const TS=[0,0.18,0.32,0.44,0.56,0.70,0.86,1];
const PAD=14,LAB=13,G=8;
const cv=createCanvas(W+PAD*2,PAD*2+TS.length*(LAB+H+G));
const ctx=cv.getContext('2d'); ctx.fillStyle='#0d0d0d'; ctx.fillRect(0,0,cv.width,cv.height);
let y=PAD;
for(const t of TS){
  ctx.fillStyle='#ffd400';ctx.font='600 9px Oswald';ctx.letterSpacing='1.1px';
  ctx.fillText('t = '+t.toFixed(2),PAD,y+9);ctx.letterSpacing='0px';y+=LAB;
  D.paintFlip(ctx,W,H,A,B,t,ST,PAD,y);
  y+=H+G;
}
fs.writeFileSync(path.join(__dirname,'renders','flip-test.png'),cv.toBuffer('image/png'));
try{fs.unlinkSync(TMP);}catch(e){}
console.log('  flip-test.png '+cv.width+'x'+cv.height+' @1:1');
})();
