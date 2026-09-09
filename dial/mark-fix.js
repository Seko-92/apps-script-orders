/** mark-fix.js — hue-sampled (what shipped) vs class-separated, at true size. */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const D=require('./discs.js'); const L=require('./logomark.js');
const BAND='#1a1a1a';
const ST={...D.BASE,pitch:3,lattice:'hex',ground:[BAND],offHi:BAND,offLo:BAND,
          onHi:'#ddd6c6',onLo:'#a49d8c'};
const W=260,H=133;
(async()=>{
  const tmp=path.join(__dirname,'renders','_mfx.png');
  execFileSync('rsvg-convert',['-w','900','-o',tmp,path.join(__dirname,'logo','fav-google.svg')]);
  const LOGO=await loadImage(tmp);
  // the shipped way: one raster, hue sampled per disc
  // ⚠ the SHIPPED path exactly: one raster, hue voted per disc inside sampleField
  const old=D.sampleField(W,H,(c,w,h)=>{const d=h*0.80;c.drawImage(LOGO,(w-d)/2,(h-d)/2,d,d);},ST);
  const fixed=await L.markField(D,ST,W,H,0.80);
  const acc=(f)=>f.filter(p=>!p.meta&&p.v===2).length;
  console.log('  accent discs — hue sampled: '+acc(old)+'   class separated: '+acc(fixed));
  const PAD=16,LAB=13,G=14;
  const cv=createCanvas(PAD*2+W*2+G,PAD*2+LAB+H), ctx=cv.getContext('2d');
  ctx.fillStyle='#0d0d0d';ctx.fillRect(0,0,cv.width,cv.height);
  ctx.fillStyle='#ffd400';ctx.font='600 9px Oswald';ctx.letterSpacing='1.1px';
  ctx.fillText('WHAT SHIPPED — hue sampled per disc',PAD,PAD+9);
  ctx.fillText('CLASS SEPARATED — the ring keeps its colour',PAD+W+G,PAD+9);
  ctx.letterSpacing='0px';
  D.paint(ctx,W,H,old,ST,PAD,PAD+LAB);
  D.paint(ctx,W,H,fixed,ST,PAD+W+G,PAD+LAB);
  fs.writeFileSync(path.join(__dirname,'renders','mark-fix.png'),cv.toBuffer('image/png'));
  fs.unlinkSync(tmp);
  console.log('  mark-fix.png');
})();
