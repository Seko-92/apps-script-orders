/** direct-motif.js — cohesion test: the DIRECT divider wearing the masthead's own motif.
 *  It is a NUMBER FORMAT change ('"▌  "@' → a disc glyph run), not a rebuild. True 1:1. */
'use strict';
const fs=require('fs'),path=require('path');
const {createCanvas}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render'); registerFonts();
const W=1136,H=34,GAP=12,LAB=15;
const rows=[
 ['NOW — a single block glyph, unrelated to anything else on the sheet','▌  DIRECT',null],
 ['DISCS — the same motif as the masthead, so the two ends of the sheet are one family','DIRECT','disc'],
];
const cv=createCanvas(W,(H+GAP+LAB)*rows.length+10),ctx=cv.getContext('2d');
ctx.fillStyle='#0d0d0d';ctx.fillRect(0,0,cv.width,cv.height);
let y=6;
for(const [label,text,motif] of rows){
  ctx.fillStyle='#ffd400';ctx.font='600 9px Oswald';ctx.letterSpacing='1.1px';
  ctx.fillText(label,4,y+9);ctx.letterSpacing='0px';y+=LAB;
  ctx.fillStyle='#ffd400';ctx.fillRect(0,y,W,H);
  let tx=12;
  if(motif){
    // four discs, lit / lit / lit / unlit — the masthead's material, in the band's own ink
    for(let i=0;i<4;i++){
      const cx=tx+5+i*11, cy=y+H/2;
      ctx.beginPath();ctx.arc(cx,cy+0.5,4.1,0,7);ctx.fillStyle='rgba(0,0,0,0.30)';ctx.fill();
      const g=ctx.createLinearGradient(0,cy-4,0,cy+4);
      if(i<3){g.addColorStop(0,'#1f1d19');g.addColorStop(1,'#0d0c0a');}
      else   {g.addColorStop(0,'#e8c93a');g.addColorStop(1,'#c9a520');}
      ctx.beginPath();ctx.arc(cx,cy,4.1,0,7);ctx.fillStyle=g;ctx.fill();
    }
    tx+=4*11+10;
  }
  ctx.fillStyle='#1a1a1a';ctx.font='600 15px Oswald';ctx.letterSpacing='1px';
  ctx.fillText(text,tx,y+22);
  ctx.font='600 8.5px Oswald';ctx.textAlign='right';
  ctx.fillText('HQMS · DIRECT ORDERS · 7 waiting',W-12,y+21);
  ctx.textAlign='left';ctx.letterSpacing='0px';
  y+=H+GAP;
}
fs.writeFileSync(path.join(__dirname,'renders','direct-motif.png'),cv.toBuffer('image/png'));
console.log('  direct-motif.png '+cv.width+'x'+cv.height+' @1:1');
