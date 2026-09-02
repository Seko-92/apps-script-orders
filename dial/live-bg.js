/** live-bg.js — row 1+2 WITHOUT the loop, so an artifact can overlay the real GIFs on top.
 *  Measured widths, 2026-09-02: A103 B70 C87 | D232 E307 | F130 G100 H107. */
'use strict';
const fs=require('fs'), path=require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const { fmtMins, fmtClock } = require('./src/draw');
registerFonts();
const COL={A:103,B:70,C:87,D:232,E:307,F:130,G:100,H:107};
const DIALW=260, DH=876, TOTAL=DIALW+DH, R1=56, R2=65;
const BAND='#1a1a1a', CREAM='#fff8e7', QUIET='#e8e8e8';
const S=+(process.env.SCALE||2), OPT=(process.env.OPT||'c2');
const clock12=t=>{const m=(+String(t).slice(0,-2))*60+(+String(t).slice(-2));
  return fmtClock(m)+(m<720?' AM':' PM');};
const Q=process.env.REST
  ? {s:'rest',t:'2157',eb:0,di:0,y:8}
  : {s:'busy',t:'1414',eb:9,di:3,y:3};
const head = Q.s==='rest' ? ['the floor is asleep'] : ['picking','eBay '+Q.eb+' · Direct '+Q.di];
const lamp = Q.s==='rest' ? '#9aa3ad' : '#7ec98a';
const word = Q.s==='rest' ? 'RESTING' : 'ALIVE';

const cv=createCanvas(TOTAL*S,(R1+R2)*S), ctx=cv.getContext('2d');
const s=v=>v*S;
ctx.fillStyle=BAND; ctx.fillRect(0,0,s(TOTAL),s(R1));
ctx.fillStyle=CREAM; ctx.fillRect(s(DIALW),s(R1),s(DH),s(R2));

if(OPT!=='b'){
  const fx=DIALW+COL.D+COL.E;                        // 799 — where the loop stops
  if(OPT==='c2'){
    const hx=fx+COL.F+COL.G;                         // 1029 — H1
    ctx.fillStyle=lamp; ctx.beginPath(); ctx.arc(s(fx+13),s(28),s(4),0,7); ctx.fill();
    ctx.fillStyle=QUIET; ctx.font=`400 ${s(11.5)}px Oswald`;
    ctx.fillText(head.join(' · '), s(fx+23), s(32));
    ctx.strokeStyle='#33302b'; ctx.lineWidth=Math.max(1,s(1));
    ctx.beginPath(); ctx.moveTo(s(hx),s(8)); ctx.lineTo(s(hx),s(R1-8)); ctx.stroke();
    ctx.fillStyle=lamp; ctx.font=`500 ${s(9)}px Oswald`; ctx.letterSpacing=`${s(1.1)}px`;
    ctx.fillText(word, s(hx+11), s(24)); ctx.letterSpacing='0px';
    ctx.fillStyle=QUIET; ctx.font=`400 ${s(10.5)}px Oswald`;
    ctx.fillText(clock12(Q.t), s(hx+11), s(40));
  } else {                                            // c1 — one merged readout
    ctx.fillStyle=lamp; ctx.beginPath(); ctx.arc(s(fx+14),s(21),s(4),0,7); ctx.fill();
    ctx.fillStyle=QUIET; ctx.font=`400 ${s(11.5)}px Oswald`;
    ctx.fillText(head.join(' · '), s(fx+25), s(25));
    ctx.fillStyle='#9a9daa'; ctx.font=`400 ${s(10)}px Oswald`;
    ctx.fillText(word+' · '+clock12(Q.t)+' · '+fmtMins(Q.y)+' ago', s(fx+25), s(42));
  }
}
// row 2 — eBay label + the two Pick ID cells, untouched by every option
ctx.fillStyle='#111'; ctx.font=`700 ${s(26)}px "Noto Sans"`;
let lx=DIALW+18;
for(const [ch,col] of [['e','#e53238'],['b','#0064d2'],['a','#f5af02'],['y','#86b817']]){
  ctx.fillStyle=col; ctx.fillText(ch,s(lx),s(R1+44)); lx+=ctx.measureText(ch).width/S; }
const px=DIALW+COL.D+COL.E;
for(const [x,w,lab,val] of [[px,COL.F+COL.G,'PICK ID · SHIPPING','Shipping - Yassin 1'],
                            [px+COL.F+COL.G,COL.H,'ADJUSTMENT','AShamma 2']]){
  ctx.strokeStyle='#e8dfc8'; ctx.lineWidth=Math.max(1,s(1));
  ctx.beginPath(); ctx.moveTo(s(x),s(R1)); ctx.lineTo(s(x),s(R1+R2)); ctx.stroke();
  ctx.fillStyle='#9a9280'; ctx.font=`600 ${s(7.5)}px Oswald`; ctx.letterSpacing=`${s(1.2)}px`;
  ctx.fillText(lab,s(x+10),s(R1+22)); ctx.letterSpacing='0px';
  ctx.fillStyle='#1a1a1a'; ctx.font=`400 ${s(11)}px Oswald`; ctx.fillText(val,s(x+10),s(R1+42));
}
const out=`bg-${OPT}${process.env.REST?'-rest':''}.png`;
fs.writeFileSync(path.join(__dirname,'renders',out), cv.toBuffer('image/png'));
console.log('  '+out+'  '+TOTAL+'x'+(R1+R2)+' @'+S+'x');
