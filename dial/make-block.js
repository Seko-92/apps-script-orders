/**
 * make-block.js — option C's art: the 260×121 brand block for A1:C2.
 *
 * ⭐⭐ THE BLEND IS THE BRIEF. Ground is the sheet's own BRAND.ink #1a1a1a, flat, and the
 *    unlit discs ARE that colour, so they are never drawn. The file has no visible boundary —
 *    it is the band, plus light. Row one reads as one unbroken block and the discs simply
 *    appear out of it.
 * ⚠⚠ rgb565 CANNOT REPRESENT #1a1a1a (0x1a lands on 24 or 27), so the quantised ground came
 *    back #26251e — twelve levels out, which is the pasted-on rectangle again. The palette
 *    entry nearest the ground is snapped to the exact band colour after quantising.
 * ⚠ A TICKER IS THE WRONG MOTION AT 2:1 — measured: only six letters of "MOTOR SERVICE" fit
 *   across 260px. The block gets its own event instead: the mark holds, a diagonal flip wave
 *   crosses it, the stacked wordmark holds, and it flips back. Two states, one motion.
 */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {GIFEncoder,quantize,applyPalette}=require('gifenc');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const D=require('./discs.js');
registerFonts();

const BAND='#1a1a1a';
const W=+(process.env.W||260), H=+(process.env.H||121);
const ST={...D.BASE, pitch:3, lattice:'hex', ground:[BAND],
          offHi:BAND, offLo:BAND, onHi:'#ddd6c6', onLo:'#a49d8c'};
const REST_MS=+(process.env.REST||10000), HOLD_MS=+(process.env.HOLD||2600);
const TMP=path.join(__dirname,'renders','_bk.png');
execFileSync('rsvg-convert',['-w','1200','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);

(async()=>{
const LOGO=await loadImage(TMP);
const roundel=(c,w,h)=>{const d=h*0.78; c.drawImage(LOGO,(w-d)/2,(h-d)/2,d,d);};
const stacked=(c,w,h)=>{
  const lines=['HQ','MOTOR','SERVICE'];
  let size=Math.round(h*0.26);
  while(size>7){ c.font='600 '+size+'px Oswald';
    if(Math.max(...lines.map(l=>c.measureText(l).width))<=w*0.80) break; size--; }
  c.fillStyle='#fff'; c.textAlign='center'; c.textBaseline='middle';
  const lh=size*1.14, y0=h/2-lh;
  lines.forEach((l,i)=>c.fillText(l,w/2,y0+i*lh));
};
/** the field for a composition — cached, since these never change frame to frame */
const fieldOf=(fn)=>D.sampleField(W,H,fn,ST);
const A=fieldOf(roundel), B=fieldOf(stacked);

/** a diagonal flip wave from field a to field b. No churn at the head — the disc grid
 *  quantises the edge anyway, and churn is what makes a file expensive. */
function wave(a,b,t){
  const span=W+H*0.7, head=t*span*1.18-H*0.35;
  return a.map((p,i)=>{
    if(p.meta) return p;
    const q=b[i];
    return {...p, v:(head-(p.x+p.y*0.7))>0 ? q.v : p.v};
  });
}
const cv=createCanvas(W,H), ctx=cv.getContext('2d');
const px=(f)=>{ D.paint(ctx,W,H,f,ST,0,0); return ctx.getImageData(0,0,W,H).data; };

const restPx=px(A), holdPx=px(B);
const both=new Uint8Array(restPx.length*2); both.set(restPx,0); both.set(holdPx,restPx.length);
const pal=quantize(both,6,{format:'rgb565'});
const G=[0x1a,0x1a,0x1a];
let gi=0,gd=Infinity;
pal.forEach((c,i)=>{const d=(c[0]-G[0])**2+(c[1]-G[1])**2+(c[2]-G[2])**2; if(d<gd){gd=d;gi=i;}});
pal[gi]=[...G,...(pal[gi].length>3?[pal[gi][3]]:[])];

const gif=GIFEncoder(); let n=0;
const emit=(data,delay)=>{ gif.writeFrame(applyPalette(data,pal,'rgb565'),W,H,
  {palette:n===0?pal:undefined, delay, repeat:n===0?0:undefined}); n++; };

const FPS=12, WAVE=0.75;
const wf=Math.round(WAVE*FPS);
emit(restPx, REST_MS);                                   // the rest — ONE frame
for(let i=1;i<=wf;i++) emit(px(wave(A,B,i/wf)),Math.round(1000/FPS));
emit(holdPx, HOLD_MS);
for(let i=1;i<=wf;i++) emit(px(wave(B,A,i/wf)),Math.round(1000/FPS));
gif.finish();

const buf=Buffer.from(gif.bytes());
const out=process.env.OUT||'block-blend.gif';
fs.writeFileSync(path.join(__dirname,'renders',out),buf);
try{fs.unlinkSync(TMP);}catch(e){}
console.log(`  ${out}  ${W}x${H}  ${n} frames  ${(buf.length/1024).toFixed(0)} KB  `
  +`loop ${((REST_MS+HOLD_MS+wf*2*1000/FPS)/1000).toFixed(1)}s`);
})();
