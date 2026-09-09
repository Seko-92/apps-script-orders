/**
 * make-hero.js — the mark is the hero. Bezel motion, 2px pitch, the logo's own words.
 *
 * ⭐⭐ THE MARK NEVER LEAVES. The block holds the roundel for the whole loop; the only motion
 *    is an accent highlight riding once around its ring. Nothing is ever hidden, so it cannot
 *    read as broken — which is the failure mode of every reveal-style idea on a masthead.
 * ⭐ THE BEZEL IS CHEAP BY CONSTRUCTION. It changes only which discs are ACCENT, never which
 *   are lit, so frame to frame the encoder carries a moving band and nothing else.
 * ⚠ BOTH BOARDS AT THE SAME PITCH. At 2px vs 3px the two disc sizes sit 260px apart on one
 *   row and stop reading as one object.
 * ⚠ THE WORDMARK IS "HIGH QUALITY MOTOR SERVICE" — the logo's own words. "HQ MOTOR SERVICE"
 *   is the initials, which is what shipped before by mistake.
 */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {GIFEncoder,quantize,applyPalette}=require('gifenc');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const D=require('./discs.js');
registerFonts();

/* ⚠⚠ 2.5px IS THE FLOOR, AND THE LIMIT IS GEOMETRIC — NOT THE PALETTE. At pitch 2 a disc is
      1.62px across on a 1.73px row spacing, so adjacent discs merge under rasterisation and
      the board stops reading as discs at all: it becomes solid streaks. Measured at 6, 12 and
      20 colours — MORE COLOURS DOES NOT FIX IT, because the problem is that there is no gap
      left to draw. 2.5px gives 61 rows across the mark against 3px's 51, and the discs still
      resolve. Do not go finer. */
const BAND='#1a1a1a', PITCH=+(process.env.PITCH||2.5);
const ST={...D.BASE,pitch:PITCH,lattice:'hex',ground:[BAND],offHi:BAND,offLo:BAND,
          onHi:'#ddd6c6',onLo:'#a49d8c'};
const CUT_FPS=24;
const TMP=path.join(__dirname,'renders','_hh.png');
execFileSync('rsvg-convert',['-w','1400','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);

const ang=(p,w,h)=>{ let a=Math.atan2(p.y-h/2,p.x-w/2)+Math.PI/2;
                     return ((a%(2*Math.PI))+2*Math.PI)%(2*Math.PI)/(2*Math.PI); };

function star(c,cx,cy,r){
  c.beginPath();
  for(let i=0;i<10;i++){ const a=-Math.PI/2+i*Math.PI/5, rr=i%2?r*0.42:r;
    c[i?'lineTo':'moveTo'](cx+Math.cos(a)*rr, cy+Math.sin(a)*rr); }
  c.closePath(); c.fill();
}
const line=(t,hs)=>(c,w,h)=>{
  let z=Math.round(h*(hs||0.72));
  while(z>7){c.font='600 '+z+'px Oswald'; if(c.measureText(t).width<=w*0.92)break; z--;}
  c.fillStyle='#fff';c.textAlign='center';c.textBaseline='middle';c.fillText(t,w/2,h/2+1);
};
const starLine=(t)=>(c,w,h)=>{
  let z=Math.round(h*0.72);
  while(z>7){c.font='600 '+z+'px Oswald'; if(c.measureText(t).width+h*1.05<=w*0.92)break; z--;}
  const tw=c.measureText(t).width, gap=h*0.24, sr=h*0.33;
  const x0=(w-(sr*2+gap+tw))/2;
  c.fillStyle='#fff'; star(c,x0+sr,h/2,sr);
  c.textAlign='left'; c.textBaseline='middle'; c.fillText(t,x0+sr*2+gap,h/2+1);
};

function enc(W,H,frames){
  const cv=createCanvas(W,H), ctx=cv.getContext('2d');
  const px=(f)=>{ D.paint(ctx,W,H,f,ST,0,0); return ctx.getImageData(0,0,W,H).data; };
  const flip=(a,b,t,o)=>{ D.paintFlip(ctx,W,H,a,b,t,ST,0,0,o); return ctx.getImageData(0,0,W,H).data; };
  const key=frames.filter(f=>f.key).map(f=>px(f.field));
  const buf=new Uint8Array(key.reduce((n,s)=>n+s.length,0));
  let at=0; for(const s of key){ buf.set(s,at); at+=s.length; }
  const pal=quantize(buf,6,{format:'rgb565'});
  const G=[0x1a,0x1a,0x1a]; let gi=0,gd=Infinity;
  pal.forEach((c,i)=>{const d=(c[0]-G[0])**2+(c[1]-G[1])**2+(c[2]-G[2])**2; if(d<gd){gd=d;gi=i;}});
  pal[gi]=[...G,...(pal[gi].length>3?[pal[gi][3]]:[])];
  const gif=GIFEncoder(); let n=0, ms=0;
  for(const f of frames){
    const data = f.flip ? flip(f.flip[0],f.flip[1],f.flip[2],f.opt) : px(f.field);
    gif.writeFrame(applyPalette(data,pal,'rgb565'),W,H,
      {palette:n===0?pal:undefined, delay:f.ms, repeat:n===0?0:undefined});
    n++; ms+=f.ms;
  }
  gif.finish();
  return {buf:Buffer.from(gif.bytes()), n, ms};
}

(async()=>{
const LOGO=await loadImage(TMP);

/* ── THE BLOCK ── the mark, and one revolution of light around its ring ─────────── */
{
  const W=260,H=133;
  const mark=D.sampleField(W,H,(c,w,h)=>{const d=h*0.92;c.drawImage(LOGO,(w-d)/2,(h-d)/2,d,d);},ST);
  const bezel=(ph)=>mark.map(p=>{
    if(p.meta||!p.v) return p;
    let d=Math.abs(ang(p,W,H)-ph); d=Math.min(d,1-d);
    // a soft head with a short tail, so it reads as light travelling, not a block sliding
    return (d<0.045 || (ph-ang(p,W,H)+1)%1 < 0.085) ? {...p,v:2} : p;
  });
  const REST=16000, SWEEP=3.0, nf=Math.round(SWEEP*CUT_FPS), step=Math.round(1000/CUT_FPS);
  const frames=[{field:mark, ms:REST, key:true}];
  for(let k=1;k<=nf;k++) frames.push({field:bezel(k/nf), ms:step, key:k===Math.round(nf*0.3)});
  const r=enc(W,H,frames);
  fs.writeFileSync(path.join(__dirname,'renders','banner-v8.gif'),r.buf);
  console.log(`  banner-v8.gif   ${W}x${H}  ${r.n} frames  ${(r.buf.length/1024).toFixed(0)} KB  ${(r.ms/1000).toFixed(1)}s`);
}

/* ── THE STRIP ── the words, and one trip to Houston ────────────────────────────── */
{
  const W=539,H=68;
  const a=D.sampleField(W,H,line('HIGH QUALITY MOTOR SERVICE',0.70),ST);
  const b=D.sampleField(W,H,starLine('HOUSTON, TEXAS'),ST);
  const step=Math.round(1000/CUT_FPS);
  const f1=Math.round(1.6*CUT_FPS), f2=Math.round(1.5*CUT_FPS);
  const frames=[{field:a, ms:13500, key:true}];
  for(let k=1;k<=f1;k++) frames.push({flip:[a,b,k/f1], opt:{shape:'columns',riffle:3}, ms:step});
  frames.push({field:b, ms:2400, key:true});
  for(let k=1;k<=f2;k++) frames.push({flip:[b,a,k/f2], opt:{shape:'columns',riffle:1}, ms:step});
  const r=enc(W,H,frames);
  fs.writeFileSync(path.join(__dirname,'renders','strip-v6.gif'),r.buf);
  console.log(`  strip-v6.gif    ${W}x${H}   ${r.n} frames  ${(r.buf.length/1024).toFixed(0)} KB  ${(r.ms/1000).toFixed(1)}s`);
}
try{fs.unlinkSync(TMP);}catch(e){}
})();
