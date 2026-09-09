/**
 * concepts.js — DIRECTIONS, not a build. Key frames at true 1:1 so a direction can be
 * agreed before any GIF is rendered.
 *
 * ⚠⚠ ONE IDEA IS ALREADY RULED OUT AND IT IS WORTH RECORDING: INVERTING the board (lit
 *    becomes unlit) cannot work here. Unlit discs ARE the band colour and are never drawn,
 *    so an inverted frame lights the whole rectangle — the blend we just spent a round
 *    getting right would break for the length of the effect.
 */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const D=require('./discs.js');
registerFonts();

const BAND='#1a1a1a';
const base=(pitch)=>({...D.BASE,pitch,lattice:'hex',ground:[BAND],offHi:BAND,offLo:BAND,
                      onHi:'#ddd6c6',onLo:'#a49d8c'});
const W=260,H=133, SW=539,SH=68;
const TMP=path.join(__dirname,'renders','_cx.png');
execFileSync('rsvg-convert',['-w','1200','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);

/* ── per-disc geometry helpers, all relative to the mark's own centre ─────────────── */
const ang=(p,w,h)=>{ let a=Math.atan2(p.y-h/2,p.x-w/2)+Math.PI/2; // 0 at 12 o'clock, clockwise
                     return ((a%(2*Math.PI))+2*Math.PI)%(2*Math.PI)/(2*Math.PI); };
const rad=(p,w,h)=>Math.hypot((p.x-w/2)/(h/2),(p.y-h/2)/(h/2));

/** an accent highlight riding around the mark's ring */
const bezel=(f,ph,w,h)=>f.map(p=>{
  if(p.meta||!p.v) return p;
  let d=Math.abs(ang(p,w,h)-ph); d=Math.min(d,1-d);
  return d<0.055 ? {...p,v:2} : p;
});
/** a band of light crossing the board */
const scan=(f,ph,w,h)=>f.map(p=>{
  if(p.meta||!p.v) return p;
  return Math.abs(p.x - ph*w) < w*0.07 ? {...p,v:2} : p;
});
/** the ring drawing itself, clockwise from 12 */
const draw=(f,ph,w,h)=>f.map(p=>{
  if(p.meta) return p;
  return ang(p,w,h) <= ph ? p : {...p,v:0};
});

(async()=>{
const LOGO=await loadImage(TMP);
const roundel=(s)=>(c,w,h)=>{const d=h*(s||0.80); c.drawImage(LOGO,(w-d)/2,(h-d)/2,d,d);};
const line=(t,hs)=>(c,w,h)=>{
  let z=Math.round(h*(hs||0.74));
  while(z>7){c.font='600 '+z+'px Oswald'; if(c.measureText(t).width<=w*0.90)break; z--;}
  c.fillStyle='#fff';c.textAlign='center';c.textBaseline='middle';c.fillText(t,w/2,h/2+1);
};

const rows=[];
/* ── A · how fine should the mark be? ─────────────────────────────────────────── */
for(const p of [3,2.5,2]){
  const st=base(p), f=D.sampleField(W,H,roundel(0.80),st);
  const m=f[f.length-1];
  rows.push({label:'A · PITCH '+p+'px  → '+m.cols+'×'+m.rows+' discs, mark ≈ '+
             Math.round(0.80*m.rows)+' rows', tiles:[{W,H,st,f}]});
}
/* ── B · four ways to move, mark-first ────────────────────────────────────────── */
const st3=base(3), mark=D.sampleField(W,H,roundel(0.80),st3);
rows.push({label:'B1 · BEZEL — an accent highlight rides around the ring. The mark never leaves.',
  tiles:[0,0.15,0.35,0.6,0.85].map(ph=>({W,H,st:st3,f:bezel(mark,ph,W,H)}))});
rows.push({label:'B2 · SCAN — a band of light crosses the board and the mark answers it.',
  tiles:[0.05,0.3,0.5,0.7,0.95].map(ph=>({W,H,st:st3,f:scan(mark,ph,W,H)}))});
rows.push({label:'B3 · DRAW — the mark builds itself clockwise from twelve, then holds.',
  tiles:[0.2,0.45,0.7,0.9,1].map(ph=>({W,H,st:st3,f:draw(mark,ph,W,H)}))});
rows.push({label:'B4 · SCALE — the mark alone at three sizes. Bigger is not automatically better.',
  tiles:[0.72,0.86,0.98].map(s=>({W,H,st:st3,f:D.sampleField(W,H,roundel(s),st3)}))});
/* ── C · the wordmark, which I have had wrong ─────────────────────────────────── */
const sst=base(3);
rows.push({label:'C1 · what I built  —  "HQ MOTOR SERVICE"',
  tiles:[{W:SW,H:SH,st:sst,f:D.sampleField(SW,SH,line('HQ MOTOR SERVICE'),sst)}]});
rows.push({label:'C2 · the logo\'s own words  —  "HIGH QUALITY MOTOR SERVICE"',
  tiles:[{W:SW,H:SH,st:sst,f:D.sampleField(SW,SH,line('HIGH QUALITY MOTOR SERVICE'),sst)}]});
rows.push({label:'C3 · two lines, the way the logo locks up',
  tiles:[{W:SW,H:SH,st:sst,f:D.sampleField(SW,SH,(c,w,h)=>{
    const a='HIGH QUALITY', b='MOTOR SERVICE';
    let z=Math.round(h*0.40);
    while(z>7){c.font='600 '+z+'px Oswald';
      if(Math.max(c.measureText(a).width,c.measureText(b).width)<=w*0.60)break; z--;}
    c.fillStyle='#fff';c.textAlign='center';c.textBaseline='middle';
    c.fillText(a,w/2,h/2-z*0.58); c.fillText(b,w/2,h/2+z*0.58);
  },sst)}]});

const PAD=16,LAB=14,G=10;
let cw=0, ch=PAD*2;
for(const r of rows){ let x=0; for(const t of r.tiles) x+=t.W+G;
  cw=Math.max(cw,x-G); ch+=LAB+Math.max(...r.tiles.map(t=>t.H))+G+6; }
const cv=createCanvas(cw+PAD*2,ch), ctx=cv.getContext('2d');
ctx.fillStyle='#0d0d0d'; ctx.fillRect(0,0,cv.width,cv.height);
let y=PAD;
for(const r of rows){
  ctx.fillStyle='#ffd400'; ctx.font='600 9.5px Oswald'; ctx.letterSpacing='1.05px';
  ctx.fillText(r.label,PAD,y+9); ctx.letterSpacing='0px'; y+=LAB;
  let x=PAD;
  for(const t of r.tiles){ D.paint(ctx,t.W,t.H,t.f,t.st,x,y); x+=t.W+G; }
  y+=Math.max(...r.tiles.map(t=>t.H))+G+6;
}
fs.writeFileSync(path.join(__dirname,'renders','concepts.png'),cv.toBuffer('image/png'));
try{fs.unlinkSync(TMP);}catch(e){}
console.log('  concepts.png '+cv.width+'x'+cv.height+' @1:1');
})();
