/** diag-sampler.js — the SHIPPED path (compose at grid resolution, 1 pixel per disc) vs
 *  AREA SAMPLING (compose at full pixel size, average under each disc). Same content, 1:1. */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const M=require('./src/board-matrix'); const D=require('./discs.js');
registerFonts();
const W=539,H=56,TMP=path.join(__dirname,'renders','_mk2.png');
execFileSync('rsvg-convert',['-w','900','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);

(async()=>{
const img=await loadImage(TMP);

/* ---- EXACTLY the shipped brandMark(): composes into a cols x rows canvas ---- */
function shippedMark(cols,rows){
  const off=createCanvas(cols,rows),c=off.getContext('2d');
  c.clearRect(0,0,cols,rows);
  const d=rows*0.98, text='MOTOR SERVICE', gap=rows*0.30;
  let size=Math.round(rows*0.95);
  while(size>5){c.font='600 '+size+'px Oswald';
    if(d+gap+c.measureText(text).width<=cols*0.92)break;size--;}
  const x0=(cols-(d+gap+c.measureText(text).width))/2;
  c.drawImage(img,x0,(rows-d)/2,d,d);
  c.fillStyle='#fff';c.textBaseline='middle';c.textAlign='left';
  c.fillText(text,x0+d+gap,rows/2);
  const px=c.getImageData(0,0,cols,rows).data,f=[];
  for(let y=0;y<rows;y++){const r=[];for(let x=0;x<cols;x++){const i=(y*cols+x)*4;
    if(px[i+3]<=90){r.push(0);continue;}
    r.push((px[i]>150&&px[i+1]>110&&px[i+2]<120)?2:1);}f.push(r);}
  return {f,size};
}
/* ⚠ drawMatrix centres on the canvas ORIGIN and ignores any offset, so the shipped path
   has to be rendered to its own canvas and composited. */
function shippedPaint(ctx,pitch,oy){
  const cols=Math.floor(W/pitch),rows=Math.floor(H/pitch);
  const {f,size}=shippedMark(cols,rows);
  const t=createCanvas(W,H), tc=t.getContext('2d');
  const g=tc.createLinearGradient(0,0,0,H);
  g.addColorStop(0,'#26221c');g.addColorStop(0.14,'#141210');g.addColorStop(1,'#100e0c');
  tc.fillStyle=g;tc.fillRect(0,0,W,H);
  M.drawMatrix(tc,{scale:1,w:W,h:H,field:f,pitch});
  ctx.drawImage(t,PAD,oy);
  return {cols,rows,size};
}

/* the area-sampled lockup, drawn at real pixel size */
const markStrip=(c,w,h)=>{
  const d=h*0.94,gap=h*0.26,text='MOTOR SERVICE';
  let size=Math.round(h*0.74);
  while(size>6){c.font='600 '+size+'px Oswald';
    if(d+gap+c.measureText(text).width<=w*0.92)break;size--;}
  const x0=(w-(d+gap+c.measureText(text).width))/2;
  c.drawImage(img,x0,(h-d)/2,d,d);
  c.fillStyle='#fff';c.textBaseline='middle';c.textAlign='left';
  c.fillText(text,x0+d+gap,h/2+1);
};

const PAD=20,LAB=16,GAP=13;
const ROWS=[
 ['① SHIPPED — grid-resolution compose · 1px per disc · 4px square', 'shipped',4],
 ['② SHIPPED renderer at 3px pitch (pitch alone)',                   'shipped',3],
 ['③ AREA SAMPLED — full-res compose · 4px square · off discs still hidden','area',
   {pitch:4}],
 ['④ AREA SAMPLED · 3px hex · off discs visible · quiet on-tone',   'area',
   {pitch:3,lattice:'hex',ground:['#1a1a1a'],offHi:'#2a2723',offLo:'#171412',
    onHi:'#ddd6c6',onLo:'#a49d8c'}],
];
const cv=createCanvas(W+PAD*2,PAD*2+ROWS.length*(LAB+H+GAP));
const ctx=cv.getContext('2d');
ctx.fillStyle='#0d0d0d';ctx.fillRect(0,0,cv.width,cv.height);
let y=PAD;
for(const [label,kind,arg] of ROWS){
  let note='';
  if(kind==='shipped'){ const r=shippedPaint(ctx,arg,y+LAB);
    note='  ('+r.cols+'×'+r.rows+', text @'+r.size+'px in a '+r.rows+'-row canvas)'; }
  else { const st={...D.BASE,...arg};
    const f=D.sampleField(W,H,markStrip,st); const m=f[f.length-1];
    D.paint(ctx,W,H,f,st,PAD,y+LAB); note='  ('+m.cols+'×'+m.rows+')'; }
  ctx.fillStyle='#ffd400';ctx.font='600 9px Oswald';ctx.letterSpacing='1.1px';
  ctx.fillText(label+note,PAD,y+9);ctx.letterSpacing='0px';
  y+=LAB+H+GAP;
}
fs.writeFileSync(path.join(__dirname,'renders','diag-sampler.png'),cv.toBuffer('image/png'));
try{fs.unlinkSync(TMP);}catch(e){}
console.log('  diag-sampler.png '+cv.width+'x'+cv.height+' @1:1');
})();
