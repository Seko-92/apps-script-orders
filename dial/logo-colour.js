/** logo-colour.js — the mark in mono vs in the logo's own yellow, both sizes. */
'use strict';
const fs=require('fs'), path=require('path'), {execFileSync}=require('child_process');
const { createCanvas, loadImage } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const M=require('./src/board-matrix'); registerFonts();
const S=3, TMP=path.join(__dirname,'renders','_lg.png');

async function svgImg(f,w){ execFileSync('rsvg-convert',['-w',String(w),'-o',TMP,
  path.join(__dirname,'logo',f)]); return await loadImage(TMP); }

/** Threshold alpha into 1, then re-read the SAME pixels and promote the yellow ones to 2.
 *  ⚠ Hue is decided per DISC, not per source pixel — a disc is one object and can only be
 *    one colour, so the sample at its centre is the honest answer. */
function fieldFrom(c, cols, rows, colour){
  const px=c.getImageData(0,0,cols,rows).data, f=[];
  for(let y=0;y<rows;y++){const r=[];
    for(let x=0;x<cols;x++){
      const i=(y*cols+x)*4;
      if(px[i+3]<=90){ r.push(0); continue; }
      if(!colour){ r.push(1); continue; }
      const R=px[i],G=px[i+1],B=px[i+2];
      r.push((R>150 && G>110 && B<120) ? 2 : 1);      // the logo's yellow vs its ink
    } f.push(r);}
  return f;
}
async function blockMark(colour){
  const cols=Math.floor(260/4), rows=Math.floor(121/4);
  const img=await svgImg('fav-google.svg',900);
  const o=createCanvas(cols,rows), c=o.getContext('2d');
  c.clearRect(0,0,cols,rows); const d=rows*0.94;
  c.drawImage(img,(cols-d)/2,(rows-d)/2,d,d);
  return fieldFrom(c,cols,rows,colour);
}
async function stripMark(colour){
  const cols=Math.floor(539/4), rows=Math.floor(56/4);
  const img=await svgImg('fav-google.svg',900);
  const o=createCanvas(cols,rows), c=o.getContext('2d');
  c.clearRect(0,0,cols,rows);
  const d=rows*0.98, gap=rows*0.30, text='MOTOR SERVICE';
  let size=Math.round(rows*0.95);
  while(size>5){c.font='600 '+size+'px Oswald';
    if(d+gap+c.measureText(text).width<=cols*0.92)break; size--;}
  const tw=c.measureText(text).width, x0=(cols-(d+gap+tw))/2;
  c.drawImage(img,x0,(rows-d)/2,d,d);
  c.fillStyle='#fff'; c.textBaseline='middle'; c.textAlign='left';
  c.fillText(text,x0+d+gap,rows/2);
  return fieldFrom(c,cols,rows,colour);
}
function tile(ctx,x,y,w,h,field){
  const p=createCanvas(w*S,h*S),pc=p.getContext('2d');
  const g=pc.createLinearGradient(0,0,0,h*S);
  g.addColorStop(0,'#26221c');g.addColorStop(0.14,'#141210');g.addColorStop(1,'#100e0c');
  pc.fillStyle=g;pc.fillRect(0,0,p.width,p.height);
  M.drawMatrix(pc,{scale:S,w,h,field});ctx.drawImage(p,x*S,y*S);
}
(async()=>{
  const CASES=[
    ['BLOCK · mono',   260,121, await blockMark(false)],
    ['BLOCK · the logo’s own yellow', 260,121, await blockMark(true)],
    ['STRIP · mono',   539,56,  await stripMark(false)],
    ['STRIP · the logo’s own yellow', 539,56,  await stripMark(true)]
  ];
  const PAD=20,LAB=17,GAP=16,Wmax=539;
  const cv=createCanvas((PAD*2+Wmax)*S,(PAD*2+CASES.reduce((a,c)=>a+LAB+c[2]+GAP,0))*S);
  const ctx=cv.getContext('2d');ctx.fillStyle='#0b0b0b';ctx.fillRect(0,0,cv.width,cv.height);
  let y=PAD;
  for(const [lab,W,H,f] of CASES){
    ctx.fillStyle='#ffd400';ctx.font=`600 ${10*S}px Oswald`;ctx.letterSpacing=`${1.4*S}px`;
    ctx.fillText(lab.toUpperCase(),PAD*S,(y+9)*S);ctx.letterSpacing='0px';y+=LAB;
    tile(ctx,PAD,y,W,H,f); y+=H+GAP;
  }
  fs.writeFileSync(path.join(__dirname,'renders','logo-colour.png'),cv.toBuffer('image/png'));
  fs.unlinkSync(TMP); console.log('  logo-colour.png');
})();
