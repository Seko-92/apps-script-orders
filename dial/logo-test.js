/** logo-test.js — the REAL mark as discs, at true aspect, plus a lockup composed for 16:1. */
'use strict';
const fs=require('fs'), path=require('path'), {execFileSync}=require('child_process');
const { createCanvas, loadImage } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const M=require('./src/board-matrix'); registerFonts();
const S=3, TMP=path.join(__dirname,'renders','_lg.png');

/** ⚠⚠ RASTERISE AT THE SVG'S OWN ASPECT. Passing both -w and -h to rsvg-convert FORCES both
 *    and distorts — the first cut squashed a 1:1 roundel into a 2.17:1 oval and it stopped
 *    being the logo. Give it width only and let height follow. */
async function raster(svg, targetW) {
  const vb = /viewBox="([\d.\s-]+)"/.exec(fs.readFileSync(path.join(__dirname,'logo',svg),'utf8'));
  const [,, vw, vh] = vb[1].trim().split(/\s+/).map(Number);
  execFileSync('rsvg-convert', ['-w', String(Math.round(targetW)), '-o', TMP,
                                path.join(__dirname,'logo',svg)]);
  const img = await loadImage(TMP);
  return { img, ar: vw / vh };
}
/** Fit an image inside cols x rows without distorting, at an optional scale, threshold alpha. */
async function logoField(svg, cols, rows, cut, fill, dx) {
  const { img, ar } = await raster(svg, 900);
  const off = createCanvas(cols, rows), c = off.getContext('2d');
  c.clearRect(0,0,cols,rows);
  const box = cols/rows;
  let w,h; if (ar > box) { w = cols*fill; h = w/ar; } else { h = rows*fill; w = h*ar; }
  c.drawImage(img, (cols-w)/2 + (dx||0), (rows-h)/2, w, h);
  const px=c.getImageData(0,0,cols,rows).data, f=[];
  for(let y=0;y<rows;y++){const r=[];for(let x=0;x<cols;x++)r.push(px[(y*cols+x)*4+3]>cut);f.push(r);}
  return f;
}
/** ⭐ A LOCKUP COMPOSED FOR 16:1 — the roundel at true circle on the left, the name set beside
 *  it on ONE line. Squashing a 3:1 lockup into 16:1 destroys it; rebuilding it for the shape
 *  keeps both halves at their own best size. */
async function stripLockup(cols, rows, cut) {
  const { img, ar } = await raster('fav-google.svg', 900);
  const off = createCanvas(cols, rows), c = off.getContext('2d');
  c.clearRect(0,0,cols,rows);
  const d = rows * 0.98, size = Math.round(rows * 0.62);
  c.font = '600 ' + size + 'px Oswald';
  const text = 'HIGH QUALITY MOTOR SERVICE';
  const tw = c.measureText(text).width, gap = rows * 0.34;
  const total = d*ar + gap + tw, x0 = (cols - total)/2;
  c.drawImage(img, x0, (rows-d)/2, d*ar, d);
  c.fillStyle = '#fff'; c.textBaseline = 'middle';
  c.fillText(text, x0 + d*ar + gap, rows/2);
  const px=c.getImageData(0,0,cols,rows).data, f=[];
  for(let y=0;y<rows;y++){const r=[];for(let x=0;x<cols;x++)r.push(px[(y*cols+x)*4+3]>cut);f.push(r);}
  return { field:f, note:'roundel ' + Math.round(d) + ' discs + text @' + size };
}
function tile(ctx,x,y,w,h,field){
  const p=createCanvas(w*S,h*S), pc=p.getContext('2d');
  const g=pc.createLinearGradient(0,0,0,h*S);
  g.addColorStop(0,'#26221c');g.addColorStop(0.14,'#141210');g.addColorStop(1,'#100e0c');
  pc.fillStyle=g;pc.fillRect(0,0,p.width,p.height);
  M.drawMatrix(pc,{scale:S,w,h,field}); ctx.drawImage(p,x*S,y*S);
}
(async()=>{
  const P=4, bc=Math.floor(260/P), br=Math.floor(121/P), sc=Math.floor(539/P), sr=Math.floor(56/P);
  const lock = await stripLockup(sc, sr, 90);
  const CASES=[
    ['BLOCK · roundel, true circle',       260,121, await logoField('fav-google.svg',bc,br,90,0.94)],
    ['BLOCK · wide lockup, true aspect',   260,121, await logoField('new-logo.svg',bc,br,90,0.96)],
    ['BLOCK · full lockup, true aspect',   260,121, await logoField('new-logo-google.svg',bc,br,90,0.96)],
    ['STRIP · roundel alone (too small)',  539,56,  await logoField('fav-google.svg',sc,sr,90,0.96)],
    ['STRIP · wide lockup, true aspect',   539,56,  await logoField('new-logo.svg',sc,sr,90,0.96)],
    ['STRIP · COMPOSED for 16:1 — '+lock.note, 539,56, lock.field]
  ];
  const PAD=20,LAB=17,GAP=16, Wmax=Math.max(...CASES.map(c=>c[1]));
  const cv=createCanvas((PAD*2+Wmax)*S,(PAD*2+CASES.reduce((a,c)=>a+LAB+c[2]+GAP,0))*S);
  const ctx=cv.getContext('2d'); ctx.fillStyle='#0b0b0b'; ctx.fillRect(0,0,cv.width,cv.height);
  let y=PAD;
  for (const [lab,W,H,field] of CASES){
    ctx.fillStyle='#ffd400';ctx.font=`600 ${10*S}px Oswald`;ctx.letterSpacing=`${1.4*S}px`;
    ctx.fillText(lab.toUpperCase(),PAD*S,(y+9)*S); ctx.letterSpacing='0px'; y+=LAB;
    tile(ctx,PAD,y,W,H,field); y+=H+GAP;
  }
  fs.writeFileSync(path.join(__dirname,'renders','logo-test.png'),cv.toBuffer('image/png'));
  fs.unlinkSync(TMP); console.log('  logo-test.png  ·  '+lock.note);
})();
