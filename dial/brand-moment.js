/** brand-moment.js — the converge → mark → disperse sequence, as a filmstrip. */
'use strict';
const fs=require('fs'), path=require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const A=require('./src/ambient'), P=require('./src/patterns'), M=require('./src/board-matrix');
registerFonts();
const S=+(process.env.SCALE||2), W=+(process.env.W||539), H=+(process.env.H||56);
const { cols, rows } = A.gridSize(W,H);

const {execFileSync}=require('child_process');
const {loadImage}=require('@napi-rs/canvas');
async function brandMark(cols,rows){
  const tmp=path.join(__dirname,'renders','_bm.png');
  execFileSync('rsvg-convert',['-w','900','-o',tmp,path.join(__dirname,'logo','fav-google.svg')]);
  const img=await loadImage(tmp);
  const o=createCanvas(cols,rows), c=o.getContext('2d'); c.clearRect(0,0,cols,rows);
  const wide=cols/rows>6, d=rows*(wide?0.98:0.94); let note;
  if(wide){ const text='MOTOR SERVICE', gap=rows*0.30; let size=Math.round(rows*0.95);
    while(size>5){c.font='600 '+size+'px Oswald';
      if(d+gap+c.measureText(text).width<=cols*0.92)break; size--;}
    const x0=(cols-(d+gap+c.measureText(text).width))/2;
    c.drawImage(img,x0,(rows-d)/2,d,d);
    c.fillStyle='#fff';c.textBaseline='middle';c.textAlign='left';
    c.fillText(text,x0+d+gap,rows/2); note='roundel + MOTOR SERVICE @'+size;
  } else { c.drawImage(img,(cols-d)/2,(rows-d)/2,d,d); note='roundel '+Math.round(d)+' discs'; }
  const px=c.getImageData(0,0,cols,rows).data,f=[];
  for(let y=0;y<rows;y++){const r=[];for(let x=0;x<cols;x++){const i=(y*cols+x)*4;
    if(px[i+3]<=90){r.push(0);continue;}
    r.push((px[i]>150&&px[i+1]>110&&px[i+2]<120)?2:1);} f.push(r);}
  try{fs.unlinkSync(tmp);}catch(e){}
  return {f,d:note};
}
const fd=fn=>{const o=createCanvas(cols,rows),c=o.getContext('2d');
  c.clearRect(0,0,cols,rows);fn(c);const px=c.getImageData(0,0,cols,rows).data,f=[];
  for(let y=0;y<rows;y++){const r=[];for(let x=0;x<cols;x++)r.push(px[(y*cols+x)*4+3]>110);f.push(r);}return f;};
function bloom(a,b,t){const cx=(cols-1)/2,cy=(rows-1)/2,out=[];
  for(let y=0;y<rows;y++){const ny=cy?(y-cy)/cy:0,row=[];
    for(let x=0;x<cols;x++){const nx=cx?(x-cx)/cx:0;
      row.push(Math.sqrt(nx*nx+ny*ny)/Math.SQRT2<=t?b[y][x]:a[y][x]);}out.push(row);}return out;}

(async()=>{
const MK=await brandMark(cols,rows);
const A_END=P.ripple(cols,rows,1);
const B_START=fd(c=>P.pendulum(c,cols,rows,0));
const STEPS=[['pattern ends',A_END],['bloom .25',bloom(A_END,MK.f,.25)],
  ['bloom .60',bloom(A_END,MK.f,.60)],['THE MARK',MK.f],
  ['disperse .40',bloom(MK.f,B_START,.40)],['next pattern',B_START]];
const PAD=20,LAB=15,GAP=9;
const cv=createCanvas((PAD*2+W)*S,(PAD*2+STEPS.length*(LAB+H+GAP))*S);
const ctx=cv.getContext('2d'); ctx.fillStyle='#0b0b0b';ctx.fillRect(0,0,cv.width,cv.height);
ctx.fillStyle='#ffd400';ctx.font=`600 ${10*S}px Oswald`;
let y=PAD;
STEPS.forEach(([lab,field])=>{
  ctx.fillStyle='#ffd400';ctx.font=`600 ${9.5*S}px Oswald`;ctx.letterSpacing=`${1.4*S}px`;
  ctx.fillText(lab.toUpperCase(),PAD*S,(y+9)*S); ctx.letterSpacing='0px'; y+=LAB;
  const p=createCanvas(W*S,H*S),pc=p.getContext('2d');
  const g=pc.createLinearGradient(0,0,0,H*S);
  g.addColorStop(0,'#26221c');g.addColorStop(0.14,'#141210');g.addColorStop(1,'#100e0c');
  pc.fillStyle=g;pc.fillRect(0,0,p.width,p.height);
  M.drawMatrix(pc,{scale:S,w:W,h:H,field});
  ctx.drawImage(p,PAD*S,y*S); y+=H+GAP;
});
fs.writeFileSync(path.join(__dirname,'renders',(process.env.OUT||'brand-moment')+'.png'),cv.toBuffer('image/png'));
console.log('  grid '+cols+'x'+rows+'   mark: '+MK.d);
})();
