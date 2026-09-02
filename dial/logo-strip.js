'use strict';
const fs=require('fs'), path=require('path'), {execFileSync}=require('child_process');
const { createCanvas, loadImage } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const M=require('./src/board-matrix'); registerFonts();
const S=3, TMP=path.join(__dirname,'renders','_lg.png');
const W=539,H=56,P=4, cols=Math.floor(W/P), rows=Math.floor(H/P);

async function roundel(){
  execFileSync('rsvg-convert',['-w','900','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);
  return await loadImage(TMP);
}
function thresh(c){const px=c.getImageData(0,0,cols,rows).data,f=[];
  for(let y=0;y<rows;y++){const r=[];for(let x=0;x<cols;x++)r.push(px[(y*cols+x)*4+3]>90);f.push(r);}return f;}
function fresh(){const o=createCanvas(cols,rows),c=o.getContext('2d');
  c.clearRect(0,0,cols,rows);c.fillStyle='#fff';c.textBaseline='middle';return c;}

async function textOnly(text){
  const c=fresh(); let size=Math.round(rows*0.95);
  while(size>5){c.font='600 '+size+'px Oswald'; if(c.measureText(text).width<=cols*0.88)break; size--;}
  c.textAlign='center'; c.fillText(text, cols/2, rows/2);
  return {f:thresh(c), n:'"'+text+'" @'+size};
}
async function withMark(text){
  const img=await roundel(), c=fresh();
  const d=rows*0.98, gap=rows*0.30;
  let size=Math.round(rows*0.95);
  while(size>5){c.font='600 '+size+'px Oswald';
    if(d + gap + c.measureText(text).width <= cols*0.92) break; size--; }
  const tw=c.measureText(text).width, total=d+gap+tw, x0=(cols-total)/2;
  c.drawImage(img,x0,(rows-d)/2,d,d);
  c.textAlign='left'; c.fillText(text, x0+d+gap, rows/2);
  return {f:thresh(c), n:'roundel + "'+text+'" @'+size};
}
function tile(ctx,x,y,field){
  const p=createCanvas(W*S,H*S),pc=p.getContext('2d');
  const g=pc.createLinearGradient(0,0,0,H*S);
  g.addColorStop(0,'#26221c');g.addColorStop(0.14,'#141210');g.addColorStop(1,'#100e0c');
  pc.fillStyle=g;pc.fillRect(0,0,p.width,p.height);
  M.drawMatrix(pc,{scale:S,w:W,h:H,field});ctx.drawImage(p,x*S,y*S);
}
(async()=>{
  const CASES=[ await textOnly('HQ MOTOR SERVICE'),
                await withMark('MOTOR SERVICE'),
                await withMark('HIGH QUALITY MOTOR SERVICE'),
                await withMark('HQ MOTOR SERVICE') ];
  const PAD=20,LAB=17,GAP=14;
  const cv=createCanvas((PAD*2+W)*S,(PAD*2+CASES.length*(LAB+H+GAP))*S);
  const ctx=cv.getContext('2d');ctx.fillStyle='#0b0b0b';ctx.fillRect(0,0,cv.width,cv.height);
  let y=PAD;
  for(const c of CASES){
    ctx.fillStyle='#ffd400';ctx.font=`600 ${10*S}px Oswald`;ctx.letterSpacing=`${1.3*S}px`;
    ctx.fillText(c.n.toUpperCase(),PAD*S,(y+9)*S);ctx.letterSpacing='0px';y+=LAB;
    tile(ctx,PAD,y,c.f); y+=H+GAP;
  }
  fs.writeFileSync(path.join(__dirname,'renders','logo-strip.png'),cv.toBuffer('image/png'));
  fs.unlinkSync(TMP); CASES.forEach(c=>console.log('  '+c.n));
})();
