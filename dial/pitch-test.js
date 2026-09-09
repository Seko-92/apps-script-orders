/** pitch-test.js — is 14 discs the problem, or is PITCH the problem?
 *  The strip is 56px tall. At the shipped 4px pitch that is 14 discs; finer pitch buys rows. */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render'); const P=require('./src/patterns'); registerFonts();
const W=539,H=56,S=1,TMP=path.join(__dirname,'renders','_p.png');

/** drawMatrix with a caller-supplied pitch, so this can compare 4 / 3 / 2.5. */
function draw(ctx,pitch,field,ox,oy){
  const R=pitch*0.405, rows=field.length, cols=field[0].length;
  const x0=ox+(W-cols*pitch)/2+pitch/2, y0=oy+(H-rows*pitch)/2+pitch/2;
  const g=ctx.createLinearGradient(0,oy,0,oy+H);
  g.addColorStop(0,'#26221c');g.addColorStop(0.14,'#141210');g.addColorStop(1,'#100e0c');
  ctx.fillStyle=g;ctx.fillRect(ox,oy,W,H);
  for(let y=0;y<rows;y++)for(let x=0;x<cols;x++){
    const on=field[y][x],cx=x0+x*pitch,cy=y0+y*pitch;
    if(on){ctx.beginPath();ctx.arc(cx,cy+0.5,R,0,7);ctx.fillStyle='rgba(0,0,0,0.55)';ctx.fill();}
    const lg=ctx.createLinearGradient(0,cy-R,0,cy+R);
    const acc=on===2;
    lg.addColorStop(0,acc?'#ffdc00':(on?'#f4f0e4':'#26221c'));
    lg.addColorStop(1,acc?'#d9ab00':(on?'#cdc7b6':'#15120f'));
    ctx.beginPath();ctx.arc(cx,cy,R,0,7);ctx.fillStyle=lg;ctx.fill();
  }
}
async function markField(cols,rows){
  execFileSync('rsvg-convert',['-w','900','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);
  const img=await loadImage(TMP);
  const o=createCanvas(cols,rows),c=o.getContext('2d');c.clearRect(0,0,cols,rows);
  const d=rows*0.98,gap=rows*0.30,text='MOTOR SERVICE';
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
function tickerField(cols,rows){
  const T=P.tickerStrip(createCanvas,cols,rows,'HQ MOTOR SERVICE · HOUSTON · ');
  const o=createCanvas(cols,rows),c=o.getContext('2d');
  c.clearRect(0,0,cols,rows);P.ticker(c,cols,rows,0.18,null,T);
  const px=c.getImageData(0,0,cols,rows).data,f=[];
  for(let y=0;y<rows;y++){const r=[];for(let x=0;x<cols;x++)r.push(px[(y*cols+x)*4+3]>110);f.push(r);}
  return f;
}
(async()=>{
  const PITCHES=[4,3,2.5];
  const PAD=18,LAB=15,GAP=10;
  const cv=createCanvas(W+PAD*2,PAD*2+PITCHES.length*2*(LAB+H+GAP));
  const ctx=cv.getContext('2d');ctx.fillStyle='#0b0b0b';ctx.fillRect(0,0,cv.width,cv.height);
  let y=PAD;
  for(const p of PITCHES){
    const cols=Math.floor(W/p),rows=Math.floor(H/p);
    for(const [what,fld,extra] of [['MARK',(await markField(cols,rows)),null],
                                   ['TICKER',{f:tickerField(cols,rows)},null]]){
      ctx.fillStyle='#ffd400';ctx.font='600 9px Oswald';ctx.letterSpacing='1.3px';
      ctx.fillText('PITCH '+p+'px → '+cols+'x'+rows+' discs   ·   '+what+
        (fld.size?'  (text @'+fld.size+')':''),PAD,y+9);
      ctx.letterSpacing='0px';y+=LAB;
      draw(ctx,p,fld.f,PAD,y); y+=H+GAP;
    }
  }
  fs.writeFileSync(path.join(__dirname,'renders','pitch-test.png'),cv.toBuffer('image/png'));
  try{fs.unlinkSync(TMP);}catch(e){}
  console.log('  pitch-test.png — at TRUE 1:1, the size it lives at');
})();
