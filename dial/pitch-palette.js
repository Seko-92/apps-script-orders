/** pitch-palette.js — a disc has to survive QUANTISATION, not just rendering. How fine can
 *  the pitch go before the shading that makes it read as a disc is encoded away? */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {GIFEncoder,quantize,applyPalette}=require('gifenc');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const D=require('./discs.js');
const BAND='#1a1a1a', W=260,H=133;
const TMP=path.join(__dirname,'renders','_pp.png');
execFileSync('rsvg-convert',['-w','1400','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);
(async()=>{
const LOGO=await loadImage(TMP);
const draw=(c,w,h)=>{const d=h*0.92;c.drawImage(LOGO,(w-d)/2,(h-d)/2,d,d);};
const cases=[];
for(const [pitch,cols] of [[3,6],[2.5,6],[2.5,10],[2,6],[2,12],[2,20],[1.8,20]]){
  const ST={...D.BASE,pitch,lattice:'hex',ground:[BAND],offHi:BAND,offLo:BAND,
            onHi:'#ddd6c6',onLo:'#a49d8c'};
  const cv=createCanvas(W,H),ctx=cv.getContext('2d');
  const f=D.sampleField(W,H,draw,ST);
  D.paint(ctx,W,H,f,ST,0,0);
  const data=ctx.getImageData(0,0,W,H).data;
  const pal=quantize(data,cols,{format:'rgb565'});
  const G=[0x1a,0x1a,0x1a];let gi=0,gd=Infinity;
  pal.forEach((c,i)=>{const d=(c[0]-G[0])**2+(c[1]-G[1])**2+(c[2]-G[2])**2;if(d<gd){gd=d;gi=i;}});
  pal[gi]=[...G,...(pal[gi].length>3?[pal[gi][3]]:[])];
  const gif=GIFEncoder();
  gif.writeFrame(applyPalette(data,pal,'rgb565'),W,H,{palette:pal,delay:100,repeat:0});
  gif.finish();
  const buf=Buffer.from(gif.bytes());
  const nm='_pp-'+pitch+'-'+cols+'.gif';
  fs.writeFileSync(path.join(__dirname,'renders',nm),buf);
  cases.push({pitch,cols,kb:buf.length/1024,nm,rows:f[f.length-1].rows});
  console.log(`  pitch ${String(pitch).padEnd(4)} palette ${String(cols).padStart(2)}  `
    +`${String(f[f.length-1].rows).padStart(2)} rows  ${(buf.length/1024).toFixed(0).padStart(3)} KB/frame`);
}
fs.writeFileSync(path.join(__dirname,'renders','_pp.json'),JSON.stringify(cases));
fs.unlinkSync(TMP);
})();
