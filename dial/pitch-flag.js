/**
 * pitch-flag.js — two questions answered by looking.
 *   1 · what "pitch" means, at 1:1 and magnified
 *   2 · can a flag survive on a one-accent board?
 */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const D=require('./discs.js');
registerFonts();
const BAND='#1a1a1a';
const base=(p)=>({...D.BASE,pitch:p,lattice:'hex',ground:[BAND],offHi:BAND,offLo:BAND,
                  onHi:'#ddd6c6',onLo:'#a49d8c'});
const W=260,H=133, SW=539,SH=68;
const TMP=path.join(__dirname,'renders','_pf.png');
execFileSync('rsvg-convert',['-w','1200','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);

/* ── flags, drawn as the most legible ONE-COLOUR reading of each ─────────────────── */
function usFlag(c,w,h){
  const fw=Math.min(w*0.30, h*1.9), fh=fw/1.9, x=(w-fw)/2, y=(h-fh)/2;
  c.fillStyle='#fff';
  for(let i=0;i<13;i+=2) c.fillRect(x, y+i*fh/13, fw, fh/13);      // the light stripes
  const cw=fw*0.4, ch=fh*7/13;
  c.fillRect(x,y,cw,ch);                                            // canton, solid
  c.globalCompositeOperation='destination-out';                     // stars punched out
  for(let r=0;r<9;r++) for(let s=0;s<(r%2?5:6);s++){
    const sx=x+cw*((s+(r%2?1:0.5))/6), sy=y+ch*((r+0.5)/9);
    c.beginPath(); c.arc(sx,sy,Math.max(0.6,ch*0.035),0,7); c.fill();
  }
  c.globalCompositeOperation='source-over';
}
function star(c,cx,cy,r){
  c.beginPath();
  for(let i=0;i<10;i++){ const a=-Math.PI/2+i*Math.PI/5, rr=i%2?r*0.42:r;
    c[i?'lineTo':'moveTo'](cx+Math.cos(a)*rr, cy+Math.sin(a)*rr); }
  c.closePath(); c.fill();
}
function txFlag(c,w,h){
  const fw=Math.min(w*0.30, h*1.5), fh=fw/1.5, x=(w-fw)/2, y=(h-fh)/2;
  c.fillStyle='#fff';
  c.fillRect(x+fw/3, y, fw*2/3, fh/2);                              // the white bar
  star(c, x+fw/6, y+fh/2, fh*0.30);                                 // the lone star
}
function loneStar(c,w,h){
  const t='HOUSTON, TEXAS';
  let z=Math.round(h*0.72);
  while(z>7){c.font='600 '+z+'px Oswald'; if(c.measureText(t).width+h*1.1<=w*0.90)break; z--;}
  const tw=c.measureText(t).width, gap=h*0.26, sr=h*0.34;
  const x0=(w-(sr*2+gap+tw))/2;
  c.fillStyle='#fff'; star(c, x0+sr, h/2, sr);
  c.textAlign='left'; c.textBaseline='middle'; c.fillText(t, x0+sr*2+gap, h/2+1);
}

(async()=>{
const LOGO=await loadImage(TMP);
const roundel=(s)=>(c,w,h)=>{const d=h*(s||0.92); c.drawImage(LOGO,(w-d)/2,(h-d)/2,d,d);};

const PAD=16,LAB=14,G=12;
const rows=[];
const P=[3,2.5,2];
rows.push({label:'PITCH — the gap between disc centres. Smaller gap = more, smaller discs = finer mark. TRUE SIZE:',
  tiles:P.map(p=>{const st=base(p),f=D.sampleField(W,H,roundel(0.92),st);
    return {W,H,st,f,cap:p+'px · '+Math.round(0.92*f[f.length-1].rows)+' rows across the mark'};})});
rows.push({label:'THE SAME THREE, MAGNIFIED 4× — this is the only thing "pitch" changes:',
  tiles:P.map(p=>{const st=base(p),f=D.sampleField(W,H,roundel(0.92),st);
    return {W,H,st,f,zoom:4,crop:[0.30,0.18,0.34,0.5]};})});
rows.push({label:'THE US FLAG on the strip (26 disc rows) and on the block (51). One accent colour, no red or blue.',
  tiles:[{W:SW,H:SH,st:base(3),f:D.sampleField(SW,SH,usFlag,base(3))},
         {W,H,st:base(2),f:D.sampleField(W,H,usFlag,base(2))}]});
rows.push({label:'THE TEXAS FLAG — three shapes instead of fifty stars.',
  tiles:[{W:SW,H:SH,st:base(3),f:D.sampleField(SW,SH,txFlag,base(3))},
         {W,H,st:base(2),f:D.sampleField(W,H,txFlag,base(2))}]});
rows.push({label:'THE LONE STAR beside the words — one shape, and it reads at any size.',
  tiles:[{W:SW,H:SH,st:base(3),f:D.sampleField(SW,SH,loneStar,base(3))},
         {W:SW,H:SH,st:base(2),f:D.sampleField(SW,SH,loneStar,base(2))}]});

let cw=0,ch=PAD*2;
for(const r of rows){ let x=0,hh=0;
  for(const t of r.tiles){ const w=t.zoom?t.W*t.crop[2]*t.zoom:t.W, h=t.zoom?t.H*t.crop[3]*t.zoom:t.H;
    x+=w+G; hh=Math.max(hh,h); }
  cw=Math.max(cw,x-G); ch+=LAB+hh+G+14; }
const cv=createCanvas(cw+PAD*2,ch), ctx=cv.getContext('2d');
ctx.fillStyle='#0d0d0d'; ctx.fillRect(0,0,cv.width,cv.height);
let y=PAD;
for(const r of rows){
  ctx.fillStyle='#ffd400'; ctx.font='600 9.5px Oswald'; ctx.letterSpacing='1.05px';
  ctx.fillText(r.label,PAD,y+9); ctx.letterSpacing='0px'; y+=LAB;
  let x=PAD, hh=0;
  for(const t of r.tiles){
    if(t.zoom){
      const tmp=createCanvas(t.W,t.H), tc=tmp.getContext('2d');
      D.paint(tc,t.W,t.H,t.f,t.st,0,0);
      const [cx,cy,cwf,chf]=t.crop;
      ctx.imageSmoothingEnabled=false;
      ctx.drawImage(tmp, t.W*cx,t.H*cy,t.W*cwf,t.H*chf, x,y, t.W*cwf*t.zoom,t.H*chf*t.zoom);
      const w=t.W*cwf*t.zoom, h=t.H*chf*t.zoom;
      ctx.strokeStyle='#3a352c'; ctx.strokeRect(x+.5,y+.5,w-1,h-1);
      x+=w+G; hh=Math.max(hh,h);
    } else {
      D.paint(ctx,t.W,t.H,t.f,t.st,x,y);
      if(t.cap){ ctx.fillStyle='#8a8272'; ctx.font='400 8.5px Oswald';
                 ctx.fillText(t.cap,x+2,y+t.H+11); }
      x+=t.W+G; hh=Math.max(hh,t.H);
    }
  }
  y+=hh+G+14;
}
fs.writeFileSync(path.join(__dirname,'renders','pitch-flag.png'),cv.toBuffer('image/png'));
try{fs.unlinkSync(TMP);}catch(e){}
console.log('  pitch-flag.png '+cv.width+'x'+cv.height);
})();
