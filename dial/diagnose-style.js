/** diagnose-style.js — is the board unreadable because of PITCH, or because the OFF discs
 *  are invisible? Rendered at TRUE 1:1, the size it lives at. */
'use strict';
const fs=require('fs'), path=require('path'), {execFileSync}=require('child_process');
const {createCanvas, loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const D=require('./discs.js');
registerFonts();

const STRIP={w:539,h:56}, TMP=path.join(__dirname,'renders','_mk.png');
execFileSync('rsvg-convert',['-w','900','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);

(async()=>{
const img=await loadImage(TMP);

/** the strip lockup: roundel + MOTOR SERVICE, drawn at real pixel size */
const markStrip=(c,w,h)=>{
  const d=h*0.94, gap=h*0.26, text='MOTOR SERVICE';
  let size=Math.round(h*0.74);
  while(size>6){c.font='600 '+size+'px Oswald';
    if(d+gap+c.measureText(text).width<=w*0.92)break;size--;}
  const tw=c.measureText(text).width, x0=(w-(d+gap+tw))/2;
  c.drawImage(img,x0,(h-d)/2,d,d);
  c.fillStyle='#fff'; c.textBaseline='middle'; c.textAlign='left';
  c.fillText(text,x0+d+gap,h/2+1);
};

const STYLES=[
 ['AS SHIPPED — off disc #26221c on a ground that starts #26221c',
  {pitch:4}],
 ['SAME FIELD · off discs made visible — the grid appears',
  {pitch:4, ground:['#1a1a1a'], offHi:'#2e2a25', offLo:'#171412'}],
 ['+ HEX LATTICE — half-pitch offset, 0.866 row spacing (buys rows too)',
  {pitch:4, lattice:'hex', ground:['#1a1a1a'], offHi:'#2e2a25', offLo:'#171412'}],
 ['+ 3px PITCH — the plan’s number',
  {pitch:3, lattice:'hex', ground:['#1a1a1a'], offHi:'#2a2723', offLo:'#171412'}],
 ['+ QUIET ON-TONE — warm grey, not #f4f0e4 (an LED sign vs an instrument)',
  {pitch:3, lattice:'hex', ground:['#1a1a1a'], offHi:'#2a2723', offLo:'#171412',
   onHi:'#ddd6c6', onLo:'#a49d8c'}],
 ['+ CORNER VIGNETTE — square file, round-looking object',
  {pitch:3, lattice:'hex', ground:['#1a1a1a'], offHi:'#2a2723', offLo:'#171412',
   onHi:'#ddd6c6', onLo:'#a49d8c', vignette:1.06}],
];

const PAD=20, LAB=16, GAP=13;
const cv=createCanvas(STRIP.w+PAD*2, PAD*2+STYLES.length*(LAB+STRIP.h+GAP));
const ctx=cv.getContext('2d');
ctx.fillStyle='#0d0d0d'; ctx.fillRect(0,0,cv.width,cv.height);
let y=PAD;
for(const [label,over] of STYLES){
  const st={...D.BASE,...over};
  const f=D.sampleField(STRIP.w,STRIP.h,markStrip,st);
  const meta=f[f.length-1];
  ctx.fillStyle='#ffd400'; ctx.font='600 9px Oswald'; ctx.letterSpacing='1.1px';
  ctx.fillText(label+'   ('+meta.cols+'×'+meta.rows+')',PAD,y+9);
  ctx.letterSpacing='0px'; y+=LAB;
  D.paint(ctx,STRIP.w,STRIP.h,f,st,PAD,y); y+=STRIP.h+GAP;
}
fs.writeFileSync(path.join(__dirname,'renders','diag-style.png'),cv.toBuffer('image/png'));
try{fs.unlinkSync(TMP);}catch(e){}
console.log('  diag-style.png  '+cv.width+'x'+cv.height+' @1:1');
})();
