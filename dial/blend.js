/**
 * blend.js — CAN YOU SEE WHERE THE IMAGE IS?
 *
 * The brief: row one should read as ONE unbroken black band, and the discs should simply
 * appear out of it. That makes the test binary — place the art on the real band colour
 * (BRAND.ink #1a1a1a) and look for the rectangle. A caret above each strip marks where the
 * image actually starts and stops, so you can check your own eye against the truth.
 *
 * ⚠ Today the art's ground is a GRADIENT #26221c → #141210 → #100e0c and the band is a FLAT
 *   #1a1a1a. Three tones, none of them the band. That is the block you can see.
 */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const D=require('./discs.js');
registerFonts();

const BAND='#1a1a1a';                       // BRAND.ink — the row's real colour
const W=1136, R1=56, IMG_X=0, IMG_W=799;    // the rail case: image spans A1:E1
const TMP=path.join(__dirname,'renders','_bl.png');
execFileSync('rsvg-convert',['-w','1200','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);

const VARIANTS=[
 ['① TODAY — gradient ground #26221c→#100e0c, right edge faded to #1a1a1a',
  {ground:['#26221c','#141210','#100e0c'], offHi:'#26221c', offLo:'#15120f', seam:true}],
 ['② FLAT #1a1a1a ground · off discs clearly visible (a board, sitting on the sheet)',
  {ground:[BAND], offHi:'#2a2723', offLo:'#171412'}],
 ['③ FLAT #1a1a1a ground · off discs barely there — a hint of material, no rectangle',
  {ground:[BAND], offHi:'#201e1b', offLo:'#181614'}],
 ['④ FLAT #1a1a1a ground · off discs INVISIBLE — only the lit discs exist',
  {ground:[BAND], offHi:BAND, offLo:BAND}],
];

(async()=>{
const LOGO=await loadImage(TMP);
const lockup=(c,w,h)=>{
  const t='HQ MOTOR SERVICE', d=h*0.90, gap=h*0.26;
  let z=Math.round(h*0.74);
  while(z>6){c.font='600 '+z+'px Oswald'; if(d+gap+c.measureText(t).width<=w*0.90)break; z--;}
  const x0=(w-(d+gap+c.measureText(t).width))/2;
  c.drawImage(LOGO,x0,(h-d)/2,d,d);
  c.fillStyle='#fff'; c.textBaseline='middle'; c.textAlign='left';
  c.fillText(t,x0+d+gap,h/2+1);
};
const PAD=16, LAB=15, MARK=11, GAP=20;
const cv=createCanvas(W+PAD*2, PAD*2+VARIANTS.length*(LAB+MARK+R1+GAP));
const ctx=cv.getContext('2d');
ctx.fillStyle='#0d0d0d'; ctx.fillRect(0,0,cv.width,cv.height);
let y=PAD;
for(const [label,over] of VARIANTS){
  ctx.fillStyle='#ffd400'; ctx.font='600 9px Oswald'; ctx.letterSpacing='1.1px';
  ctx.fillText(label,PAD,y+9); ctx.letterSpacing='0px'; y+=LAB;

  // the truth markers — where the image really begins and ends
  ctx.strokeStyle='#585248'; ctx.lineWidth=1;
  for(const mx of [IMG_X, IMG_X+IMG_W]){
    ctx.beginPath(); ctx.moveTo(PAD+mx+0.5,y+2); ctx.lineTo(PAD+mx+0.5,y+MARK-2); ctx.stroke();
  }
  ctx.fillStyle='#585248'; ctx.font='400 8px Oswald';
  ctx.fillText('image starts',PAD+IMG_X+4,y+MARK-2);
  ctx.textAlign='right'; ctx.fillText('image ends',PAD+IMG_X+IMG_W-4,y+MARK-2); ctx.textAlign='left';
  y+=MARK;

  // the WHOLE row painted in the sheet's own band colour first — the image goes ON it
  ctx.fillStyle=BAND; ctx.fillRect(PAD,y,W,R1);

  const st={...D.BASE, pitch:3, lattice:'hex', onHi:'#ddd6c6', onLo:'#a49d8c', ...over};
  const tile=createCanvas(IMG_W,R1), tc=tile.getContext('2d');
  D.paint(tc,IMG_W,R1,D.sampleField(IMG_W,R1,lockup,st),st,0,0);
  if(over.seam){
    const g=tc.createLinearGradient(IMG_W-26,0,IMG_W,0);
    g.addColorStop(0,'rgba(26,26,26,0)'); g.addColorStop(1,BAND);
    tc.fillStyle=g; tc.fillRect(IMG_W-26,0,26,R1);
  }
  ctx.drawImage(tile,PAD+IMG_X,y);

  // the live cells that sit to the right of the image, so the join is judged in context
  ctx.fillStyle='#e8e8e8'; ctx.font='400 11.5px Oswald';
  ctx.fillText('the floor is asleep',PAD+IMG_W+12,y+24);
  ctx.fillText('waiting: 14',PAD+IMG_W+12,y+41);
  ctx.fillStyle='#9aa3ad'; ctx.beginPath(); ctx.arc(PAD+IMG_W+241,y+24,4,0,7); ctx.fill();
  ctx.fillStyle='#9a9386'; ctx.font='400 11px Oswald';
  ctx.fillText('RESTING',PAD+IMG_W+251,y+27); ctx.fillText('4:33 PM',PAD+IMG_W+251,y+44);
  y+=R1+GAP;
}
fs.writeFileSync(path.join(__dirname,'renders','blend.png'),cv.toBuffer('image/png'));
try{fs.unlinkSync(TMP);}catch(e){}
console.log('  blend.png  '+cv.width+'x'+cv.height+' @1:1');
})();
