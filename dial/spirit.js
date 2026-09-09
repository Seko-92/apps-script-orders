/**
 * spirit.js — the three structures the user shortlisted (C · F · E), each with the WHOLE
 * system applied, so "do the puzzles come together" is judgeable instead of arguable.
 *
 * ⚠⚠ HONEST CELL TYPOGRAPHY. A cell whose value is a FORMULA carries exactly ONE font size,
 *    weight and colour — rich text does not survive a recalc, and every readout here has to
 *    self-update with no trigger. So there is no big-number-over-small-label inside one cell.
 *    Hierarchy comes from cell-to-cell size differences, and the coloured lamp is an EMOJI,
 *    which renders in its own colour regardless of the cell's font colour (the trick E1 and
 *    G1 already use today).
 */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const D=require('./discs.js');
registerFonts();

const COL={A:103,B:70,C:87,D:232,E:307,F:130,G:100,H:107};
const X={}; { let x=0; for(const k of 'ABCDEFGH'){X[k]=x; x+=COL[k];} }
const W=1136, R1=56, R2=65, R3=36, RD=30, RB=34;
const INK='#1a1a1a', CREAM='#fff8e7', PAPER='#fff', YEL='#ffd400';
const QUIET='#e8e8e8', DIM='#9a9386', LBL='#807a6d', AMB='#e0a97a', LAMP='#9aa3ad';
const FINE={...D.BASE,pitch:3,lattice:'hex',ground:[INK],
            offHi:'#2a2723',offLo:'#171412',onHi:'#ddd6c6',onLo:'#a49d8c'};
const TMP=path.join(__dirname,'renders','_sp.png');
execFileSync('rsvg-convert',['-w','1200','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);
let LOGO=null;

/* ── THE MOTIF — four discs, three at rest and one lit. Same object on both bands; only
      the ground decides which tones read as off and on. This is a NUMBER FORMAT on the
      sheet, not an image: it costs nothing, and it moves with its row. ─────────────── */
function motif(ctx,x,cy,onDark){
  for(let i=0;i<4;i++){
    const cx=x+5+i*11, r=4.1;
    ctx.beginPath();ctx.arc(cx,cy+0.5,r,0,7);
    ctx.fillStyle=onDark?'rgba(0,0,0,0.55)':'rgba(0,0,0,0.28)';ctx.fill();
    const g=ctx.createLinearGradient(0,cy-r,0,cy+r);
    if(i<3){ if(onDark){g.addColorStop(0,'#2e2a25');g.addColorStop(1,'#171412');}
             else      {g.addColorStop(0,'#211f1a');g.addColorStop(1,'#0d0c0a');} }
    else   { if(onDark){g.addColorStop(0,'#ffdc00');g.addColorStop(1,'#d9ab00');}
             else      {g.addColorStop(0,'#fff3b0');g.addColorStop(1,'#e8c93a');} }
    ctx.beginPath();ctx.arc(cx,cy,r,0,7);ctx.fillStyle=g;ctx.fill();
  }
  return x+4*11+12;
}
function ebayMark(ctx,cx,cy,size){
  ctx.font=`700 ${size}px "Noto Sans"`;ctx.textAlign='left';ctx.textBaseline='alphabetic';
  const chs=[['e','#e53238'],['b','#0064d2'],['a','#f5af02'],['y','#86b817']];
  let tw=0;for(const [ch] of chs)tw+=ctx.measureText(ch).width;
  let x=cx;for(const [ch,col] of chs){ctx.fillStyle=col;ctx.fillText(ch,x,cy);x+=ctx.measureText(ch).width;}
  return tw;
}
function board(ctx,x,y,w,h,compose){
  D.paint(ctx,w,h,D.sampleField(w,h,compose,FINE),FINE,x,y);
}
const roundel=(s)=>(c,w,h)=>{const d=h*s;c.drawImage(LOGO,(w-d)/2,(h-d)/2,d,d);};
const lockup=(t,s,band)=>(c,w,h)=>{
  // `band` confines the lockup to a slice of the canvas, so a two-row board can keep its
  // lower strip clear for the eBay mark instead of the mark landing on the wordmark.
  const top=band?band[0]*h:0, hh=band?(band[1]-band[0])*h:h;
  const d=hh*s,gap=hh*0.26;let z=Math.round(hh*0.74);
  while(z>6){c.font='600 '+z+'px Oswald';if(d+gap+c.measureText(t).width<=w*0.90)break;z--;}
  const x0=(w-(d+gap+c.measureText(t).width))/2;
  c.drawImage(LOGO,x0,top+(hh-d)/2,d,d);
  c.fillStyle='#fff';c.textBaseline='middle';c.textAlign='left';
  c.fillText(t,x0+d+gap,top+hh/2+1);
};

/* ── the readouts. ONE size per cell — this is what a formula cell can really do. ──── */
function cell(ctx,x,lines,size,fill,lamp){
  ctx.textAlign='left';ctx.textBaseline='alphabetic';
  let lx=x+12;
  if(lamp){ctx.fillStyle=lamp;ctx.beginPath();ctx.arc(x+16,lines.length>1?20:R1/2-1,4,0,7);ctx.fill();lx=x+26;}
  ctx.font=`400 ${size}px Oswald`;ctx.fillStyle=fill;
  if(lines.length===1){ctx.fillText(lines[0],lx,R1/2+size*0.36);}
  else{lines.forEach((L,i)=>ctx.fillText(L,i?x+12:lx,24+i*17));}
}
function panel(ctx){
  cell(ctx,X.D,['14 TO GRAB'],19,YEL);
  cell(ctx,X.E,['4:33 PM · resting · 4m ago'],11.5,DIM,LAMP);   // E1 keeps h:mm AM/PM
  cell(ctx,X.F,['the floor is asleep','eBay 7 · Direct 7'],11.5,QUIET);
  cell(ctx,X.H,['oldest','10h 55m'],11.5,AMB);
}
function quietReadout(ctx){
  cell(ctx,X.F,['the floor is asleep','waiting: 14'],11.5,QUIET);
  cell(ctx,X.H-2,['RESTING','4:33 PM'],11,DIM,LAMP);
}
function pickIds(ctx,oy){
  for(const [x,lab,val] of [[X.F,'PICK ID · SHIPPING','Shipping - Yassin 1'],
                            [X.H,'ADJUSTMENT','AShamma 2']]){
    ctx.strokeStyle='#2e2a25';ctx.lineWidth=1;
    ctx.beginPath();ctx.moveTo(x+0.5,oy);ctx.lineTo(x+0.5,oy+R2);ctx.stroke();
    ctx.fillStyle='#8a8272';ctx.font='600 7.5px Oswald';ctx.letterSpacing='1.2px';
    ctx.fillText(lab,x+10,oy+22);ctx.letterSpacing='0px';
    ctx.fillStyle=YEL;ctx.font='400 11px Oswald';ctx.fillText(val,x+10,oy+42);
  }
}

/* ── everything below the banner ───────────────────────────────────────────────────── */
const HEADS=[['A','◆ SKU'],['B','# QTY'],['C','LOCATION'],['D','SALES ORDER'],
             ['E','NOTE'],['F','STATUS'],['G','◫ HAND'],['H','◪ LEFT']];
const EB=[['157860','1','D-36','27-15085-56706','Ship Without Head Gasket','PENDING','0'],
          ['194622','3','E-16','27-15085-56706','↳ from KIT-157860','PENDING','36'],
          ['168111','1','L-70/C-63','27-15085-56706','↳ from KIT-157860','PENDING','12']];
const DR=[['166212','6','B-45','SO-25314','','PREPARING','15'],
          ['163710','1','B-52','SO-25314','','PREPARING','30']];
function headerRow(ctx,y){
  ctx.fillStyle=INK;ctx.fillRect(0,y,W,R3);
  ctx.fillStyle=YEL;ctx.font='600 10px Oswald';ctx.letterSpacing='0.6px';
  for(const [c,t] of HEADS){
    ctx.textAlign=(c==='A'||c==='C'||c==='D')?'left':'center';
    ctx.fillText(t,ctx.textAlign==='left'?X[c]+10:X[c]+COL[c]/2,y+R3/2+4);}
  ctx.letterSpacing='0px';ctx.textAlign='left';return y+R3;
}
function dataRows(ctx,y,rows,statusFill,statusInk){
  rows.forEach((r,i)=>{
    ctx.fillStyle=i%2?CREAM:PAPER;ctx.fillRect(0,y,W,RD);
    ctx.fillStyle=statusFill;ctx.fillRect(X.F,y,COL.F,RD);
    r.forEach((v,j)=>{const c='ABCDEFGH'[j];if(!v)return;
      const mono=j===0||j===3;
      ctx.font=c==='F'?'600 10px Oswald':(mono?'600 10.5px "JetBrains Mono"':'400 10.5px "Noto Sans"');
      ctx.fillStyle=c==='F'?statusInk:'#1d1d1b';
      ctx.textAlign=(c==='A'||c==='C')?'left':'center';
      ctx.fillText(v,ctx.textAlign==='left'?X[c]+12:X[c]+COL[c]/2,y+RD/2+4);});
    ctx.textAlign='left';y+=RD;});
  return y;
}
function directBand(ctx,y,withMotif){
  ctx.fillStyle=YEL;ctx.fillRect(0,y,W,RB);
  let tx=12;
  if(withMotif) tx=motif(ctx,tx,y+RB/2,false);
  else{ctx.fillStyle=INK;ctx.font='600 15px Oswald';ctx.fillText('▌',tx,y+22);tx+=22;}
  ctx.fillStyle=INK;ctx.font='600 15px Oswald';ctx.letterSpacing='1px';
  ctx.fillText('DIRECT',tx,y+22);
  ctx.font='600 8.5px Oswald';ctx.textAlign='right';
  ctx.fillText('HQMS · DIRECT ORDERS · 7 waiting',W-12,y+21);
  ctx.textAlign='left';ctx.letterSpacing='0px';return y+RB;
}

/* ── the three shortlisted structures ──────────────────────────────────────────────── */
const OPTS=[
{id:'C', label:'C · BLOCK + PANEL',
 draw(ctx){
   ctx.fillStyle=INK;ctx.fillRect(0,0,W,R1+R2);
   if(!process.env.NOBOARD) board(ctx,0,0,260,121,roundel(0.78));
   panel(ctx);
   // row 2 right of the block IS the eBay table's label band — same grammar as DIRECT
   ebayMark(ctx,300,R1+R2/2+9,28);
   pickIds(ctx,R1);
 }},
{id:'F', label:'F · PANEL UP, NAMEPLATE DOWN',
 draw(ctx){
   ctx.fillStyle=INK;ctx.fillRect(0,0,W,R1+R2);
   cell(ctx,0,['14 TO GRAB'],19,YEL);
   cell(ctx,X.D,['oldest 10h 55m'],13,AMB);
   cell(ctx,X.E,['4:33 PM · resting · 4m ago'],11.5,DIM,LAMP);
   cell(ctx,X.F,['the floor is asleep','eBay 7 · Direct 7'],11.5,QUIET);
   cell(ctx,X.H,['shipped','0 today'],11.5,QUIET);
   board(ctx,0,R1,799,R2,lockup('HQ MOTOR SERVICE',0.58));
   ebayMark(ctx,688,R1+R2/2+9,26);      // floats ON the board, in its own clear zone
   pickIds(ctx,R1);
 }},
{id:'E', label:'E · FULL BOARD',
 draw(ctx){
   ctx.fillStyle=INK;ctx.fillRect(0,0,W,R1+R2);
   board(ctx,0,0,799,121,lockup('HQ MOTOR SERVICE',0.86,[0,0.54]));
   quietReadout(ctx);
   // the board's lower strip does row 2's job — motif + mark, same grammar as DIRECT
   const ex=motif(ctx,20,R1+R2/2+3,true);
   ebayMark(ctx,ex,R1+R2/2+12,26);
   pickIds(ctx,R1);
 }},
];

(async()=>{
  LOGO=await loadImage(TMP);
  const H=R1+R2+R3+RD*3+RB+R3+RD*2;
  for(const o of OPTS){
    const cv=createCanvas(W,H),ctx=cv.getContext('2d');
    ctx.fillStyle='#fff';ctx.fillRect(0,0,W,H);
    o.draw(ctx);
    let y=headerRow(ctx,R1+R2);
    y=dataRows(ctx,y,EB,'#ffcdd2','#b71c1c');
    y=directBand(ctx,y,true);
    y=headerRow(ctx,y);
    dataRows(ctx,y,DR,'#ffd400','#1a1a1a');
    fs.writeFileSync(path.join(__dirname,'renders',`spirit-${o.id}.png`),cv.toBuffer('image/png'));
    console.log(`  spirit-${o.id}.png  ${o.label.padEnd(30)} ${W}x${H} @1:1`);
  }
  try{fs.unlinkSync(TMP);}catch(e){}
})();
