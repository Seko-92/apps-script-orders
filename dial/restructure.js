/**
 * restructure.js — five structures for row one, each drawn at TRUE 1:1 with the real table
 * under it. The axis is one question: HOW MUCH OF ROW ONE IS DECORATION?
 *
 * ⚠ HONEST CELL CONSTRAINT, respected throughout: a cell whose value is a FORMULA can carry
 *   exactly ONE font size — rich text does not survive a recalc. So hierarchy here comes from
 *   cell-to-cell size differences, never from mixed sizes inside one cell. That is already how
 *   D1 (headline) and E1 (pulse) work today.
 * ⚠ F1:G1 / H1 and both Pick ID cells are NEVER covered in any option.
 */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const M=require('./src/board-matrix');
const D=require('./discs.js');
const S=require('./sheet-ctx.js');
registerFonts();

const TMP=path.join(__dirname,'renders','_rs.png');
execFileSync('rsvg-convert',['-w','1200','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);

/* the refined style — every finding from the diagnosis, applied */
const FINE={...D.BASE, pitch:3, lattice:'hex', ground:['#1a1a1a'],
            offHi:'#2a2723', offLo:'#171412', onHi:'#ddd6c6', onLo:'#a49d8c'};

/* ── readout typography ────────────────────────────────────────────────────
 * Row 1 is 56px. Label baseline 21, value baseline 44 — one rhythm everywhere, so four
 * different readouts still read as one instrument row.
 * ⚠⚠ E1 IS LOAD-BEARING. ActivityLog.js regex-parses `h:mm AM/PM` out of it for the board
 *    heartbeat, the sidebar pulse and /status. Any option that FREES E1 must leave the
 *    System Pulse formula showing — which is a win, not a cost: it is live information that
 *    the strip is currently covering up.
 */
const QUIET='#e8e8e8', DIM='#9a9386', LAMP='#9aa3ad', LBL='#807a6d';
/** ONE size, weight and colour per cell — the honest limit of a formula cell. Two lines are
 *  fine; two SIZES are not. Hierarchy lives in the differences between cells. */
function stat(ctx,x,w,lines,size,fill,lamp){
  ctx.textAlign='left'; ctx.textBaseline='alphabetic';
  let lx=x+12;
  if(lamp){ ctx.fillStyle=lamp; ctx.beginPath();
            ctx.arc(x+16,lines.length>1?20:27,4,0,7); ctx.fill(); lx=x+26; }
  ctx.font=`400 ${size}px Oswald`; ctx.fillStyle=fill;
  if(lines.length===1) ctx.fillText(lines[0],lx,28+size*0.36);
  else lines.forEach((L,i)=>ctx.fillText(L,i?x+12:lx,24+i*17));
}
function prose(ctx,x,w,l1,l2,lamp){
  ctx.textAlign='left'; ctx.textBaseline='alphabetic';
  let lx=x+12;
  if(lamp){ ctx.fillStyle=lamp; ctx.beginPath(); ctx.arc(x+15,20,3.5,0,7); ctx.fill(); lx=x+24; }
  ctx.font='400 12px Oswald'; ctx.fillStyle=QUIET; ctx.fillText(l1,lx,24);
  if(l2){ ctx.font='400 10.5px Oswald'; ctx.fillStyle=DIM; ctx.fillText(l2,x+12,42); }
}
const FIG='600 29px Oswald';          // the headline figure — the answer to "what do I do"
const SUB='600 15px Oswald';
/** the four-cell instrument row that options C and D both use */
function panel(ctx){
  stat(ctx,X.D,COL.D,['14 TO GRAB'],19,'#ffd400');
  // ⚠ E1 keeps a value containing `h:mm AM/PM` — ActivityLog.js regex-parses it.
  stat(ctx,X.E,COL.E,['4:33 PM · resting · 4m ago'],11.5,DIM,LAMP);
  stat(ctx,X.F,COL.F+COL.G,['the floor is asleep','eBay 7 · Direct 7'],11.5,QUIET);
  stat(ctx,X.H,COL.H,['oldest','10h 55m'],11.5,'#e0a97a');
}
/* ── board content, drawn at real pixel size then area-sampled ─────────────── */
let LOGO=null;
const roundel=(scale,shiftX)=>(c,w,h)=>{
  const d=h*(scale||0.9);
  c.drawImage(LOGO,(w-d)/2+(shiftX||0),(h-d)/2,d,d);
};
const lockup=(text,hs)=>(c,w,h)=>{
  const d=h*(hs||0.94), gap=h*0.26;
  let size=Math.round(h*0.74);
  while(size>6){c.font='600 '+size+'px Oswald';
    if(d+gap+c.measureText(text).width<=w*0.90)break;size--;}
  const x0=(w-(d+gap+c.measureText(text).width))/2;
  c.drawImage(LOGO,x0,(h-d)/2,d,d);
  c.fillStyle='#fff';c.textBaseline='middle';c.textAlign='left';
  c.fillText(text,x0+d+gap,h/2+1);
};
const stacked=(a,b)=>(c,w,h)=>{               // roundel left, two lines right
  const d=h*0.72, gap=h*0.14;
  let sa=Math.round(h*0.34), sb=Math.round(h*0.20);
  c.font='600 '+sa+'px Oswald'; const wa=c.measureText(a).width;
  c.font='500 '+sb+'px Oswald'; const wb=c.measureText(b).width;
  const tw=Math.max(wa,wb), x0=(w-(d+gap+tw))/2;
  c.drawImage(LOGO,x0,(h-d)/2,d,d);
  c.fillStyle='#fff'; c.textBaseline='middle'; c.textAlign='left';
  c.font='600 '+sa+'px Oswald'; c.fillText(a,x0+d+gap,h/2-h*0.14);
  c.font='500 '+sb+'px Oswald'; c.fillText(b,x0+d+gap,h/2+h*0.16);
};
function board(ctx,x,y,w,h,compose,style){
  const st={...(style||FINE)};
  const f=D.sampleField(w,h,compose,st);
  D.paint(ctx,w,h,f,st,x,y);
  return f[f.length-1];
}
/* the SHIPPED renderer, for an honest baseline */
function shippedBoard(ctx,x,y,w,h,cols,rows,compose){
  const off=createCanvas(cols,rows),c=off.getContext('2d');
  c.clearRect(0,0,cols,rows); compose(c,cols,rows);
  const px=c.getImageData(0,0,cols,rows).data,fl=[];
  for(let yy=0;yy<rows;yy++){const r=[];for(let xx=0;xx<cols;xx++){const i=(yy*cols+xx)*4;
    if(px[i+3]<=90){r.push(0);continue;}
    r.push((px[i]>150&&px[i+1]>110&&px[i+2]<120)?2:1);}fl.push(r);}
  const t=createCanvas(w,h),tc=t.getContext('2d');
  const g=tc.createLinearGradient(0,0,0,h);
  g.addColorStop(0,'#26221c');g.addColorStop(0.14,'#141210');g.addColorStop(1,'#100e0c');
  tc.fillStyle=g;tc.fillRect(0,0,w,h);
  M.drawMatrix(tc,{scale:1,w,h,field:fl,pitch:4});
  const seam=tc.createLinearGradient(w-26,0,w,0);
  seam.addColorStop(0,'rgba(26,26,26,0)');seam.addColorStop(1,'#1a1a1a');
  tc.fillStyle=seam;tc.fillRect(w-26,0,26,h);
  ctx.drawImage(t,x,y);
}

/* ── the five structures ───────────────────────────────────────────────────── */
const {X,COL,W,R1,R2}=S;
const OPTS=[
{ id:'A', name:'AS IS', sub:'block A1:C2 + strip D1:E1 — 799 of 1136px is board (70%). Shipped renderer, shipped 4px pitch.',
  draw(ctx){
    ctx.fillStyle=S.INK; ctx.fillRect(0,0,W,R1+R2);
    shippedBoard(ctx,0,0,260,121,65,30,(c,w,h)=>{
      const d=h*0.94; c.drawImage(LOGO,(w-d)/2,(h-d)/2,d,d); });
    shippedBoard(ctx,X.D,0,539,56,134,14,(c,w,h)=>{
      const d=h*0.98,gap=h*0.30,t='MOTOR SERVICE';
      let z=Math.round(h*0.95);
      while(z>5){c.font='600 '+z+'px Oswald'; if(d+gap+c.measureText(t).width<=w*0.92)break;z--;}
      const x0=(w-(d+gap+c.measureText(t).width))/2;
      c.drawImage(LOGO,x0,(h-d)/2,d,d);
      c.fillStyle='#fff';c.textBaseline='middle';c.textAlign='left';
      c.fillText(t,x0+d+gap,h/2); });
    prose(ctx,X.F,COL.F+COL.G,'the floor is asleep','waiting: 14');
    prose(ctx,X.H-2,COL.H,'RESTING','4:33 PM',LAMP);
    S.ebay(ctx,X.D+50,R1+44,26);
    S.pickIds(ctx,R1,true);
  }},

{ id:'A2', name:'AS IS · REFINED', sub:'the SAME structure — only the disc style changes: one shared #1a1a1a ground, no seam fade, visible off discs, 3px hex, area sampling, quiet on-tone.',
  draw(ctx){
    ctx.fillStyle=S.INK; ctx.fillRect(0,0,W,R1+R2);
    board(ctx,0,0,260,121,roundel(0.78));
    board(ctx,X.D,0,539,56,lockup('MOTOR SERVICE',0.90));
    prose(ctx,X.F,COL.F+COL.G,'the floor is asleep','waiting: 14');
    prose(ctx,X.H-2,COL.H,'RESTING','4:33 PM',LAMP);
    S.ebay(ctx,X.D+50,R1+44,26);
    S.pickIds(ctx,R1,true);
  }},

{ id:'B', name:'ONE RAIL', sub:'a single image across A1:E1 — 799×56. One pitch, no seam, no second object. Row 2 handed back.',
  draw(ctx){
    ctx.fillStyle=S.INK; ctx.fillRect(0,0,W,R1);
    ctx.fillStyle=S.INK; ctx.fillRect(0,R1,W,R2);
    board(ctx,0,0,799,56,lockup('HQ MOTOR SERVICE',0.90));
    prose(ctx,X.F,COL.F+COL.G,'the floor is asleep','waiting: 14');
    prose(ctx,X.H-2,COL.H,'RESTING','4:33 PM',LAMP);
    S.ebay(ctx,399,R1+44,30);
    S.pickIds(ctx,R1,true);
  }},

{ id:'C', name:'BLOCK + PANEL', sub:'image A1:C2 only — 260×121, the aspect a roundel actually wants. D1:H1 becomes an instrument row.',
  draw(ctx){
    ctx.fillStyle=S.INK; ctx.fillRect(0,0,W,R1);
    ctx.fillStyle=S.INK; ctx.fillRect(0,R1,W,R2);
    board(ctx,0,0,260,121,roundel(0.78));
    panel(ctx);
    S.ebay(ctx,X.D+80,R1+44,30);
    S.pickIds(ctx,R1,true);
  }},

{ id:'D', name:'BADGE', sub:'image A1:C1 only — 260×56. Row 1 becomes the sheet’s own cockpit; row 2 goes fully back to the eBay zone.',
  draw(ctx){
    ctx.fillStyle=S.INK; ctx.fillRect(0,0,W,R1);
    ctx.fillStyle=S.INK; ctx.fillRect(0,R1,W,R2);
    board(ctx,0,0,260,56,roundel(0.94));
    panel(ctx);
    S.ebay(ctx,300,R1+44,30);
    S.pickIds(ctx,R1,true);
  }},

{ id:'F', name:'PANEL + NAMEPLATE BELOW', sub:'row 1 carries NO image — five readouts across all 1136px. The board moves down to A2:E2 (799×65, taller than row 1) and the eBay mark floats on it.',
  draw(ctx){
    ctx.fillStyle=S.INK; ctx.fillRect(0,0,W,R1+R2);
    stat(ctx,0,260,['14 TO GRAB'],19,'#ffd400');
    stat(ctx,X.D,COL.D,['oldest 10h 55m'],13,'#e0a97a');
    stat(ctx,X.E,COL.E,['4:33 PM · resting · 4m ago'],11.5,DIM,LAMP);
    stat(ctx,X.F,COL.F+COL.G,['the floor is asleep','eBay 7 · Direct 7'],11.5,QUIET);
    stat(ctx,X.H,COL.H,['shipped','0 today'],11.5,QUIET);
    board(ctx,0,R1,799,R2,lockup('HQ MOTOR SERVICE',0.60));
    S.ebay(ctx,690,R1+42,24);
    S.pickIds(ctx,R1,true);
  }},

{ id:'E', name:'FULL BOARD', sub:'one image across A1:E2 — 799×121. Covers the eBay cell; 1.6× the pixels of the current pair.',
  draw(ctx){
    ctx.fillStyle=S.INK; ctx.fillRect(0,0,W,R1+R2);
    board(ctx,0,0,799,121,lockup('HQ MOTOR SERVICE',0.62));
    prose(ctx,X.F,COL.F+COL.G,'the floor is asleep','waiting: 14');
    prose(ctx,X.H-2,COL.H,'RESTING','4:33 PM',LAMP);
    S.pickIds(ctx,R1,true);
  }},
];

(async()=>{
  LOGO=await loadImage(TMP);
  const CTXH=R1+R2+S.R3+S.RD*3+34;
  for(const o of OPTS){
    const cv=createCanvas(W,CTXH), ctx=cv.getContext('2d');
    ctx.fillStyle='#ffffff'; ctx.fillRect(0,0,W,CTXH);
    o.draw(ctx);
    S.drawTable(ctx,R1+R2);
    fs.writeFileSync(path.join(__dirname,'renders',`opt-${o.id}.png`),cv.toBuffer('image/png'));
    console.log(`  opt-${o.id}.png  ${o.name.padEnd(14)} ${W}x${CTXH} @1:1`);
  }
  try{fs.unlinkSync(TMP);}catch(e){}
})();
