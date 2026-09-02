/**
 * row-one-options.js — B vs C for extending the ambient loop across D1:H1.
 *
 * ⚠ WIDTHS: A/B/C are MEASURED (103+70+87=260). D..H are NOT — they come from
 *   banner-mock's constants and Schema's layout comment, which DISAGREE
 *   (D 232 vs 233, E 307 vs 306, F:H 337 vs 344). Both are comments, and this
 *   project has already lost a session to a stale width constant. diagnoseBanner()
 *   must confirm before any art ships.
 */
'use strict';
const fs = require('fs'), path = require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const { fmtMins, fmtClock } = require('./src/draw');
const A = require('./src/ambient');
const P = require('./src/patterns');
const M = require('./src/board-matrix');
registerFonts();

const COL = { A:103, B:70, C:87, D:232, E:307, F:130, G:100, H:107 };
const DIALW = COL.A + COL.B + COL.C;                 // 260  MEASURED
const DH    = COL.D + COL.E + COL.F + COL.G + COL.H; // 876  UNVERIFIED
const TOTAL = DIALW + DH;
const R1 = 56, R2 = 65;
const BAND='#1a1a1a', CREAM='#fff8e7', QUIET='#e8e8e8', YEL='#ffd400', REST_ACCENT='#7e8894';
const DAY = [0,0,0,0,0,0,0,0,1,6,14,19,11,16,22,13,7,2,0,0,0,0,0,0];
const upTo = (h) => DAY.map((n,i)=> i<=h ? n : 0);

/** One flip-disc field, drawn to fill w x h exactly — same recipe as make-shuffle's frame(). */
function discs(ctx, x, y, w, h, patName, ph, S) {
  const { cols, rows } = A.gridSize(w, h);
  const mark   = A.markField(cols, rows, [{ text:'HQ', size: Math.min(19, rows*0.62), weight:'600' }]);
  const TSTRIP = P.tickerStrip(createCanvas, cols, rows, 'HQ MOTOR SERVICE · HOUSTON · ');
  const fromDraw = (fn) => {
    const off = createCanvas(cols, rows), c = off.getContext('2d');
    c.clearRect(0,0,cols,rows); fn(c);
    const px = c.getImageData(0,0,cols,rows).data, f=[];
    for (let ry=0; ry<rows; ry++){ const r=[]; for(let rx=0; rx<cols; rx++) r.push(px[(ry*cols+rx)*4+3]>110); f.push(r); }
    return f;
  };
  const FIELD = {
    piston:()=>fromDraw(c=>P.piston(c,cols,rows,ph)),
    ticker:()=>fromDraw(c=>P.ticker(c,cols,rows,ph,null,TSTRIP)),
    mark:  ()=>A.composeAmbient('mark',ph,cols,rows,mark),
    belt:  ()=>fromDraw(c=>P.belt(c,cols,rows,ph)),
    night: ()=>fromDraw(c=>P.night(c,cols,rows,ph)),
    inlinefour:()=>fromDraw(c=>P.inlineFour(c,cols,rows,ph)),
    wave:  ()=>fromDraw(c=>P.wave(c,cols,rows,ph)),
    aisle: ()=>fromDraw(c=>P.aisle(c,cols,rows,ph)),
    sweep: ()=>fromDraw(c=>P.sweep(c,cols,rows,ph))
  };
  const p = createCanvas(w*S, h*S), pc = p.getContext('2d');
  const g = pc.createLinearGradient(0,0,0,h*S);
  g.addColorStop(0,'#26221c'); g.addColorStop(0.14,'#141210'); g.addColorStop(1,'#100e0c');
  pc.fillStyle=g; pc.fillRect(0,0,p.width,p.height);
  M.drawMatrix(pc, { scale:S, w, h, field: FIELD[patName]() });
  ctx.drawImage(p, x*S, y*S);
}

const clock12 = (t) => { const m=(+String(t).slice(0,-2))*60 + (+String(t).slice(-2));
                         return fmtClock(m) + (m<720?' AM':' PM'); };
function headline(q){ const sp='eBay '+q.eb+' · Direct '+q.di;
  if(q.s==='rest')  return ['the floor is asleep', q.g>0?'waiting: '+q.g:''];
  if(q.s==='busy')  return ['picking', sp];
  return ['all caught up', sp]; }
function pulse(q){
  if(q.s==='rest') return ['#9aa3ad','⚪ RESTING · '+clock12(q.t)+' · '+fmtMins(+q.y)+' ago'];
  return ['#7ec98a','🟢 ALIVE · '+clock12(q.t)+' · '+fmtMins(+q.y)+' ago']; }

function readouts(ctx,q,S,parts){
  const s=(v)=>v*S;
  if(parts.headline){
    const ln=headline(q).filter(Boolean);
    ctx.fillStyle=QUIET; ctx.font=`400 ${s(12.5)}px Oswald`;
    ln.forEach((t,i)=>ctx.fillText(t, s(DIALW+16), s(ln.length===1?33:24+i*16)));
  }
  if(parts.pulse){
    const [lamp,txt]=pulse(q), ex=DIALW+COL.D;
    ctx.fillStyle=lamp; ctx.beginPath(); ctx.arc(s(ex+14),s(28),s(4.5),0,Math.PI*2); ctx.fill();
    ctx.fillStyle=QUIET; ctx.font=`400 ${s(11.5)}px Oswald`; ctx.fillText(txt,s(ex+25),s(32));
  }
  if(parts.curve){
    const fx=DIALW+COL.D+COL.E, fw=COL.F+COL.G+COL.H;
    const vals=q.s==='rest'?DAY:upTo(Math.floor(+String(q.t).slice(0,-2)));
    const max=Math.max(1,...vals), bw=(fw-28)/24;
    ctx.fillStyle=q.s==='rest'?REST_ACCENT:YEL;
    vals.forEach((v,i)=>{ const h=Math.max(v>0?1.5:0,(v/max)*36);
      if(h>0) ctx.fillRect(s(fx+14+i*bw), s(48-h), s(bw-1.6), s(h)); });
    ctx.fillStyle='#8a8f98'; ctx.font=`500 ${s(8)}px Oswald`; ctx.letterSpacing=`${s(1.3)}px`;
    ctx.fillText(q.s==='rest'?'YESTERDAY':'TODAY', s(fx+14), s(15)); ctx.letterSpacing='0px';
  }
  // OPTION C2 — headline in F1:G1, pulse alone in H1. Both are MIRRORS reading the same
  // __SparkData cells E1 reads; E1 itself keeps its real formula, covered but alive.
  if(parts.split){
    const fx=DIALW+COL.D+COL.E, hx=fx+COL.F+COL.G;
    const [lamp,ptxt]=pulse(q), hl=headline(q).filter(Boolean);
    ctx.fillStyle=lamp; ctx.beginPath(); ctx.arc(s(fx+13),s(28),s(4),0,Math.PI*2); ctx.fill();
    ctx.fillStyle=QUIET; ctx.font=`400 ${s(11.5)}px Oswald`;
    ctx.fillText(hl.join(' · '), s(fx+23), s(32));
    ctx.strokeStyle='#33302b'; ctx.lineWidth=Math.max(1,s(1));
    ctx.beginPath(); ctx.moveTo(s(hx),s(8)); ctx.lineTo(s(hx),s(R1-8)); ctx.stroke();
    ctx.fillStyle=lamp; ctx.font=`500 ${s(9)}px Oswald`; ctx.letterSpacing=`${s(1.1)}px`;
    ctx.fillText(ptxt.replace(/^[^A-Z]*/,'').split(' · ')[0], s(hx+11), s(24));
    ctx.letterSpacing='0px';
    ctx.fillStyle=QUIET; ctx.font=`400 ${s(10.5)}px Oswald`;
    ctx.fillText(clock12(q.t), s(hx+11), s(40));
  }
  // OPTION C1 — one consolidated readout living in the existing F1:H1 merge.
  if(parts.combined){
    const fx=DIALW+COL.D+COL.E;
    const [lamp,ptxt]=pulse(q), hl=headline(q).filter(Boolean);
    ctx.fillStyle=lamp; ctx.beginPath(); ctx.arc(s(fx+14),s(21),s(4),0,Math.PI*2); ctx.fill();
    ctx.fillStyle=QUIET; ctx.font=`400 ${s(11.5)}px Oswald`;
    ctx.fillText(hl.join(' · '), s(fx+25), s(25));
    ctx.fillStyle='#9a9daa'; ctx.font=`400 ${s(10)}px Oswald`;
    ctx.fillText(ptxt.replace(/^[⚪🟢]\s*/,''), s(fx+25), s(42));
  }
}

function row2(ctx,S){
  const s=(v)=>v*S;
  ctx.fillStyle=CREAM; ctx.fillRect(s(DIALW), s(R1), s(DH), s(R2));
  ctx.fillStyle='#111'; ctx.font=`700 ${s(26)}px "Noto Sans"`;
  let lx=DIALW+18;
  for(const [ch,col] of [['e','#e53238'],['b','#0064d2'],['a','#f5af02'],['y','#86b817']]){
    ctx.fillStyle=col; ctx.fillText(ch,s(lx),s(R1+44)); lx+=ctx.measureText(ch).width/S; }
  const px=DIALW+COL.D+COL.E;
  for(const [x,w,lab,val] of [[px,COL.F+COL.G,'PICK ID · SHIPPING','Shipping - Yassin 1'],
                              [px+COL.F+COL.G,COL.H,'ADJUSTMENT','AShamma 2']]){
    ctx.strokeStyle='#e8dfc8'; ctx.lineWidth=Math.max(1,s(1));
    ctx.beginPath(); ctx.moveTo(s(x),s(R1)); ctx.lineTo(s(x),s(R1+R2)); ctx.stroke();
    ctx.fillStyle='#9a9280'; ctx.font=`600 ${s(7.5)}px Oswald`; ctx.letterSpacing=`${s(1.2)}px`;
    ctx.fillText(lab,s(x+10),s(R1+22)); ctx.letterSpacing='0px';
    ctx.fillStyle='#1a1a1a'; ctx.font=`400 ${s(11)}px Oswald`; ctx.fillText(val,s(x+10),s(R1+42));
  }
}

/** opt: 'now' | 'B' | 'C' */
function drawRow(ctx,q,S,opt,pat){
  const s=(v)=>v*S;
  ctx.fillStyle=BAND; ctx.fillRect(0,0,s(TOTAL),s(R1));
  row2(ctx,S);
  discs(ctx,0,0,DIALW,R1+R2,pat,q.ph,S);              // A1:C2 — unchanged in every option
  if(opt==='B')  discs(ctx,DIALW,0,DH,R1,pat,q.ph+0.5,S);
  if(opt!=='B' && opt!=='now') discs(ctx,DIALW,0,COL.D+COL.E,R1,pat,q.ph+0.5,S);
  readouts(ctx,q,S,
    opt==='now'? {headline:1,pulse:1,curve:1} :
    opt==='B'  ? {} :
    opt==='C1' ? {combined:1} : {split:1});
}

const S = +(process.env.SCALE||1);
const CASES=[
  ['BUSY · 2:14 PM',  {s:'busy', t:'1414', g:12, eb:9, di:3, y:3,  ph:0.30}],
  ['REST · 9:57 PM',  {s:'rest', t:'2157', g:0,  eb:0, di:0, y:8,  ph:0.62}]
];
const OPTS=[
  ['NOW — loop A1:C2 only. Headline D1, pulse E1, day curve F1:H1.','now'],
  ['C1 — ONE CELL. Loop takes D1:E1 (539px); state + split + pulse combined in the F1:H1 merge (337px).','C1'],
  ['C2 — TWO CELLS. Loop takes D1:E1; state + split in F1:G1 (230px), pulse alone in H1 (107px).','C2'],
  ['B — FULL COVER. Loop takes all of D1:H1 (876px). Every cell still computes; none is visible.','B']];

const PAD=22, LAB=18, GAP=10, CGAP=30;
const W=(TOTAL+PAD*2)*S;
const H=(PAD*2 + CASES.length*(CASES.length? (14 + OPTS.length*(LAB+R1+R2+GAP)+CGAP):0))*S;
const cv=createCanvas(W,H), ctx=cv.getContext('2d');
ctx.fillStyle='#0b0b0b'; ctx.fillRect(0,0,W,H);
let y=PAD;
for(const [cname,q] of CASES){
  ctx.fillStyle=YEL; ctx.font=`600 ${11*S}px Oswald`; ctx.letterSpacing=`${1.6*S}px`;
  ctx.fillText(cname, PAD*S, (y+10)*S); ctx.letterSpacing='0px';
  y+=14;
  for(const [label,opt] of OPTS){
    ctx.fillStyle='#7d776c'; ctx.font=`500 ${9.5*S}px Oswald`;
    ctx.fillText(label, PAD*S, (y+10)*S);
    y+=LAB;
    ctx.save(); ctx.translate(PAD*S, y*S); drawRow(ctx,q,S,opt,process.env.PAT||'inlinefour'); ctx.restore();
    y+=R1+R2+GAP;
  }
  y+=CGAP;
}
fs.writeFileSync(path.join(__dirname,'renders',`row-one-options${S===1?'':'@'+S+'x'}.png`),
                 cv.toBuffer('image/png'));
console.log(`row-one-options  ${TOTAL}px wide (A1:C2 ${DIALW} MEASURED + D:H ${DH} UNVERIFIED) at ${S}x`);
