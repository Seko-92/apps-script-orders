/** sheet-ctx.js — draws the REAL sheet around row one, at true 1:1, so the masthead is
 *  judged where it lives: above a dense table, at a glance. Measured widths 2026-09-02. */
'use strict';
const COL={A:103,B:70,C:87,D:232,E:307,F:130,G:100,H:107};
const X={}; { let x=0; for(const k of 'ABCDEFGH'){X[k]=x; x+=COL[k];} }
const W=1136, R1=56, R2=65, R3=36, RD=30;
const INK='#1a1a1a', CREAM='#fff8e7', PAPER='#ffffff', YEL='#ffd400';
const HEADS=[['A','◆ SKU'],['B','# QTY'],['C','LOCATION'],['D','SALES ORDER'],
             ['E','NOTE'],['F','STATUS'],['G','◫ HAND'],['H','◪ LEFT']];
const DATA=[
 ['157860','1','D-36','27-15085-56706','Ship Without Head Gasket','PENDING','0',''],
 ['194622','3','E-16','27-15085-56706','↳ from KIT-157860','PENDING','36',''],
 ['168111','1','L-70/C-63','27-15085-56706','↳ from KIT-157860','PENDING','12',''],
];
/** everything BELOW row one — header band + banded data rows */
function drawTable(ctx, oy){
  ctx.fillStyle=INK; ctx.fillRect(0,oy,W,R3);
  ctx.fillStyle=YEL; ctx.font='600 10px Oswald'; ctx.letterSpacing='0.6px';
  for(const [c,t] of HEADS){
    ctx.textAlign = (c==='A'||c==='C'||c==='D') ? 'left' : 'center';
    ctx.fillText(t, ctx.textAlign==='left'?X[c]+10:X[c]+COL[c]/2, oy+R3/2+4);
  }
  ctx.letterSpacing='0px'; ctx.textAlign='left';
  let y=oy+R3;
  DATA.forEach((r,i)=>{
    ctx.fillStyle=i%2?CREAM:PAPER; ctx.fillRect(0,y,W,RD);
    ctx.fillStyle='#ffcdd2'; ctx.fillRect(X.F,y,COL.F,RD);
    r.forEach((v,j)=>{
      const c='ABCDEFGH'[j];
      if(!v) return;
      const mono = j===0||j===3;
      ctx.font = c==='F' ? '600 10px Oswald'
        : (mono ? '600 10.5px "JetBrains Mono"' : '400 10.5px "Noto Sans"');
      ctx.fillStyle = c==='F' ? '#b71c1c' : (c==='D' ? '#1d1d1b' : '#1d1d1b');
      ctx.textAlign = (c==='A'||c==='C') ? 'left' : 'center';
      ctx.fillText(v, ctx.textAlign==='left'?X[c]+12:X[c]+COL[c]/2, y+RD/2+4);
    });
    ctx.textAlign='left'; y+=RD;
  });
  // the DIRECT band, so the yellow that the masthead has to live with is in frame
  ctx.fillStyle=YEL; ctx.fillRect(0,y,W,34);
  ctx.fillStyle='#1a1a1a'; ctx.font='600 15px Oswald'; ctx.letterSpacing='1px';
  ctx.fillText('▌  DIRECT', 12, y+22);
  ctx.font='600 8.5px Oswald'; ctx.textAlign='right';
  ctx.fillText('HQMS · DIRECT ORDERS · 7 waiting', W-12, y+21);
  ctx.textAlign='left'; ctx.letterSpacing='0px';
  return y+34;
}
/** the eBay wordmark, drawn where a caller wants it */
function ebay(ctx,cx,cy,size){
  ctx.font=`700 ${size}px "Noto Sans"`; ctx.textAlign='left'; ctx.textBaseline='alphabetic';
  const chs=[['e','#e53238'],['b','#0064d2'],['a','#f5af02'],['y','#86b817']];
  let tw=0; for(const [ch] of chs) tw+=ctx.measureText(ch).width;
  let x=cx-tw/2;
  for(const [ch,col] of chs){ ctx.fillStyle=col; ctx.fillText(ch,x,cy); x+=ctx.measureText(ch).width; }
}
/** the two Pick ID cells — never covered, in every option */
function pickIds(ctx,oy,dark){
  for(const [x,w,lab,val] of [[X.F,COL.F+COL.G,'PICK ID · SHIPPING','Shipping - Yassin 1'],
                              [X.H,COL.H,'ADJUSTMENT','AShamma 2']]){
    ctx.strokeStyle=dark?'#2e2a25':'#e8dfc8'; ctx.lineWidth=1;
    ctx.beginPath(); ctx.moveTo(x+0.5,oy); ctx.lineTo(x+0.5,oy+R2); ctx.stroke();
    ctx.fillStyle=dark?'#8a8272':'#9a9280'; ctx.font='600 7.5px Oswald'; ctx.letterSpacing='1.2px';
    ctx.fillText(lab,x+10,oy+22); ctx.letterSpacing='0px';
    ctx.fillStyle=dark?'#ffd400':'#1a1a1a'; ctx.font='400 11px Oswald';
    ctx.fillText(val,x+10,oy+42);
  }
}
module.exports={COL,X,W,R1,R2,R3,RD,INK,CREAM,PAPER,YEL,drawTable,ebay,pickIds};
