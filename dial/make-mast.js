/**
 * make-mast.js — the masthead art, both boards, one style.
 *
 * ⭐⭐ THE BLEND: ground is the sheet's own BRAND.ink #1a1a1a, flat, and unlit discs ARE that
 *    colour so they are never drawn. The file has no boundary to see.
 * ⭐⭐ THE FLIP: discs rotate through edge-on rather than snapping. See discs.paintFlip.
 * ⭐ THE WAVE MATCHES THE SHAPE — the block (a round mark) flips radially from its centre,
 *   the strip (a band) flips along its length.
 * ⭐ THE REST IS ONE FRAME. GIF carries a per-frame delay up to 655s, so a long still costs
 *   nothing; all the bytes go into the motion, where they show.
 *
 * ⚠⚠ NO ROUNDEL IN THE STRIP. At 26 disc rows the double ring and the H|Q inside it cannot
 *    resolve, and it sat 260px from the block's copy of the same mark rendered at 44 rows —
 *    the same thing done twice, once badly. The block owns the mark; the strip owns the words.
 * ⚠⚠ rgb565 CANNOT REPRESENT #1a1a1a — the palette entry nearest the ground is snapped to the
 *    exact band colour after quantising, or the image comes back a different black.
 */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {GIFEncoder,quantize,applyPalette}=require('gifenc');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const D=require('./discs.js');
registerFonts();

const BAND='#1a1a1a';
const ST={...D.BASE,pitch:3,lattice:'hex',ground:[BAND],offHi:BAND,offLo:BAND,
          onHi:'#ddd6c6',onLo:'#a49d8c'};
const FPS=24, FLIP=1.15, REST=11000, HOLD=2800;
const TMP=path.join(__dirname,'renders','_mm.png');
execFileSync('rsvg-convert',['-w','1400','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);

/** one line of Oswald, sized up to the height ceiling then shrunk to fit the width */
const line=(t,hs,wf)=>(c,w,h)=>{
  let z=Math.round(h*(hs||0.72));
  while(z>7){c.font='600 '+z+'px Oswald'; if(c.measureText(t).width<=w*(wf||0.88))break; z--;}
  c.fillStyle='#fff';c.textAlign='center';c.textBaseline='middle';c.fillText(t,w/2,h/2+1);
};
const stack=(lines,hs)=>(c,w,h)=>{
  let z=Math.round(h*(hs||0.26));
  while(z>7){c.font='600 '+z+'px Oswald';
    if(Math.max(...lines.map(l=>c.measureText(l).width))<=w*0.82)break; z--;}
  c.fillStyle='#fff';c.textAlign='center';c.textBaseline='middle';
  const lh=z*1.16, y0=h/2-lh*(lines.length-1)/2;
  lines.forEach((l,i)=>c.fillText(l,w/2,y0+i*lh));
};

function build(o){
  const {W,H,states,radial,out}=o;
  const cv=createCanvas(W,H), ctx=cv.getContext('2d');
  const F=states.map(s=>D.sampleField(W,H,s.draw,ST));

  const still=(f)=>{ D.paint(ctx,W,H,f,ST,0,0); return ctx.getImageData(0,0,W,H).data; };
  const tween=(a,b,t)=>{ D.paintFlip(ctx,W,H,a,b,t,ST,0,0,{radial}); 
                         return ctx.getImageData(0,0,W,H).data; };

  // ⚠ the palette must see EVERY state — a colour that only exists in state 2 is otherwise
  //   snapped to the nearest tone in state 1 and silently never appears (the 09-03 bug).
  const samples=F.map(still);
  const both=new Uint8Array(samples.reduce((n,s)=>n+s.length,0));
  let off=0; for(const s of samples){ both.set(s,off); off+=s.length; }
  const pal=quantize(both,6,{format:'rgb565'});
  const G=[0x1a,0x1a,0x1a];
  let gi=0,gd=Infinity;
  pal.forEach((c,i)=>{const d=(c[0]-G[0])**2+(c[1]-G[1])**2+(c[2]-G[2])**2; if(d<gd){gd=d;gi=i;}});
  pal[gi]=[...G,...(pal[gi].length>3?[pal[gi][3]]:[])];

  const gif=GIFEncoder(); let n=0;
  const emit=(data,delay)=>{ gif.writeFrame(applyPalette(data,pal,'rgb565'),W,H,
    {palette:n===0?pal:undefined, delay, repeat:n===0?0:undefined}); n++; };

  const nf=Math.round(FLIP*FPS), step=Math.round(1000/FPS);
  let ms=0;
  F.forEach((f,i)=>{
    const hold=states[i].hold;
    emit(samples[i],hold); ms+=hold;                      // the still — ONE frame
    const nxt=F[(i+1)%F.length];
    for(let k=1;k<=nf;k++){ emit(tween(f,nxt,k/nf),step); ms+=step; }
  });
  gif.finish();
  const buf=Buffer.from(gif.bytes());
  fs.writeFileSync(path.join(__dirname,'renders',out),buf);
  console.log(`  ${out.padEnd(16)} ${W}x${H}  ${n} frames  ${(buf.length/1024).toFixed(0)} KB  `
    +`loop ${(ms/1000).toFixed(1)}s`);
}

(async()=>{
  const LOGO=await loadImage(TMP);
  const roundel=(c,w,h)=>{const d=h*0.80; c.drawImage(LOGO,(w-d)/2,(h-d)/2,d,d);};

  build({ W:260, H:133, radial:true, out:'banner-v6.gif', states:[
    { draw:roundel,                                  hold:REST },
    { draw:stack(['HQ','MOTOR','SERVICE'],0.25),     hold:HOLD },
  ]});
  build({ W:539, H:68, radial:false, out:'strip-v4.gif', states:[
    { draw:line('HQ MOTOR SERVICE',0.74),            hold:REST },
    { draw:line('HOUSTON, TEXAS',0.74),              hold:HOLD },
  ]});
  try{fs.unlinkSync(TMP);}catch(e){}
})();
