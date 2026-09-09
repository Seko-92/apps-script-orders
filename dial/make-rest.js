/**
 * make-rest.js — THE RESTING BOARD. The mark is home; motion is an interlude.
 *
 * ⭐⭐ THE HOLD COSTS ONE FRAME. A long still stretch does NOT need N identical frames —
 *    GIF carries a per-frame delay (centiseconds, up to 655s), so "rest for 9 seconds" is a
 *    SINGLE frame with delay:9000. That is why resting is not a trade against file size:
 *    it is the cheapest thing the format can do, and the whole 811 KB problem goes with it.
 *
 * ⚠ PALETTE MUST BE SAMPLED FROM THE MARK **AND** A MOTION FRAME (the 09-03 gotcha) —
 *   gifenc quantises one sample and maps every later frame into it.
 */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {GIFEncoder,quantize,applyPalette}=require('gifenc');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const P=require('./src/patterns'); const D=require('./discs.js');
registerFonts();

/* ⭐⭐ THE BLEND IS THE BRIEF. The ground is the sheet's own BRAND.ink #1a1a1a, FLAT — no
   gradient, no edge fade — and the unlit discs are that same colour, so they are not drawn
   at all. The file therefore has no visible boundary: it is the band, plus light.
   ⚠ The old three-tone gradient (#26221c → #141210 → #100e0c) was a fourth, fifth and sixth
     black that the sheet never had, which is exactly what made the image read as a block
     pasted on top. */
const FINE={...D.BASE,pitch:3,lattice:'hex',ground:['#1a1a1a'],
            offHi:'#1a1a1a',offLo:'#1a1a1a',onHi:'#ddd6c6',onLo:'#a49d8c'};
const TMP=path.join(__dirname,'renders','_rest.png');
execFileSync('rsvg-convert',['-w','1200','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);

function build(W,H,text,scale,out,restMs){
  return loadImage(TMP).then(LOGO=>{
    const st=FINE, cv=createCanvas(W,H), ctx=cv.getContext('2d');
    const lockup=(c,w,h)=>{
      const d=h*(scale||0.90), gap=h*0.26;
      let s=Math.round(h*0.74);
      while(s>6){c.font='600 '+s+'px Oswald';
        if(d+gap+c.measureText(text).width<=w*0.90)break;s--;}
      const x0=(w-(d+gap+c.measureText(text).width))/2;
      c.drawImage(LOGO,x0,(h-d)/2,d,d);
      c.fillStyle='#fff';c.textBaseline='middle';c.textAlign='left';
      c.fillText(text,x0+d+gap,h/2+1);
    };
    /* ⚠ THE TICKER IS DRAWN AT FULL PIXEL SIZE, not at grid resolution and scaled.
       Composing at cols x rows and upscaling quantises twice — the letterforms arrive
       already chunky and the disc sampler then re-chunks them, so the moving text reads
       visibly softer than the mark that shares the board. Same lockup, same crispness. */
    const TXT='HQ MOTOR SERVICE  ·  HOUSTON, TEXAS  ·  ';
    const tick=(ph)=>(c,w,h)=>{
      c.font='600 '+Math.round(h*0.74)+'px Oswald';
      c.fillStyle='#fff'; c.textBaseline='middle'; c.textAlign='left';
      const tw=c.measureText(TXT).width;
      let x=w-((ph%1)+1)%1*tw;
      for(x-=tw; x<w; x+=tw) c.fillText(TXT,x,h/2+1);
    };
    const px=(compose)=>{ const f=D.sampleField(W,H,compose,st);
      D.paint(ctx,W,H,f,st,0,0); return ctx.getImageData(0,0,W,H).data; };

    const markPx=px(lockup);
    const tickPx=px(tick(0.3));
    const both=new Uint8Array(markPx.length*2);
    both.set(markPx,0); both.set(tickPx,markPx.length);
    /* ⚠⚠ rgb565 CANNOT REPRESENT #1a1a1a. It keeps 5 bits of red and blue, so 0x1a (26)
       lands on 24 or 27 and the ground comes back a different black — measured at #1b1b1b on
       the rail and #26251e on the block, which is the pasted-on rectangle all over again.
       The encoder reports nothing; only sampling a decoded frame finds it.
       ⭐ THE FIX: quantise as normal, then SNAP the palette entry nearest the ground to the
       exact band colour. Near-ground pixels that mapped to it were within a few levels
       anyway, so nothing else moves — and the ground is now exact by construction. */
    const pal=quantize(both,6,{format:'rgb565'});
    const G=[0x1a,0x1a,0x1a];
    let gi=0,gd=Infinity;
    pal.forEach((c,i)=>{const d=(c[0]-G[0])**2+(c[1]-G[1])**2+(c[2]-G[2])**2;
                        if(d<gd){gd=d;gi=i;}});
    pal[gi]=[...G,...(pal[gi].length>3?[pal[gi][3]]:[])];

    const gif=GIFEncoder(); let n=0;
    const emit=(data,delay)=>{ gif.writeFrame(applyPalette(data,pal,'rgb565'),W,H,
      {palette:n===0?pal:undefined, delay, repeat:n===0?0:undefined}); n++; };

    emit(markPx, restMs);                       // ← the whole rest, in ONE frame
    const FPS=10, RUN=3.2, FADE=0.5;
    const fadeF=Math.round(FADE*FPS), runF=Math.round(RUN*FPS);
    for(let i=0;i<fadeF;i++) emit(px(tick(-0.06+i/fadeF*0.06)),100);
    for(let i=0;i<runF;i++)  emit(px(tick(i/runF)),100);
    for(let i=0;i<fadeF;i++) emit(px(lockup),100);
    gif.finish();
    const buf=Buffer.from(gif.bytes());
    fs.writeFileSync(path.join(__dirname,'renders',out),buf);
    console.log(`  ${out.padEnd(18)} ${W}x${H}  ${n} frames  ${(buf.length/1024).toFixed(0)} KB   `
      +`loop ${((restMs+ (fadeF*2+runF)*100)/1000).toFixed(1)}s`);
  });
}
(async()=>{
  await build(+(process.env.SW||539),+(process.env.SH||68),'MOTOR SERVICE',0.90,process.env.SOUT||'strip-v3.gif',9000);
  try{fs.unlinkSync(TMP);}catch(e){}
})();
