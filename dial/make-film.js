/**
 * make-film.js — row one as a short film.
 *
 * ⭐⭐ THE SHAPE OF IT. One board is the ANCHOR and one is the NARRATIVE. The block holds the
 *    mark and moves twice; the strip tells the story. That is composition, not economy — two
 *    objects both performing read as noise, and the sheet underneath is somebody's work.
 *
 * ⭐ EVERY CUT IS A DIFFERENT CUT. columns · diagonal · radial · dissolve, with a riffle count
 *   per cut. On a board whose states are only shapes, the CUT is where the personality lives.
 *
 * ⭐ TWO FRAME RATES, ON PURPOSE. Transitions run at 24fps because a rotation wants to be
 *   smooth. The engine runs at 15 — a real flip-dot cannot physically go faster, and pretending
 *   otherwise costs bytes to look less like the thing it is.
 *
 * ⭐ THE RESTS ARE ONE FRAME EACH. A 14-second hold costs a single frame, so every byte in the
 *   file is spent on motion, where it shows.
 *
 * ⚠ BOTH FILMS ARE THE SAME LENGTH. They are separate files and WILL drift apart — nothing can
 *   sync two GIFs — but equal loops mean the drift never accumulates into one board looking
 *   stuck while the other performs.
 * ⚠⚠ rgb565 cannot represent #1a1a1a; the ground entry is snapped after quantising or the
 *    blend breaks. The palette is sampled from EVERY scene, or a colour that exists in only
 *    one of them is silently mapped away.
 */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {GIFEncoder,quantize,applyPalette}=require('gifenc');
const {createCanvas,loadImage}=require('@napi-rs/canvas');
const {registerFonts}=require('./src/render');
const P=require('./src/patterns');
const D=require('./discs.js');
registerFonts();

const BAND='#1a1a1a';
const ST={...D.BASE,pitch:3,lattice:'hex',ground:[BAND],offHi:BAND,offLo:BAND,
          onHi:'#ddd6c6',onLo:'#a49d8c'};
const CUT_FPS=24, RUN_FPS=15;
const TMP=path.join(__dirname,'renders','_mf.png');
execFileSync('rsvg-convert',['-w','1400','-o',TMP,path.join(__dirname,'logo','fav-google.svg')]);

/* ── compositions ───────────────────────────────────────────────────────────────────── */
const line=(t,hs)=>(c,w,h)=>{
  let z=Math.round(h*(hs||0.74));
  while(z>7){c.font='600 '+z+'px Oswald'; if(c.measureText(t).width<=w*0.88)break; z--;}
  c.fillStyle='#fff';c.textAlign='center';c.textBaseline='middle';c.fillText(t,w/2,h/2+1);
};
const stack=(lines,hs)=>(c,w,h)=>{
  let z=Math.round(h*(hs||0.25));
  while(z>7){c.font='600 '+z+'px Oswald';
    if(Math.max(...lines.map(l=>c.measureText(l).width))<=w*0.82)break; z--;}
  c.fillStyle='#fff';c.textAlign='center';c.textBaseline='middle';
  const lh=z*1.16, y0=h/2-lh*(lines.length-1)/2;
  lines.forEach((l,i)=>c.fillText(l,w/2,y0+i*lh));
};

/* ⚠ A PATTERN DRAWS IN GRID UNITS, NOT PIXELS. patterns.js sizes its strokes off cols/rows,
     so it has to be rendered into a cols x rows canvas and read back per DISC INDEX — not
     drawn at full size and area-sampled like type, and not scaled up (which quantises twice
     and arrives already chunky). */
function patternField(name, phase, W, H) {
  const { pts, cols, rows } = D.lattice(W, H, ST);
  const off=createCanvas(cols,rows), c=off.getContext('2d');
  c.clearRect(0,0,cols,rows);
  P[name](c, cols, rows, phase);
  const px=c.getImageData(0,0,cols,rows).data;
  const f=pts.map(p=>({...p, v: px[((p.r*cols)+p.c)*4+3] > 110 ? 1 : 0}));
  f.push({cols,rows,meta:true});
  return f;
}

function film(o){
  const {W,H,scenes,out}=o;
  const cv=createCanvas(W,H), ctx=cv.getContext('2d');
  // resolve each scene to its opening field (a run scene opens on phase 0)
  const open = scenes.map(s => s.kind==='run'
    ? patternField(s.pattern, 0, W, H)
    : D.sampleField(W,H,s.draw,ST));
  const close = scenes.map((s,i) => s.kind==='run'
    ? patternField(s.pattern, (s.secs*RUN_FPS-1)/(s.secs*RUN_FPS)*s.turns, W, H)
    : open[i]);

  const still=(f)=>{ D.paint(ctx,W,H,f,ST,0,0); return ctx.getImageData(0,0,W,H).data; };

  // ⚠ every scene must be in the palette sample, or a tone unique to one of them vanishes
  const samples=open.map(still);
  const buf=new Uint8Array(samples.reduce((n,s)=>n+s.length,0));
  let at=0; for(const s of samples){ buf.set(s,at); at+=s.length; }
  const pal=quantize(buf,6,{format:'rgb565'});
  const G=[0x1a,0x1a,0x1a];
  let gi=0,gd=Infinity;
  pal.forEach((c,i)=>{const d=(c[0]-G[0])**2+(c[1]-G[1])**2+(c[2]-G[2])**2; if(d<gd){gd=d;gi=i;}});
  pal[gi]=[...G,...(pal[gi].length>3?[pal[gi][3]]:[])];

  const gif=GIFEncoder(); let n=0, ms=0;
  const emit=(data,delay)=>{ gif.writeFrame(applyPalette(data,pal,'rgb565'),W,H,
    {palette:n===0?pal:undefined, delay, repeat:n===0?0:undefined}); n++; ms+=delay; };

  scenes.forEach((s,i)=>{
    if(s.kind==='run'){
      const nf=Math.round(s.secs*RUN_FPS), step=Math.round(1000/RUN_FPS);
      for(let k=0;k<nf;k++){
        D.paint(ctx,W,H,patternField(s.pattern,(k/nf)*s.turns,W,H),ST,0,0);
        emit(ctx.getImageData(0,0,W,H).data, step);
      }
    } else {
      emit(samples[i], s.hold);
    }
    const cut=s.cut, nf=Math.round(cut.secs*CUT_FPS), step=Math.round(1000/CUT_FPS);
    const a=close[i], b=open[(i+1)%scenes.length];
    for(let k=1;k<=nf;k++){
      D.paintFlip(ctx,W,H,a,b,k/nf,ST,0,0,
                  {shape:cut.shape,riffle:cut.riffle,all:cut.all});
      emit(ctx.getImageData(0,0,W,H).data, step);
    }
  });
  gif.finish();
  const bytes=Buffer.from(gif.bytes());
  fs.writeFileSync(path.join(__dirname,'renders',out),bytes);
  console.log(`  ${out.padEnd(16)} ${(W+'x'+H).padEnd(8)} ${String(n).padStart(3)} frames  `
    +`${(bytes.length/1024).toFixed(0).padStart(4)} KB   ${(ms/1000).toFixed(1)}s`);
  return bytes.length;
}

(async()=>{
  const LOGO=await loadImage(TMP);
  const roundel=(c,w,h)=>{const d=h*0.80; c.drawImage(LOGO,(w-d)/2,(h-d)/2,d,d);};

  /* ── THE STRIP · the marquee ─────────────────────────────────────────────────────
     A 7:1 band 26 discs tall is a departure board and nothing else. Its beats are a name,
     a place, and the board re-reading itself — which is the one thing a Solari does that
     needs no new words. ⚠ The engine was HERE first and it did not read: five cylinders at
     26 rows are brackets, and the crowns move about two discs. Measured, then moved. */
  const s = film({ W:539, H:68, out:'strip-v5.gif', scenes:[
    { kind:'still', draw:line('HQ MOTOR SERVICE'), hold:15640,
      cut:{shape:'columns',  riffle:3, secs:1.6} },   // left to right, in reading order
    { kind:'still', draw:line('HOUSTON, TEXAS'),   hold:2800,
      cut:{shape:'diagonal', riffle:1, secs:1.2} },
    { kind:'still', draw:line('HQ MOTOR SERVICE'), hold:700,
      cut:{shape:'columns',  riffle:5, secs:2.0, all:true} },   // ← the refresh, in place
  ]});

  /* ── THE BLOCK · the mark, and the business ──────────────────────────────────────
     2:1 and 51 discs tall is the one shape on row one with vertical room, so the engine
     lives here: two cylinders, bores, con-rods and cranks all legible, crowns stroking a
     third of the board's height. It rests on the mark for two thirds of the loop. */
  const b = film({ W:260, H:133, out:'banner-v7.gif', scenes:[
    { kind:'still', draw:roundel, hold:14300,
      cut:{shape:'radial',   riffle:1, secs:1.2} },   // a round mark opens from its centre
    { kind:'still', draw:stack(['HQ','MOTOR','SERVICE'],0.24), hold:2400,
      cut:{shape:'diagonal', riffle:1, secs:1.0} },
    { kind:'run',   pattern:'inlineFour', secs:3.6, turns:5,
      cut:{shape:'dissolve', riffle:1, secs:1.4} },   // it runs, then dissolves back to the mark
  ]});
  console.log(`\n  total ${((s+b)/1024/1024).toFixed(2)} MB`);
  try{fs.unlinkSync(TMP);}catch(e){}
})();
