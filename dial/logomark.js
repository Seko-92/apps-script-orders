/**
 * logomark.js — the mark, rendered so its own structure survives.
 *
 * ⚠⚠ THE LOGO IS DRAWN FOR A LIGHT GROUND. Its outer ring and its HQ letterforms are
 *    #1d1d1b — on our #1a1a1a band they are invisible. So a faithful render shows almost
 *    nothing, and the disc version has to be a REVERSED treatment: every shape the logo
 *    draws becomes a lit disc, and only the brand yellow becomes an accent disc.
 *
 * ⚠⚠ AND SAMPLING HUE PER DISC LOSES THE RING. The yellow ring is a thin stroke; averaged
 *    under a 3px disc footprint it reads closer to the pale tone than to #ffdc00, so the
 *    brand colour quietly drained out of the mark. Fixed by rasterising the two colour
 *    CLASSES separately — .st1 (yellow) and .st0 (dark) — and deciding each disc from the
 *    yellow mask first. Exact, and it cannot lose a thin stroke.
 */
'use strict';
const fs=require('fs'),path=require('path'),{execFileSync}=require('child_process');
const {createCanvas,loadImage}=require('@napi-rs/canvas');

const SRC=path.join(__dirname,'logo','fav-google.svg');

/** rasterise the mark with ONE colour class visible */
async function classMask(keep, px){
  const svg=fs.readFileSync(SRC,'utf8');
  const drop = keep==='st1' ? 'st0' : 'st1';
  const out=svg.replace(new RegExp('\\.'+drop+'\\s*\\{[^}]*\\}'), '.'+drop+' { fill: none; }');
  const tmp=path.join(__dirname,'renders','_lm-'+keep+'.svg');
  const png=path.join(__dirname,'renders','_lm-'+keep+'.png');
  fs.writeFileSync(tmp,out);
  execFileSync('rsvg-convert',['-w',String(px),'-o',png,tmp]);
  const img=await loadImage(png);
  fs.unlinkSync(tmp);
  return img;
}

/**
 * A field for the mark at a given disc lattice.
 * @returns per-disc 0 (off) | 1 (lit) | 2 (brand yellow)
 */
async function markField(D, ST, W, H, scale){
  const px=900;
  const [yellow, dark] = await Promise.all([classMask('st1',px), classMask('st0',px)]);
  const d=H*(scale||0.80), x0=(W-d)/2, y0=(H-d)/2;
  const draw=(img)=>(c,w,h)=>{ c.drawImage(img,x0,y0,d,d); };
  const fy=D.sampleField(W,H,draw(yellow),ST);
  const fd=D.sampleField(W,H,draw(dark),ST);
  // yellow wins where it exists — the brand colour is the thing that must not be lost
  return fy.map((p,i)=> p.meta ? p : ({...p, v: p.v ? 2 : (fd[i].v ? 1 : 0)}));
}
module.exports={ classMask, markField };
