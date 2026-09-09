/**
 * discs.js — a disc renderer with the STYLE pulled out as data, so treatments can be
 * compared at true size instead of argued about.
 *
 * ⭐ AREA-SAMPLED, not grid-composed. The shipped renderer draws the composition into a
 *   cols x rows canvas (text at ~18px on an 18-row canvas) and reads one pixel per disc.
 *   Here the composition is drawn at FULL pixel size and each disc averages the alpha under
 *   its own footprint — which is what a real display does, and it survives fine strokes.
 */
'use strict';
const { createCanvas } = require('@napi-rs/canvas');

const BASE = {
  pitch: 4, r: 0.405, lattice: 'square',
  ground: ['#26221c', '#141210', '#100e0c'],
  offHi: '#26221c', offLo: '#15120f',
  onHi: '#f4f0e4', onLo: '#cdc7b6',
  accHi: '#ffdc00', accLo: '#d9ab00',
  shadow: 0.55, vignette: 0, threshold: 0.42
};

/** disc centres for a style over a w x h canvas */
function lattice(w, h, st) {
  const P = st.pitch;
  const vy = st.lattice === 'hex' ? P * 0.866 : P;
  const cols = Math.floor(w / P), rows = Math.floor(h / vy);
  const x0 = (w - cols * P) / 2 + P / 2;
  const y0 = (h - rows * vy) / 2 + vy / 2;
  const pts = [];
  for (let y = 0; y < rows; y++) {
    const off = (st.lattice === 'hex' && y % 2) ? P / 2 : 0;
    for (let x = 0; x < cols; x++) {
      let cx = x0 + x * P + off;
      if (cx > w - P * 0.3) cx -= cols * P;      // wrap the odd row's overhang
      pts.push({ x: cx, y: y0 + y * vy, r: y, c: x });
    }
  }
  return { pts, cols, rows };
}

/** average alpha (and hue vote) under each disc → 0 | 1 | 2 */
function sampleField(w, h, compose, st) {
  const cv = createCanvas(w, h), c = cv.getContext('2d');
  c.clearRect(0, 0, w, h);
  compose(c, w, h);
  const px = c.getImageData(0, 0, w, h).data;
  const { pts, cols, rows } = lattice(w, h, st);
  const R = Math.max(1, Math.round(st.pitch * 0.5));
  return pts.map(p => {
    let a = 0, n = 0, acc = 0;
    for (let dy = -R; dy <= R; dy++) for (let dx = -R; dx <= R; dx++) {
      const X = Math.round(p.x) + dx, Y = Math.round(p.y) + dy;
      if (X < 0 || Y < 0 || X >= w || Y >= h) continue;
      const i = (Y * w + X) * 4;
      a += px[i + 3]; n++;
      if (px[i + 3] > 120 && px[i] > 150 && px[i + 1] > 110 && px[i + 2] < 120) acc++;
    }
    const lit = n && (a / n / 255) > st.threshold;
    return { ...p, v: lit ? (acc > n * 0.25 ? 2 : 1) : 0 };
  }).concat([{ cols, rows, meta: true }]);
}

function paint(ctx, w, h, field, st, ox, oy) {
  ox = ox || 0; oy = oy || 0;
  const g = st.ground.length > 1
    ? (() => { const q = ctx.createLinearGradient(0, oy, 0, oy + h);
        q.addColorStop(0, st.ground[0]);
        if (st.ground.length === 3) q.addColorStop(0.14, st.ground[1]);
        q.addColorStop(1, st.ground[st.ground.length - 1]); return q; })()
    : st.ground[0];
  ctx.fillStyle = g; ctx.fillRect(ox, oy, w, h);

  const R = st.pitch * st.r;
  for (const p of field) {
    if (p.meta) continue;
    // corner vignette — keep discs out of the corners so the rectangle stops announcing itself
    if (st.vignette) {
      const nx = (p.x / w - 0.5) * 2, ny = (p.y / h - 0.5) * 2;
      const d = Math.hypot(nx * (h / w > 0.3 ? 1 : 0.35), ny);
      if (d > st.vignette) continue;
    }
    const cx = ox + p.x, cy = oy + p.y, on = p.v;
    if (on && st.shadow) {
      ctx.beginPath(); ctx.arc(cx, cy + 0.5, R, 0, 7);
      ctx.fillStyle = `rgba(0,0,0,${st.shadow})`; ctx.fill();
    }
    const acc = on === 2;
    const lg = ctx.createLinearGradient(0, cy - R, 0, cy + R);
    lg.addColorStop(0, acc ? st.accHi : (on ? st.onHi : st.offHi));
    lg.addColorStop(1, acc ? st.accLo : (on ? st.onLo : st.offLo));
    ctx.beginPath(); ctx.arc(cx, cy, R, 0, 7); ctx.fillStyle = lg; ctx.fill();
  }
}
module.exports = { BASE, lattice, sampleField, paint };

/**
 * ⭐⭐ THE FLIP. A flip-disc does not SNAP — it rotates about a horizontal axis, passing
 *    through edge-on where it is a thin line, and shows its other face on the way out. That
 *    rotation is the whole charm of the medium and it is what "not smooth" was pointing at:
 *    the old wipe threw discs from one state to the other in a single frame.
 *
 * ⚠ ONLY DISCS THAT CHANGE FLIP. A real board leaves the rest alone — and it is also far
 *   cheaper, because an unchanged disc contributes no delta for the encoder to carry.
 * ⚠ THE WAVE FRONT IS STAGGERED, NOT A LINE. Each disc takes its cue from a diagonal plus a
 *   deterministic per-disc jitter, so the front reads as a board catching up with itself
 *   rather than a ruler being dragged across it.
 */
function paintFlip(ctx, w, h, A, B, t, st, ox, oy, o) {
  o = o || {};
  const dur   = o.dur   === undefined ? 0.34 : o.dur;    // one disc's flip, as a fraction of t
  const skew  = o.skew  === undefined ? 0.62 : o.skew;   // how much the diagonal leads
  const jit   = o.jit   === undefined ? 0.10 : o.jit;    // stagger, so the front is soft
  ox = ox || 0; oy = oy || 0;

  const g = st.ground.length > 1
    ? (() => { const q = ctx.createLinearGradient(0, oy, 0, oy + h);
        q.addColorStop(0, st.ground[0]); q.addColorStop(1, st.ground[st.ground.length - 1]);
        return q; })()
    : st.ground[0];
  ctx.fillStyle = g; ctx.fillRect(ox, oy, w, h);

  const R = st.pitch * st.r;
  const span = w + h * skew;
  const face = (v, hi) => v === 2 ? (hi ? st.accHi : st.accLo)
                        : v       ? (hi ? st.onHi  : st.onLo)
                                  : null;                // an unlit disc is the ground: undrawn

  for (let i = 0; i < A.length; i++) {
    const p = A[i]; if (p.meta) continue;
    const a = p.v, b = B[i].v;
    let v = a, ry = R;

    /* ⭐ THE REFRESH. `all` flips EVERY disc, including the ones whose face does not change,
       so a board can riffle in place and land on exactly what it was showing. That is what a
       real departure board does when it re-reads itself, and it is the only beat here that
       needs no new words to be worth watching. */
    if (a !== b || o.all) {
      // deterministic jitter — same every render, so the loop is reproducible
      const n = ((p.c * 73856093) ^ (p.r * 19349663)) >>> 0;
      /* ⭐⭐ THE CUT VOCABULARY. A film does not cut the same way twice, and on a disc board
         the CUT is the only place personality can live — the states themselves are just
         shapes. Each shape decides WHEN a given disc takes its turn:
           diagonal — a wave along the band's length          (the default)
           radial   — centre outward; for a round mark
           columns  — column by column, left to right, the way a real departure board
                      updates in reading order. The most Solari of the four.
           dissolve — no order at all; every disc on its own clock. Reads as a fade. */
      const rnd = (n % 1000) / 1000;
      let pos;
      switch (o.shape) {
        case 'radial':
          pos = Math.hypot((p.x - w / 2) / (w / 2), (p.y - h / 2) / (h / 2)) / Math.SQRT2; break;
        case 'columns':
          pos = p.x / w; break;
        case 'dissolve':
          pos = rnd; break;
        default:
          pos = (p.x + p.y * skew) / span;
      }
      const lead = pos * (1 - dur) + (o.shape === 'dissolve' ? 0 : (rnd - 0.5) * jit);
      const u = Math.max(0, Math.min(1, (t - lead) / dur));
      if (u > 0 && u < 1) {
        /* ⭐ THE RIFFLE. A split-flap runs through the whole alphabet before it lands, and
           that search is the sound and the sight of a Solari board changing. A flip-dot has
           only two faces, so the equivalent is to let the disc flutter a few times before it
           settles. `riffle` is how many half-turns it takes to get there — 1 is a plain flip. */
        /* The physics, so this stays right if anyone touches it: the disc sweeps
           θ = u · turns · π radians. Its projected height is |cos θ| — full at 0, zero at
           θ=π/2 where it is edge-on, full again at π. You see face A through the first
           quarter-turn and face B through the second, so the face index is θ/(π/2) floored.
           At u→1 that index is 2·turns−1, which is ODD for ANY turns, so the disc always
           lands on b. (An earlier cut swapped on floor(phase) and landed back on a.) */
        const turns = o.riffle || 1;
        const theta = Math.PI * u * turns;
        ry = Math.max(0.35, R * Math.abs(Math.cos(theta)));
        v = (Math.floor(u * turns * 2) % 2 === 0) ? a : b;
      } else v = u <= 0 ? a : b;
    }
    if (!v) continue;

    const cx = ox + p.x, cy = oy + p.y;
    if (st.shadow) {
      ctx.beginPath(); ctx.ellipse(cx, cy + 0.5, R, ry, 0, 0, 7);
      ctx.fillStyle = `rgba(0,0,0,${st.shadow})`; ctx.fill();
    }
    const lg = ctx.createLinearGradient(0, cy - ry, 0, cy + ry);
    lg.addColorStop(0, face(v, true)); lg.addColorStop(1, face(v, false));
    ctx.beginPath(); ctx.ellipse(cx, cy, R, ry, 0, 0, 7);
    ctx.fillStyle = lg; ctx.fill();
  }
}
module.exports.paintFlip = paintFlip;
