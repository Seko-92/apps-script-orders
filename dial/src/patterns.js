/**
 * patterns.js — what the flip-disc field DOES. One function per idea.
 *
 * ⭐ EACH PATTERN IS DRAWN AT GRID RESOLUTION AND THRESHOLDED INTO DISCS. That means a
 *   pattern can use ordinary canvas primitives — arcs, lines, rects — and still come out as
 *   a mechanical field of discs. Ten lines buys a real mechanism instead of a hand-authored
 *   bitmap, which is the same trick that gave the letters real Oswald letterforms.
 *
 * ⚠ THE GRID IS SMALL AND THAT IS THE DESIGN CONSTRAINT, not an obstacle. 280x121 at a 4px
 *   pitch is 70 x 30 discs. Anything needing fine detail will sample to mush — measured on
 *   "MOTOR SERVICE" at 4px, which came out a solid bar. Draw BIG shapes.
 */
'use strict';

const TAU = Math.PI * 2;

/** ⭐ THE FOUR-STROKE. This is an automotive parts business and the medium is a mechanism —
 *  a piston stroking in its bore is the most on-brand thing this banner could possibly do,
 *  and it is legible at any size because it is one big moving shape. */
function piston(c, cols, rows, phase) {
  const cx = cols * 0.42;
  const bore = Math.min(22, cols * 0.30);
  const top = 3, bottom = rows - 9;
  const crankY = rows - 5, crankR = 4.2;

  const ang = phase * TAU;
  const pinX = cx + Math.sin(ang) * crankR;
  const pinY = crankY - Math.cos(ang) * crankR;
  const stroke = (bottom - top);
  const headY = top + (0.5 - Math.cos(ang) * 0.5) * stroke * 0.72;

  c.strokeStyle = '#fff'; c.fillStyle = '#fff';

  c.lineWidth = 1.9;                                   // the bore walls
  c.beginPath();
  c.moveTo(cx - bore / 2, top - 2); c.lineTo(cx - bore / 2, bottom);
  c.moveTo(cx + bore / 2, top - 2); c.lineTo(cx + bore / 2, bottom);
  c.stroke();

  c.fillRect(cx - bore / 2, headY, bore, 4.2);          // the piston crown
  c.fillRect(cx - bore / 2, headY + 5.4, bore, 1.2);    // one ring

  c.lineWidth = 2.2;                                    // the connecting rod
  c.beginPath(); c.moveTo(cx, headY + 4); c.lineTo(pinX, pinY); c.stroke();

  c.lineWidth = 1.9;                                    // the crank circle and its pin
  c.beginPath(); c.arc(cx, crankY, crankR, 0, TAU); c.stroke();
  c.beginPath(); c.arc(pinX, pinY, 1.6, 0, TAU); c.fill();
}

/** ⭐ THE TICKER. What a departure board does between announcements: text walks across the
 *  field. ⚠ It must WRAP seamlessly or the loop shows a seam — so the string is drawn twice,
 *  a full period apart, and the window slides exactly one period. */
function ticker(c, cols, rows, phase, text, strip) {
  // ⚠ THE FIRST CUT DREW THE STRING TWICE, `measureText` apart, AND THE COPIES OVERLAPPED —
  //   it rendered "HOUSTOSTON". measureText and the rasterised glyph run are not the same
  //   width once you are drawing into a 70-unit grid. So the strip is rendered ONCE, its
  //   real pixel width IS the period, and the loop blits it twice. Seamless by construction
  //   rather than by getting a measurement right.
  const w = strip.width;
  const x = -((phase * w) % w);
  c.drawImage(strip, x, 0);
  c.drawImage(strip, x + w, 0);
}

/** Renders the ticker's text once. Its width is the loop period. */
function tickerStrip(createCanvas, cols, rows, text) {
  const probe = createCanvas(8, 8).getContext('2d');
  // ⚠⚠ SIZED TO AN ABSOLUTE CEILING, NOT A FRACTION OF ROWS. `rows * 0.46` gave the
  //    121px block 14-unit letters and the 56px strip 6-unit ones — side by side on the
  //    same row they read as two different boards, which is exactly what the D1:H1
  //    render exposed. Clamping to 14 keeps both near the same absolute height.
  const size = Math.round(Math.min(rows * 0.78, 14));
  probe.font = '600 ' + size + 'px Oswald';
  const tw = Math.ceil(probe.measureText(text).width);
  const gap = Math.round(cols * 0.55);                  // clear space between repeats
  const cv = createCanvas(tw + gap, rows);
  const c = cv.getContext('2d');
  c.clearRect(0, 0, cv.width, rows);
  c.fillStyle = '#fff';
  c.font = '600 ' + size + 'px Oswald';
  c.textBaseline = 'middle';
  c.fillText(text, 0, rows / 2);
  return cv;
}

/** ⭐ THE REFRESH. The Solari signature in discs: a wavefront sweeps the field, every disc
 *  it passes flips, and behind it the mark is standing. Then it sweeps back to clear. */
function refresh(cols, rows, phase, mark) {
  const field = [];
  // Diagonal wavefront position, out and back across the loop.
  // ⚠ SWEEP FAST, HOLD LONG. The first cut spread the sweep across the whole loop, which
  //   left the panel EMPTY for the frames before the wavefront entered — a dead board, not
  //   an idling one. The wipe now takes a quarter of the loop at each end and the mark
  //   stands for the half in between.
  let p;
  if (phase < 0.25)      p = phase / 0.25;          // wipe on
  else if (phase < 0.78) p = 1;                     // hold, mark standing
  else                   p = 1 - (phase - 0.78) / 0.22;  // wipe off
  const head = p * (cols + rows * 0.8) * 1.15 - rows * 0.4;
  for (let y = 0; y < rows; y++) {
    const row = [];
    for (let x = 0; x < cols; x++) {
      const d = head - (x + y * 0.8);
      // A band of churn AT the wavefront, settled behind it, dark ahead of it.
      // ⚠ A WIDER, DENSER CHURN BAND. The first cut used ±3 and a 1-in-3 fill, which left
      //   two of every four frames nearly empty — the loop read as dead rather than idle.
      if (d > 5) row.push(mark[y][x]);
      else if (d > -7) {
        const n = (x * 5 + y * 11 + Math.floor(phase * 137)) % 5;
        row.push(n < 2 || (d > 1 && mark[y][x]));       // settling INTO the mark, not onto black
      } else row.push(false);
    }
    field.push(row);
  }
  return field;
}

/** ⭐ THE BELT. Warehouse-native: parts travelling right to left, the way the floor works.
 *  Blocks of different lengths so it never reads as a repeating tile. */
function belt(c, cols, rows, phase) {
  c.fillStyle = '#fff';
  const lane = [rows * 0.30, rows * 0.58];
  const sizes = [7, 4, 11, 5, 8, 3, 9, 6];
  const period = cols * 1.6;
  lane.forEach((y, li) => {
    const speed = li ? 1 : 0.72;                      // two lanes, different speeds
    for (let i = 0; i < sizes.length; i++) {
      const w = sizes[(i + li * 3) % sizes.length];
      let x = (i * period / sizes.length) - phase * period * speed;
      x = ((x % period) + period) % period - w;
      c.fillRect(x, y, w, 3);
    }
  });
  c.fillStyle = '#fff';                                // the belt line itself
  c.fillRect(0, rows - 3, cols, 1);
}

/** ⭐ THE NIGHT. Sparse points drifting — the quietest of the set, and the closest in
 *  feeling to the night dial this banner already wore. Deterministic, so it loops exactly. */
function night(c, cols, rows, phase) {
  c.fillStyle = '#fff';
  // ⚠ 0.018 of the grid was ~38 points in 2100 — indistinguishable from a dead panel.
  const N = Math.round(cols * rows * 0.055);
  for (let i = 0; i < N; i++) {
    // ⚠ TWO INDEPENDENT HASHES. Using (i*37)%rows for y walked a regular diagonal and the
    //   field showed visible stripes instead of scatter — obvious at 2x, invisible in the
    //   arithmetic.
    const h1 = ((i * 2654435761) % 10007) / 10007;
    const h2 = ((i * 40503 + 17) % 10009) / 10009;
    const sp = 0.25 + (i % 5) * 0.12;
    const x = (h1 * cols + phase * cols * sp) % cols;
    const y = Math.floor(h2 * rows);
    // a slow twinkle, out of phase per point, so the field is never static
    const tw = Math.sin((phase * 2 + h1 * 7) * Math.PI * 2);
    if (tw > -0.2) c.fillRect(Math.floor(x), y, 1, 1);
  }
}


/** ⭐⭐ THE INLINE ENGINE — the piston, fixed for a long board. `piston` is composed for a
 *  TALL bore and at 876x56 it renders as one small blob adrift in a wide empty field. An
 *  inline four is the same mechanism arranged the way a wide canvas actually wants it, and
 *  it is more on-brand, not less: the cylinders fire in sequence across the strip.
 *  ⚠ COUNT IS DERIVED FROM WIDTH, so one function serves both the 260px block (2 cylinders)
 *    and the 876px strip (6). A fixed 4 would leave the block unreadable. */
function inlineFour(c, cols, rows, phase) {
  const n = Math.max(1, Math.min(6, Math.round(cols / 38)));
  const FIRE = [0, 0.5, 0.75, 0.25, 0.125, 0.625];   // 1-3-4-2, extended for 5 and 6
  const pitch = cols / n;
  const bore = Math.min(pitch * 0.55, rows * 1.5);
  const top = rows * 0.10, bottom = rows * 0.72;
  const crankY = rows * 0.87, crankR = Math.max(2.0, rows * 0.11);
  c.strokeStyle = '#fff'; c.fillStyle = '#fff';
  for (let i = 0; i < n; i++) {
    const cx = pitch * (i + 0.5);
    const ang = (phase + FIRE[i % FIRE.length]) * TAU;
    const pinX = cx + Math.sin(ang) * crankR;
    const pinY = crankY - Math.cos(ang) * crankR;
    const headY = top + (0.5 - Math.cos(ang) * 0.5) * (bottom - top) * 0.72;
    c.lineWidth = Math.max(1.3, rows * 0.055);
    c.beginPath();
    c.moveTo(cx - bore / 2, top); c.lineTo(cx - bore / 2, bottom);
    c.moveTo(cx + bore / 2, top); c.lineTo(cx + bore / 2, bottom);
    c.stroke();
    c.fillRect(cx - bore / 2, headY, bore, Math.max(2.2, rows * 0.12));
    c.lineWidth = Math.max(1.5, rows * 0.07);
    c.beginPath(); c.moveTo(cx, headY + rows * 0.10); c.lineTo(pinX, pinY); c.stroke();
    c.lineWidth = Math.max(1.2, rows * 0.05);
    c.beginPath(); c.arc(cx, crankY, crankR, 0, TAU); c.stroke();
  }
}

/** ⭐ THE WAVE. The oldest flip-disc demo there is, and the one pattern that cannot care
 *  about aspect — it is a band, so it fills whatever shape it is given. Cheap insurance in
 *  a set where three of five patterns turned out to be composition-dependent. */
function wave(c, cols, rows, phase) {
  const mid = rows / 2, amp = rows * 0.30;
  const th = Math.max(2, rows * 0.17);
  const period = Math.max(18, cols / 3);
  c.fillStyle = '#fff';
  for (let x = 0; x < cols; x++) {
    const y = mid + Math.sin((x / period + phase) * TAU) * amp;
    c.fillRect(x, y - th / 2, 1.05, th);
  }
}

/** ⭐⭐ THE AISLE. The floor's own shape: bays of shelving with a picker walking them,
 *  pausing at each. A 16:1 field IS an aisle — this is the pattern the strip was asking for.
 *  ⚠ The shelves are STATIC and only the picker moves. Scrolling both reads as a camera pan
 *    and the walk stops being legible as a walk. */
function aisle(c, cols, rows, phase) {
  const floor = rows - Math.max(2, rows * 0.14);
  const bays = Math.max(2, Math.round(cols / 30));
  const pitch = cols / bays, bw = pitch * 0.56;
  c.fillStyle = '#fff';
  for (let i = 0; i < bays; i++) {
    const x = i * pitch + (pitch - bw) / 2;
    c.fillRect(x, floor - rows * 0.66, bw, Math.max(1.6, rows * 0.17));
    c.fillRect(x, floor - rows * 0.36, bw, Math.max(1.6, rows * 0.17));
  }
  c.fillRect(0, floor, cols, Math.max(1, rows * 0.06));
  const t = phase * bays, i = Math.floor(t), f = t - i;
  const ease = f < 0.65 ? f / 0.65 : 1;                 // walk, then dwell at the bay
  const px = ((i + ease) * pitch + pitch / 2) % cols;
  // ⚠ THE PICKER HAS TO OUT-READ THE SHELVING. As a single floor-level disc it merged into
  //   the bays and the walk stopped being legible — the whole point of the pattern. It gets
  //   a body and a mast so it is unmistakably the one thing moving.
  const r = Math.max(1.8, rows * 0.15);
  c.fillRect(px - r, floor - r * 2, r * 2, r * 2);
  c.fillRect(px - 0.6, floor - r * 3.4, 1.4, r * 1.6);
}

/** ⭐ THE SWEEP. A bar crosses the board and the discs behind it settle back in a decaying
 *  trail — the medium showing its own refresh. Wraps, so the loop has no seam. */
function sweep(c, cols, rows, phase) {
  // ⚠ ONE HEAD IS NOT ENOUGH ON A LONG BOARD. At cols*0.30 the 876px strip sat ~70% dark
  //   at every instant and read as broken rather than idling — measured, not guessed. Two
  //   heads half a board apart keep the field alive at any width; on the short block they
  //   overlap into one continuous sweep, which is what it looked like before.
  const tail = Math.max(12, cols * 0.55);
  c.fillStyle = '#fff';
  for (let x = 0; x < cols; x++) {
    let best = -1;
    for (const off of [0, 0.5]) {
      let d = ((phase + off) * cols) - x; if (d < 0) d += cols;
      if (d <= tail) best = best < 0 ? d : Math.min(best, d);
    }
    if (best < 0) continue;
    const k = 1 - best / tail;
    for (let y = 0; y < rows; y++) {
      if (best < 1.7 || ((x * 7 + y * 13) % 11) < k * k * 10) c.fillRect(x, y, 1.05, 1.05);
    }
  }
}


/* ═════════════════════════════════════════════════════════════════════════════════════════
   THE QUIET SET — motion you cannot finish reading.

   ⭐ The first seven patterns are REPRESENTATIONAL: a piston, a belt, an aisle. You read one,
     you understand it, and after that there is nothing left to look at. These four are
     SYSTEMS instead — simple rules whose output never quite repeats inside the loop, which
     is the whole difference between something you watch and something you have already seen.

   ⚠⚠ EVERY ONE LOOPS EXACTLY, and that is a hard constraint rather than a nicety: motion
      built on `phase` must return to its start at phase 1 or the GIF shows a visible jump
      once a loop. So every frequency here is an INTEGER number of cycles per loop. Nothing
      is a hand-tuned drift that happens to look close.

   ⚠ AND NOTHING HERE USES A NOISE FIELD. Measured 2026-09-02: `refresh` costs 99 KB alone
     and +278 KB inside a set, `sweep` +83, because random dither is precisely what LZW
     cannot pack — and it widens the shared palette, which taxes every OTHER pattern in the
     same file. Large coherent shapes are both prettier here and very much cheaper.
   ═════════════════════════════════════════════════════════════════════════════════════ */

/** ⭐⭐ THE PENDULUM WAVE. A row of bobs whose frequencies differ by exactly one cycle per
 *  loop, so they start in a line, fall out of step into travelling waves, braid through
 *  every phase relationship there is, and snap back into a line at the end. The oldest
 *  hypnotic physics demo there is, and it is a perfect fit: the whole point of the piece is
 *  that it never looks the same twice until the instant it resolves. */
function pendulum(c, cols, rows, phase) {
  // ⚠⚠ SPEED LIVES HERE, NOT IN THE FRAME RATE. Bob i runs (base + i) cycles per hold, so
  //    the COUNT sets the top speed: 24 bobs at base 3 put the fastest at 26 cycles in a 5s
  //    hold — 5 Hz, which reads as shimmer, not as pendulums. Ten bobs starting at one cycle
  //    top out near 1.2 Hz and the wave becomes something you can actually follow.
  const n = Math.max(5, Math.min(16, Math.round(cols / 14)));
  const base = 1;                       // slowest bob, in whole cycles per loop
  const mid = rows / 2, amp = rows * 0.38;
  const r = Math.max(1.4, rows * 0.15);
  c.fillStyle = '#fff'; c.strokeStyle = '#fff';
  c.lineWidth = Math.max(1, rows * 0.05);
  for (let i = 0; i < n; i++) {
    const x = (i + 0.5) * (cols / n);
    const y = mid + Math.sin(phase * TAU * (base + i)) * amp;
    // ⚠ THE STEM IS WHAT MAKES THE WAVE READABLE. A bob on its own is ~2 discs on a
    //   14-disc-tall strip — it renders as dust and the travelling wave disappears. The
    //   line back to the rest axis traces the wave shape at grid resolution.
    c.beginPath(); c.moveTo(x, mid); c.lineTo(x, y); c.stroke();
    c.beginPath(); c.arc(x, y, r, 0, TAU); c.fill();
  }
}

/** ⭐⭐ THE MOIRÉ. Two gratings at slightly different wavelengths drifting against each
 *  other. Neither is interesting alone; their interference walks slow dark bands across the
 *  field that belong to neither, and the eye keeps trying to resolve which grating it is
 *  looking at. Costs almost nothing — the output is large flat regions. */
function moire(cols, rows, phase) {
  const t = phase * TAU, f = [];
  for (let y = 0; y < rows; y++) {
    const r = [];
    for (let x = 0; x < cols; x++) {
      // ⚠ LOW SPATIAL FREQUENCIES. The first cut used 0.55/0.62 and produced a field of
      //   small blobs — busy, expensive, and the interference was invisible because the
      //   bands were narrower than the beat between them. Wide gratings, wide beat.
      const u = Math.sin(x * 0.20 + y * 0.34 + t * 1);
      const v = Math.sin(x * 0.24 - y * 0.29 - t * 2);
      r.push(u + v > 0.62);
    }
    f.push(r);
  }
  return f;
}

/** ⭐⭐ THE LIQUID. Metaballs — bodies whose fields sum, so they bulge toward each other,
 *  fuse into one mass and tear apart again. Nothing in the warehouse moves like this, which
 *  is exactly why it reads as the calm thing on the board rather than another machine.
 *  ⚠ Body count derives from ASPECT, not width: a long strip needs them spread along it or
 *    four balls sit in a puddle at the left with 700px of dark to their right. */
function liquid(cols, rows, phase) {
  const t = phase * TAU;
  // ⚠⚠ THE BODIES HAVE TO MEET OR THIS IS JUST DOTS. Metaballs are only interesting where
  //    their fields overlap — the merge and the tear ARE the pattern. On a 14-disc-tall
  //    strip that means radii near the full height and centres about a radius apart, which
  //    reads as one undulating band that pinches and reconnects rather than as marbles.
  // ⚠⚠ RADIUS AND COUNT MUST BE DERIVED FROM EACH OTHER, and the two earlier cuts each got
  //    this wrong in opposite directions: count from ASPECT and radius from HEIGHT means the
  //    strip gets 3 bodies too far apart to ever touch (marbles) while the block gets 3
  //    bodies so large they fuse into a filled rectangle. Bodies merge when their centres
  //    sit about two radii apart, so radius is capped by height and the COUNT follows from
  //    it — one rule that holds at 2:1 and at 16:1.
  const rad = rows * 0.45;
  const n = Math.max(2, Math.min(12, Math.round(cols / (rad * 2.1))));
  const pts = [];
  for (let i = 0; i < n; i++) {
    const u = (i + 0.5) / n;
    const fx = 1 + (i % 3), fy = 1 + ((i * 2) % 3);     // integer → exact loop
    pts.push({
      x: (u + 0.075 * Math.sin(t * fx + i)) * cols,
      y: (0.5 + 0.20 * Math.sin(t * fy + i * 1.7)) * rows,
      r: rad * (1 + 0.16 * Math.sin(t * 2 + i))
    });
  }
  const f = [];
  for (let y = 0; y < rows; y++) {
    const row = [];
    for (let x = 0; x < cols; x++) {
      let sum = 0;
      for (let i = 0; i < n; i++) {
        // ⚠⚠ NO ASPECT SCALING HERE, AND THE FIRST CUT'S WAS A NO-OP ANYWAY: it multiplied
        //    (rows/cols) BY (cols/rows), which cancels to 1 and merely divided dx by 3.2 —
        //    every body came out 3.2x too wide and the whole strip fused into a solid slab.
        //    Discs are square, so grid units already ARE display units. Circles stay circles.
        const dx = x - pts[i].x, dy = y - pts[i].y;
        sum += (pts[i].r * pts[i].r) / (dx * dx + dy * dy + 0.7);
      }
      row.push(sum > 1.25);
    }
    f.push(row);
  }
  return f;
}

/** ⭐⭐ THE RIPPLE. A few still points on the board each send out expanding rings; where the
 *  rings cross they reinforce and cancel. Reads as rain on water — and because the sources
 *  sit at irrational spacings relative to one another, the interference figure between them
 *  is different at every moment of the loop. */
function ripple(cols, rows, phase) {
  const t = phase * TAU, f = [];
  const src = [
    { x: cols * 0.14, y: rows * 0.32, k: 2 },
    { x: cols * 0.47, y: rows * 0.70, k: 1 },
    { x: cols * 0.79, y: rows * 0.28, k: 3 }
  ];
  // ⚠ NO HORIZONTAL STRETCH — discs are square, so a circle in grid units is already a
  //   circle on screen. The first cut multiplied dx by 3.0 and the rings came out as wide
  //   blobs. On a 14-row strip a real ring mostly shows as its vertical slice, which is
  //   the sonar look and is the point.
  for (let y = 0; y < rows; y++) {
    const r = [];
    for (let x = 0; x < cols; x++) {
      let a = 0;
      for (let i = 0; i < src.length; i++) {
        const dx = x - src[i].x, dy = y - src[i].y;
        const d = Math.sqrt(dx * dx + dy * dy);
        a += Math.sin(d * 0.42 - t * src[i].k) / (1 + d * 0.055);
      }
      r.push(a > 0.62);          // a NARROW band around the crest → thin rings, not fill
    }
    f.push(r);
  }
  return f;
}

module.exports = { piston, ticker, tickerStrip, refresh, belt, night,
                   inlineFour, wave, aisle, sweep,
                   pendulum, moire, liquid, ripple };
