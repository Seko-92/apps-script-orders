/**
 * draw.js — THE DIAL. One drawing module, two hosts.
 *
 *   local  : ../render-states.js writes PNGs so a human can look before anything ships
 *   server : ./server.js answers GET /dial?... with the same bytes
 *
 * ⚠⚠ THE PROCESS RULE THIS FILE EXISTS TO SERVE (2026-08-31). Every other surface in this
 *    project renders headlessly and can be judged before it ships. THE GOOGLE SHEET CANNOT
 *    — the user's eyes are its first render. So the drawing lives here, in something that
 *    produces a PNG on this machine, and the sheet gets the formula only after they have
 *    seen the picture. The round that skipped this step was rejected wholesale.
 *
 * ⚠ NOTHING IS DERIVED HERE. Every number arrives in the query string, already computed by
 *   __SparkData. This is a RENDERER, not a feature — if the dial and the D1 headline ever
 *   disagree, the bug is upstream of this file.
 *
 * GEOMETRY. Everything is authored in FINAL CSS PIXELS (280 x 121) and multiplied by
 * `scale` on the way out, so a 2x render is the identical drawing at twice the resolution
 * rather than a second set of numbers to keep in sync.
 *
 * THE CLOCK FACE IS A REAL 12-HOUR FACE. The now-hand sits at the true clock angle, the
 * shift track traces 9am -> 5pm where the hands actually travel (270deg round through 0 to
 * 150deg, 240deg of sweep), and each worked hour lights a pip at its own hour mark. A 9-17
 * shift never wraps onto itself on a 12-hour face, which is what makes this legible rather
 * than merely decorative.
 */
'use strict';

const { paletteFor, BAND } = require('./palette');

const WIDTH  = 280;
const HEIGHT = 121;

// The dial occupies a 120-unit design space, drawn at 103px — the artifact's own numbers.
const DIAL_PX   = 109;
const VIEW      = 120;
const PAD       = 11;   // block padding, left and right
const GAP       = 17;   // dial -> flank
const SHIFT_OPEN  = 9;  // matches __SparkData!A13's off-hours definition, and the board's
const SHIFT_CLOSE = 17;

/** Sheet-identical minute formatting. Mirrors _fmtMinsExpr in BrandTheme.js.
 *  ⚠ ROUND THE TOTAL FIRST, THEN DECOMPOSE — rounding the remainder on its own prints
 *    "5h 60m", which was live on the banner for minutes after install. */
function fmtMins(min) {
  if (min == null || !isFinite(min) || min < 0) return '—';
  const r = Math.round(min);
  if (r < 60) return r + 'm';
  return Math.floor(r / 60) + 'h ' + (r % 60) + 'm';
}

/** Minutes-of-day -> "9:57". 12-hour, no meridiem — the face already says which half. */
function fmtClock(minOfDay) {
  const m = ((Math.round(minOfDay) % 1440) + 1440) % 1440;
  let h = Math.floor(m / 60) % 12;
  if (h === 0) h = 12;
  return h + ':' + String(m % 60).padStart(2, '0');
}

/** Clock angle in radians, 0 = 12 o'clock, clockwise. A 12-hour face: 720 min = full turn. */
function clockAngle(minOfDay) {
  return (((minOfDay % 720) + 720) % 720) / 720 * Math.PI * 2;
}

/* ------------------------------------------------------------------------------------ *
 * canvas helpers — all in design px, scaled once at the edge
 * ------------------------------------------------------------------------------------ */

function makePen(ctx, scale) {
  const S = (v) => v * scale;
  return {
    S,
    /** A radial line between two radii at a clock angle, measured from the dial centre. */
    spoke(cx, cy, ang, r0, r1, color, width, cap) {
      ctx.save();
      ctx.strokeStyle = color;
      ctx.lineWidth = S(width);
      ctx.lineCap = cap || 'round';
      ctx.beginPath();
      ctx.moveTo(S(cx + Math.sin(ang) * r0), S(cy - Math.cos(ang) * r0));
      ctx.lineTo(S(cx + Math.sin(ang) * r1), S(cy - Math.cos(ang) * r1));
      ctx.stroke();
      ctx.restore();
    },
    arc(cx, cy, r, a0, a1, color, width) {
      ctx.save();
      ctx.strokeStyle = color;
      ctx.lineWidth = S(width);
      ctx.lineCap = 'round';
      ctx.beginPath();
      // canvas angles run from +x; clock angles run from -y. Rotate by -90deg.
      ctx.arc(S(cx), S(cy), S(r), a0 - Math.PI / 2, a1 - Math.PI / 2);
      ctx.stroke();
      ctx.restore();
    },
    ring(cx, cy, r, color, width) {
      ctx.save();
      ctx.strokeStyle = color;
      ctx.lineWidth = S(width);
      ctx.beginPath();
      ctx.arc(S(cx), S(cy), S(r), 0, Math.PI * 2);
      ctx.stroke();
      ctx.restore();
    },
    dot(cx, cy, r, color) {
      ctx.save();
      ctx.fillStyle = color;
      ctx.beginPath();
      ctx.arc(S(cx), S(cy), S(r), 0, Math.PI * 2);
      ctx.fill();
      ctx.restore();
    },
    /** Pie wedge between two clock angles — the wait, as a shape. */
    wedge(cx, cy, r, a0, a1, color) {
      ctx.save();
      ctx.fillStyle = color;
      ctx.beginPath();
      ctx.moveTo(S(cx), S(cy));
      ctx.arc(S(cx), S(cy), S(r), a0 - Math.PI / 2, a1 - Math.PI / 2);
      ctx.closePath();
      ctx.fill();
      ctx.restore();
    },
    /** Text with real letter-spacing. napi-rs/canvas honours ctx.letterSpacing, but it is
     *  measured INTO the width, so a centred tracked string needs the trailing space
     *  removed or it sits half a gap to the left. */
    text(str, x, y, { size, weight, color, align, track }) {
      ctx.save();
      ctx.fillStyle = color;
      ctx.font = `${weight} ${S(size)}px Oswald`;
      const sp = track ? S(track) : 0;
      if (sp) ctx.letterSpacing = sp + 'px';
      ctx.textBaseline = 'alphabetic';
      let px = S(x);
      if (align === 'center') px -= (ctx.measureText(str).width - sp) / 2;
      else if (align === 'right') px -= ctx.measureText(str).width - sp;
      ctx.textAlign = 'left';
      ctx.fillText(str, px, S(y));
      if (sp) ctx.letterSpacing = '0px';
      ctx.restore();
    },
    width(str, { size, weight, track }) {
      ctx.save();
      ctx.font = `${weight} ${S(size)}px Oswald`;
      if (track) ctx.letterSpacing = S(track) + 'px';
      const w = ctx.measureText(str).width - (track ? S(track) : 0);
      ctx.letterSpacing = '0px';
      ctx.restore();
      return w / scale;
    }
  };
}

/* ------------------------------------------------------------------------------------ *
 * the ground
 * ------------------------------------------------------------------------------------ */

function drawGround(ctx, scale, pal) {
  const g = ctx.createLinearGradient(0, 0, 0, HEIGHT * scale);
  g.addColorStop(0, pal.bg[0]);
  g.addColorStop(0.7, pal.bg[1]);
  g.addColorStop(1, pal.bg[2]);
  ctx.fillStyle = g;
  ctx.fillRect(0, 0, WIDTH * scale, HEIGHT * scale);

  // ⚠ THE RIGHT EDGE MUST DISSOLVE INTO THE BAND. The cell to the right is flat #1a1a1a;
  //   a hard seam there is what makes a merged image read as a sticker instead of part of
  //   the banner. Same reason the v5 faces carry a seam vignette.
  const seam = ctx.createLinearGradient((WIDTH - 26) * scale, 0, WIDTH * scale, 0);
  seam.addColorStop(0, 'rgba(26,26,26,0)');
  seam.addColorStop(1, BAND);
  ctx.fillStyle = seam;
  ctx.fillRect((WIDTH - 26) * scale, 0, 26 * scale, HEIGHT * scale);

  // A breath of the state's own light along the bottom, so the block has a floor.
  // ⚠ KEPT UNDER 0.06. At 0.10 this read as a light leak along the bottom of every warm
  //   palette — obvious at true size, invisible at 4x, which is exactly why the contact
  //   sheet renders at 1:1.
  const glow = ctx.createLinearGradient(0, (HEIGHT - 13) * scale, 0, HEIGHT * scale);
  glow.addColorStop(0, hexA(pal.accent, 0));
  glow.addColorStop(1, hexA(pal.accent, 0.055));
  ctx.fillStyle = glow;
  ctx.fillRect(0, (HEIGHT - 13) * scale, WIDTH * scale, 13 * scale);
}

module.exports = {
  WIDTH, HEIGHT, DIAL_PX, VIEW, PAD, GAP, SHIFT_OPEN, SHIFT_CLOSE,
  fmtMins, fmtClock, clockAngle, makePen, drawGround, paletteFor
};

/* ------------------------------------------------------------------------------------ *
 * the face
 * ------------------------------------------------------------------------------------ */

// Radii, in design px, derived from the 103px dial box. Authored here once so the parts
// cannot drift apart the way three hardcoded copies of a shelf comparator did.
const CX = PAD + DIAL_PX / 2;         // 64.5
const CY = HEIGHT / 2;                // 60.5
const R  = { rim: 47.2, glow: 49.0, tickOut: 47.2, tickIn: 43.8,
             pipOut: 40.6, pipIn: 34.3, arc: 47.2, wedge: 30.2,
             // ⚠ THE HANDS STOP SHORT OF THE PIP RING, deliberately. The now-hand always
             //   points at the current hour, which is always where the newest pip is — so
             //   full-length hands permanently bury the most recent thing the ring records.
             //   Outer ring = the day so far. Inner zone = the wait. They never overlap.
             oldHand: 32.4, nowHand: 27.9, hub: 2.3 };

// ⚠ FOUR, NOT EIGHT. The first render put half-hour ticks at the pips' own radius and the
//   two rings fused into one dashed mess. The ticks exist to say "this is a clock"; the
//   pips carry the day. Different jobs, different radii, different weights.
const TICK_HOURS = [0, 3, 6, 9];

/**
 * @param {CanvasRenderingContext2D} ctx
 * @param {object} st  the state — every field already computed upstream
 * @param {number} scale
 */
function drawDial(ctx, st, scale) {
  const pal = paletteFor(st.verdict);
  const pen = makePen(ctx, scale);
  const open = st.verdict !== 'rest';
  // ⭐⭐ ONE ARC-SHAPED THING AT A TIME. The first render drew the shift arc AND the wait
  //    wedge in the same accent, concentric, on a 103px face — two overlapping orange
  //    curves that nobody could tell apart. They compete for the same visual channel, and
  //    the wedge is worth more (it is the only DURATION on this sheet, and the only number
  //    that costs money when it grows), so the arc yields to it whenever there is a wait.
  const showWait = open && st.oldestMin > 0 && (st.verdict === 'busy' || st.verdict === 'late');
  // ⭐ AND HANDS NEVER SHARE THE FACE WITH TEXT. Squeezed under the hub, "3h 12m" collided
  //   with the 5 o'clock pip and read at 13.5px; in the flank it reads at 17.5px with room.
  //   The face shows the clock when it has nothing better to show, and shows the wait as a
  //   SHAPE when it does — which is what a chronograph does with its subdials.
  const faceText = !showWait;

  drawGround(ctx, scale, pal);

  // --- the ground glow under the face -------------------------------------------------
  const g = ctx.createRadialGradient(CX * scale, CY * scale, R.rim * 0.55 * scale,
                                     CX * scale, CY * scale, R.glow * scale);
  g.addColorStop(0, 'rgba(0,0,0,0)');
  g.addColorStop(1, hexA(pal.accent, 0.07));
  ctx.fillStyle = g;
  ctx.beginPath();
  ctx.arc(CX * scale, CY * scale, R.glow * scale, 0, Math.PI * 2);
  ctx.fill();

  // --- rim + the four hour marks ------------------------------------------------------
  pen.ring(CX, CY, R.rim, pal.rim, 1);
  for (const h of TICK_HOURS) {
    pen.spoke(CX, CY, h / 12 * Math.PI * 2, R.tickOut, R.tickIn, pal.tick, 1.2, 'butt');
  }

  // --- the shift: 9am -> 5pm, where the hands actually travel --------------------------
  // Drawn as a HERO when there is no wedge (the day's progress is then the only story) and
  // as a bare TRACK when there is (the window still wants marking; it just stops shouting).
  // ⚠ Never at rest: before 9 the track is a promise, and the resting face is about the day
  //   just finished, not the one not yet started.
  if (open) {
    const a0 = clockAngle(SHIFT_OPEN * 60);
    const a1 = a0 + (SHIFT_CLOSE - SHIFT_OPEN) / 12 * Math.PI * 2;
    pen.arc(CX, CY, R.arc, a0, a1, showWait ? pal.rim : pal.track, showWait ? 1.7 : 4.1);
    if (!showWait) {
      const done = Math.min(Math.max(st.nowMin, SHIFT_OPEN * 60), SHIFT_CLOSE * 60);
      if (done > SHIFT_OPEN * 60) {
        const aNow = a0 + (done - SHIFT_OPEN * 60) / 720 * Math.PI * 2;
        pen.arc(CX, CY, R.arc, a0, aNow, pal.accent, 4.1);
        pen.dot(CX + Math.sin(aNow) * R.arc, CY - Math.cos(aNow) * R.arc, 3.0, pal.accent);
      }
    }
  }

  // --- the day's pips: one mark per worked hour, at its own hour position --------------
  // A 9-17 shift spans 9,10,11,12,1,2,3,4 on a 12-hour face — 240deg with nothing wrapping
  // onto itself, which is the only reason this is legible rather than merely decorative.
  const hours = st.showPips === false ? [] : (st.hours || []);
  const nowH  = Math.floor(st.nowMin / 60);
  for (let h = 0; st.showPips !== false && h < 24; h++) {
    // ⚠ THE FUTURE IS NEVER LIT. A count on an hour that has not happened yet means the
    //   data is wrong (clock drift, a stale __Published copy) — and a face that cheerfully
    //   paints it would be the "reassuring label on a dangerous state" this codebase rules
    //   is a bug. At rest the whole day is behind us, so every hour counts as past.
    const past    = st.verdict === 'rest' ? true : h <= nowH;
    const worked  = past && (hours[h] || 0) > 0;
    const inShift = h >= SHIFT_OPEN && h < SHIFT_CLOSE;
    if (!worked && !(inShift && past)) continue;
    pen.spoke(CX, CY, (h % 12) / 12 * Math.PI * 2, R.pipOut, R.pipIn,
              worked ? pal.accent : pal.track, worked ? 2.1 : 1.8);
  }

  // --- the wait, as an angle ----------------------------------------------------------
  if (showWait) {
    // ⚠ CLAMPED UNDER A FULL TURN. Past 12h the wedge laps itself and reads as a SHORTER
    //   wait than it is — a reassuring shape on the worst state this face can show.
    const waited = Math.min(st.oldestMin, 719);
    const aNow   = clockAngle(st.nowMin);
    const aOld   = clockAngle(st.nowMin - waited);
    ctx.save();
    ctx.globalCompositeOperation = 'lighter';
    pen.wedge(CX, CY, R.wedge, aOld, aNow, hexA(pal.accent, 0.20));
    ctx.restore();
    pen.spoke(CX, CY, aOld, 0, R.oldHand, pal.accent, 2);
    pen.spoke(CX, CY, aNow, 0, R.nowHand, hexA(pal.ink, 0.82), 1.5);
    pen.dot(CX, CY, R.hub, pal.ink);
  } else if (!open) {
    // A bezel mark for NOW, outside the pip ring so it cannot be mistaken for one — and
    // only when nothing else marks it. With the floor open the arc's bright head is now,
    // and two markers on one face is one too many.
    // ⚠ The first render put a filled dot at the pips' own radius and it read as a smudge.
    pen.spoke(CX, CY, clockAngle(st.nowMin), R.tickOut + 0.6, R.tickIn - 0.6, pal.lead, 2.2);
  }

  // --- the centre: one big fact, one caption ------------------------------------------
  if (faceText) {
    pen.text(st.big, CX, CY - 2, { size: 19, weight: 200, color: pal.ink, align: 'center' });
    pen.text(st.caption, CX, CY + 10,
             { size: 6.6, weight: 500, color: pal.label, align: 'center', track: 1.5 });
  }

  drawFlank(pen, st, pal);
}

/** Three facts down the right side. Value big and coloured, label small and quiet. */
function drawFlank(pen, st, pal) {
  const x = PAD + DIAL_PX + GAP;              // 128
  const rows = st.flank;
  const ROW = 28.5, GAPY = 8;
  let y = (HEIGHT - (rows.length * ROW + (rows.length - 1) * GAPY)) / 2;
  for (const r of rows) {
    pen.text(r.value, x, y + 14.6, {
      size: 17.5, weight: 500, color: r.tone === 'quiet' ? pal.dim : pal.lead, align: 'left'
    });
    pen.text(r.label.toUpperCase(), x, y + 26.4, {
      size: 7.6, weight: 400, color: pal.label, align: 'left', track: 1.05
    });
    y += ROW + GAPY;
  }
}

/** #rrggbb + alpha -> rgba(). Kept tiny and local; there is no colour library here. */
function hexA(hex, a) {
  const n = parseInt(hex.slice(1), 16);
  return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
}

module.exports.drawDial = drawDial;
module.exports.hexA = hexA;
module.exports.R = R;
module.exports.CX = CX;
module.exports.CY = CY;
