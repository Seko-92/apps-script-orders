/**
 * board.js — THE SPLIT-FLAP PRIMITIVE. One flap cell, and the stepper that drives it.
 *
 * ⚠ PURE. No canvas creation, no fonts, no fs, no query parsing. Everything here is either
 *   arithmetic (the stepper) or a draw call against a ctx handed in. That is what makes
 *   `test-board.js` able to assert the mechanic without rendering anything.
 *
 * ⚠⚠ NOTHING IS DERIVED HERE — same contract as draw.js. The board renders the state it is
 *    GIVEN. The server never remembers a previous state either: Apps Script sends BOTH
 *    `from` and `to` in the query string, which is what lets /dial stay stateless and
 *    therefore safe to leave unauthenticated. A caller can only ever animate text they
 *    themselves supplied.
 *
 * GEOMETRY. Authored in FINAL CSS PIXELS and multiplied by `scale` on the way out, so a 2x
 * render is the same drawing at twice the resolution rather than a second set of numbers.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * THE SIX MOVES THIS FILE EXISTS TO HONOUR (from the approved design review). The first
 * attempt at this board violated 1, 5 and 6 and read as a terminal emulator, not Solari:
 *
 *   1 · NO MONOSPACE. A split-flap is monospaced by CONSTRUCTION — every flap is the same
 *       width — but the letterforms are a condensed grotesque. Oswald, centred in a fixed
 *       grid. This is the single biggest upgrade over a naive version.
 *   2 · FIELDS, NOT A SENTENCE. Labels are engraved into the housing and never move; only
 *       VALUES sit on flaps.
 *   3 · ONLY WHAT CHANGES MOVES. A real board does not flip a character that already reads
 *       correctly. `cellState` returns phase 0 for an unchanged cell, forever.
 *   4 · FLIP THROUGH THE ALPHABET. R -> W passes S T U V. Characters with further to travel
 *       land LAST, which is the organic settle that signs the medium.
 *   5 · FIXED LIGHT SOURCE. Top halves catch light from above, bottom halves take bounce.
 *       There is deliberately NO sweeping sheen — a moving highlight reads as a screen
 *       effect; a fixed one reads as an object in a room.
 *   6 · BE STILL. The board does nothing most of the time, and the stillness is what makes
 *       a flip land. It is also honest: the data does not change every four seconds.
 */
'use strict';

/** ⚠ ORDER IS THE MECHANIC, not a lookup table. The gap between two characters IS the
 *  number of physical flaps that must fall, so this string decides how long each cell
 *  travels. Space leads so a blanked cell settles fastest. */
const ALPHABET = ' ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789·:/-';

/** ⚠⚠ A NUMBER DRUM CARRIES ONLY DIGITS, and this is not pedantry — it is what the thing
 *  looks like. Spinning a COUNT field through the full alphabet renders "6QO" mid-settle,
 *  which reads as garbage or a fault on a banner someone glances at. A physical board has
 *  different drums for different fields; so does this one. It also settles faster, because
 *  the longest journey is 10 flaps instead of 40. */
const DIGITS = ' 0123456789';
const N = ALPHABET.length;

// Seconds. Kept small and named because trap 7 caps the whole loop at 8s: a fresh GIF
// always plays from frame 0, so a long loop makes every swap a visible snap.
const CASCADE_S = 0.045;   // stagger between adjacent cells
const STEP_S    = 0.055;   // one character flip
const HOLD_S    = 1.30;    // stillness at the end, so the board is READ, not watched

const idxOf = (ch, abc) => {
  const i = (abc || ALPHABET).indexOf(String(ch || ' ').toUpperCase());
  return i < 0 ? 0 : i;                       // anything unmappable becomes a blank flap
};

/** How many flaps must fall for this cell. 0 means it never moves (move 3). */
function distance(from, to, abc) {
  const A = abc || ALPHABET;
  return (idxOf(to, A) - idxOf(from, A) + A.length) % A.length;
}

/**
 * The whole animation in one function.
 * @returns {{cur:string,next:string,phase:number,moving:boolean}}
 *   phase 0..1 through the CURRENT flip. phase 0 + moving false = settled, draw it static.
 */
function cellState(i, from, to, t, abc) {
  const A = abc || ALPHABET;
  const d = distance(from, to, A);
  const settled = { cur: A[idxOf(to, A)], next: A[idxOf(to, A)], phase: 0, moving: false };
  if (d === 0) return settled;                          // move 3 — never flips

  const start = i * CASCADE_S;
  const local = t - start;
  if (local <= 0) {
    const c = A[idxOf(from, A)];
    return { cur: c, next: c, phase: 0, moving: false }; // not started yet
  }
  if (local >= d * STEP_S) return settled;               // landed

  const step = Math.floor(local / STEP_S);
  const phase = (local - step * STEP_S) / STEP_S;
  return {
    cur:  A[(idxOf(from, A) + step) % A.length],
    next: A[(idxOf(from, A) + step + 1) % A.length],
    phase,
    moving: true
  };
}

/** Total seconds for a whole word to settle, plus the deliberate stillness after it. */
function loopSeconds(from, to, abc) {
  let longest = 0;
  const n = Math.max(String(from).length, String(to).length);
  for (let i = 0; i < n; i++) {
    const d = distance(String(from)[i] || ' ', String(to)[i] || ' ', abc);
    if (d) longest = Math.max(longest, i * CASCADE_S + d * STEP_S);
  }
  return longest + HOLD_S;
}

/**
 * The falling card's vertical scale, 0..1, with a 6% settle bounce.
 *
 * ⚠ NEVER A RECTANGLE SQUASHING IN PLACE. That is what every cheap imitation does and it
 *   is instantly readable as fake. The card rotates about the SEAM, so its scale is the
 *   cosine of a real angle — and it OVERSHOOTS and settles, because real flaps bounce.
 *   Continuous at u = 0.86 by construction: the base reaches exactly 1.0 there.
 */
function flapScale(u) {
  if (u <= 0.86) return Math.sin((u / 0.86) * (Math.PI / 2));
  const k = (u - 0.86) / 0.14;
  return 1 - 0.06 * Math.sin(k * Math.PI);
}

module.exports = {
  ALPHABET, DIGITS, N, CASCADE_S, STEP_S, HOLD_S,
  idxOf, distance, cellState, loopSeconds, flapScale
};


/* ═══════════════════════════════════════════════════════════════════════════
   THE DRAWING HALF. Still pure in the sense that matters: it never creates a
   canvas or loads a font, it only issues calls against a ctx handed in.
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * ⚠ CREAM ON MATTE BLACK, AND BRAND YELLOW IS RESERVED FOR EXACTLY ONE THING:
 *   a field that needs a person. The board is monochrome until it isn't.
 *
 *   This is the codebase's own oldest ruling — YELLOW MEANS ACT — expressed in a new
 *   medium. It is the same argument that moved red off NOT FOUND and refused to paint a
 *   calm hold red. Do not spend `act` on decoration; the moment it appears twice it stops
 *   meaning anything.
 */
const BOARD = {
  housing:   '#141210',
  housingHi: '#26221c',   // the lit top edge of the housing — fixed light, from above
  tileTopA:  '#2b2721',   // top half, lit
  tileTopB:  '#221f1a',
  tileBotA:  '#1a1714',   // bottom half, in shade, with a little bounce at its foot
  tileBotB:  '#232019',
  seam:      '#0b0a09',
  seamLit:   '#3d382f',   // ONE lit pixel under the seam. This single line is most of
                          // what makes a flat rectangle read as a physical card.
  ink:       '#f0ece0',
  inkDim:    '#8b8578',
  engrave:   '#6b6459',
  engraveHi: '#2b2722',
  act:       '#ffd400'
};

/** Rounded tile body path, in device px. */
function tilePath(ctx, X, Y, W, H, R) {
  ctx.beginPath();
  ctx.roundRect(X, Y, W, H, R);
}

/**
 * One half of a settled tile. `which` is 'top' or 'bottom'.
 * Split into halves because the flip animates them independently about the seam.
 */
function drawHalf(ctx, o, ch, which) {
  const S = (v) => v * o.scale;
  const X = S(o.x), Y = S(o.y), W = S(o.w), H = S(o.h);
  const R = S(o.radius == null ? 2.2 : o.radius);
  const half = H / 2;
  const top = which === 'top';

  ctx.save();
  // Clip to this half of the rounded tile, so the card edges stay square at the seam and
  // rounded at the outside — which is how a real flap is cut.
  ctx.beginPath();
  ctx.rect(X, top ? Y : Y + half, W, half);
  ctx.clip();
  tilePath(ctx, X, Y, W, H, R);
  const g = ctx.createLinearGradient(0, top ? Y : Y + half, 0, top ? Y + half : Y + H);
  g.addColorStop(0, top ? BOARD.tileTopA : BOARD.tileBotA);
  g.addColorStop(1, top ? BOARD.tileTopB : BOARD.tileBotB);
  ctx.fillStyle = g;
  ctx.fill();

  // The character. Drawn ONCE per half against a clip, so the glyph is split exactly at
  // the seam the way a printed flap is — never two separate half-glyphs that can drift.
  ctx.fillStyle = o.ink || BOARD.ink;
  ctx.font = `${o.weight || 500} ${S(o.fontPx)}px Oswald`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(String(ch), X + W / 2, Y + H / 2 + S(o.baseline || 0));
  ctx.restore();
}

/** The seam: one dark line with a single lit line beneath it. */
function drawSeam(ctx, o) {
  const S = (v) => v * o.scale;
  const X = S(o.x), W = S(o.w), seamY = S(o.y + o.h / 2);
  const lw = Math.max(1, S(0.9));
  ctx.save();
  ctx.fillStyle = BOARD.seam;
  ctx.fillRect(X, seamY - lw / 2, W, lw);
  ctx.fillStyle = BOARD.seamLit;
  ctx.fillRect(X, seamY + lw / 2, W, Math.max(1, S(0.5)));
  ctx.restore();
}

/**
 * ONE FLAP CELL, at any phase. This is the whole mechanic.
 *
 *   phase 0   -> 0.5 : the top half of the OLD character collapses about the seam
 *                      (scaleY = cos), revealing the NEW top already standing behind it.
 *                      The falling card throws a travelling shadow on the half below.
 *   phase 0.5 -> 1   : a card unfolds DOWNWARD from the seam carrying the NEW bottom half,
 *                      overshooting 6% and settling.
 */
function drawFlapCell(ctx, o) {
  const S = (v) => v * o.scale;
  const seamY = S(o.y + o.h / 2);
  const { cur, next, phase, moving } = o;

  // ⚠ BOTH ENDS ARE SETTLED STATES, and they show DIFFERENT characters. phase 0 is the
  //   old card at rest; phase >= 1 is the new one landed. cellState never emits phase 1
  //   (it returns `settled` instead), but a caller constructing a frame by hand can — the
  //   8x anatomy render did, and drew an R where an S had just landed. Defend the boundary
  //   rather than trusting every caller to know that.
  if (!moving || !phase) {
    drawHalf(ctx, o, cur, 'top');
    drawHalf(ctx, o, cur, 'bottom');
    drawSeam(ctx, o);
    return;
  }
  if (phase >= 1) {
    drawHalf(ctx, o, next, 'top');
    drawHalf(ctx, o, next, 'bottom');
    drawSeam(ctx, o);
    return;
  }

  if (phase < 0.5) {
    const u = phase / 0.5;
    drawHalf(ctx, o, next, 'top');        // already standing behind the falling card
    drawHalf(ctx, o, cur,  'bottom');

    // The travelling shadow. It is strongest as the card starts to fall and fades as the
    // card flattens toward the seam — the cue that something is moving ABOVE the surface
    // rather than the surface itself changing.
    const shade = 0.42 * (1 - u);
    if (shade > 0.01) {
      const X = S(o.x), W = S(o.w), H = S(o.h);
      const g = ctx.createLinearGradient(0, seamY, 0, seamY + H * 0.34);
      g.addColorStop(0, `rgba(0,0,0,${shade.toFixed(3)})`);
      g.addColorStop(1, 'rgba(0,0,0,0)');
      ctx.save();
      ctx.beginPath();
      ctx.rect(X, seamY, W, H / 2);
      ctx.clip();
      ctx.fillStyle = g;
      ctx.fillRect(X, seamY, W, H / 2);
      ctx.restore();
    }

    ctx.save();                            // the collapsing OLD top, hinged at the seam
    ctx.translate(0, seamY);
    ctx.scale(1, Math.cos(u * (Math.PI / 2)));
    ctx.translate(0, -seamY);
    drawHalf(ctx, o, cur, 'top');
    ctx.restore();
  } else {
    const u = (phase - 0.5) / 0.5;
    drawHalf(ctx, o, next, 'top');
    drawHalf(ctx, o, cur,  'bottom');      // the old bottom, still showing until covered

    ctx.save();                            // the NEW bottom unfolding down from the seam
    ctx.translate(0, seamY);
    ctx.scale(1, flapScale(u));
    ctx.translate(0, -seamY);
    drawHalf(ctx, o, next, 'bottom');
    ctx.restore();
  }
  drawSeam(ctx, o);
}

/**
 * A whole field of flaps. `cells` is the fixed grid — its LENGTH never changes with the
 * value, because a Solari board has a fixed number of physical cells and short words
 * simply leave blanks. That constraint is the medium, not a limitation to design around.
 */
function drawFlapWord(ctx, o) {
  const cells = o.cells;
  for (let i = 0; i < cells; i++) {
    const st = cellState(i, (o.from[i] || ' '), (o.to[i] || ' '), o.t, o.abc);
    drawFlapCell(ctx, Object.assign({}, o, {
      x: o.x + i * (o.w + o.gap),
      cur: st.cur, next: st.next, phase: st.phase, moving: st.moving
    }));
  }
}

/** Engraved into the housing: a dark letter with one lit pixel under it. Never moves. */
function drawEngraved(ctx, o) {
  const S = (v) => v * o.scale;
  ctx.save();
  ctx.font = `${o.weight || 500} ${S(o.fontPx)}px Oswald`;
  ctx.textAlign = o.align || 'left';
  ctx.textBaseline = 'middle';
  ctx.fillStyle = BOARD.engraveHi;                       // the lit lip below the cut
  ctx.fillText(o.text, S(o.x), S(o.y) + Math.max(1, S(0.8)));
  ctx.fillStyle = o.ink || BOARD.engrave;                // the cut itself
  ctx.fillText(o.text, S(o.x), S(o.y));
  ctx.restore();
}

module.exports.BOARD = BOARD;
module.exports.drawFlapCell = drawFlapCell;
module.exports.drawFlapWord = drawFlapWord;
module.exports.drawEngraved = drawEngraved;
module.exports.drawHalf = drawHalf;
