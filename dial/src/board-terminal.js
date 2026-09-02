/**
 * board-terminal.js — THE TERMINAL. Row 1 as a Solari split-flap board.
 *
 * TWO IMAGES, because row 1 is not one continuous surface:
 *
 *   ANCHOR  280 x 121  over A1:C2 — the state, one figure, and the maker's plate
 *   STRIP   876 x  56  over D1:H1 — one figure, then TWO TRANSPARENT WINDOWS
 *
 * ⚠⚠ THE WINDOWS ARE THE WHOLE ARCHITECTURE, AND THEY ONLY EXIST BECAUSE OF A MEASUREMENT.
 *    The 2026-09-02 gate proved a transparent PNG lets the cell underneath show through.
 *    So the strip does NOT redraw E1's pulse or the F1:H1 curve — it cuts holes and lets
 *    the LIVE CELLS show through an engraved bezel.
 *
 *    That is not a cosmetic choice, it retires a trap. The original design covered D1:H1
 *    opaquely and re-rendered the pulse and the curve from the published tick — which is
 *    TWO THINGS COMPUTING ONE VALUE, this project's number-one recurring bug class
 *    (A-9/A-50, _fmtMinsExpr, the hold rule in three files). A window cannot drift from
 *    the cell, because it IS the cell.
 *
 * ⚠ THE SIZES ARE LOAD-BEARING ON ROW HEIGHTS AND COLUMN WIDTHS.
 *      280 = A 107 + B  70 + C 103                  (MASTHEAD.imgW, asserted by setupMasthead)
 *      876 = D 232 + E 307 + F 130 + G 100 + H 107
 *      56  = MASTHEAD.rowHeight
 *    ⚠ The plan said 880 for the strip; it is 876. Measured off applyBrandTheme's own
 *      setColumnWidth calls, not estimated. A 4px overhang would put the board's right
 *      edge into column I on every install. The installer must ASSERT these and refuse
 *      rather than hang a misaligned board.
 *
 * ⚠⚠ F2:G2 AND H2 MUST NEVER BE COVERED — they are the Pick ID dropdowns, and a floating
 *    image swallows clicks as well as covering pixels. Breaking them is a floor outage on
 *    the print/pick accountability gate, not a cosmetic bug. THAT is why the strip is 56
 *    tall (row 1 only) and not 121.
 *
 * DESIGN LINEAGE: the split-flap board was Gino Valle's 1956 *Cifra* for Solari di Udine,
 * and "solari" became the generic term because of them. The reference belongs HERE, in a
 * comment — never on the plate. The plate says HQ MOTOR SERVICE.
 */
'use strict';

const { paletteFor, BAND } = require('./palette');
const { buildState } = require('./state');
const { fmtMins } = require('./draw');
const B = require('./board');

const ANCHOR = { w: 280, h: 121 };
const STRIP  = { w: 876, h: 56 };

// Column geometry inside the strip, in final CSS px from D1's left edge.
const COL_D = 232, COL_E = 307, COL_FH = 337;

/** ⚠ Fixed at 9 cells because a Solari board has a fixed number of PHYSICAL cells.
 *  "PAST THE LINE" is shortened to "PAST LINE" to fit — the grid does not stretch for a
 *  long word, which is the medium, not a limitation to design around. */
const STATE_CELLS = 9;
/** ⚠ EACH FIELD IS SIZED TO ITS OWN RANGE, because a Solari field is a fixed run of
 *  physical cells and dead ones read as broken. The LEAD field can carry a duration
 *  ("4h 55m" = 6), the quiet fields only ever carry counts — shipped / received / to grab
 *  / out today — so four cells covers 9999 with headroom.
 *  ⚠ Four, not three: clip() would TRUNCATE a larger number silently, and a quietly wrong
 *    count is exactly the "reassuring label on a dangerous state" this codebase rules is a
 *    bug. Better a spare cell than a plausible lie. */
const LEAD_CELLS = 6;
const FIG_CELLS  = 4;

const WORD = {
  rest:  'RESTING',
  clear: 'ALL CLEAR',
  busy:  'WORKING',
  late:  'PAST LINE',
  stale: 'NO SIGNAL'
};

/** ⚠ A PHYSICAL BOARD CANNOT SHOW A CHARACTER IT DOES NOT HAVE. state.js writes an EM
 *  DASH for a missing figure, and '—' is not on the flaps — so idxOf() fell through to a
 *  blank and OLDEST / TO GRAB rendered as SIX EMPTY CELLS at true size, reading as a
 *  broken field rather than an empty one. The flap alphabet carries a HYPHEN; use it.
 *  ⭐ Caught only by rendering at scale 1. At 2x the eye fills the gap in. */
const NONE = '-';

const clip = (s, n) => String(s == null ? '' : s).toUpperCase().slice(0, n);
/** ⚠ ALIGNMENT IS A FLAP DECISION, not a text decision. A figure right-aligns so its
 *  blanks LEAD — that is how a real board carries a number, and it means a 8 -> 12 change
 *  moves the cells nearest the eye. A word centres so a short state does not sit against
 *  one edge of its grid with a run of dead cells trailing it. */
const padFig  = (s, n) => clip(s, n).padStart(n, ' ');
const padWord = (s, n) => {
  const v = clip(s, n), left = Math.floor((n - v.length) / 2);
  return ' '.repeat(left) + v + ' '.repeat(n - v.length - left);
};

/**
 * What the board SHOWS. Deliberately built on top of buildState so the verdict and the
 * three flank figures are chosen in exactly ONE place — the board never picks its own.
 */
function buildBoard(q) {
  const st = buildState(q);
  const n  = (v) => (v == null || v === '' || isNaN(Number(v)) ? null : Number(v));

  const oldest = n(q.o), toGrab = n(q.g), out = n(q.p);

  // ⚠⚠ THE LABELS ARE FIXED. THIS IS MOVE 2, AND THE FIRST CUT BROKE IT.
  //    That version reused buildState's flank, whose LABELS change with the verdict — so
  //    the engraved word read UNTIL OPEN at rest and OLDEST WAITING when busy. An engraved
  //    label that changes is a lie about the medium: engraving is cut once, into metal,
  //    and it is the fixed thing a value is read AGAINST. A departure board changes the
  //    times, never the word DEPARTURES.
  //
  //    So the board carries FOUR PERMANENT FIELDS. Every one is meaningful in every
  //    verdict, and a field with nothing to say shows a blank grid rather than borrowing
  //    another field's meaning.
  //
  // ⚠ THE COST OF THAT, STATED PLAINLY: the dial's rest-state countdown ("4h 55m UNTIL
  //   OPEN") has NO fixed home here — it is meaningless during a shift, so a permanent
  //   field for it would sit blank all day. It is the one figure this design drops.
  //   The user rules on that; do not quietly reinstate it as a changing label.
  return {
    verdict: st.verdict,
    word: padWord(WORD[st.verdict] || 'ALL CLEAR', STATE_CELLS),

    // ⭐ THE ONE FIELD THAT MAY TURN YELLOW, AND ONLY WHEN IT NEEDS A PERSON.
    //   busy/late mean work is aging, which is the only state on this board that asks
    //   somebody to move. Everywhere else the board stays monochrome — which is what makes
    //   the yellow mean something the day it appears. Same ruling that moved red off
    //   NOT FOUND and refused to paint a calm hold red.
    lead: { label: 'OLDEST',
            value: padFig(oldest == null ? NONE : fmtMins(oldest), LEAD_CELLS),
            act: st.verdict === 'busy' || st.verdict === 'late' },
    figA: { label: 'TO GRAB',
            value: padFig(toGrab == null ? NONE : String(toGrab), FIG_CELLS), act: false },
    figB: { label: 'OUT TODAY',
            value: padFig(out == null ? NONE : String(out), FIG_CELLS), act: false }
  };
}

/**
 * Housing plate with a lit top edge — fixed light from above, never a sweep.
 *
 * ⚠⚠ FILLS EDGE TO EDGE, OPAQUE, WITH NO ROUNDED CORNERS — and that is not a style choice.
 *    The first version used roundRect, which left the four corners TRANSPARENT. In a
 *    canvas render that is invisible against a dark page; in the SHEET the cell showed
 *    through and the face landed with a purple fringe around it. `drawGround` in draw.js
 *    has always done a plain full fillRect for exactly this reason.
 *    ⭐ Caught by the user on the live sheet, not by any render here — the standing rule
 *      again: the sheet cannot be rendered headlessly, so it is the only honest test.
 *
 * ⚠ AND IT CARRIES THE SAME RIGHT-EDGE SEAM VIGNETTE AS THE DIAL. The cell to the right is
 *   flat #1a1a1a; a hard seam there makes a merged image read as a sticker pasted onto the
 *   banner instead of part of it.
 */
function housing(ctx, x, y, w, h, scale) {
  const S = (v) => v * scale;
  ctx.save();
  const g = ctx.createLinearGradient(0, S(y), 0, S(y + h));
  g.addColorStop(0, B.BOARD.housingHi);
  g.addColorStop(0.14, B.BOARD.housing);
  g.addColorStop(1, '#100e0c');
  ctx.fillStyle = g;
  ctx.fillRect(S(x), S(y), S(w), S(h));

  const seam = ctx.createLinearGradient(S(x + w - 26), 0, S(x + w), 0);
  seam.addColorStop(0, 'rgba(26,26,26,0)');
  seam.addColorStop(1, BAND);
  ctx.fillStyle = seam;
  ctx.fillRect(S(x + w - 26), S(y), S(26), S(h));
  ctx.restore();
}

/** An engraved label above a field of flaps. */
function field(ctx, o) {
  B.drawEngraved(ctx, {
    scale: o.scale, x: o.x, y: o.y, text: String(o.label || '').toUpperCase(),
    fontPx: o.labelPx, weight: 500, ink: B.BOARD.engrave
  });
  B.drawFlapWord(ctx, {
    scale: o.scale, x: o.x, y: o.y + o.gap, w: o.cw, h: o.ch, gap: o.cgap,
    cells: o.cells, from: o.from, to: o.to, t: o.t, abc: o.abc,
    fontPx: o.charPx, weight: 500, ink: o.act ? B.BOARD.act : B.BOARD.ink,
    baseline: o.baseline || 0, radius: 2
  });
}

/**
 * THE ANCHOR — A1:C2, AND IN v1 THIS IS THE WHOLE PRODUCT.
 *
 * ⭐⭐ ONE IMAGE, NOT TWO. The strip below is built and tested but NOT installed in v1, and
 *    the reason is a cost correction the user caught: the only thing that may sit on flaps
 *    is something that changes in DISCRETE STEPS. The state word does — five values, a
 *    handful of changes a day. Every figure on this banner (oldest, to grab, out today, the
 *    pulse, the curve) changes CONTINUOUSLY, so putting any of them on flaps makes the image
 *    dirty every minute and turns ~8 swaps a day into ~1,440.
 *
 *    ⚠ That is also an HONESTY point, not only a cost one: a continuously-ticking number
 *      rendered as flaps is a lie about the mechanism. A real board flips when a value
 *      genuinely changes to a new discrete value; it does not tick.
 *
 *    So D1, E1 and F1:H1 are LEFT ALONE — untouched, uncovered, working exactly as they do
 *    today at formula speed for zero Apps Script. The board takes the one cell whose content
 *    is genuinely discrete, and frames nothing else.
 *
 * ⚠ It is OPAQUE, and it covers the =IMAGE() dial, which stays in the cell underneath
 *   untouched. ROLLBACK = delete one floating image. Zero cell writes, nothing to restore.
 */
function drawAnchor(ctx, from, to, t, scale) {
  const s = scale || 1;
  const S = (v) => v * s;
  housing(ctx, 0, 0, ANCHOR.w, ANCHOR.h, s);

  // 9 cells at 27 wide + 2 gap = 259; centred in 280 leaves 10.5 each side.
  field(ctx, {
    scale: s, x: 10.5, y: 15, label: 'STATE', labelPx: 8, gap: 9,
    cells: STATE_CELLS, from: from.word, to: to.word, t,
    cw: 27, ch: 58, cgap: 2, charPx: 36, baseline: 1
  });

  // The maker's plate. A real board carries one; this one carries OURS. The Solari
  // reference is design lineage and lives in this file's header, never on the plate.
  ctx.save();
  ctx.strokeStyle = '#2e2a24';
  ctx.lineWidth = Math.max(1, S(0.8));
  ctx.beginPath();
  ctx.roundRect(S(10.5), S(94), S(259), S(19), S(2));
  ctx.stroke();
  ctx.restore();
  B.drawEngraved(ctx, { scale: s, x: 140, y: 100.5, text: 'HQ MOTOR SERVICE',
                        fontPx: 8.5, weight: 600, align: 'center', ink: '#7c7466' });
  B.drawEngraved(ctx, { scale: s, x: 140, y: 109, text: 'HOUSTON',
                        fontPx: 6.5, weight: 400, align: 'center', ink: '#585144' });
}

/**
 * THE STRIP — D1:H1. ⏸ BUILT AND TESTED, NOT INSTALLED IN v1.
 *
 * Kept because the alpha windows are a proven capability worth having on the shelf, and
 * because it costs nothing sitting here. It becomes worth installing only if D1's headline
 * is ever reworded to carry FIGURES rather than the state — right now the board would say
 * RESTING in 36px flaps while D1 says "the floor is asleep" 20px to its right, which is the
 * exact D1-repeating-the-dial fault fixed on 2026-09-02. ⛔ Do not install both until that
 * formula question is settled, and settle it on its own — never bundle a look with a repair.
 *
 * ⚠ clearRect is what makes the window a real hole. The canvas starts transparent, the
 *   housing fills it, and the window is CUT back out — so the alpha the gate measured is
 *   what reaches Sheets.
 */
function drawStrip(ctx, from, to, t, scale) {
  const s = scale || 1;
  const S = (v) => v * s;
  housing(ctx, 0, 0, STRIP.w, STRIP.h, s);

  [['figA', 12], ['figB', 122]].forEach(function (pair) {
    const k = pair[0];
    field(ctx, {
      scale: s, x: pair[1], y: 12, label: to[k].label, labelPx: 7.5, gap: 6,
      cells: FIG_CELLS, from: from[k].value, to: to[k].value, t,
      cw: 22, ch: 26, cgap: 2, charPx: 17, baseline: 0.3, act: to[k].act
    });
  });

  // ── the two windows ───────────────────────────────────────────────────────
  const wins = [
    { x: COL_D + 6,          w: COL_E - 12,  label: 'SYSTEM PULSE' },
    { x: COL_D + COL_E + 6,  w: COL_FH - 12, label: 'THE DAY' }
  ];
  for (const win of wins) {
    const y = 15, h = STRIP.h - y - 7;
    ctx.clearRect(S(win.x), S(y), S(win.w), S(h));           // ← the actual hole

    ctx.save();                                               // recessed bezel around it
    ctx.strokeStyle = '#0a0908';
    ctx.lineWidth = Math.max(1, S(1.2));
    ctx.beginPath();
    ctx.roundRect(S(win.x) - S(1.2), S(y) - S(1.2), S(win.w) + S(2.4), S(h) + S(2.4), S(2));
    ctx.stroke();
    ctx.strokeStyle = '#2f2b24';                              // one lit lip, from above
    ctx.lineWidth = Math.max(1, S(0.7));
    ctx.beginPath();
    ctx.moveTo(S(win.x) - S(1.2), S(y) - S(1.6));
    ctx.lineTo(S(win.x + win.w) + S(1.2), S(y) - S(1.6));
    ctx.stroke();
    ctx.restore();

    B.drawEngraved(ctx, { scale: s, x: win.x, y: 8.5, text: win.label,
                          fontPx: 6.5, weight: 500, ink: '#585144' });
  }
}

module.exports = {
  ANCHOR, STRIP, COL_D, COL_E, COL_FH,
  STATE_CELLS, LEAD_CELLS, FIG_CELLS, WORD, buildBoard, drawAnchor, drawStrip
};


/* ═══════════════════════════════════════════════════════════════════════════
   THE FIGURES FACE — the alternative the true-size render argued for.
   ═══════════════════════════════════════════════════════════════════════════

   ⚠⚠ WHY THIS EXISTS. Rendering the state face at true size showed the state word
      appearing THREE TIMES across row 1: the board says RESTING, D1 says "the floor is
      asleep", and E1's pulse says RESTING again. That is the D1-repeating-the-dial fault
      of 2026-09-02, one level up — and no amount of drawing fixes it, because the
      redundancy is in WHAT the board says, not how.

   ⭐ AND THE FIGURES ARE THE BETTER FLAP SUBJECT ANYWAY. A flap is for a value that steps
      to a new DISCRETE value — which is exactly what a count does when an order ships.
      "30 -> 31" is the honest use of the mechanism. It also puts back what the dial
      actually carried, in a new medium, instead of duplicating what the cells already say.

   ⚠ COST: a count steps a few dozen times a day rather than a handful, so this is ~30-60
     swaps/day instead of ~8. Still off the trigger budget entirely (doPost is not a
     trigger) and still a rounding error against runPublishTick's 1,440.
*/
function buildFigures(q) {
  const st = buildState(q);
  const n = (v) => (v == null || v === '' || isNaN(Number(v)) ? null : Number(v));
  const cell = (v) => padFig(v == null ? NONE : String(v), 3);
  return {
    verdict: st.verdict,
    // ⚠ Both are COUNTS, deliberately. No duration goes on these flaps — a duration ticks
    //   every minute, which would make the image dirty every minute and is also a lie
    //   about the mechanism.
    a: { label: 'TO GRAB',   value: cell(n(q.g)), act: st.verdict === 'busy' || st.verdict === 'late' },
    b: { label: 'OUT TODAY', value: cell(n(q.p)), act: false }
  };
}

function drawAnchorFigures(ctx, from, to, t, scale) {
  const s = scale || 1;
  const S = (v) => v * s;
  housing(ctx, 0, 0, ANCHOR.w, ANCHOR.h, s);

  [['a', 12], ['b', 145]].forEach(function (pair) {
    const k = pair[0];
    field(ctx, {
      scale: s, x: pair[1], y: 15, label: to[k].label, labelPx: 8, gap: 9,
      cells: 3, from: from[k].value, to: to[k].value, t, abc: B.DIGITS,
      cw: 38, ch: 58, cgap: 2, charPx: 40, baseline: 1, act: to[k].act
    });
  });

  ctx.save();
  ctx.strokeStyle = '#2e2a24';
  ctx.lineWidth = Math.max(1, S(0.8));
  ctx.beginPath();
  ctx.roundRect(S(12), S(94), S(256), S(19), S(2));
  ctx.stroke();
  ctx.restore();
  B.drawEngraved(ctx, { scale: s, x: 140, y: 100.5, text: 'HQ MOTOR SERVICE',
                        fontPx: 8.5, weight: 600, align: 'center', ink: '#7c7466' });
  B.drawEngraved(ctx, { scale: s, x: 140, y: 109, text: 'HOUSTON',
                        fontPx: 6.5, weight: 400, align: 'center', ink: '#585144' });
}

module.exports.buildFigures = buildFigures;
module.exports.drawAnchorFigures = drawAnchorFigures;


/**
 * drawBoardFace — the SETTLED board, for the `=IMAGE()` path.
 *
 * ⭐⭐ THIS IS STEP 1, AND IT COSTS NOTHING. A split-flap board standing still is still a
 *    split-flap board — the tiles, the seams, the housing, the plate. So the face ships
 *    through the SAME `=IMAGE()` formula the dial uses today: the numbers ride in the query
 *    string, this container draws them, Sheets fetches the PNG. **Zero Apps Script.**
 *
 *    The animated version is the same drawing at phase > 0, and it is a separate decision
 *    with a separate cost, because motion is the ONLY thing `=IMAGE()` cannot do (it renders
 *    frame one of a GIF and never advances — settled 2026-08-30). Shipping the look first
 *    means living with the design before paying for machinery to move it.
 *
 * ⚠ from === to, so every cell is settled and nothing is mid-flip. `t` is irrelevant.
 */
function drawBoardFace(ctx, q, scale) {
  const b = buildFigures(q);
  drawAnchorFigures(ctx, b, b, 1e6, scale || 1);
}

module.exports.drawBoardFace = drawBoardFace;
