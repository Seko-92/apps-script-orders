/**
 * make-alpha-probe.js — generates the ALPHA half of the phase-1 gate.
 *
 * Plan item 1.6 asked the second question and no probe ever asked it:
 *   "Does a TRANSPARENT PNG let the cell underneath show through, or does Sheets
 *    composite it onto an opaque ground?"
 *
 * The answer decides between three shapes (animated board / static frame / dead),
 * so the test has to be unambiguous. This draws a 232x56 PNG sized to cover D1
 * exactly — the live headline cell, cream text on the #1a1a1a band — carrying:
 *
 *   · a 3px MAGENTA opaque frame        -> proves the image rendered at all
 *   · an opaque WHITE reference swatch  -> the "is the hole this colour?" control
 *   · an opaque BLACK reference swatch  -> the other composite candidate
 *   · a 50%-alpha grey bar              -> is alpha BINARY or 8-bit?
 *   · everything else alpha 0           -> THE HOLE. D1's text must show through it.
 *
 * Reading it needs no measurement, only eyes:
 *   headline text readable inside the frame -> ALPHA PASSES THROUGH  -> static frame is live
 *   hole matches the WHITE swatch           -> composited onto white -> static frame is dead
 *   hole is flat black, no text             -> composited onto black -> static frame is dead
 */
const { createCanvas } = require('@napi-rs/canvas');

const W = 232, H = 56;                 // exactly D1 (col D = 232, MASTHEAD.rowHeight = 56)
const cv = createCanvas(W, H);
const g = cv.getContext('2d');

g.clearRect(0, 0, W, H);               // start fully transparent — the hole is the default

// 3px magenta frame, opaque. Unmistakably ours, and it proves the PNG drew.
g.strokeStyle = '#ff00ff';
g.lineWidth = 3;
g.strokeRect(1.5, 1.5, W - 3, H - 3);

// ⚠ EVERY reference element lives in the RIGHT THIRD. D1's headline is LEFT-aligned and
//   runs ~130px, so keeping the left 190px a clean hole is what makes the read unambiguous:
//   the text either shows through or it does not, with nothing of ours on top of it.
g.fillStyle = '#ffffff'; g.fillRect(W - 40, 6, 14, 14);   // white candidate
g.fillStyle = '#000000'; g.fillRect(W - 22, 6, 14, 14);   // black candidate

// A 50%-alpha block: if it renders as a blend, alpha is 8-bit; if it vanishes or goes
// solid, Sheets is treating alpha as binary. Free to ask while we are here.
g.fillStyle = 'rgba(128,128,128,0.5)';
g.fillRect(W - 40, 26, 32, 22);

const b64 = cv.toBuffer('image/png').toString('base64');
console.log('bytes:', cv.toBuffer('image/png').length);
console.log('data:image/png;base64,' + b64);
