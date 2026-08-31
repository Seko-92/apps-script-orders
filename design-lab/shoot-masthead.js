/**
 * shoot-masthead.js — render the five masthead faces and assemble them into GIFs.
 *
 * The renderer problem, solved with tooling already on this machine: design the faces as
 * HTML, step a DETERMINISTIC t (0..1) so every frame is reproducible, screenshot each,
 * and let ffmpeg assemble the GIF. No service on the VPS, no /api route, no Caddyfile edit.
 *
 * ⚠⚠ SETTLED 2026-08-30: GOOGLE SHEETS DOES NOT ANIMATE GIFs IN =IMAGE(). It renders the
 *    FIRST FRAME as a still. Tested with a known-good animated GIF; it sat motionless.
 *    So the ceiling really is ~1 frame per MINUTE (formula recalc), and the faces are
 *    five stills picked live by formula — which is still the whole feature.
 *
 *    Consequence: we ship PNG. A GIF was 180KB to display 9KB of content. GIF assembly
 *    is kept behind GIF=1 purely so the finding can be re-tested cheaply if Sheets ever
 *    changes; nothing reads those files.
 *
 * ⚠ Frames are deterministic BY CONSTRUCTION — masthead.html computes everything from ?t=
 *   rather than running a CSS animation. A harness that screenshots a live animation
 *   catches whatever frame the compositor happened to be on (the page.clock lesson).
 *
 *   node shoot-masthead.js            all five
 *   node shoot-masthead.js held late  only these
 */
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const path = require('path'), fs = require('fs');

const STATES = ['clear', 'busy', 'late', 'stale', 'held', 'rest'];
const ONLY   = process.argv.slice(2).filter(a => STATES.includes(a));
const PICK   = ONLY.length ? ONLY : STATES;

const FRAMES = Number(process.env.FRAMES || 36);   // 36 @ 12fps = a 3s loop
const FPS    = Number(process.env.FPS || 12);
const SCALE  = Number(process.env.SCALE || 4);     // deviceScaleFactor → 1120x176 output
const VER    = process.env.VER || 'v4';

// ⚠⚠ THESE MUST EQUAL masthead.html's OWN width/height. They are a second copy of a
//    number that lives in the CSS, so they are ASSERTED against the real .face below
//    rather than trusted — a silent mismatch would letterbox or stretch all 144 faces
//    and the contact sheet would look plausible while being wrong.
const W = Number(process.env.W || 287);
const H = Number(process.env.H || 56);

const SRC  = 'file://' + path.join(__dirname, 'masthead.html');
const TMP  = path.join(__dirname, 'renders', 'mast-frames');
const OUT  = path.join(__dirname, 'renders', 'mast');

(async () => {
  fs.rmSync(TMP, { recursive: true, force: true });
  fs.mkdirSync(TMP, { recursive: true });
  fs.mkdirSync(OUT, { recursive: true });

  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: W, height: H }, deviceScaleFactor: SCALE });
  const errs = [];
  p.on('pageerror', e => errs.push(String(e)));
  p.on('console', m => { if (m.type() === 'error') errs.push(m.text()); });

  // ⚠ load ONCE. Navigating per frame waits on Google Fonts 180 times and times out.
  await p.goto(SRC, { waitUntil: 'networkidle' });
  await p.waitForSelector('html[data-ready="1"]');
  await p.evaluate(() => document.fonts.ready);

  // ⚠ The renderer and the stylesheet must agree about the canvas. Fail LOUDLY here
  //   rather than emit 144 subtly-wrong PNGs — the whole set would still open, still
  //   look like a masthead, and be the wrong shape in the cell.
  const face = await p.evaluate(() => {
    const r = document.getElementById('face').getBoundingClientRect();
    return { w: Math.round(r.width), h: Math.round(r.height) };
  });
  if (face.w !== W || face.h !== H) {
    throw new Error('viewport/CSS drift: .face is ' + face.w + 'x' + face.h +
                    ' but the renderer is shooting ' + W + 'x' + H +
                    '. Reconcile masthead.html with W/H in this file.');
  }
  console.log(`canvas ${W}x${H} @${SCALE}x → ${W*SCALE}x${H*SCALE}px  ·  version ${VER}`);

  // ---- the hour-lit set: every state at every Houston hour ---------------------------
  // ⭐ The masthead cannot MOVE — a floating image scrolls off the frozen banner, and
  //    =IMAGE() shows a GIF's first frame only. But it can be LIT, and light is the one
  //    thing that reads correctly at ~1 frame per minute, because a day changes slowly
  //    anyway. This is the Floor Board's night dial reaching the sheet at last.
  let n = 0;
  for (const state of PICK) {
    for (let h = 0; h < 24; h++) {
      const stillT = (state === 'clear' || state === 'busy' || state === 'rest') ? 0.5 : 0.25;
      await p.evaluate(([s, t, hh]) => window.render(s, t, hh), [state, stillT, h]);
      await p.screenshot({
        path: path.join(OUT, `${state}-h${String(h).padStart(2, '0')}-${VER}.png`)
      });
      n++;
    }
    console.log(`  ${state.padEnd(6)} → 24 hours`);

    if (!process.env.GIF) continue;
    // GIFs are hour-less on purpose: they only matter on a surface that scrolls, and
    // 24 x 6 animated sets would be ~25MB for a case we do not have yet.
    for (let i = 0; i < FRAMES; i++) {
      await p.evaluate(([s, t, hh]) => window.render(s, t, hh), [state, i / FRAMES, 12]);
      await p.screenshot({ path: path.join(TMP, `${state}-${String(i).padStart(3, '0')}.png`) });
    }
    const gif = path.join(OUT, `${state}-${VER}.gif`);
    execFileSync('ffmpeg', ['-y', '-loglevel', 'error', '-framerate', String(FPS),
      '-i', path.join(TMP, `${state}-%03d.png`),
      '-vf', 'split[a][b];[a]palettegen=max_colors=128:stats_mode=full[p];[b][p]paletteuse=dither=sierra2_4a',
      '-loop', '0', gif]);
  }
  console.log(`  ${n} hour-lit stills`);

  // contact sheet — all five stacked, so the set can be judged as a family in one look
  await b.close();

  const sheetPath = path.join(OUT, '_contact.png');
  // contact sheet at a representative working hour
  execFileSync('convert', [
    ...PICK.map(s => path.join(OUT, `${s}-h11-${VER}.png`)),
    '-append', sheetPath
  ]);

  if (errs.length) { console.log('\n⚠ page errors:'); errs.forEach(e => console.log('   ' + e)); }
  console.log(`\n${PICK.length} faces → ${OUT}`);
  console.log(`contact sheet → ${sheetPath}`);
})();
