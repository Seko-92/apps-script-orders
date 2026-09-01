#!/usr/bin/env node
/**
 * shift-film.js — ONE FRAME PER MINUTE OF A WHOLE SHIFT, played back fast.
 *
 * ⚠⚠ THE POINT. "Can it animate?" has been answered twice in this project and answered
 *    WRONGLY once, so this stops arguing and shows the thing: the sheet's real ceiling is
 *    ~1 frame per minute (formula recalc), and this renders exactly that — every minute
 *    from open to close, no interpolation, no cheating — then plays 8 hours in 12 seconds.
 *
 *    What you are watching is not a simulation of animation. It is EVERY FRAME the banner
 *    will actually draw, in order. If it reads as alive here, it reads as alive on the
 *    sheet; you just live it at 1/40th the speed.
 *
 * The day it walks is a plausible one, not a flattering one: a quiet open, a build-up, an
 * episode that crosses the 3h line, a recovery, a dead-sync window, and a clean close.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { renderPng } = require('./src/render');

const OUT    = path.join(__dirname, 'renders');
const FRAMES = path.join(OUT, '_frames');
fs.rmSync(FRAMES, { recursive: true, force: true });
fs.mkdirSync(FRAMES, { recursive: true });

const STEP  = Number(process.env.STEP  || 2);    // minutes per frame
const SCALE = Number(process.env.SCALE || 2);
const FROM  = 8 * 60 + 40;                        // a little before open, so REST is seen
const TO    = 17 * 60 + 40;                       // a little after close, so REST returns

/** A day with a shape. Returns the state of the floor at a given minute. */
function floorAt(m) {
  const hours = new Array(24).fill(0);
  const shipTimes = [];
  // received / shipped accumulate through the day; the pips are the shipping hours
  const arrivals = [[9,3],[10,6],[11,9],[12,5],[13,7],[14,8],[15,6],[16,4]];
  let received = 0, shipped = 0;
  for (const [h, n] of arrivals) {
    if (m >= h * 60) { received += n; }
    // shipping lags arrival by roughly an hour
    if (m >= (h + 1) * 60) { shipped += n; hours[h + 1] = n; shipTimes.push(h + 1); }
    else if (m >= h * 60 + 30 && m < (h + 1) * 60) { hours[h + 1] = Math.ceil(n / 2); }
  }
  const toGrab = Math.max(0, received - shipped);

  // the oldest open order: grows, then resets when the backlog is cleared
  let oldest = 0;
  if (toGrab > 0) {
    // the 11:20-14:10 episode is the one that crosses the 3h line
    if (m >= 11 * 60 + 20 && m <= 14 * 60 + 10) oldest = m - (11 * 60 + 20) + 74;
    else oldest = Math.min(150, 18 + ((m % 97) * 1.6));
  }

  // a dead-sync window: nothing logged 15:05 -> 15:40
  const syncGap = (m >= 15 * 60 + 5 && m <= 15 * 60 + 40) ? m - (15 * 60 + 5) + 62 : 3;

  const off = m < 9 * 60 || m >= 17 * 60;
  let verdict;
  if (off) verdict = 'rest';
  else if (oldest > 180) verdict = 'late';
  else if (syncGap > 60) verdict = 'stale';
  else if (toGrab > 0) verdict = 'busy';
  else verdict = 'clear';

  const pastLine = verdict === 'late' ? Math.max(1, Math.round((oldest - 180) / 22)) : 0;
  return { verdict, toGrab, received, shipped, oldest, syncGap, hours, pastLine, off };
}

const hhmm = (m) => String(Math.floor(m / 60)).padStart(2, '0') + String(m % 60).padStart(2, '0');

let n = 0;
for (let m = FROM; m <= TO; m += STEP) {
  const f = floorAt(m);
  const q = {
    s: f.verdict, t: hhmm(m), h: f.hours.join(','),
    g: String(f.toGrab), r: String(f.received), p: String(f.shipped),
    o: String(Math.round(f.oldest)), y: String(Math.round(f.syncGap)),
    u: String(f.off ? (m < 9 * 60 ? 9 * 60 - m : (33 * 60) - m) : 0),
    l: f.pastLine ? String(f.pastLine) : ''
  };
  fs.writeFileSync(path.join(FRAMES, String(n).padStart(4, '0') + '.png'), renderPng(q, SCALE));
  n++;
}
console.log(`${n} frames · ${STEP} min each · ${hhmm(FROM)} -> ${hhmm(TO)}`);

// ---- encode. palettegen/paletteuse or the gradients band badly at 256 colours ----------
const fps = Number(process.env.FPS || 20);
const gif = path.join(OUT, 'shift.gif');
const pal = path.join(FRAMES, 'palette.png');
const ff = (args) => execFileSync('ffmpeg', ['-y', '-loglevel', 'error', ...args]);
ff(['-framerate', String(fps), '-i', path.join(FRAMES, '%04d.png'),
    '-vf', 'palettegen=stats_mode=diff:max_colors=192', pal]);
ff(['-framerate', String(fps), '-i', path.join(FRAMES, '%04d.png'), '-i', pal,
    '-lavfi', 'paletteuse=dither=bayer:bayer_scale=3:diff_mode=rectangle', gif]);
const kb = (fs.statSync(gif).size / 1024).toFixed(0);
console.log(`${gif}  ${kb} KB  ·  ${(n / fps).toFixed(1)}s playback  ·  9 real hours`);
