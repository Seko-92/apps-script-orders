/**
 * test-masthead.js — the masthead's formulas, asserted against the REAL BrandTheme.js.
 *
 * ⚠⚠ THE HEADLINE ASSERTION IS E1. ActivityLog.js regex-parses "h:mm AM/PM" out of that
 *    cell into cockpit.lastSyncMinutes, which drives the Floor Board's heartbeat dot,
 *    the sidebar System Pulse, /status and the published tick. If that format is ever
 *    dropped the board's heartbeat dies with NO error anywhere — the exact class of
 *    silent failure this project keeps getting bitten by. Both regexes are extracted
 *    from the shipped files rather than retyped, so they cannot drift.
 */
const fs = require('fs'), path = require('path'), vm = require('vm');
const R = f => fs.readFileSync(path.join(__dirname, '..', f), 'utf8');

let pass = 0, fail = 0;
const ok = (name, cond, got) => {
  if (cond) { pass++; console.log('  ✓ ' + name); }
  else { fail++; console.log('  ✗ ' + name + (got !== undefined ? '  → got ' + JSON.stringify(got) : '')); }
};

// ---- a sheet that records what was written, per range ---------------------------------
function fakeSheet() {
  const w = {};
  // ⚠ Ranges that were STYLED, not just written. The 2026-09-02 live failure was a styling
  //   call on an undefined address — every formula was fine, so a writes-only fake could
  //   never have seen it.
  const touched = new Set();
  const cell = ref => ({
    setFormula(f) { w[ref] = f; return this; },
    setBorder() { touched.add(ref); return this; },
    setBackground(){return this;}, setFontColor(){return this;}, setFontFamily(){return this;},
    setFontSize(){return this;},   setFontWeight(){return this;}, setWrap(){return this;},
    setHorizontalAlignment(){return this;}, setVerticalAlignment(){return this;},
    setFormulas(f) { w[ref] = f; return this; },
    setDataValidation(){return this;}, getDataValidation(){return null;},
    merge(){return this;}, breakApart(){return this;}
  });
  return {
    writes: w, touched: touched,
    getRange: (a, b, c, d) => {
      // ⚠ THROW ON A NULL ADDRESS, exactly as Sheets does. A fake that quietly accepts
      //   undefined cannot see a bug that the real getRange refuses outright.
      if (a === undefined || a === null) throw new Error('Argument cannot be null: a1Notation');
      const ref = typeof a === 'string' ? a : 'R' + a + 'C' + b + '+' + c + 'x' + d;
      touched.add(ref);
      return cell(ref);
    },
    setRowHeight(){}, hideSheet(){}, getName(){ return '__SparkData'; }
  };
}

// ---- load the real file ---------------------------------------------------------------
const sandbox = {
  console, Date, Math, String, Number, JSON, RegExp, Object, Array,
  SPREADSHEET_ID: 'x', MAIN_SHEET_NAME: 'All orders',
  Schema: { dataStartRow: 4, dataWidth: 10, cellSyncTime: 'E1', cellMasthead: 'A1',
            cellStats: 'D1', cellDayCurve: 'F1',
            cellEmployeeId: 'F2', cellAdjustmentId: 'H2',
            pickIdA1: function (which) { return which === 'adjustment' ? 'H2' : 'F2'; } },
  SpreadsheetApp: { flush(){}, openById: () => ({ getSheetByName: () => null, insertSheet: () => null }) }
};
vm.createContext(sandbox);
// BrandTheme references Sheets services at call time only, so the whole file loads clean.
vm.runInContext(R('BrandTheme.js'), sandbox);

console.log('\nA · the four zones land in the right cells');
const sh = fakeSheet();
sandbox._setSystemPulseBannerFormulas(sh);
const W = sh.writes;
ok('A1 holds the face image',      /^=IFERROR\(IMAGE\(/.test(W.A1 || ''), (W.A1||'').slice(0,28));
ok('D1 holds the headline',        /^=IF\('__SparkData'!A6="rest"/.test(W.D1 || ''));
ok('E1 holds the pulse',           /^=IF\('__SparkData'!A4<0/.test(W.E1 || ''));
ok('F1 holds the SPARKLINE',       /SPARKLINE\('__SparkData'!A1:X1/.test(W.F1 || ''));
ok('nothing was written to B1/C1', !W.B1 && !W.C1);
// ⚠ the extension must match what is actually SHIPPED. Sheets shows only a GIF's first
//   frame (tested 2026-08-30), so these are PNG stills — and a config that drifts from
//   the files on the VPS turns every face into the text chip, silently, behind a 200.
// ⚠ THE ART CONTRACT MOVED. With the dial there is no filename to drift — the risk is the
//   ROUTE, and a missing route under this host answers 200 with the Floor Board's HTML,
//   which IMAGE() cannot decode. So pin the endpoint and the mode-4 dimensions instead.
if (sandbox.MASTHEAD.dial) {
  ok('A1 points at the dial renderer', (W.A1 || '').indexOf(sandbox.MASTHEAD.dialUrl) > -1);
  ok('⚠ mode 4 with explicit w/h, so nothing letterboxes',
     new RegExp(',4,' + sandbox.MASTHEAD.dialH + ',' + sandbox.MASTHEAD.dialW + '\\)').test(W.A1 || ''),
     (W.A1||'').slice(-40));
  ['s=', '&t=', '&o=', '&g=', '&r=', '&p=', '&u=', '&y=', '&l='].forEach(function (k) {
    ok('the URL carries ' + JSON.stringify(k), (W.A1 || '').indexOf('"' + k + '"') > -1 ||
       (W.A1 || '').indexOf(k) > -1);
  });
} else {
  ok('the face URL uses MASTHEAD.ext', new RegExp('\\.' + sandbox.MASTHEAD.ext + '"').test(W.A1 || ''), (W.A1||'').slice(-46));
  ok('⚠ and it is NOT .gif', !/\.gif"/.test(W.A1 || ''));
}

console.log('\nB · ⚠⚠ E1 still speaks the format the Floor Board parses');
// the REAL regex, lifted out of the shipped ActivityLog.js — never retyped
const alSrc = R('ActivityLog.js');
// ⚠ Find the CODE line, not the comment above it — an assertion made against source
//   text must never match documentation (the 2026-08-21 openById lesson).
const heartLine = alSrc.split('\n').find(l =>
  l.includes('AM|PM') && l.includes('.match(') && !/^\s*(\/\/|\*)/.test(l));
const reLit = heartLine && heartLine.match(/\/.*\/i/);
ok('extracted the heartbeat regex from ActivityLog.js', !!reLit, reLit && reLit[0]);
const HEART = eval(reLit[0]);
ok('E1 formula still calls TEXT(...,"h:mm AM/PM")', /TEXT\('__SparkData'!A3,"h:mm AM\/PM"\)/.test(W.E1 || ''));
// the strings that formula can actually produce
const LIVE    = '⏱ 🟢 ALIVE · 9:57 PM · 8m ago';
const STALE   = '⏱ 🔴 STALE · 8:03 PM · 2h 14m ago';
const OFFLINE = '⏱ OFFLINE · no activity logged';
ok('a healthy pulse parses → heartbeat lives', HEART.test(LIVE), LIVE.match(HEART) && LIVE.match(HEART)[0]);
ok('a stale pulse still parses',               HEART.test(STALE));
ok('OFFLINE yields no time — unchanged from today', !HEART.test(OFFLINE));

console.log('\nC · the verdict is computed ONCE and shared');
const ss = { getSheetByName: () => sh, insertSheet: () => sh };
sandbox._ensureSparkData(ss);
const S = sh.writes;
ok('A6 is the verdict',            /^=IF\(A13,"rest"/.test(S.A6 || ''), (S.A6||'').slice(0,22));
ok('late outranks stale',          (S.A6||'').indexOf('"late"') < (S.A6||'').indexOf('"stale"'));
ok('stale outranks busy',          (S.A6||'').indexOf('"stale"') < (S.A6||'').indexOf('"busy"'));
ok('⚠ A4 = -1 (log unreadable) lands on stale, not clear', /OR\(A4<0,A4>60\)/.test(S.A6 || ''));
ok('the face reads A6',            /'__SparkData'!A6/.test(W.A1 || ''));
ok('the headline reads A6 too',    /'__SparkData'!A6/.test(W.D1 || ''));

console.log('\nD · every published read degrades rather than throwing');
['A7','A9','A10'].forEach(r => ok(r + ' is IFERROR-wrapped', /^=IFERROR\(/.test(S[r] || '')));
ok('A8 (queue) reads the sheet, not __Published', /COUNTIF\('All orders'!F4:F/.test(S.A8 || ''));
ok('the face falls back to the text chip', /,"HQ"\)$/.test(W.A1 || ''));
ok('the curve falls back to blank',        /,""\)$/.test(W.F1 || ''));

// a real tick, and the null case that must read as "nothing pending"
const tick = '{"cockpit":{"shippedToday":16,"receivedToday":109,"oldestPendingMinutes":252}}';
ok('the published regex finds a real oldest-pending', /"oldestPendingMinutes":(\d+)/.exec(tick)[1] === '252');
ok('null oldest-pending → no match → IFERROR → 0',    !/"oldestPendingMinutes":(\d+)/.test('{"oldestPendingMinutes":null}'));

console.log('\n⚠ E · off-hours — the sheet never knew, and that made the masthead lie');
const V = S.A6 || '', off = S.A13 || '', opens = S.A14 || '';
ok('A13 tests off-hours at all',            /HOUR\(NOW\(\)\)<9/.test(off) && /HOUR\(NOW\(\)\)>=17/.test(off));
ok('⚠ and it counts WEEKENDS off, like the board — not just the hour, like the sidebar',
   /WEEKDAY\(NOW\(\)\)=1/.test(off) && /WEEKDAY\(NOW\(\)\)=7/.test(off));
ok('rest outranks late',   V.indexOf('"rest"') < V.indexOf('"late"'));
ok('rest outranks stale',  V.indexOf('"rest"') < V.indexOf('"stale"'));
ok('rest outranks busy',   V.indexOf('"rest"') < V.indexOf('"busy"'));
ok('a quiet night can no longer read as a dead pipeline',
   V.indexOf('A13') < V.indexOf('A4<0'));
ok('"opens in" skips the weekend',          /WEEKDAY\([\s\S]*?\)=7,2/.test(opens));
ok('opens-in can never go negative',        /^=MAX\(0,/.test(opens));
// ⚠ The countdown moved ONTO the dial, so D1 no longer prints it — but it must still name
//   the state, or an operator glancing at the words alone learns nothing.
ok('D1 has a rest branch that names the state',
   /A6="rest"[\s\S]*?"the floor is asleep"/.test(W.D1 || ''));
// ⚠ ASSERT THE BEHAVIOUR, NOT THE ORDERING. The first version of this pinned
//   `A6="rest",IF(A8>0` — the exact shape the branch happened to have — so moving the
//   count to a SUFFIX ("opens in 19h 44m · 12 waiting") failed a test whose stated
//   subject was preserved perfectly. A test that describes the implementation instead
//   of the promise blocks the refactor it should have been protecting.
ok('D1 still reports what is waiting overnight',
   /A6="rest"[\s\S]*?A8>0[\s\S]*?waiting/.test(W.D1 || ''), W.D1);

// ⭐ THE SPLIT-VOICE RULE (2026-08-31). The face carries the STATE in words; D1 carries
//   NUMBERS. Before this, both narrated the same thing in two typefaces 287px apart —
//   on the clear verdict they used the SAME WORDS, "nothing waiting". That redundancy is
//   the "two objects" feeling, so it is pinned here rather than left to taste.
// ⚠⚠ THE RULE, AND IT IS THE ONE THAT KEEPS THE BANNER HONEST: D1 names the STATE and says
//    WHERE; the DIAL carries the FIGURES. Drawn at true width, the old D1 read
//    "12 to grab / 14 out today" while the dial said exactly that 200px away, and 'stale'
//    printed one duration three times across D1, the dial and the pulse.
// ⚠ This caught a REGRESSION during the very commit that introduced the rule — the 'late'
//   branch was written as "3 past the 3h line", which is the dial's own flank label. A
//   rule with no test is a rule that lasts one branch.
['to grab', 'out today', 'until open', 'oldest waiting', 'past the line',
 'nothing waiting', 'nothing has landed', 'orders being picked',
 'TO GRAB', 'OLDEST', 'LAST SEEN'].forEach(function (phrase) {
  ok('D1 does not repeat the dial: ' + JSON.stringify(phrase),
     (W.D1 || '').toLowerCase().indexOf(phrase.toLowerCase()) === -1);
});
ok('⚠ D1 is lowercase — the face does the shouting',
   // ⚠ Strip FUNCTION NAMES before judging the prose. TEXT( is four capitals and is not
   //   something anyone reads on the banner — the assertion is about the words, not the code.
   !/[A-Z]{4,}/.test((W.D1 || '').replace(/'__SparkData'!A\d+|IFERROR|REGEXEXTRACT|VALUE|CHAR|TEXT|IF/g, '')));
// ⚠ LET() threw "Formula parse error" on this live sheet (2026-06-05). It must never
//   reappear in a banner formula, however tempting the repeated base expression is.
ok('⚠ no LET() in any masthead formula',
   !Object.values(S).concat(Object.values(W)).some(f => typeof f === 'string' && /\bLET\s*\(/.test(f)));

console.log('\nF · and the date math itself, not just its spelling');
// The same arithmetic A14 encodes, run in JS. A regex on formula TEXT proves the
// characters are present; this proves the answer is right.
function opensInMin(d) {
  const base = new Date(d);
  base.setDate(d.getDate() + (d.getHours() < 9 ? 0 : 1));
  base.setHours(0, 0, 0, 0);
  const wd = base.getDay() + 1;                       // WEEKDAY() type 1: Sun=1 .. Sat=7
  const open = new Date(base);
  open.setDate(base.getDate() + (wd === 7 ? 2 : wd === 1 ? 1 : 0));
  open.setHours(9, 0, 0, 0);
  return Math.max(0, Math.round((open - d) / 60000));
}
const at = (y, m, day, h, mi) => new Date(y, m, day, h, mi);
// 2026-08-29 is a Saturday — the very moment in the screenshot
ok('Sat 03:45 → opens Monday 9am',  opensInMin(at(2026,7,29,3,45))  === 2 * 1440 + 5 * 60 + 15,
   opensInMin(at(2026,7,29,3,45)));
ok('Fri 18:00 → skips the weekend', opensInMin(at(2026,7,28,18,0))  === 2 * 1440 + 15 * 60,
   opensInMin(at(2026,7,28,18,0)));
ok('Tue 07:00 → 2h, same morning',  opensInMin(at(2026,8,1,7,0))    === 120, opensInMin(at(2026,8,1,7,0)));
ok('Tue 18:00 → 15h, next morning', opensInMin(at(2026,8,1,18,0))   === 900, opensInMin(at(2026,8,1,18,0)));
ok('never negative during work hours', opensInMin(at(2026,8,1,11,0)) >= 0);

console.log('\nG · the age formatter — "5h 60m" was live on the banner');
// The arithmetic _fmtMinsExpr encodes, run in JS. Same reasoning as section F: a regex
// over formula text proves the characters are there; this proves the ANSWER is right.
const fmt = x => { const r = Math.round(x);
  return r < 60 ? r + 'm' : Math.trunc(r / 60) + 'h ' + (r % 60) + 'm'; };
ok('359.6m → 6h 0m, not 5h 60m', fmt(359.6) === '6h 0m', fmt(359.6));
ok('59.6m  → 1h 0m, not 60m',    fmt(59.6)  === '1h 0m', fmt(59.6));
ok('8.4m   → 8m',                fmt(8.4)   === '8m',    fmt(8.4));
ok('134m   → 2h 14m',            fmt(134)   === '2h 14m',fmt(134));
let sixty = null;
for (let x = 0; x < 5000; x += 0.1) { const t = fmt(x); if (/\b60m$/.test(t)) { sixty = [x, t]; break; } }
ok('⚠ no input anywhere can render "…60m"', sixty === null, sixty);
ok('the formula rounds the TOTAL before splitting', /INT\(ROUND\(/.test(S.A11 || ''), (S.A11||'').slice(0,40));

console.log('\nH · the pulse must not contradict the face');
ok('E1 reads the SAME off-hours flag the verdict does', /'__SparkData'!A13/.test(W.E1 || ''));
ok('E1 has a RESTING tier',                     /RESTING/.test(W.E1 || ''));
ok('⚠ and it outranks STALE inside E1 too',
   (W.E1 || '').indexOf('RESTING') < (W.E1 || '').indexOf('STALE'));
ok('E1 still carries h:mm AM/PM after the change',
   /TEXT\('__SparkData'!A3,"h:mm AM\/PM"\)/.test(W.E1 || ''));
ok('and the real heartbeat regex still parses a RESTING line',
   HEART.test('⏱ ⚪ RESTING · 9:57 PM · 6h 0m ago'));

console.log('\nI · the resting curve wears yesterday, not an empty today');
const curve = W.F1 || '', yRow = (S['R2C1+1x24'] || [[]])[0] || [];
// ⚠ assert the harness can SEE the row before asserting anything about it — every
//   check below uses .every(), which is vacuously true on an empty array. That is the
//   blindness that let the command palette survive two emoji sweeps.
ok('the harness can see yesterday\'s row at all', yRow.length > 0, yRow.length);
ok('the curve branches on the SAME off-hours flag', /IF\('__SparkData'!A13/.test(curve));
ok('resting reads row 2 (yesterday)',   /A13,SPARKLINE\('__SparkData'!A2:X2/.test(curve));
ok('working reads row 1 (today)',       /SPARKLINE\('__SparkData'!A1:X1/.test(curve));
ok('⚠ and it goes COOL when resting, so the banner rests together',
   new RegExp('A2:X2[^)]*' + sandbox.MASTHEAD.restAccent.replace('#','#')).test(curve));
ok('the live curve stays brand yellow', /A1:X1[\s\S]*#ffd400/.test(curve));
ok('yesterday is 24 buckets',           yRow.length === 24, yRow.length);
ok('yesterday counts TODAY()-1',        /TODAY\(\)-1\+0\/24/.test(yRow[0] || ''));
ok('the last bucket closes at midnight',/TODAY\(\)-1\+24\/24/.test(yRow[23] || ''));
ok('each bucket degrades to 0',         yRow.length === 24 && yRow.every(f => /^=IFERROR\(/.test(f)));
ok('⚠ the face and its curve share ONE rest tone',
   curve.indexOf(sandbox.MASTHEAD.restAccent) > -1);

console.log('\nJ · the face keeps the hour');
const face = W.A1 || '';
ok('the URL carries the Houston hour',
   sandbox.MASTHEAD.dial
     ? /TEXT\(HOUR\(NOW\(\)\),"00"\)&TEXT\(MINUTE\(NOW\(\)\),"00"\)/.test(face)
     : /"-h"&TEXT\(HOUR\(NOW\(\)\),"00"\)/.test(face));
ok('⚠ zero-padded, so h09 is not h9',   /"00"/.test(face));
ok('it still carries the state first',  face.indexOf('A6') < face.indexOf('HOUR(NOW())'));
ok('and still falls back to the chip',  /,"HQ"\)$/.test(face));
ok('version is the current set',        /^v\d+$/.test(sandbox.MASTHEAD.version), sandbox.MASTHEAD.version);
// HOUR(NOW()) is spreadsheet-timezone, and setupActivityLogSheet pins that to Chicago —
// the same clock A13 uses, so the light and the off-hours verdict can never disagree.
ok('⚠ light and off-hours read the SAME clock', /HOUR\(NOW\(\)\)/.test(face) && /HOUR\(NOW\(\)\)/.test(S.A13 || ''));

console.log('\nK · row 2 is the eBay table\'s NAME, not a canvas');
// ⚠⚠ The sky cost the eBay table its label. The sheet is TWO stacked tables and row 2 is
//    the counterpart to the gold "▌ DIRECT" divider — that cell was never empty space.
ok('the sky is OFF',                    sandbox.MASTHEAD.sky === false);
ok('so A2 is left to setupEbayLogo',    !(W.A2 || '').length, W.A2);
ok('the switch still exists for later', 'sky' in sandbox.MASTHEAD);
ok('⚠ the day CURVE stays brand yellow — data is not decorated by the clock',
   /A1:X1[\s\S]*#ffd400/.test(W.F1 || '') && !/HOUR\(NOW\(\)\)/.test(W.F1 || ''));

// ==========================================================================================
console.log('\nL · ⚠⚠ the row-2 styler must survive BOTH layouts');
// 2026-09-02, on the live sheet: setupMasthead reported
//     ✗ logo zone + Pick ID badges — Exception: Argument cannot be null: a1Notation
// and both Pick ID dropdowns silently kept the cream banner styling instead of their dark
// badges. Cause: `var adj = Schema.pickIdA1('adjustment')` had been declared INSIDE the
// non-dial branch, so with the dial on it was hoisted-but-unassigned and the badge loop
// called getRange(undefined). Every formula was correct, which is exactly why the existing
// writes-only assertions all passed — this one DRIVES the function instead of reading it.
[true, false].forEach(function (dialOn) {
  const label = dialOn ? 'dial' : 'face';
  const prev = sandbox.MASTHEAD.dial;
  sandbox.MASTHEAD.dial = dialOn;
  const s2 = fakeSheet();
  let threw = null;
  try { sandbox._styleBannerRow2(s2); } catch (e) { threw = String((e && e.message) || e); }
  ok('_styleBannerRow2 does not throw · ' + label, threw === null, threw);
  ok('both Pick ID cells get styled · ' + label,
     s2.touched.has('F2') && s2.touched.has('H2'), [...s2.touched].join(','));
  ok('the logo zone matches the layout · ' + label,
     s2.touched.has(dialOn ? 'D2:E2' : 'A2:E2'), [...s2.touched].join(','));
  sandbox.MASTHEAD.dial = prev;
});

console.log('\n' + (fail ? '✗ ' + fail + ' FAILED' : '✓ all') + ' · ' + pass + ' passed\n');
process.exit(fail ? 1 : 0);
