// ============================================================================
// test-timezone.js — the floor is in Houston, wherever you are reading from
//
// WHY THIS EXISTS (2026-09-03, reported from use)
//   Order Lookup rendered every Activity Log timestamp with getHours()/getDate().
//   Those are LOCAL-timezone getters, so read from Riyadh (UTC+3) every warehouse
//   event shifted by 8-9 hours: a pick logged 9:01 AM on the floor displayed as
//   5:01 PM, and the timeline read as though the day happened in the evening.
//
// ⚠ THE WHOLE POINT: this file runs with TZ FORCED to something that is NOT Houston.
//   A formatter that leans on the machine clock CANNOT pass. Under TZ=America/Chicago
//   the broken and the fixed versions are indistinguishable, which is exactly why the
//   bug survived — it was only ever visible to someone reading from another country.
//
// Formatters are LIFTED FROM THE SHIPPED FILES, never re-typed.
//
// Run:  node test-timezone.js
// HEAD: SRC=/tmp/head node test-timezone.js
// ============================================================================
'use strict';
process.env.TZ = 'Asia/Riyadh';          // ⚠ must be set before any Date work

const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(SRC, f), 'utf8');

let failed = 0, passed = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` +
    (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  ok ? passed++ : failed++;
};
const section = n => console.log('\n' + n);

/** Lift a named function (plus anything it needs) out of a shipped file. */
function lift(file, names, ctx) {
  const src = read(file);
  const sb = Object.assign({ Date, Intl, String, Number, parseInt, JSON, console }, ctx || {});
  vm.createContext(sb);
  names.forEach(n => {
    let i = src.indexOf('function ' + n + '(');
    if (i === -1) i = src.indexOf('var ' + n + ' ');
    if (i === -1) { sb['__missing_' + n] = true; return; }
    if (src.startsWith('var ', i)) {                    // a const line
      const end = src.indexOf(';', i) + 1;
      vm.runInContext(src.slice(i, end), sb, { filename: file });
      return;
    }
    let d = 0, started = false;
    for (let j = i; j < src.length; j++) {
      if (src[j] === '{') { d++; started = true; }
      else if (src[j] === '}') { d--; if (started && d === 0) {
        vm.runInContext(src.slice(i, j + 1), sb, { filename: file + '#' + n }); return; } }
    }
  });
  return sb;
}

console.log('\n' + '='.repeat(74));
console.log('  timezone — the floor is in Houston (this process is NOT)');
console.log('='.repeat(74));
console.log('  process TZ: ' + process.env.TZ +
            '  ·  offset now: UTC' + (-new Date().getTimezoneOffset() / 60 >= 0 ? '+' : '') +
            (-new Date().getTimezoneOffset() / 60));

// A real Houston moment: 2026-09-03 09:01 CDT = 14:01 UTC.
const NINE_OH_ONE = Date.UTC(2026, 8, 3, 14, 1);
// And an evening one that crosses the DATE line in Riyadh: 8:30 PM CDT = 01:30 UTC next day.
const EVENING     = Date.UTC(2026, 8, 4, 1, 30);

t('sanity · this process really is NOT on Houston time',
  new Date(NINE_OH_ONE).getHours() !== 9, true);

// ============================================================================
section('A · Sidebar — Order Lookup timeline (_olTime)');
// ============================================================================
{
  const sb = lift('Sidebar.html', ['HQ_TZ', '_hqParts', '_olTime']);
  if (sb.__missing__olTime || !sb._olTime) {
    t('A0 _olTime exists', false, true);
  } else {
    t('A1 ⭐ 9:01 AM on the floor reads as 9:01 AM, not 5:01 PM',
      sb._olTime(NINE_OH_ONE), '9/3 9:01AM');
    t('A2 ⭐ an evening event keeps HOUSTON\'s date, not the reader\'s next day',
      sb._olTime(EVENING), '9/3 8:30PM');
    t('A3 a missing timestamp is a dash, never "NaN"', sb._olTime(null), '—');
  }
}

// ============================================================================
section('B · Order Case modal (fmtWhen) — the surface Look Up actually opens');
// ============================================================================
{
  const sb = lift('OrderCaseModal.html', ['OC_TZ', 'fmtWhen']);
  if (!sb.fmtWhen) {
    t('B0 fmtWhen exists', false, true);
  } else {
    t('B1 ⭐ 9:01 AM on the floor reads as 9:01 AM',
      sb.fmtWhen(NINE_OH_ONE), '9/3 9:01am');
    t('B2 ⭐ the evening event keeps Houston\'s date',
      sb.fmtWhen(EVENING), '9/3 8:30pm');
    t('B3 blank is a dash', sb.fmtWhen(0), '—');
  }
}

// ============================================================================
section('C · ⚠ DURATIONS MUST NOT BE "FIXED" — they are timezone-independent');
// ============================================================================
{
  const sb = lift('Sidebar.html', ['_olAgo'], { Date });
  if (!sb._olAgo) { t('C0 _olAgo exists', false, true); }
  else {
    const ninetyMinAgo = Date.now() - 90 * 60000;
    t('C1 90 minutes ago is 1h 30m in every zone', sb._olAgo(ninetyMinAgo), '1h 30m ago');
    t('C2 …and it still handles no timestamp', sb._olAgo(null), '—');
  }
}

// ============================================================================
section('D · THE SOURCE CONTRACT — no local getter decides a displayed time');
// ============================================================================
{
  const strip = s => s.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '');
  const grab = (file, name) => {
    const src = strip(read(file));
    const i = src.indexOf('function ' + name + '(');
    if (i === -1) return '';
    let d = 0, started = false;
    for (let j = i; j < src.length; j++) {
      if (src[j] === '{') { d++; started = true; }
      else if (src[j] === '}') { d--; if (started && d === 0) return src.slice(i, j + 1); }
    }
    return '';
  };
  const olTime  = grab('Sidebar.html', '_olTime');
  const fmtWhen = grab('OrderCaseModal.html', 'fmtWhen');

  t('D1 ⚠ _olTime uses no local date getter at all',
    /\.get(Hours|Minutes|Date|Month|Day)\(\)/.test(olTime), false);
  t('D2 ⚠ fmtWhen keeps local getters ONLY inside its Intl-absent fallback',
    /catch[\s\S]*getHours/.test(fmtWhen) &&
    !/getHours[\s\S]*try\s*\{/.test(fmtWhen), true);
  // ⚠ _olTime DELEGATES; the zone lives in the helper it calls. A first-6000-chars
  //   window missed it entirely and accused correct code.
  const hqParts = grab('Sidebar.html', '_hqParts');
  t('D3 ⭐ both name a timezone explicitly, in whichever function formats',
    /timeZone:\s*HQ_TZ/.test(hqParts) && /timeZone:\s*OC_TZ/.test(fmtWhen), true);
  t('D3b …and _olTime really does route through that helper',
    /_hqParts\(/.test(olTime), true);
  t('D4 ⚠ and the zone is America/Chicago in both files',
    /HQ_TZ\s*=\s*'America\/Chicago'/.test(strip(read('Sidebar.html'))) &&
    /OC_TZ\s*=\s*'America\/Chicago'/.test(strip(read('OrderCaseModal.html'))), true);
}

console.log('\n' + '='.repeat(74));
if (failed) { console.log(`❌ test-timezone: ${passed} passed, ${failed} failed`); process.exit(1); }
console.log(`✅ test-timezone: ${passed} passed, 0 failed`);
