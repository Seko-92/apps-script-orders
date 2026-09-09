// =====================================================================================
// test-spark-pulse.js — freshness comes from __SparkData!A4, not from parsing row 1.
//
// Two things are under test, and the SECOND is the one that matters:
//   1. the decoupling — E1 can be moved/cleared without killing the heartbeat
//   2. ⭐ the CORRECTNESS FIX — the old E1 regex read "h:mm AM/PM", which carries NO
//      DATE, and diffed minutes-of-day with a midnight wrap. So it capped at 1439 and
//      wrapped: a pipeline dead EXACTLY 24h reported 0 minutes ("🟢 ALIVE"). The longer
//      the outage, the healthier it looked.
//
// Loads the REAL ActivityLog.js + ApiMonitor.js in a VM, so it cannot drift.
//
// Prove against HEAD:
//   mkdir -p /tmp/headsp && for f in ActivityLog.js ApiMonitor.js; do \
//     git show HEAD:$f > /tmp/headsp/$f; done && SRC=/tmp/headsp node test-spark-pulse.js
// =====================================================================================
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');

let pass = 0, fail = 0;
const ok = (n, c, x) => { c ? pass++ : (fail++, console.log('  ✗ ' + n + (x !== undefined ? '  → ' + JSON.stringify(x) : ''))); };

const two = n => (n < 10 ? '0' : '') + n;
function fmt(d) {                       // "h:mm a" in UTC, deterministic
  let h = d.getUTCHours(), m = d.getUTCMinutes(), mer = h >= 12 ? 'PM' : 'AM';
  h = h % 12; if (h === 0) h = 12;
  return h + ':' + two(m) + ' ' + mer;
}

// ── the OLD algorithm, reproduced verbatim from the pre-fix ActivityLog.js ───────────
// It only ever saw "h:mm AM/PM" — no date — which is the whole bug.
function oldAlgorithm(syncAt, now) {
  const m = fmt(syncAt).match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!m) return null;
  let hour = parseInt(m[1], 10); const min = parseInt(m[2], 10);
  const mer = m[3].toUpperCase();
  if (mer === 'PM' && hour < 12) hour += 12;
  if (mer === 'AM' && hour === 12) hour = 0;
  const nowMin = now.getUTCHours() * 60 + now.getUTCMinutes();
  let diff = nowMin - (hour * 60 + min);
  if (diff < 0) diff += 1440;
  return diff;
}

function makeCtx(spark) {               // spark === null → sheet missing
  const calls = { ranges: [] };
  const sheet = spark === null ? null : {
    getRange: (a1) => { calls.ranges.push(a1); return { getValues: () => spark }; }
  };
  const ctx = {
    console: { log: () => {}, error: () => {} },
    JSON, String, Number, Date, Math, parseInt, parseFloat, isNaN, Array, Object, RegExp,
    SPREADSHEET_ID: 'x', MAIN_SHEET_NAME: 'All orders',
    ACTIVITY_LOG: { sheetName: 'Activity Log' },
    Schema: { cellSyncTime: 'E1' },
    Utilities: { formatDate: (d) => fmt(d) },
    SpreadsheetApp: { openById: () => ({ getSheetByName: (n) => (n === '__SparkData' ? sheet : null) }) }
  };
  ctx.calls = calls;
  vm.createContext(ctx);
  for (const f of ['ActivityLog.js', 'ApiMonitor.js']) {
    try { vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f }); }
    catch (e) { console.log('  ! could not load ' + f + ': ' + e.message); }
  }
  return ctx;
}

// A3..A13 = 11 rows. A3=idx0 (timestamp), A4=idx1 (minutes), A12=idx9 (label)
const rows = (at, mins, ago) => {
  const r = Array.from({ length: 11 }, () => ['']);
  r[0] = [at]; r[1] = [mins]; r[9] = [ago]; r[10] = [false];
  return r;
};

console.log('\n── A · _sparkPulse derives minutes from A3, the log timestamp');
{
  const at = new Date(Date.now() - 6 * 60000);
  const c = makeCtx(rows(at, 6.4, '6m'));
  const soft = (fn) => { try { return fn(); } catch (e) { return { _err: e.message }; } };
  const p = soft(() => c._sparkPulse(c.SpreadsheetApp.openById()));
  ok('A1 _sparkPulse exists', typeof c._sparkPulse === 'function');
  ok('A2 minutes computed from A3, floored', p && p.minutes === 6, p);
  ok('A3 timestamp returned', p && p.at && p.at.getTime() === at.getTime());
  ok('A4 human label carried', p && p.ago === '6m', p && p.ago);
  ok('A5 ONE round trip (A3:A13)', c.calls.ranges.length === 1 && c.calls.ranges[0] === 'A3:A13', c.calls.ranges);
}

console.log('── B · unreadable is null, never a reassuring zero');
{
  const c = makeCtx(rows(0, -1, ''));            // A4 = -1 sentinel
  const p = (() => { try { return c._sparkPulse(c.SpreadsheetApp.openById()); } catch (e) { return null; } })();
  ok('B1 minutes null on the -1 sentinel', p && p.minutes === null, p);
  ok('B2 does NOT report 0', !(p && p.minutes === 0));
}

console.log('── C · a missing __SparkData degrades, never throws');
{
  const c = makeCtx(null);
  let threw = false, p = null;
  try { p = c._sparkPulse(c.SpreadsheetApp.openById()); } catch (e) { threw = true; }
  ok('C1 no throw', !threw);
  ok('C2 minutes null', p && p.minutes === null, p);
}

console.log('── D · getLastSyncFromSheet no longer touches E1');
{
  const at = new Date(Date.UTC(2026, 8, 4, 16, 45));
  const c = makeCtx(rows(at, 6.4, '6m'));
  const s = (() => { try { return c.getLastSyncFromSheet(); } catch (e) { return '!' + e.message; } })();
  ok('D1 renders the clock', /4:45 PM/.test(s), s);
  ok('D2 renders the duration', /6m ago/.test(s), s);
  ok('D3 read __SparkData, not E1', c.calls.ranges.indexOf('E1') === -1, c.calls.ranges);
  const c2 = makeCtx(rows(0, -1, ''));
  const s2 = (() => { try { return c2.getLastSyncFromSheet(); } catch (e) { return '!' + e.message; } })();
  ok('D4 offline line when unreadable', /OFFLINE/.test(s2), s2);
}

console.log('── E · ⭐ THE 24-HOUR WRAP — the bug the old parse had');
{
  const now = new Date();
  const dayOld = new Date(now.getTime() - 1440 * 60000);   // EXACTLY 24h stale
  const c = makeCtx(rows(dayOld, 1440, '1d'));
  const p = (() => { try { return c._sparkPulse(c.SpreadsheetApp.openById()); } catch (e) { return null; } })();
  const old = oldAlgorithm(dayOld, now);
  ok('E1 old parse reported 0 min for a 24h-dead pipeline', old === 0, old);
  ok('E2 new reports the true 1440', p && p.minutes === 1440, p);
  ok('E3 new is past the STALE line (>60), old was not', p && p.minutes > 60 && old <= 60);

  const threeDay = new Date(Date.now() - 4320 * 60000);
  const c3 = makeCtx(rows(threeDay, 4320, '3d'));
  const p3 = (() => { try { return c3._sparkPulse(c3.SpreadsheetApp.openById()); } catch (e) { return null; } })();
  ok('E4 a 3-day outage reads 4320, not a wrapped remainder', p3 && p3.minutes === 4320, p3);
}

console.log('── F · regression net — on a FRESH sync, old and new agree');
{
  const now = new Date();
  const fresh = new Date(now.getTime() - 6 * 60000);
  const c = makeCtx(rows(fresh, 6, '6m'));
  const p = (() => { try { return c._sparkPulse(c.SpreadsheetApp.openById()); } catch (e) { return null; } })();
  ok('F1 both say 6 — the change is surgical', p && p.minutes === 6 && oldAlgorithm(fresh, now) === 6, p);
}

console.log('── G · ⚠ the two SILENT ways reading A4 fails — Gotcha #16 and a frozen NOW()');
{
  // A4 arrives as a DATE because the cell inherited a date format. parseFloat(Date)
  // is NaN, which is exactly how this class has bitten three times before.
  const at = new Date(Date.now() - 12 * 60000);
  const asDate = new Date(1899, 11, 30 + 12);        // serial 12, wearing a date format
  const c = makeCtx(rows(at, asDate, '12m'));
  const p = (() => { try { return c._sparkPulse(c.SpreadsheetApp.openById()); } catch (e) { return null; } })();
  ok('G1 a date-formatted A4 does NOT blank the heartbeat', p && p.minutes === 12, p);

  // A4 frozen at a stale value (NOW() not recalculating) — A3 still tells the truth.
  const c2 = makeCtx(rows(new Date(Date.now() - 90 * 60000), 3, '3m'));
  const p2 = (() => { try { return c2._sparkPulse(c2.SpreadsheetApp.openById()); } catch (e) { return null; } })();
  ok('G2 a frozen A4 is ignored — A3 is the source of truth', p2 && p2.minutes === 90, p2);
  ok('G3 ⚠ and it reports STALE (>60), not the frozen 3', p2 && p2.minutes > 60);

  // A3 unusable → fall back to A4, tolerating the date form.
  const c3 = makeCtx(rows(0, asDate, '12m'));
  const p3 = (() => { try { return c3._sparkPulse(c3.SpreadsheetApp.openById()); } catch (e) { return null; } })();
  ok('G4 A3 unusable → A4 fallback still recovers 12', p3 && p3.minutes === 12, p3);
}

console.log('\n' + (fail ? '✗ ' : '✓ ') + pass + ' passed, ' + fail + ' failed\n');
process.exit(fail ? 1 : 0);
