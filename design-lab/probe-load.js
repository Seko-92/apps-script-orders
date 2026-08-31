// ============================================================================
// PEAK-HOUR HEALTH PROBE — is the floor about to lose taps?
//
// WHY LATENCY AND NOT QUOTA: Apps Script publishes its limits (~90 min/day of
// runtime on a consumer account) but exposes NO API and no dashboard for what
// you have USED. There is nothing to read. And quota is the wrong number
// anyway — the script gets slow from CONTENTION long before it runs out of
// budget, and contention is what the picker actually feels. So measure the
// symptom: how long a write-path round trip takes, and how often it fails.
//
// WHAT IT MEASURES
//   `boardRadio` — a real action on the SAME path a ✓ Pick takes (tablet →
//   Caddy → n8n → Apps Script /exec) that does essentially no work and changes
//   NOTHING. So the number is pure overhead: transport + whatever the script is
//   busy with. A board TICK would not do — that is served from the published
//   cell and never reaches Apps Script.
//
// ⚠ THE PROBE IS ITSELF LOAD. Each call costs an Apps Script execution, so do
//   not leave it running all day: at one call per 20s that is ~2 minutes of the
//   daily budget per hour of probing. Run it for a window, read it, stop.
//
// THE VERDICT LINE is against the board's own 25s write bound — the point at
// which a picker's tap is abandoned and they are asked to tap again.
//
// Usage:
//   node probe-load.js               # 10 minutes, default
//   node probe-load.js 20            # 20 minutes
//   MINUTES=5 EVERY=10 node probe-load.js
// ============================================================================
'use strict';

const URL     = process.env.HQ_URL || 'https://hq.yassinqurabi.com/api/board';
const MINUTES = Number(process.argv[2] || process.env.MINUTES || 10);
const EVERY   = Number(process.env.EVERY || 20);          // seconds between calls
const BOUND   = 25;                                        // the board's write bound

const samples = [];
const started = Date.now();

function pct(sorted, p) {
  if (!sorted.length) return 0;
  return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * p))];
}

async function one() {
  const t0 = Date.now();
  let body = '', ok = false, kind = 'ok';
  try {
    const ctrl = new AbortController();
    const kill = setTimeout(() => ctrl.abort(), 45000);
    const r = await fetch(URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'boardRadio', stationId: '' }),
      signal: ctrl.signal
    });
    clearTimeout(kill);
    body = await r.text();
    ok = body.startsWith('{"ok":true');
    if (!ok) {
      // Name the failure the way the board would see it.
      if (body.includes('Server Busy'))      kind = 'LOCK TIMEOUT (Server Busy)';
      else if (body.includes('"headers"'))   kind = 'MALFORMED (proxy echoed the request)';
      else if (body.includes('proxy:'))      kind = 'PROXY ERROR';
      else                                   kind = 'UNEXPECTED BODY';
    }
  } catch (e) {
    kind = /abort/i.test(String(e)) ? 'NO ANSWER (45s)' : 'TRANSPORT (' + e.message + ')';
  }
  const secs = (Date.now() - t0) / 1000;
  samples.push({ secs, ok, kind, at: new Date().toISOString().slice(11, 19) });
  return samples[samples.length - 1];
}

function report() {
  const n = samples.length;
  if (!n) return console.log('no samples');
  const sorted = samples.map(s => s.secs).sort((a, b) => a - b);
  const good   = samples.filter(s => s.ok);
  const overB  = samples.filter(s => s.secs > BOUND);
  const mins   = ((Date.now() - started) / 60000).toFixed(1);

  console.log(`\n${'='.repeat(62)}`);
  console.log(`PEAK-HOUR PROBE — ${n} samples over ${mins} min`);
  console.log(`${'='.repeat(62)}`);
  console.log(`  succeeded        ${good.length}/${n}   (${(good.length / n * 100).toFixed(1)}%)`);
  console.log(`  median           ${pct(sorted, 0.50).toFixed(2)}s`);
  console.log(`  p95              ${pct(sorted, 0.95).toFixed(2)}s`);
  console.log(`  worst            ${sorted[sorted.length - 1].toFixed(2)}s`);
  console.log(`  past the ${BOUND}s bound  ${overB.length}   <- each one is a tap the picker loses`);

  const bad = samples.filter(s => !s.ok);
  if (bad.length) {
    const by = {};
    bad.forEach(b => { by[b.kind] = (by[b.kind] || 0) + 1; });
    console.log(`\n  failures by kind:`);
    Object.keys(by).forEach(k => console.log(`    ${String(by[k]).padStart(3)}  ${k}`));
    console.log(`\n  when:`);
    bad.forEach(b => console.log(`    ${b.at}  ${b.secs.toFixed(1)}s  ${b.kind}`));
  }

  // The verdict, in the terms that matter on the floor.
  const p95 = pct(sorted, 0.95);
  console.log('');
  if (good.length === n && p95 < 6) {
    console.log('  ✅ HEALTHY — comfortable headroom under the 25s bound.');
  } else if (overB.length === 0 && p95 < 15) {
    console.log('  🟡 WORKING, BUT LOADED — nothing lost, headroom shrinking.');
    console.log('     Worth checking Executions for what is running long.');
  } else {
    console.log('  🔴 THE FLOOR IS LOSING TAPS. Open Executions NOW, sort by');
    console.log('     duration, and look at the timestamps listed above —');
    console.log('     whatever is sitting there for 15-30s is the cause.');
  }
  console.log('');
}

(async () => {
  console.log(`Probing ${URL}`);
  console.log(`every ${EVERY}s for ${MINUTES} min — Ctrl-C any time for the report so far\n`);
  process.on('SIGINT', () => { report(); process.exit(0); });

  const until = Date.now() + MINUTES * 60000;
  while (Date.now() < until) {
    const s = await one();
    const mark = s.ok ? (s.secs > 6 ? '🟡' : '  ') : '🔴';
    console.log(`${mark} ${s.at}  ${s.secs.toFixed(2)}s  ${s.ok ? '' : s.kind}`);
    if (Date.now() < until) await new Promise(r => setTimeout(r, EVERY * 1000));
  }
  report();
})();
