// test-floor-pulse.js — the FLOOR PULSE rule lives in THREE files (board, wall, sidebar);
// they must be byte-identical, and the rule must say the right thing. (2026-10-02)
//   node test-floor-pulse.js      SRC=/other/dir overrides where the files are read
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');
let pass = 0, fail = 0;
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want);
  if (g === w) pass++; else { fail++; console.log('✗ ' + name + '\n    got  ' + g + '\n    want ' + w); }
}
const re = /\/\* ── FLOOR PULSE · BEGIN[\s\S]*?\/\* ── FLOOR PULSE · END ── \*\//;
const blocks = {};
for (const f of ['FloorBoard.html', 'wall.html', 'Sidebar.html']) {
  const m = fs.readFileSync(path.join(SRC, f), 'utf8').match(re);
  blocks[f] = m ? m[0].split('\n').map(l => l.trim()).join('\n') : null;   // indentation may differ
}
eq('found in all three files', Object.values(blocks).every(Boolean), true);
eq('board == wall', blocks['FloorBoard.html'] === blocks['wall.html'], true);
eq('board == sidebar', blocks['FloorBoard.html'] === blocks['Sidebar.html'], true);

const ctx = {}; vm.createContext(ctx); vm.runInContext(blocks['FloorBoard.html'] || '', ctx);
const fp = ctx.floorPulse, NOW = 1e12, ago = m => NOW - m * 60000;
const dev = (role, ageSec) => ({ id: role + ageSec, name: 'x', role, ageSec });

let p = fp({ floorLast: { at: ago(4) }, floorToday: 23 }, [dev('floor', 10), dev('office', 30)], NOW);
eq('a tap 4 min ago → ACTIVE, 2 screens', [p.state, p.on, p.text], ['active', 2, 'ACTIVE · last tap 4 min ago · 2 screens on']);
eq('today count carried', p.today, 23);
p = fp({ floorLast: { at: ago(70) } }, [dev('floor', 10)], NOW);
eq('screen on, no tap for 70 min → QUIET', [p.state, p.text], ['quiet', 'QUIET · last tap 1h 10m ago · 1 screen on']);
p = fp({ floorLast: { at: ago(70) } }, [dev('floor', 600)], NOW);
eq('screen silent > 2 min → OFF', [p.state, p.on], ['off', 0]);
p = fp({ floorLast: { at: ago(3) } }, [dev('remote', 5), dev('wall', 5)], NOW);
eq('the owner\'s device and the wall never count as the warehouse', p.on, 0);
eq('...but a fresh tap still reads ACTIVE', p.state, 'active');
p = fp({ floorLast: { at: ago(16) } }, null, NOW);
eq('no screen list (live fallback) → no "0 on" claim', [p.state, p.on, /screens? on/.test(p.text)], ['quiet', null, false]);
p = fp({}, [], NOW);
eq('nothing at all → OFF, no tap today', p.text, 'OFF · no tap today · 0 screens on');
p = fp({ floorLast: { at: ago(0.2) } }, [dev('floor', 1)], NOW);
eq('just now', /tapped just now/.test(p.text), true);
p = fp({ floorLast: { at: ago(15) } }, [dev('floor', 1)], NOW);
eq('15 min is still ACTIVE (inclusive)', p.state, 'active');

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
