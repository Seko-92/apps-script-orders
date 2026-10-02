// HQ Board proxy — serve the PUBLISHED tick, fall back to computing it live.
//
// THE CHANGE (2026-08-07): the board used to make Apps Script rebuild the tick
// on every cache miss, on every device. Apps Script runs ONE execution at a
// time, so cost scaled with VIEWERS — two devices already queued behind each
// other. Now Apps Script writes the finished tick to a cell whenever something
// actually changes, and this node just reads it.
//
//   cost before:  viewers x poll rate
//   cost after:   order activity
//
// THREE TIERS, cheapest first:
//   1. n8n static cache  (15s)  — repeat polls inside the window cost nothing
//   2. the PUBLISHED cell       — one small Sheets read, no Apps Script at all
//   3. Apps Script live         — the original path, unchanged
//
// Tier 3 is the SAFETY NET, not dead code. It runs when the payload is missing
// (never published), stale (publish trigger died), or malformed. The board
// therefore cannot go dark because of this change, and reverting is pointing
// the flow back at it.
//
// WRITES NEVER TOUCH TIERS 1-2. boardStatus is the ✓ Pick button; boardPart /
// boardPartLite / boardOrder are drawer lookups that must be current. Only
// boardTick is ever served from a cache or a published copy.

const sd = $getWorkflowStaticData('global');

// The webhook body, reached explicitly: $input here is the Sheets read node.
const body = ($('1. Board Request').first().json.body) || {};
const action = String(body.action || '');

const EXEC_URL = '__APPS_SCRIPT_EXEC_URL__';
const TOKEN    = '__APP_SECRET_TOKEN__';

const CACHE_MS       = 3000;              // tier 1 window (15s → 3s 2026-10-02: boards poll every 5s)
const MAX_PUB_AGE_MS = 10 * 60 * 1000;    // tier 2 trust window — 2x the 5-min
                                          // publish trigger, so one missed run
                                          // is tolerated and a dead trigger is
                                          // not.

const now = Date.now();
const isTick = (action === 'boardTick');

// ---- SCREENS (2026-10-02) ---------------------------------------------------
// Every board/wall sends {device:{id,name,role}}. Remember when each was last
// heard from — here, because polls are answered from this node's cache and never
// reach Apps Script. {action:"boardDevices"} reads the list (Apps Script's
// quiet-floor-screen check asks it); every tick carries it as _devices so the
// board can show "screens online". Requests without a device (the Chrome alerts
// extension, Apps Script's own probes) are not recorded.
const DEVICE_TTL_MS = 24 * 60 * 60 * 1000;
const ROLES = { floor: 1, wall: 1, office: 1, remote: 1, unset: 1 };   // unknown -> 'unset', never 'floor'
const dev = body.device;
if (dev && typeof dev === 'object' && dev.id) {
  sd.devices = sd.devices || {};
  sd.devices[String(dev.id).slice(0, 40)] = {
    name: String(dev.name || 'Screen').replace(/[\u0000-\u001f]/g, ' ').slice(0, 40),
    role: ROLES[dev.role] ? dev.role : 'unset',
    at: now
  };
}
function deviceList() {
  const out = [];
  const all = sd.devices || {};
  for (const id of Object.keys(all)) {
    if (now - all[id].at > DEVICE_TTL_MS) { delete all[id]; continue; }
    out.push({ id, name: all[id].name, role: all[id].role,
               ageSec: Math.round((now - all[id].at) / 1000) });
  }
  return out.sort((a, b) => a.ageSec - b.ageSec).slice(0, 30);
}
if (action === 'boardDevices') {
  return [{ json: { ok: true, devices: deviceList() } }];
}

// ---- tier 1: n8n static cache ---------------------------------------------
if (isTick && sd.tick && sd.tickAt && (now - sd.tickAt) < CACHE_MS) {
  const hit = JSON.parse(sd.tick);
  hit._cached = true;
  hit._ageMs  = now - sd.tickAt;
  hit._devices = deviceList();
  return [{ json: hit }];
}

// ---- tier 2: the published cell -------------------------------------------
// (replaces the existing `if (isTick) { try { … } catch (e) {} }` block)
let tier2 = 'skipped: not a tick';
if (isTick) {
  tier2 = 'unknown';
  try {
    const sheetOut = $input.first().json || {};
    // A failed read still produces an item here — that is the whole trap.
    if (sheetOut.error) {
      tier2 = 'sheets read FAILED: ' +
              String(JSON.stringify(sheetOut.error)).slice(0, 140);
    } else {
      // SHAPE-AGNOSTIC (2026-08-13). The Sheets REST API returns
      // {values:[["…json…"]]}; the NATIVE Google Sheets node returns a row
      // object instead. Accepting both means node 2 can be swapped from an
      // HTTP Request to a native Sheets node — which sidesteps n8n's generic
      // HTTP domain allowlist entirely — without touching this code again.
      let raw = sheetOut.values && sheetOut.values[0] && sheetOut.values[0][0];
      if (!raw) {
        raw = Object.values(sheetOut).find(
          v => typeof v === 'string' && v.indexOf('"cockpit"') !== -1);
      }
      if (!raw) {
        tier2 = 'cell empty — nothing published, or the range is wrong';
      } else {
        const tick = JSON.parse(raw);
        const pubAt = tick._publishedAt ? Date.parse(tick._publishedAt) : 0;
        if (!pubAt) {
          tier2 = 'payload has no _publishedAt';
        } else if (!tick.cockpit) {
          tier2 = 'payload parsed but has no cockpit';
        } else if ((now - pubAt) >= MAX_PUB_AGE_MS) {
          tier2 = 'stale by ' + Math.round((now - pubAt) / 1000) + 's ' +
                  '(trust window ' + Math.round(MAX_PUB_AGE_MS / 1000) + 's)';
        } else {
          tick._published      = true;
          tick._publishedAgeMs = now - pubAt;
          sd.tick   = JSON.stringify(tick);
          sd.tickAt = now;
          tick._devices = deviceList();     // after caching — never stored in sd.tick
          return [{ json: tick }];          // ← the fast path
        }
      }
    }
  } catch (e) {
    tier2 = 'threw: ' + (e.message || String(e));
  }
}

// ---- tier 3: Apps Script, exactly as before --------------------------------
let res;
try {
  res = await this.helpers.httpRequest({
    method:  'POST',
    url:     EXEC_URL,
    body:    Object.assign({}, body, { token: TOKEN }),
    json:    true,
    timeout: 90000
  });
} catch (err) {
  // Never throw: a thrown Code node returns no body and the board sees a dead
  // socket instead of something it can report.
  return [{ json: { error: 'proxy: ' + (err.message || String(err)), _tier2: tier2 } }];
}

if (isTick && res && res.cockpit) {
  sd.tick   = JSON.stringify(res);
  sd.tickAt = now;
  res._liveFallback = true;   // visible in the response when tier 2 missed
  res._tier2 = tier2;
  res._devices = deviceList();
}
return [{ json: res }];