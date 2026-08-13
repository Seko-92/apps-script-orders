// ============================================================================
// PATCH for n8n → "HQ Board Proxy" → node "3. Serve or Forward"
//
// WHY: tier 2 (the published cell) falls through SILENTLY. Node 2 is set to
// onError:continueRegularOutput + alwaysOutputData, so a failed Sheets read
// arrives as an empty item, `raw` is undefined, and the proxy quietly rebuilds
// the whole tick in Apps Script — 11-17s instead of ~1s — while reporting
// nothing anywhere. Measured 2026-08-13: EVERY uncached tick was doing this.
//
// WHAT CHANGES: nothing about the behaviour. The tiers, the trust window and
// the fallback are identical. The response simply carries `_tier2` saying WHY
// the published cell was not used, so `curl … -d '{"action":"boardTick"}'`
// answers the question instead of starting an investigation.
//
// HOW TO APPLY: open the workflow in n8n, click node "3. Serve or Forward",
// and replace ONLY the tier-2 block plus the final return. The two constants
// at the top (EXEC_URL / TOKEN) already hold real values on the live instance —
// DO NOT paste those over from the repo copy, which carries placeholders.
// ============================================================================

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
          return [{ json: tick }];          // ← the fast path, unchanged
        }
      }
    }
  } catch (e) {
    tier2 = 'threw: ' + (e.message || String(e));
  }
}

// ---- tier 3: Apps Script, exactly as before --------------------------------
// (unchanged — shown only for placement)
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
  return [{ json: { error: 'proxy: ' + (err.message || String(err)), _tier2: tier2 } }];
}

if (isTick && res && res.cockpit) {
  sd.tick   = JSON.stringify(res);
  sd.tickAt = now;
  res._liveFallback = true;
  res._tier2        = tier2;    // ← the new line that ends the guessing
}
return [{ json: res }];
