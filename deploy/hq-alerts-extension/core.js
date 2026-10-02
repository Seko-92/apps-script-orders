// ===========================================================================
// HQ ALERTS — the rules, in ONE place (2026-10-02).
//
// Used by BOTH:
//   · the /alerts tab        → served from /opt/hq-app as /alerts-core.js
//   · the Chrome extension   → bundled in this folder as core.js
// ⚠ It is ONE file deployed twice. Never fork it: two copies of a rule is how
//   A-9 sorted after A-50 in three files. Change it here, then scp + reload.
//
// Pure: no DOM, no chrome.*, no storage, no clock of its own. The caller passes
// the tick, its memory and the time, shows what comes back, and reports which
// holds were really shown (a hold that could not be shown is announced later).
// ===========================================================================
var HQAlertsCore = (function () {
  var TZ = 'America/Chicago';
  var BURST_AGE_MAX = 20;            // an order older than this (min) is not news
  var SEEN_KEEP_MS = 2 * 86400000;   // remember notified ids for 2 days

  function num(v) { var n = parseFloat(v); return isFinite(n) ? n : 0; }

  // ---------- working hours (Houston) ----------
  function houston(now) {
    var d = now ? new Date(now) : new Date();
    var h = parseInt(d.toLocaleString('en-US', { timeZone: TZ, hour: '2-digit', hour12: false }), 10) % 24;
    var wd = d.toLocaleDateString('en-US', { timeZone: TZ, weekday: 'short' });
    return { hour: h, weekday: wd !== 'Sat' && wd !== 'Sun' };
  }
  function inHours(now) { var t = houston(now); return t.weekday && t.hour >= 9 && t.hour < 17; }
  function clockNow(now) {
    return (now ? new Date(now) : new Date()).toLocaleTimeString('en-US', { timeZone: TZ, hour: 'numeric', minute: '2-digit' });
  }

  // ---------- shelf order: natural, like the pick list (A-9 before A-50) ----------
  function shelfKey(l) {
    var m = String(l || '').trim().toUpperCase().match(/^([A-Z]+)\s*-\s*(\d+)/);
    return m ? [0, m[1], parseInt(m[2], 10)] : [1, String(l || ''), 0];
  }
  function cmpShelf(a, b) {
    var x = shelfKey(a), y = shelfKey(b);
    for (var i = 0; i < 3; i++) { if (x[i] < y[i]) return -1; if (x[i] > y[i]) return 1; }
    return 0;
  }
  function shelvesOf(rows, n) {
    var seen = {}, out = [];
    rows.forEach(function (r) { var l = String(r.location || '').trim(); if (l && l !== 'NOT FOUND' && !seen[l]) { seen[l] = 1; out.push(l); } });
    return out.sort(cmpShelf).slice(0, n);
  }

  // ---------- memory ----------
  function freshSeen(s) {
    s = s && typeof s === 'object' ? s : {};
    return { orders: s.orders || {}, holds: s.holds || {} };
  }
  function prune(seen, now) {
    ['orders', 'holds'].forEach(function (k) {
      for (var id in seen[k]) if (now - seen[k][id] > SEEN_KEEP_MS) delete seen[k][id];
    });
    return seen;
  }

  // ---------- THE decision ----------
  // tick   : the published board tick
  // seen   : { orders:{id:ms}, holds:{id:ms} } — MUTATED (orders marked here)
  // primed : false on the first tick after a (re)start → only remember, announce nothing
  // now    : ms
  // Returns { notes:[…], newCount, openHolds:[…], hours }
  //   note = { kind:'hold'|'order', id, title, body, sticky, logClass, logText, holdId? }
  //   ⚠ the caller marks seen.holds[note.holdId] ONLY if the note was really shown.
  function decide(tick, seen, primed, now) {
    var t = tick || {}, rows = t.openOrders || [], held = t.held || [],
        age = (t.cockpit && t.cockpit.orderAgeMin) || {}, customers = t.customers || {};
    var notes = [], hours = inHours(now);

    // 1 · HOLDS — any time. An unanswered hold not yet announced from here.
    held.forEach(function (h) {
      var id = String(h.orderId || '');
      if (!id || h.acked || seen.holds[id]) return;
      var it = (h.items || [])[0] || {};
      var why = h.shipped ? 'label already bought' : (h.urgent ? 'being picked' : 'not started');
      notes.push({
        kind: 'hold', id: 'hq-hold-' + id, holdId: id, sticky: true,
        title: 'Hold · ' + why,
        body: id + ' · ' + (it.loc || 'no shelf') + ' · ' + (it.sku || '') + ' ×' + (it.qty || 1) +
              ((h.items || []).length > 1 ? ' (+' + (h.items.length - 1) + ' more)' : '') +
              '\n' + (h.shipped ? 'Set the box aside. ' : '') + 'Acknowledge on the board.',
        logClass: 'r', logText: 'Hold · ' + why + ' · ' + id
      });
    });

    // 2 · NEW ORDERS — PENDING order ids not seen before. The first tick after a
    // start only REMEMBERS what is already there: history is not news.
    var byOrder = {}, order = [];
    rows.forEach(function (r) {
      var id = String(r.orderId || '').trim();
      if (!id) return;
      if (!byOrder[id]) { byOrder[id] = { id: id, ch: r.channel || 'EBAY', rows: [], pending: false }; order.push(byOrder[id]); }
      byOrder[id].rows.push(r);
      if (String(r.status).toUpperCase() === 'PENDING') byOrder[id].pending = true;
    });
    var fresh = order.filter(function (o) {
      if (seen.orders[o.id]) return false;
      seen.orders[o.id] = now;
      if (!primed || !o.pending) return false;
      // ⚠ the board's list is capped, so an OLD order can scroll into view later —
      // that is not an arrival. Same 20-minute guard the sidebar's order voices use.
      if (age[o.id] !== undefined && num(age[o.id]) > BURST_AGE_MAX) return false;
      return true;
    });
    var newCount = 0;
    if (fresh.length && hours) {
      newCount = fresh.length;
      var ebay = fresh.filter(function (o) { return o.ch !== 'DIRECT' && o.ch !== 'AMAZON'; });
      // eBay — ONE notification per pull, however many landed. Closes by itself.
      if (ebay.length === 1) {
        var r0 = ebay[0].rows[0];
        notes.push({ kind: 'order', id: 'hq-ebay-' + now, sticky: false, title: 'New eBay order',
          body: (r0.location || 'no shelf') + ' · ' + r0.sku + ' ×' + (r0.qty || 1) +
                (ebay[0].rows.length > 1 ? ' (+' + (ebay[0].rows.length - 1) + ' lines)' : '') + '\n' + ebay[0].id,
          logClass: 'y', logText: 'New eBay order · ' + (r0.location || ebay[0].id) });
      } else if (ebay.length > 1) {
        var all = []; ebay.forEach(function (o) { all = all.concat(o.rows); });
        notes.push({ kind: 'order', id: 'hq-ebay-' + now, sticky: false, title: ebay.length + ' new eBay orders',
          body: 'First walk ' + shelvesOf(all, 3).join(' · ') + '\n' + all.length + ' lines',
          logClass: 'y', logText: ebay.length + ' new eBay orders' });
      }
      // Direct + Amazon — one each, stays until clicked
      fresh.filter(function (o) { return o.ch === 'DIRECT' || o.ch === 'AMAZON'; }).forEach(function (o) {
        var r1 = o.rows[0], isA = o.ch === 'AMAZON';
        var first = shelvesOf(o.rows, 1)[0];
        var line1 = (first ? first + ' · ' : '') + r1.sku + ' ×' + (r1.qty || 1) +
                    (o.rows.length > 1 ? ' (+' + (o.rows.length - 1) + ' lines)' : '');
        // second line: who it is for (Direct) or the ship-by deadline (Amazon)
        var line2 = isA ? ((String(r1.note || '').match(/ship by\s+([^·]+)/i) || [])[1] || '').trim()
                        : (customers[o.id] || '');
        if (isA && line2) line2 = 'ship by ' + line2;
        notes.push({ kind: 'order', id: 'hq-' + o.id, sticky: true,
          title: (isA ? 'New Amazon order · ' : 'New Direct order · ') + o.id,
          body: line1 + (line2 ? '\n' + line2 : ''),
          logClass: 'y', logText: (isA ? 'New Amazon order · ' : 'New Direct order · ') + o.id });
      });
    }
    prune(seen, now);
    return { notes: notes, newCount: newCount, hours: hours,
             openHolds: held.filter(function (h) { return !h.acked; }) };
  }

  var api = { decide: decide, freshSeen: freshSeen, inHours: inHours, clockNow: clockNow,
              shelvesOf: shelvesOf, cmpShelf: cmpShelf, TZ: TZ, BURST_AGE_MAX: BURST_AGE_MAX };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  return api;
})();
