#!/usr/bin/env python3
# ============================================================================
# THE PORT — FloorBoard.html v2.1 → v3.0 (tool-first), in-place surgery.
# All edits in memory, written ONCE at the end (the assert-then-write lesson:
# a failed assert here loses nothing because nothing has been written yet).
# ============================================================================
import re, sys

ROOT = '/home/yassin/Desktop/Projects/Projects/Excel Code'
s = open(f'{ROOT}/FloorBoard.html').read()
orig_len = len(s)

def cut(start_anchor, end_anchor, replacement, label):
    """Replace [start_anchor .. end_anchor) with replacement + end_anchor."""
    global s
    i = s.find(start_anchor)
    assert i != -1, f'{label}: start anchor not found'
    assert s.find(start_anchor, i + 1) == -1, f'{label}: start anchor not unique'
    j = s.find(end_anchor, i)
    assert j != -1, f'{label}: end anchor not found after start'
    s = s[:i] + replacement + s[j:]
    print(f'  ✓ {label}')

def swap(old, new, label):
    global s
    assert s.count(old) == 1, f'{label}: anchor count {s.count(old)} != 1'
    s = s.replace(old, new)
    print(f'  ✓ {label}')

# ── 1. STYLE ────────────────────────────────────────────────────────────────
new_css = open(f'{ROOT}/design-lab/port/new-style.css').read()
carried = open(f'{ROOT}/design-lab/carried-css.txt').read()
i = s.find('<style>'); j = s.find('</style>')
assert i != -1 and j != -1
s = s[:i] + '<style>\n' + new_css + '\n' + carried + '\n  ' + s[j:]
print('  ✓ stylesheet replaced')

# ── 2. BODY ─────────────────────────────────────────────────────────────────
new_body = open(f'{ROOT}/design-lab/port/new-body.html').read()
i = s.find('<body>'); j = s.find('  <script>')
assert i != -1 and j != -1 and i < j
s = s[:i] + new_body + '\n\n' + s[j:]
print('  ✓ body replaced')

# ── 3. constants block ──────────────────────────────────────────────────────
cut("    // notify-on-change state",
    "    // ---------- boot ----------",
    """    // notify-on-change state
    var prevOpen  = null;    // ebayGrab + directGrab — caught-up chime
    var prevPaid  = null;    // paidShipping count — paid chime
    var seenKeys  = {};      // feed dedupe: have we already rendered this event?
    var firstTick = true;    // suppress chimes + animation on initial paint
    var failCount = 0;
    // ON by default. A board whose job is to reach a picker who has walked
    // away cannot do it silently; the toggle lives in the ⋯ menu.
    var soundOn   = true;
    var audioCtx  = null;
    var eggClicks = 0, eggTimer = null;
    var eggArmed = false, eggArmTimer = null;   // second run of 5 taps -> arcade
    var prayerList = [];   // [{name, hhmm, mins}] for today, Houston time

""", 'constants slimmed')

# ── 4. boot + weather deletion ──────────────────────────────────────────────
cut("    // ---------- boot ----------",
    "    function tickClock() {",
    """    // ---------- boot ----------
    document.addEventListener('DOMContentLoaded', function () {
      var _snd = localStorage.getItem('floorSound');
      soundOn = (_snd === null) ? true : (_snd === '1');
      paintSoundToggle();
      document.getElementById('soundToggle').addEventListener('click', toggleSound);
      document.getElementById('fsToggle').addEventListener('click', toggleFullscreen);
      document.getElementById('beaconStrip').addEventListener('click', beaconClear);
      // Someone CAN scroll on the tablet — drop the fade once they reach the end.
      ['pickListEbay', 'pickListDirect', 'pickSplit', 'feedList'].forEach(function (id) {
        var el = document.getElementById(id);
        if (el) el.addEventListener('scroll', function () { markOverflow(el); });
      });
      document.getElementById('chanTabs').addEventListener('click', function (ev) {
        var b = ev.target && ev.target.closest ? ev.target.closest('.ct') : null;
        if (b) setChannel(b.getAttribute('data-ch'));
      });
      document.getElementById('otherNudge').addEventListener('click', function () {
        setChannel(activeChannel === 'DIRECT' ? 'EBAY' : 'DIRECT');
      });
      document.getElementById('sortSw').addEventListener('click', function (ev) {
        var o = ev.target && ev.target.closest ? ev.target.closest('.sortsw-opt') : null;
        if (o) setPickMode(o.getAttribute('data-mode'));
      });
      document.addEventListener('fullscreenchange', paintFullscreen);
      document.addEventListener('webkitfullscreenchange', paintFullscreen);
      paintFullscreen();
      var hqMark = document.querySelector('.hdr-mono');
      if (hqMark) hqMark.addEventListener('click', onHqTap);
      requestWakeLock();
      document.addEventListener('visibilitychange', function () {
        if (document.visibilityState === 'visible') requestWakeLock();
      });
      // ONE delegated handler on the split, so it covers both columns and
      // survives every rebuild of their contents.
      document.getElementById('pickSplit').addEventListener('click', function (e) {
        if (!e.target || !e.target.closest) return;
        var b = e.target.closest('.pick-do');
        if (b) { markPicked(b.getAttribute('data-order'), b.getAttribute('data-sku'), b); return; }
        var c = e.target.closest('.pc-btn');
        if (c) { askCount(c); return; }
        var a = e.target.closest('.pc-adj');
        if (a) askAdjust(a);
      });
      // Today drawer — everything ambient, one tap away.
      document.getElementById('todayBar').addEventListener('click', function () {
        document.getElementById('todaySheet').classList.toggle('open');
      });
      document.getElementById('todayClose').addEventListener('click', function () {
        document.getElementById('todaySheet').classList.remove('open');
      });
      // ⋯ menu (arcade — SAME TAB — · fullscreen · sound)
      document.getElementById('menuBtn').addEventListener('click', function (ev) {
        ev.stopPropagation();
        document.getElementById('menuPop').classList.toggle('hidden');
      });
      document.addEventListener('click', function (ev) {
        var pop = document.getElementById('menuPop');
        if (!pop.classList.contains('hidden') &&
            !(ev.target && ev.target.closest && ev.target.closest('#menuPop, #menuBtn'))) {
          pop.classList.add('hidden');
        }
      });
      radioInit();
      fetchPrayerTimes();
      setInterval(fetchPrayerTimes, 3 * 3600000);   // refresh prayer times every 3h
      tickClock();
      setInterval(tickClock, CLOCK_MS);
      poll();   // self-rescheduling — see schedulePoll(). Deliberately NOT
                // setInterval: a fixed clock is what caused the pile-up.
    });

""", 'boot rebuilt + weather deleted')

# ── 5. tickClock: drop warmth/rest calls ────────────────────────────────────
swap("""      setWarmth();
      updateNextPrayer();
      updateRestAmbient();""",
     """      updateNextPrayer();""",
     'tickClock slimmed')

# ── 6. theme..stars deletion (preserving HQ_API/HQ_NATIVE) ─────────────────
cut("    // ---------- theme ----------",
    "    function hqCall(action, body) {",
    """    // ---------- host bridge ----------
    var HQ_API    = '/api/board';
    var HQ_NATIVE = (typeof google !== 'undefined' && google.script && google.script.run);

""", 'theme/warmth/stars deleted')

# ── 7. onTick ───────────────────────────────────────────────────────────────
cut("    function onTick(tick) {",
    "    // ---------- WORKING HOURS + ADAPTIVE MODE ----------",
    """    function onTick(tick) {
      // A response is not the same as a USABLE response. If the server answered
      // with something that isn't a tick — an older /exec that doesn't know the
      // action, an error object, a proxy's own reply — painting it would show a
      // confident wall of zeros while the indicator says LIVE. That happened on
      // 2026-08-05 and looked healthy at every layer. Treat a tick with no
      // cockpit as a failure so the board SAYS it lost the line.
      if (!tick || typeof tick !== 'object' || !tick.cockpit) {
        try {
          console.warn('onTick: response is not a tick —',
                       (tick && (tick.message || tick.status)) || tick);
        } catch (e) {}
        return onFail(tick && tick.message ? tick.message : 'bad tick');
      }

      failCount = 0;
      document.getElementById('connLost').classList.remove('show');
      document.getElementById('board').classList.remove('booting');

      lastTick = tick;                   // optimistic picks re-paint from this

      var c = tick.cockpit || {};
      var a = tick.alerts  || {};
      orderAge = c.orderAgeMin || {};   // feeds the "oldest first" re-sort + band ages

      paintTopBar(c, a);
      // BEFORE the pick list — fillPickRow/fillPickHead read the maps this builds.
      buildKitIndex(tick.kits);
      // Overlay any just-tapped rows the server hasn't confirmed yet — a tick
      // served from n8n's 15s cache can still be showing the old status.
      paintPickList(applyPickOverrides((tick.openOrders || []).slice()),
                    tick.openOrdersTotal, tick.openOrdersBy);
      // RAW rows — applyPickOverrides fakes a status the server has not
      // confirmed, which would make an optimistic pick look like a fresh
      // arrival on the very next tick.
      beaconCheck(tick.openOrders || []);
      paintToday(c, tick);
      paintLive(c, tick);
      paintFooter(tick);

      firstTick = false;
    }

    // ---------- WORKING HOURS + ADAPTIVE MODE ----------""",
    'onTick rebuilt')

# ── 8. applyMode deletion (keep isOffHours + houstonHour) ──────────────────
cut("    // Space is earned by live data:",
    "    function houstonHour() {",
    "", 'applyMode deleted')

# ── 9. rest scene + hero paints deletion ────────────────────────────────────
cut("    function paintRest(tick, c) {",
    "    function pickKey(r) {",
    "", 'rest scene + hero paints deleted')

# ── 10. paintUnits → paintTopBar/paintToday/humanNote ───────────────────────
cut("    /* UNITS TO PICK — how BIG the queue is,",
    "    /* Aisle ↔ age.",
    """    // ---------- TOP BAR — the count that matters + the money caution ----------
    function paintTopBar(c, a) {
      var open = num(c.ebayGrab) + num(c.directGrab);
      var bg = document.getElementById('barGrab');
      bg.classList.toggle('clear', open === 0);
      document.getElementById('grabNum').textContent = open === 0 ? '✓' : String(open);
      document.getElementById('grabWord').textContent = open === 0 ? 'caught up' : 'to grab';
      if (!firstTick && prevOpen !== null && open === 0 && prevOpen > 0) chime('caughtUp');
      prevOpen = open;

      var paid = num(a && a.paidShipping && a.paidShipping.count);
      document.getElementById('barPaid').classList.toggle('hidden', paid <= 0);
      document.getElementById('barPaidNum').textContent = String(paid);
      if (!firstTick && prevPaid !== null && paid > prevPaid) chime('paid');
      prevPaid = paid;
    }

    // ---------- TODAY DRAWER — ambient numbers + the feed, one tap away ------
    function paintToday(c, tick) {
      var rows = tick.openOrders || [];
      var units = 0, aisles = {};
      for (var i = 0; i < rows.length; i++) {
        units += num(rows[i].qty) || 1;
        var al = _aisleOf(rows[i].location);
        if (al) aisles[al] = 1;
      }
      var nA = 0; for (var k in aisles) { if (aisles.hasOwnProperty(k)) nA++; }
      // ⚠ openOrders is capped server-side — when truncated the sum is a
      // FLOOR, not a total, and it says so with the "+".
      var capped = num(tick.openOrdersTotal) > rows.length;

      document.getElementById('tdOut').textContent = String(num(c.shippedToday));
      document.getElementById('tdIn').textContent  = String(num(c.receivedToday));
      document.getElementById('tdShipped').textContent = String(num(c.shippedToday));
      document.getElementById('tdShippedSub').textContent =
        num(c.receivedToday) > 0
          ? Math.round(num(c.shippedToday) / num(c.receivedToday) * 100) + '% of today out'
          : ' ';
      document.getElementById('tdReceived').textContent = String(num(c.receivedToday));
      document.getElementById('tdReceivedSub').textContent =
        'eBay ' + num(c.receivedEbay) + ' · Direct ' + num(c.receivedDirect);
      document.getElementById('tdUnits').textContent = String(units) + (capped && units ? '+' : '');
      document.getElementById('tdUnitsSub').textContent = rows.length
        ? (rows.length + (capped ? ' of ' + num(tick.openOrdersTotal) : '') + ' lines · ' +
           nA + (nA === 1 ? ' aisle' : ' aisles'))
        : ' ';

      paintFeed(c.timeline || []);
    }

    // "A-42" → "A". NOT FOUND rows have no aisle and must not invent one.
    function _aisleOf(loc) {
      var s = String(loc || '').trim().toUpperCase();
      if (!s || s === 'NOT FOUND') return '';
      var m = s.match(/^([A-Z]+)\\s*-/);
      return m ? m[1] : '';
    }

    /* ---------- NOTE CLASSIFIER — the ** rule, retired 2026-08-12 ----------
       Show every HUMAN note; hide only machine kit-tags. Legacy ** markers are
       stripped on display, so nothing anyone already wrote breaks — and the
       Telegram /note command keeps writing ** (its formula-injection guard)
       without the floor ever needing to know the convention existed.
       ⚠ Zoho flags (REMOVED IN ZOHO / QTY changed) surface as the row's RED
       warning — before this, the floor could not see a removed line at all. */
    function humanNote(noteText) {
      var raw = String(noteText || '').trim();
      if (!raw) return { text: '', warn: '' };
      if (raw.indexOf('↳') === 0) return { text: '', warn: '' };   // machine kit tag
      if (/ZOHO/i.test(raw) && /^⚠/.test(raw)) {
        return { text: '', warn: raw.replace(/^⚠️?\\s*/, '') };
      }
      var t = raw.replace(/\\*\\*/g, ' ').replace(/\\s+/g, ' ').trim();
      t = t.replace(/^Buyer Note:\\s*/i, '');
      return { text: t, warn: '' };
    }

    /* Aisle ↔ age.""",
    'paintUnits → topbar/today/classifier')

# ── 11. paintKitStrip → buildKitIndex ───────────────────────────────────────
cut("    function paintKitStrip(kits) {",
    "    function paintPickList(rows, total, byChannel) {",
    """    /* Kit progress lives ON the order band it belongs to (the strip was
       retired in the 2026-08 redesign — it sat above the CHANNEL tabs
       narrating the other tab's work, and duplicated what the bands and shut
       cards already said). This just indexes the tick's kits for the rows
       (hue spines) and the bands (progress chips). */
    var kitsByOrder = {};
    function buildKitIndex(kits) {
      kits = kits || [];
      kitHue = {}; kitsByOrder = {};
      for (var i = 0; i < kits.length; i++) {
        kitHue[kits[i].key] = { hue: num(kits[i].hue) || 0, dash: num(kits[i].dash) || 0 };
        var oid = String(kits[i].order || '').trim();
        if (!oid) continue;
        (kitsByOrder[oid] = kitsByOrder[oid] || []).push(kits[i]);
      }
    }
    function runHasKit(rows, start, n) {
      for (var i = start; i < start + n; i++) { if (rows[i].isKit) return true; }
      return false;
    }

""", 'kit strip → kit index')

# ── 12. fillPickRow ─────────────────────────────────────────────────────────
cut("    function fillPickRow(li, r) {",
    "    // ================= PRINT THE PICK LIST =================",
    """    function fillPickRow(li, r) {
      li.setAttribute('data-key', pickKey(r));
      li.innerHTML = '';
      // Set by withSections: this row sits under an order band, so it drops
      // its own order id rather than repeating the band on every line.
      li.classList.toggle('in-group', !!r._grouped);

      var isPrep = String(r.status).toUpperCase() === 'PREPARING';

      // ── shelf — the routing key, the biggest thing on the row ──
      var loc = document.createElement('span');
      loc.className = 'shelf';
      var locTxt = String(r.location || '').trim();
      li.classList.toggle('noloc', !locTxt || locTxt.toUpperCase() === 'NOT FOUND');
      if (!locTxt || locTxt.toUpperCase() === 'NOT FOUND') {
        // dimmed: absent, not broken — a missing shelf is the common non-event
        loc.classList.add('gone');
        loc.textContent = '—';
      } else if (locTxt.indexOf('/') !== -1) {
        // ⚠ "L-208/C-51" IS TWO SHELVES — both real, neither dropped or dimmed
        loc.classList.add('dual');
        var parts = locTxt.split('/');
        for (var pj = 0; pj < parts.length; pj++) {
          if (pj) {
            var sep = document.createElement('i');
            sep.textContent = '/';
            loc.appendChild(sep);
          }
          loc.appendChild(document.createTextNode(parts[pj].trim()));
        }
      } else {
        // "G-35 * 2" — pack-size note recedes, the aisle stays the routing key
        var pack = locTxt.match(/^(.*?)(\\s*\\*\\s*\\d+)$/);
        if (pack) {
          loc.appendChild(document.createTextNode(pack[1].trim()));
          var pk = document.createElement('small');
          pk.textContent = ' ' + pack[2].trim();
          pk.title = 'Warehouse pack size — one sellable unit is packed as ' +
                     pack[2].replace(/\\D/g, '') + ' pieces';
          loc.appendChild(pk);
        } else {
          loc.textContent = locTxt;
          if (locTxt.length > 8) loc.classList.add('longloc');
        }
      }

      // ── main: SKU line + sub line (constant two-line anatomy) ──
      var main = document.createElement('span');
      main.className = 'pick-main-cell';

      var skuLine = document.createElement('span');
      skuLine.className = 'sku-line';
      var sku = document.createElement('b');
      sku.className = 'pick-sku';
      sku.textContent = r.sku || '';
      // Tap the SKU for the part, the order id for the order. stopPropagation
      // so neither can ever reach the row's ✓ Pick button. No `title` —
      // OS tooltips covered the band on the desk board.
      sku.addEventListener('click', function (ev) {
        ev.preventDefault(); ev.stopPropagation();
        openPartDrawer(r.sku, r);
      });
      skuLine.appendChild(sku);
      if (r.isKit) {
        // An unexpanded kit CANNOT be picked off a shelf — its components
        // aren't on the sheet. Without this the picker walks to a K-* aisle
        // expecting a box and finds a decision instead.
        var kb = document.createElement('span');
        kb.className = 'pick-kit';
        kb.textContent = 'KIT';
        kb.title = 'Not expanded yet — decide: ship the box, or break it into components';
        skuLine.appendChild(kb);
      }
      main.appendChild(skuLine);

      var sub = document.createElement('span');
      sub.className = 'sub-line';
      if (!r._grouped) {
        var meta = document.createElement('span');
        meta.className = 'sub-id';
        meta.textContent = r.orderId || '';
        meta.addEventListener('click', function (ev) {
          ev.preventDefault(); ev.stopPropagation();
          openOrderDrawer(r.orderId);
        });
        sub.appendChild(meta);
      }
      var note = humanNote(r.note);
      li.classList.toggle('warned', !!note.warn);
      if (note.warn) {
        var wn = document.createElement('span');
        wn.className = 'sub-warn';
        wn.textContent = '⚠ ' + note.warn;
        wn.title = 'Flagged from Zoho — check with the office before picking this line';
        sub.appendChild(wn);
      }
      li.classList.toggle('noted', !!note.text);
      if (note.text) {
        var nt = document.createElement('span');
        nt.className = 'sub-note';
        nt.textContent = '📌 ' + note.text;
        sub.appendChild(nt);
      }
      main.appendChild(sub);

      // ── ◩ shelf count — its own named lane; buildCountLine is unchanged ──
      var handN  = (typeof r.hand === 'number') ? r.hand : null;
      var countEl = null;
      if (handN !== null && handN <= COUNT_MAX_HAND && !isPrep) {
        countEl = buildCountLine(r, handN);
        // The alarm colour belongs to a counted shelf that DISAGREES — rare,
        // actionable, and the whole reason the count exists.
        var cnt = (typeof r.left === 'number') ? r.left : null;
        li.classList.toggle('deviance',
          cnt !== null && cnt !== Math.max(0, handN - (num(r.qty) || 0)));
      } else {
        li.classList.remove('deviance');
      }

      var qty = document.createElement('span');
      var qn = num(r.qty) || 1;
      qty.className = 'pick-qty' + (qn > 1 ? ' multi' : '');
      qty.textContent = '×' + qn;

      li.appendChild(loc);
      li.appendChild(main);
      if (countEl) li.appendChild(countEl);
      li.appendChild(qty);

      // ⚠ EXACTLY ONE of these, and they share the `act` grid area.
      if (isPrep) {
        var st = document.createElement('span');
        st.className = 'pick-status prep';
        st.textContent = 'Prep';
        li.appendChild(st);
      } else {
        var doBtn = document.createElement('button');
        doBtn.className = 'pick-do';
        doBtn.textContent = '✓ Pick';
        doBtn.setAttribute('data-order', r.orderId || '');
        doBtn.setAttribute('data-sku', r.sku || '');
        li.appendChild(doBtn);
      }

      // KIT THREAD — spine only when the order holds MORE THAN ONE box;
      // otherwise the band already says which rows ship together.
      li.classList.remove('kit-c0', 'kit-c1', 'kit-dash');
      var kh = r._multiKit ? kitHue[kitKeyOf(r)] : null;
      li.classList.toggle('kitted', !!kh);
      if (kh) {
        li.classList.add('kit-c' + kh.hue);
        if (kh.dash) li.classList.add('kit-dash');
      }
    }

    // ================= PRINT THE PICK LIST =================""",
    'fillPickRow rebuilt')

# ── 13. fillPickSec + buildPickSec (slim cap notice) ───────────────────────
cut("    function fillPickSec(li, s) {",
    "    function fillPickHead(li, g) {",
    """    /* The channel word lives on the TABS now. A section item only earns
       pixels when its channel is TRUNCATED — the honesty line, kept: a pick
       list that looks complete but isn't is worse than one that admits it. */
    function fillPickSec(li, s) {
      li.setAttribute('data-key', s.key);
      li.className = 'pick-sec' + (num(s.hidden) > 0 ? ' more' : ' empty');
      li.textContent = num(s.hidden) > 0
        ? ('+' + s.hidden + ' more open on the floor — past the ' +
           (s.channel === 'DIRECT' ? 'Direct' : 'eBay') + ' cap')
        : '';
    }

    function buildPickSec(s) {
      var li = document.createElement('li');
      fillPickSec(li, s);
      return li;
    }

""", 'fillPickSec slimmed')

# ── 14. fillPickHead: kit chips + KIT·decide chip ──────────────────────────
swap("""      li.appendChild(id); li.appendChild(meta);

      // A shut card carries the shelves still to walk""",
     """      li.appendChild(id); li.appendChild(meta);

      // Kit progress lives ON the box it belongs to (round 1 — strip retired).
      var kl = g.kits || [];
      for (var kc = 0; kc < kl.length; kc++) {
        var chip = document.createElement('span');
        chip.className = 'kitchip kit-h' + (num(kl[kc].hue) % 2) +
                         (num(kl[kc].dash) ? ' kit-hd' : '');
        chip.innerHTML = '<i></i>' + esc(String(kl[kc].parent || '')) + ' · ' +
                         num(kl[kc].done) + '/' + num(kl[kc].total);
        li.appendChild(chip);
      }
      // An unexpanded kit is a DECISION — a shut card must not hide it.
      if (g.hasKitRow) {
        var kd = document.createElement('span');
        kd.className = 'kitchip kit-h0';
        kd.innerHTML = '<i></i>KIT · decide';
        kd.title = 'This order holds an unexpanded kit — ship the box, or expand it';
        li.appendChild(kd);
      }

      // A shut card carries the shelves still to walk""",
     'fillPickHead kit chips')

# ── 15. withSections head item carries kit info ────────────────────────────
swap("""              orderId: oid, channel: ch, n: n, age: orderAge[oid],""",
     """              orderId: oid, channel: ch, n: n, age: orderAge[oid],
              kits: kitsByOrder[oid] || null,
              hasKitRow: runHasKit(rows, i, n),""",
     'withSections kit fields')

# ── 16. isToolLayout: this file IS the tool now ────────────────────────────
swap("""    var TOOL_MQ = '(max-width: 1200px), (pointer: coarse) and (max-width: 1600px)';
    var activeChannel = null;          // null = not chosen yet this session
    function isToolLayout() {
      try { return window.matchMedia(TOOL_MQ).matches; } catch (e) { return false; }
    }""",
     """    var activeChannel = null;          // null = not chosen yet this session
    // Since the /wall split this file IS the tool — there is no monitor
    // breakpoint left to detect. Kept as a function because the nudge and
    // channel seeding read it.
    function isToolLayout() { return true; }""",
     'isToolLayout constant')

# ── 17. extractFloorNote + right column deletion ───────────────────────────
cut("    function extractFloorNote(noteText) {",
    "    function paintFeed(timeline) {",
    "", 'floor-note extractor + right column deleted')

# ── 18. breath deletion ─────────────────────────────────────────────────────
cut("    // ---------- ambient breath (barely-there premium) ----------",
    "    // ---------- sound ----------",
    "", 'breath deleted')

open(f'{ROOT}/FloorBoard.html', 'w').write(s)
print(f'\nPORT COMPLETE: {orig_len} → {len(s)} bytes ({orig_len - len(s):+d} removed)')
