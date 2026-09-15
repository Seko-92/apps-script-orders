/**
 * diag-order-voices.js — the sidebar's order voices, driven against the REAL Sidebar.html.
 *
 * Feeds _hqsNoteTick() synthetic ticks shaped like the published board tick
 * (openOrders + cockpit.orderAgeMin + _publishedAt) and records every sound instead of
 * playing it. What it proves:
 *   A  the first tick after load is history, not news — no sound
 *   B  a new SO number plays ITS CHANNEL's voice (Direct ≠ eBay)
 *   C  an old order scrolling into the list is not an arrival
 *   D  an order that flaps out and back in is announced once
 *   E  an order nobody has touched past its wait repeats, louder, with the strip up
 *   F  one PREPARING line means somebody has it — no repeat
 *   G  eBay waits longer than Direct
 *   H  never outside shop hours
 *   I  "I'm on it" silences it and the next tick keeps it silent
 *   J  the header bell silences arrivals but NOT repeats
 *   K  "stay quiet while I'm clicking" skips arrivals, never repeats
 *   L  garbage in localStorage falls back to the defaults
 *   M  the resting panel cannot cover the alarm
 *   N  the strip survives the cockpit condensing
 *
 * ⚠ SIDEBAR_SRC runs it against an older revision to prove the assertions bite:
 *   git show HEAD:Sidebar.html > /tmp/old.html && SIDEBAR_SRC=/tmp/old.html node diag-order-voices.js
 *   Every section FAILS SOFT — a missing function must not hide the sections after it.
 */
const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');
const SRC = process.env.SIDEBAR_SRC || path.join(__dirname, '..', 'Sidebar.html');
const HTML = fs.readFileSync(SRC, 'utf8').replace("'<?!= boardApiUrl ?>'", "''");

(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 310, height: 1400 } });
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  await p.addInitScript((seed) => {
    if (seed) try { localStorage.setItem('hqSoundSettings.v1', seed); } catch (e) {}
    const mk = (su, fa) => new Proxy({}, { get(_, k) {
      if (k === 'withSuccessHandler') return f => mk(f, fa);
      if (k === 'withFailureHandler') return f => mk(su, f);
      return () => { if (su) setTimeout(() => su(null), 0); };
    } });
    window.google = { script: { run: mk(null, null), host: { close() {}, setHeight() {} }, url: { getLocation(f) { f({ parameter: {} }); } } } };
  }, process.env.SEED || '');
  await p.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: HTML }));
  await p.goto('http://hq.test/sidebar');
  await p.waitForTimeout(1500);

  let pass = 0, fail = 0;
  const ok = (n, c, got) => { if (c) { pass++; console.log('  ✓ ' + n); } else { fail++; console.log('  ✗ ' + n + (got !== undefined ? '  → got ' + JSON.stringify(got) : '')); } };

  // one in-page driver; returns {plays, strip, head, sub, nagging} after each step
  await p.evaluate(() => {
    window.__plays = [];
    window.__open = true;
    if (typeof hqsPlay === 'function') window.hqsPlay = function (v, loud) { window.__plays.push((loud ? 'LOUD:' : '') + v); };
    if (typeof _hqsShopOpen === 'function') window._hqsShopOpen = function () { return window.__open; };
    window.__row = (ch, id, st) => ({ channel: ch, orderId: id, sku: 'X', status: st || 'PENDING' });
    window.__tick = (rows, ages) => {
      window.__plays = [];
      const t = { openOrders: rows, cockpit: { orderAgeMin: ages || {} }, _publishedAt: new Date().toISOString() };
      try { _hqsNoteTick(t); } catch (e) { window.__err = String(e); }
      const s = document.getElementById('hqsAlarm');
      return { plays: window.__plays.slice(), strip: !!(s && !s.hidden), head: s ? document.getElementById('hqsAlarmHead').textContent : null,
               sub: s ? document.getElementById('hqsAlarmSub').textContent : null, nagging: document.body.classList.contains('hq-nagging') };
    };
  });
  const has = await p.evaluate(() => typeof _hqsNoteTick === 'function');
  ok('the order-voices module is present', has);
  // ⚠ Many assertions below are NEGATIVE ("no sound"). Against a sidebar without the module
  //   nothing ever plays, so they would pass vacuously — every one is ANDed with `has`.
  const okRaw = ok;
  const okH = (n, c, got) => okRaw(n, has && c, got);
  const R = (ch, id, st) => ({ channel: ch, orderId: id, sku: 'X', status: st || 'PENDING' });
  const tick = (rows, ages) => p.evaluate(([r, a]) => window.__tick(r, a), [rows, ages]);

  console.log('\nA · the first tick is history, not news');
  let s = await tick([R('DIRECT', 'SO-1'), R('EBAY', '11-1-1')], { 'SO-1': 3, '11-1-1': 2 });
  okH('no sound on the first tick', s.plays.length === 0, s.plays);

  console.log('\nB · a new SO number plays its channel\'s voice');
  s = await tick([R('DIRECT', 'SO-1'), R('EBAY', '11-1-1'), R('DIRECT', 'SO-2')], { 'SO-1': 3, '11-1-1': 2, 'SO-2': 1 });
  await p.waitForTimeout(100);
  let plays = await p.evaluate(() => window.__plays);
  okH('a new Direct order plays the Direct voice (arp)', plays.join() === 'arp', plays);
  await tick([R('DIRECT', 'SO-1'), R('EBAY', '11-1-1'), R('DIRECT', 'SO-2'), R('EBAY', '22-2-2')], { 'SO-1': 3, '11-1-1': 2, 'SO-2': 1, '22-2-2': 0 });
  await p.waitForTimeout(100);
  plays = await p.evaluate(() => window.__plays);
  okH('a new eBay order plays the eBay voice (bright)', plays.join() === 'bright', plays);
  okH('⚠ the two defaults differ', 'arp' !== 'bright');
  const badge = await p.evaluate(() => { const x = document.querySelector('#muteBtn .badge'); return x ? x.textContent : ''; });
  okH('arrivals still count on the bell badge', badge === '2', badge);

  console.log('\nC · an old order scrolling into the list is not an arrival');
  s = await tick([R('DIRECT', 'SO-1'), R('EBAY', '11-1-1'), R('DIRECT', 'SO-2'), R('EBAY', '22-2-2'), R('EBAY', '33-3-3')], { 'SO-1': 3, '11-1-1': 2, 'SO-2': 1, '22-2-2': 0, '33-3-3': 25 });
  await p.waitForTimeout(100);
  plays = await p.evaluate(() => window.__plays);
  okH('an order 25 min old makes no arrival sound', plays.length === 0, plays);

  console.log('\nD · a flap is announced once');
  await tick([R('DIRECT', 'SO-1')], { 'SO-1': 3 });
  s = await tick([R('DIRECT', 'SO-1'), R('DIRECT', 'SO-2')], { 'SO-1': 3, 'SO-2': 2 });
  await p.waitForTimeout(100);
  plays = await p.evaluate(() => window.__plays);
  okH('SO-2 dropping out and coming back does not chime again', plays.length === 0, plays);

  console.log('\nE · nobody has it past its wait → repeat, louder, strip up');
  s = await tick([R('DIRECT', 'SO-9'), R('DIRECT', 'SO-9')], { 'SO-9': 31 });
  okH('the strip is shown', s.strip);
  okH('it names the order', /SO-9/.test(s.head || ''), s.head);
  okH('it says the channel, lines and wait', /^Direct · 2 lines · 31 min$/.test(s.sub || ''), s.sub);
  okH('the first repeat plays at once, LOUD, in the Direct voice', s.plays[0] === 'LOUD:arp', s.plays);
  okH('the body is marked nagging', s.nagging);
  const beats = await p.evaluate(async () => {
    if (typeof _hqsRestartNag !== 'function') return [];
    window.__plays = []; hqs.every = 0.25; _hqsRestartNag();
    await new Promise(r => setTimeout(r, 900));
    return window.__plays.slice();
  });
  okH('it keeps repeating on its interval', beats.length >= 2 && beats.every(x => x === 'LOUD:arp'), beats);

  console.log('\nF · one PREPARING line means somebody has it');
  s = await tick([R('DIRECT', 'SO-9', 'PENDING'), R('DIRECT', 'SO-9', 'PREPARING')], { 'SO-9': 45 });
  okH('no strip', !s.strip);
  okH('no repeat', s.plays.length === 0, s.plays);
  const stopped = await p.evaluate(async () => { window.__plays = []; await new Promise(r => setTimeout(r, 700)); return window.__plays.slice(); });
  okH('⚠ and the interval really stopped', stopped.length === 0, stopped);

  console.log('\nG · eBay waits longer than Direct');
  s = await tick([R('EBAY', '44-4-4')], { '44-4-4': 45 });
  okH('eBay at 45 min: no repeat yet', !s.strip && s.plays.length === 0, s);
  s = await tick([R('EBAY', '44-4-4')], { '44-4-4': 61 });
  okH('eBay at 61 min: repeats in the eBay voice', s.strip && s.plays[0] === 'LOUD:bright', s.plays);

  console.log('\nH · never outside shop hours');
  await p.evaluate(() => { window.__open = false; });
  s = await tick([R('EBAY', '44-4-4')], { '44-4-4': 90 });
  okH('closed: strip hidden and silent', !s.strip && s.plays.length === 0, s);
  await p.evaluate(() => { window.__open = true; });

  console.log('\nI · "I\'m on it"');
  s = await tick([R('DIRECT', 'SO-7')], { 'SO-7': 40 });
  okH('(repeating first)', s.strip);
  const after = await p.evaluate(() => { window.__plays = []; if (typeof hqsOnIt !== 'function') return { strip: true, nag: true }; hqsOnIt(); const st = document.getElementById('hqsAlarm'); return { strip: !!st && !st.hidden, nag: document.body.classList.contains('hq-nagging') }; });
  okH('the strip goes away at once', !after.strip && !after.nag, after);
  s = await tick([R('DIRECT', 'SO-7')], { 'SO-7': 41 });
  okH('⚠ the next tick keeps it quiet', !s.strip && s.plays.length === 0, s);
  s = await tick([R('DIRECT', 'SO-7'), R('DIRECT', 'SO-8')], { 'SO-7': 41, 'SO-8': 35 });
  okH('a DIFFERENT unattended order still repeats', s.strip && /SO-8/.test(s.head), s.head);
  await tick([], {});

  console.log('\nJ · the header bell silences arrivals, not repeats');
  await p.evaluate(() => { isMuted = true; });
  await tick([R('DIRECT', 'SO-20')], { 'SO-20': 2 });
  await tick([R('DIRECT', 'SO-20'), R('DIRECT', 'SO-21')], { 'SO-20': 2, 'SO-21': 1 });
  await p.waitForTimeout(100);
  plays = await p.evaluate(() => window.__plays);
  okH('muted: a new order is silent', plays.length === 0, plays);
  s = await tick([R('DIRECT', 'SO-20')], { 'SO-20': 33 });
  okH('⚠ muted: an unattended order STILL repeats', s.plays[0] === 'LOUD:arp', s.plays);
  await p.evaluate(() => { isMuted = false; });
  await tick([], {});

  console.log('\nK · stay quiet while I\'m clicking');
  await p.evaluate(() => { window._hqsLastHuman = Date.now(); });
  await tick([R('EBAY', '55-5-5'), R('EBAY', '66-6-6')], { '55-5-5': 1, '66-6-6': 1 });
  await tick([R('EBAY', '55-5-5'), R('EBAY', '66-6-6'), R('EBAY', '77-7-7')], { '55-5-5': 1, '66-6-6': 1, '77-7-7': 0 });
  await p.waitForTimeout(100);
  plays = await p.evaluate(() => window.__plays);
  okH('clicked a moment ago: the arrival is skipped', plays.length === 0, plays);
  s = await tick([R('EBAY', '55-5-5')], { '55-5-5': 70 });
  okH('clicked a moment ago: a repeat still plays', s.plays[0] === 'LOUD:bright', s.plays);
  await tick([], {});

  console.log('\nL · settings');
  const def = await p.evaluate(() => typeof hqs === 'object' ? JSON.stringify({ vol: hqs.vol, d: hqs.nag.direct.after, e: hqs.nag.ebay.after, dv: hqs.arrive.direct.voice, ev: hqs.arrive.ebay.voice }) : null);
  okH('defaults: 85% volume, Direct after 30, eBay after 60, arp vs bright',
     def === JSON.stringify({ vol: 85, d: 30, e: 60, dv: 'arp', ev: 'bright' }), def);
  const bad = await p.evaluate(() => {
    if (typeof _hqsLoad !== 'function') return null;
    localStorage.setItem('hqSoundSettings.v1', JSON.stringify({ vol: 900, every: 7, nag: { direct: { after: 11, on: 'yes' } }, arrive: { ebay: { voice: 'kazoo' } } }));
    const x = _hqsLoad();
    localStorage.removeItem('hqSoundSettings.v1');
    return JSON.stringify({ vol: x.vol, every: x.every, d: x.nag.direct.after, don: x.nag.direct.on, ev: x.arrive.ebay.voice });
  });
  okH('⚠ garbage in storage falls back to the defaults, key by key',
     bad === JSON.stringify({ vol: 85, every: 60, d: 30, don: true, ev: 'bright' }), bad);
  const cardOk = await p.evaluate(() => ['hqsVol', 'hqsEvery', 'hqsQuiet', 'hqsHours', 'hqsArrDirectOn', 'hqsArrDirectVoice',
    'hqsArrEbayVoice', 'hqsNagDirectAfter', 'hqsNagEbayAfter'].every(id => !!document.getElementById(id)));
  okH('the card has every control', cardOk);
  const houstonOnly = await p.evaluate(() => { const c = document.querySelector('.card[data-id="sound"]'); return c ? !/Riyadh/i.test(c.textContent) && /Houston/.test(c.textContent) : false; });
  okH('the card names Houston time only', houstonOnly);

  console.log('\nM · the resting panel cannot cover the alarm');
  await tick([R('DIRECT', 'SO-30')], { 'SO-30': 50 });
  const may = await p.evaluate(() => typeof _rpMayShow === 'function' ? _rpMayShow() : null);
  okH('_rpMayShow() is false while an order repeats', may === false, may);

  console.log('\nN · the strip survives the cockpit condensing');
  const cond = await p.evaluate(() => { try { _setCockpitCondensed(true); } catch (e) {} const s = document.getElementById('hqsAlarm'); return s ? s.getBoundingClientRect().height : 0; });
  await p.waitForTimeout(400);
  const cond2 = await p.evaluate(() => { const s = document.getElementById('hqsAlarm'); return s ? s.getBoundingClientRect().height : 0; });
  okH('the strip keeps its height when condensed', cond > 20 && cond2 > 20, [cond, cond2]);
  const outside = await p.evaluate(() => { const s = document.getElementById('hqsAlarm'); return !!s && !s.closest('#modules'); });
  okH('the strip sits outside #modules (never inside a collapsible card)', outside);
  await tick([], {});

  const err = await p.evaluate(() => window.__err || null);
  okH('no error inside _hqsNoteTick', !err, err);
  okH('no page errors', errs.length === 0, errs.slice(0, 3));
  console.log('\n' + (fail ? '✗ ' + fail + ' FAILED' : '✓ all') + ' · ' + pass + ' passed\n');
  await b.close();
  process.exit(fail ? 1 : 0);
})();
