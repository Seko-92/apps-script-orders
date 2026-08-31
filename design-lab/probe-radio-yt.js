// =====================================================================================
// probe-radio-yt.js — does the YouTube-backed station actually play ON THE REAL BOARD?
//
// Parsing proves nothing. This serves the REAL FloorBoard.html from a real origin (the
// documented pattern — setContent leaves an opaque origin where localStorage THROWS and
// kills the whole script), boots it against a stubbed tick, then drives the radio the
// way a picker does and asks the YouTube player what state it reached.
//
// ⚠ youtube.com is deliberately NOT intercepted — the point is real network.
// =====================================================================================
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');
const MOCK = require('./mock-tick.js');

const html = fs.readFileSync(path.join(__dirname, '..', 'FloorBoard.html'), 'utf8')
               .replace(/<\?!?=?\s*boardApiUrl\s*\?>/g, '/api/board');

let pass = 0, fail = 0;
const ok = (n, c, x) => { c ? pass++ : (fail++, console.log('  ✗ ' + n + (x !== undefined ? '  → ' + JSON.stringify(x) : ''))); };

(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1280, height: 800 }, timezoneId: 'America/Chicago' });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));

  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      let res = { ok: false };
      if (body.action === 'boardTick')   res = Object.assign({ ok: true }, MOCK);
      if (body.action === 'boardStatus') res = { ok: true };
      if (body.action === 'boardRadio')  res = { ok: true, nowPlaying: '' };
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());

  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.waitForTimeout(2500);

  console.log('\n  probe-radio-yt.js — the station the floor asked for\n');

  // ── the station exists, and is the yt kind ──────────────────────────────────
  const info = await page.evaluate(() => {
    const i = radioStations.findIndex(s => s.yt);
    return { idx: i, total: radioStations.length,
             name: i >= 0 ? radioStations[i].name : null,
             yt:   i >= 0 ? radioStations[i].yt   : null,
             group: i >= 0 ? radioStations[i].group : null };
  });
  ok('the station is registered', info.idx >= 0, info);
  ok('it carries a yt id, not a url', !!info.yt, info.yt);
  ok('it sits in the Arabic group', info.group === 'Arabic', info.group);

  // ── ⚠ NOTHING is loaded until it is chosen — the whole cost argument ────────
  const before = await page.evaluate(() => ({
    iframes: document.querySelectorAll('iframe[src*="youtube"]').length,
    apiTag:  !!document.querySelector('script[src*="iframe_api"]')
  }));
  ok('LAZY · no iframe before the station is picked', before.iframes === 0, before);
  ok('LAZY · the API script is not fetched either', before.apiTag === false, before);

  // ── drive it the way a picker does ──────────────────────────────────────────
  await page.evaluate(i => playStation(i), info.idx);
  await page.waitForTimeout(14000);

  const after = await page.evaluate(() => {
    const st = (window.ytPlayer && ytPlayer.getPlayerState) ? ytPlayer.getPlayerState() : null;
    return {
      // ⚠ YT.Player REPLACES #ytHost with the iframe rather than nesting one
      //   inside it. Counting children of #ytHost reports 0 forever — a harness
      //   bug that looked exactly like a product bug on the first run.
      iframes: document.querySelectorAll('iframe[src*="youtube"]').length,
      // ⚠ The API swaps the DIV for an IFRAME that KEEPS THE SAME ID. So the host
      //   is neither gone nor a parent — it IS the iframe. Two wrong assertions in
      //   a row before checking the DOM; the board was right both times.
      hostTag: (document.getElementById('ytHost') || {}).tagName || null,
      state: st,
      muted: (window.ytPlayer && ytPlayer.isMuted) ? ytPlayer.isMuted() : null,
      playing: window.radioPlaying,
      meta: (document.getElementById('radioMeta') || {}).textContent,
      audioPaused: document.getElementById('radioAudio').paused
    };
  });
  const NAMES = { '-1':'unstarted','0':'ended','1':'PLAYING','2':'paused','3':'buffering','5':'cued' };
  ok('the iframe is created on demand', after.iframes === 1, after.iframes);
  ok('and the off-screen host IS that iframe now', after.hostTag === 'IFRAME', after.hostTag);
  ok('the player reaches PLAYING or buffering', after.state === 1 || after.state === 3,
     NAMES[String(after.state)] || after.state);
  ok('it is UNMUTED — a muted radio is a broken radio', after.muted === false, after.muted);
  ok('the board considers the radio on', after.playing === true, after.playing);
  ok('the station name shows in the footer', /Sunna/.test(after.meta || ''), after.meta);
  ok('⚠ the <audio> element is NOT also running', after.audioPaused === true, after.audioPaused);

  // ── switching away must stop it — no two sources at once ────────────────────
  const other = await page.evaluate(() => radioStations.findIndex(s => !s.yt));
  await page.evaluate(i => playStation(i), other);
  await page.waitForTimeout(3000);
  const sw = await page.evaluate(() => ({
    yt: (window.ytPlayer && ytPlayer.getPlayerState) ? ytPlayer.getPlayerState() : null
  }));
  ok('switching to an audio station stops the YouTube one', sw.yt !== 1, NAMES[String(sw.yt)] || sw.yt);

  // ══ PASTE-YOUR-OWN ═══════════════════════════════════════════════════════
  console.log('\n  ── paste a link\n');

  // ---- the sniffer, all five link shapes plus every refusal ----------------
  const P = await page.evaluate(() => {
    const t = u => { const r = _radioParse(u); return { yt: r.yt || null, url: r.url || null, err: !!r.err, warn: !!r.warn }; };
    return {
      watch:  t('https://www.youtube.com/watch?v=Rs7St51oDDc'),
      short:  t('https://youtu.be/Rs7St51oDDc'),
      live:   t('https://www.youtube.com/live/Rs7St51oDDc'),
      extra:  t('https://www.youtube.com/watch?app=desktop&v=Rs7St51oDDc&t=90'),
      stream: t('https://radio.mp3islam.com/listen/quran_radio/radio.mp3'),
      hls:    t('https://example.com/live/x.m3u8'),
      http:   t('http://insecure.example.com/stream.mp3'),
      junk:   t('hello there'),
      ytbad:  t('https://www.youtube.com/feed/subscriptions'),
      empty:  t('   ')
    };
  });
  ok('watch?v= link', P.watch.yt === 'Rs7St51oDDc', P.watch);
  ok('youtu.be short link', P.short.yt === 'Rs7St51oDDc', P.short);
  ok('/live/ link', P.live.yt === 'Rs7St51oDDc', P.live);
  ok('id found among other query params', P.extra.yt === 'Rs7St51oDDc', P.extra);
  ok('a plain stream url takes the audio path', !!P.stream.url && !P.stream.yt, P.stream);
  ok('HLS is allowed but WARNED about', !!P.hls.url && P.hls.warn === true, P.hls);
  ok('⚠ http is refused — the board is https, it would fail silently', P.http.err, P.http);
  ok('non-links are refused', P.junk.err, P.junk);
  ok('a YouTube page with no video id is refused', P.ytbad.err, P.ytbad);
  ok('empty is refused', P.empty.err, P.empty);

  // ---- add one, and it should join the list AND persist --------------------
  const nBefore = await page.evaluate(() => radioStations.length);
  const added = await page.evaluate(() =>
    radioAddCustom('https://radio.mp3islam.com/listen/sudais/radio.mp3'));
  await page.waitForTimeout(1500);
  const st = await page.evaluate(() => {
    const i = radioStations.findIndex(s => s.mine);
    let saved = [];
    try { saved = JSON.parse(localStorage.getItem('hqRadioCustom') || '[]'); } catch (e) {}
    return { len: radioStations.length, idx: i, group: i >= 0 ? radioStations[i].group : null,
             saved: saved.length, playingIdx: radioIdx, playing: radioPlaying };
  });
  ok('the station is added', st.len === nBefore + 1, { before: nBefore, after: st.len });
  ok('it lands in its own group', st.group === 'Yours', st.group);
  ok('it starts playing immediately', st.playing === true && st.playingIdx === st.idx, st);
  ok('and it is persisted to localStorage', st.saved === 1, st.saved);
  ok('add reports no error', !added.err, added);

  // ---- the popup renders it with a remove affordance -----------------------
  const pop = await page.evaluate(() => { buildRadioPop(); const p = document.getElementById('radioPop');
    return { items: p.querySelectorAll('.radio-pop-item').length,
             customs: p.querySelectorAll('.radio-pop-item.custom').length,
             xs: p.querySelectorAll('.radio-pop-x').length,
             box: !!p.querySelector('#radioAddBox'), btn: !!p.querySelector('#radioAddBtn') }; });
  ok('the paste box is in the popup', pop.box && pop.btn, pop);
  ok('the saved station renders as removable', pop.customs === 1 && pop.xs === 1, pop);
  ok('built-ins get no ✕', pop.items > pop.customs, pop);

  // ---- pasting the SAME link twice must not duplicate ----------------------
  const dup = await page.evaluate(() => {
    radioAddCustom('https://radio.mp3islam.com/listen/sudais/radio.mp3');
    return radioStations.filter(s => s.mine).length;
  });
  ok('the same link twice does not duplicate', dup === 1, dup);

  // ---- remove --------------------------------------------------------------
  const rm = await page.evaluate(() => {
    const i = radioStations.findIndex(s => s.mine);
    radioRemoveCustom(i);
    let saved = [];
    try { saved = JSON.parse(localStorage.getItem('hqRadioCustom') || '[]'); } catch (e) {}
    return { mine: radioStations.filter(s => s.mine).length, saved: saved.length,
             builtins: radioStations.length };
  });
  ok('remove drops it', rm.mine === 0, rm);
  ok('and clears it from storage', rm.saved === 0, rm);
  ok('⚠ built-in stations are untouched by remove', rm.builtins === nBefore, { before: nBefore, now: rm.builtins });

  ok('no page errors', errors.length === 0, errors.slice(0, 2));
  console.log('\n  ' + pass + ' passed, ' + fail + ' failed\n');
  await b.close();
  process.exit(fail ? 1 : 0);
})();
