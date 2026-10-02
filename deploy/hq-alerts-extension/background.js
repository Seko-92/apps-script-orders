// ===========================================================================
// HQ Alerts — Chrome extension background worker (2026-10-02).
//
// Same job as the /alerts tab, without the tab: Chrome wakes this worker every
// 30 seconds, it reads the board's published tick, and shows what core.js
// decides. Read-only — it never writes to the sheet.
//
// ⚠ The RULES are in core.js, which is the SAME file the /alerts tab loads.
//   This file only does the plumbing: polling, memory, notifications, badge.
// ⚠ A service worker forgets everything between wakes, so every piece of state
//   lives in chrome.storage.local, never in a variable that must survive.
// ===========================================================================
importScripts('core.js');

var ORIGIN = 'https://hq.yassinqurabi.com';
var API = ORIGIN + '/api/board';
var BOARD = ORIGIN + '/';
var POLL_MIN = 0.5;          // 30 s — the shortest alarm Chrome allows (120+)
var FETCH_TIMEOUT_MS = 20000;
var LOG_KEEP = 12;

function get(keys) { return chrome.storage.local.get(keys); }
function set(obj) { return chrome.storage.local.set(obj); }

// ---------- lifecycle ----------
// A (re)start only REMEMBERS what is already open: history is not news.
function start() {
  chrome.alarms.create('poll', { periodInMinutes: POLL_MIN });
  return set({ primed: false }).then(poll);
}
chrome.runtime.onInstalled.addListener(start);
chrome.runtime.onStartup.addListener(start);
chrome.alarms.onAlarm.addListener(function (a) { if (a.name === 'poll') poll(); });

// ---------- notifications ----------
function show(n) {
  var opts = {
    type: 'basic',
    iconUrl: n.kind === 'hold' ? 'icons/hold-128.png' : 'icons/new-128.png',
    title: n.title, message: n.body,
    requireInteraction: !!n.sticky, silent: true, priority: n.sticky ? 2 : 0
  };
  // resolves true ONLY when Chrome accepted it — the log must not lie
  return new Promise(function (resolve) {
    try {
      chrome.notifications.create(n.id, opts, function () { resolve(!chrome.runtime.lastError); });
    } catch (e) { resolve(false); }
  });
}
chrome.notifications.onClicked.addListener(function (id) {
  chrome.notifications.clear(id);
  openBoard();
});
function openBoard() {
  chrome.tabs.query({ url: ORIGIN + '/*' }, function (tabs) {
    var board = (tabs || []).filter(function (t) { return !/\/(alerts|wall|kits|arcade|video)\b/.test(t.url || ''); })[0];
    if (board) { chrome.tabs.update(board.id, { active: true }); chrome.windows.update(board.windowId, { focused: true }); }
    else chrome.tabs.create({ url: BOARD });
  });
}

// ---------- the toolbar badge: the state, even with no notification on screen ----------
function paintBadge(st) {
  var holds = st.openHolds || [], unseen = st.unseen || 0, text = '', bg = '#ffd400', fg = '#111', title;
  if (st.fails >= 2) { text = '!'; bg = '#666'; fg = '#fff'; title = 'HQ Alerts · cannot reach the server'; }
  else if (holds.length) { text = String(holds.length); bg = '#d32f2f'; fg = '#fff'; title = 'HOLD · ' + holds.join(', '); }
  else if (unseen > 0) { text = String(unseen); title = unseen + ' new order' + (unseen > 1 ? 's' : '') + ' · HQ'; }
  else title = 'HQ Alerts · all clear';
  chrome.action.setBadgeText({ text: text });
  chrome.action.setBadgeBackgroundColor({ color: bg });
  if (chrome.action.setBadgeTextColor) chrome.action.setBadgeTextColor({ color: fg });
  chrome.action.setTitle({ title: title });
}

function addLog(log, cls, text, now) {
  log.unshift({ t: HQAlertsCore.clockNow(now), c: cls, x: text });
  return log.slice(0, LOG_KEEP);
}

// ---------- the poll ----------
var busy = false;   // in-memory only: guards an overlap within one wake
function poll() {
  if (busy) return Promise.resolve();
  busy = true;
  var ctl = new AbortController(), timer = setTimeout(function () { ctl.abort(); }, FETCH_TIMEOUT_MS);
  return fetch(API, { method: 'POST', headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({ action: 'boardTick' }), signal: ctl.signal })
    .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
    .then(function (t) {
      if (!t || !t.cockpit) throw new Error('no tick');
      return get(['seen', 'primed', 'log', 'unseen']).then(function (st) {
        var now = Date.now(), seen = HQAlertsCore.freshSeen(st.seen), log = st.log || [];
        var d = HQAlertsCore.decide(t, seen, !!st.primed, now);
        return d.notes.reduce(function (p, n) {
          return p.then(function () {
            return show(n).then(function (ok) {
              if (!ok) return;
              // a hold is remembered only once actually shown
              if (n.holdId) seen.holds[n.holdId] = now;
              log = addLog(log, n.logClass, n.logText, now);
            });
          });
        }, Promise.resolve()).then(function () {
          var next = { seen: seen, primed: true, log: log, fails: 0, lastOk: now, hours: d.hours,
                       unseen: (st.unseen || 0) + d.newCount,
                       openHolds: d.openHolds.map(function (h) { return String(h.orderId); }) };
          return set(next).then(function () { paintBadge(next); });
        });
      });
    })
    .catch(function () {
      return get(['fails', 'openHolds', 'unseen']).then(function (st) {
        var s = { fails: (st.fails || 0) + 1, openHolds: st.openHolds, unseen: st.unseen };
        return set({ fails: s.fails }).then(function () { paintBadge(s); });
      });
    })
    .then(function () { clearTimeout(timer); busy = false; }, function () { clearTimeout(timer); busy = false; });
}

// ---------- popup commands ----------
chrome.runtime.onMessage.addListener(function (msg, _sender, reply) {
  if (!msg) return;
  if (msg.cmd === 'seen') {
    get(['openHolds', 'fails']).then(function (st) {
      st.unseen = 0;
      set({ unseen: 0 }).then(function () { paintBadge(st); reply({ ok: true }); });
    });
    return true;
  }
  if (msg.cmd === 'test') {
    show({ kind: 'order', id: 'hq-test', sticky: false, title: 'Test · HQ Alerts',
           body: 'This PC shows notifications. You are all set.' })
      .then(function (ok) { reply({ ok: ok }); });
    return true;
  }
  if (msg.cmd === 'poll') { poll().then(function () { reply({ ok: true }); }); return true; }
  if (msg.cmd === 'board') { openBoard(); reply({ ok: true }); }
});
