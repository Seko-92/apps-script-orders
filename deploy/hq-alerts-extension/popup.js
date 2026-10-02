// HQ Alerts popup — shows what the background worker knows. Changes nothing on the sheet.
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
  return { '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;' }[c]; }); }
function ago(ms) {
  if (!ms) return 'never';
  var s = Math.round((Date.now() - ms) / 1000);
  return s < 60 ? s + ' s ago' : Math.round(s / 60) + ' min ago';
}
function paint() {
  chrome.storage.local.get(['fails', 'lastOk', 'hours', 'openHolds', 'log'], function (st) {
    var big = document.getElementById('big'), sub = document.getElementById('sub');
    if ((st.fails || 0) >= 2) { big.className = 'big bad'; big.textContent = 'Cannot reach the server'; sub.textContent = 'Retrying every 30 seconds · last good ' + ago(st.lastOk); }
    else if (!st.lastOk) { big.className = 'big rest'; big.textContent = 'Starting…'; sub.textContent = ''; }
    else if (st.hours) { big.className = 'big ok'; big.textContent = 'Listening'; sub.textContent = 'New orders and holds · checked ' + ago(st.lastOk); }
    else { big.className = 'big rest'; big.textContent = 'Holds only until 9 AM'; sub.textContent = 'Houston time · checked ' + ago(st.lastOk); }
    var h = st.openHolds || [], hb = document.getElementById('holds');
    hb.hidden = !h.length;
    hb.innerHTML = '<b>⚠ ' + h.length + ' unanswered hold' + (h.length > 1 ? 's' : '') + '</b> · ' + esc(h.join(', '));
    var log = st.log || [];
    document.getElementById('log').innerHTML = log.length
      ? log.map(function (e) { return '<li><time>' + esc(e.t) + '</time><span class="' + esc(e.c) + '">' + esc(e.x) + '</span></li>'; }).join('')
      : '<li><span class="none">nothing yet</span></li>';
  });
}
// opening the popup counts as having seen the new orders
chrome.runtime.sendMessage({ cmd: 'seen' }, function () { paint(); });
chrome.runtime.sendMessage({ cmd: 'poll' }, function () { paint(); });
document.getElementById('btnBoard').addEventListener('click', function () { chrome.runtime.sendMessage({ cmd: 'board' }); window.close(); });
document.getElementById('btnTest').addEventListener('click', function () {
  chrome.runtime.sendMessage({ cmd: 'test' }, function (r) {
    document.getElementById('sub').textContent = r && r.ok ? 'Test sent — check the corner of the screen.' : 'Chrome could not show it — check Windows notification settings for Chrome.';
  });
});
paint();
