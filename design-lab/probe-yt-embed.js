// probe-yt-embed.js — can the board legitimately embed this stream?
// Renders the OFFICIAL YouTube IFrame Player API and asks the player itself.
// Error 101/150 = the OWNER disabled embedding → we stop, we do not work around it.
const { chromium } = require('playwright');
const VID = process.argv[2] || 'Rs7St51oDDc';

(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  const log = [];
  p.on('console', m => log.push(m.text()));

  await p.route('http://hq.test/**', r => r.fulfill({
    status: 200, contentType: 'text/html; charset=utf-8',
    body: `<!doctype html><html><body><div id="yt"></div>
    <script src="https://www.youtube.com/iframe_api"></script>
    <script>
      window.__r = { state: null, error: null, ready: false };
      function onYouTubeIframeAPIReady() {
        new YT.Player('yt', {
          videoId: '${VID}', height: '200', width: '320',
          playerVars: { autoplay: 1, mute: 1, playsinline: 1 },
          events: {
            onReady:       e => { window.__r.ready = true; e.target.playVideo(); },
            onStateChange: e => { window.__r.state = e.data; },
            onError:       e => { window.__r.error = e.data; }
          }
        });
      }
    </script></body></html>`
  }));

  await p.goto('http://hq.test/', { waitUntil: 'domcontentloaded' });
  await p.waitForTimeout(12000);
  const r = await p.evaluate(() => window.__r);

  const STATE = { '-1':'unstarted', '0':'ended', '1':'PLAYING', '2':'paused', '3':'buffering', '5':'cued' };
  const ERR   = { 2:'invalid id', 5:'html5 error', 100:'video not found / private',
                  101:'EMBEDDING DISABLED BY OWNER', 150:'EMBEDDING DISABLED BY OWNER' };

  console.log('\n  video      ' + VID);
  console.log('  api ready  ' + r.ready);
  console.log('  state      ' + (r.state === null ? 'none' : (STATE[String(r.state)] || r.state)));
  console.log('  error      ' + (r.error === null ? 'none' : (r.error + ' — ' + (ERR[r.error] || '?'))));
  console.log('\n  VERDICT    ' + (
    r.error === 101 || r.error === 150
      ? '✗ owner disabled embedding — STOP, do not work around it'
      : (r.error !== null ? '✗ not playable: ' + r.error
      : (r.state === 1 || r.state === 3 ? '✓ EMBEDS AND PLAYS' : '? inconclusive — state ' + r.state))) + '\n');
  await b.close();
})();
