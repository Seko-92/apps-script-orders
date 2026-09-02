/**
 * server.js — GET /dial?… -> a PNG of the banner dial.
 *
 * ⭐⭐ THE SERVER HOLDS NO DATA AND READS NOTHING. Every number arrives in the query string,
 *    already computed by __SparkData. That is what makes an UNAUTHENTICATED endpoint the
 *    right shape here: Google's image fetcher cannot send a header, so the route has to be
 *    open — and it is safe to be open precisely because a caller can only ever render
 *    numbers they themselves supplied. No sheet access, no filesystem, no network, no
 *    credentials in the image.
 *
 * ⚠ IT IS STILL A PUBLIC ENDPOINT ON A BOX THAT RUNS n8n AND THE BOARD, so it is bounded
 *   on purpose: query length capped, scale clamped, one render per request, nothing
 *   spawned. A dial is not worth a resource-exhaustion surface.
 *
 * ⚠⚠ /healthz ANSWERS WITH A SHAPE, NOT A STATUS. A path that does not exist under
 *    hq.yassinqurabi.com answers **HTTP 200 with the Floor Board's 303 KB of HTML** —
 *    Caddy's try_files. So "did the deploy land" can never be answered by a status code
 *    here; the verify greps the BODY. Same success-status/failure-content shape as Zoho's
 *    200-with-an-error-body and the /exec error page n8n banked as success for weeks.
 */
'use strict';

const http = require('http');
const { renderPng, renderGif, registerFonts, assertFontWeights, FACES } = require('./render');
const { VERDICTS } = require('./state');
const { WIDTH, HEIGHT } = require('./draw');

const PORT      = Number(process.env.PORT || 8099);
const MAX_QUERY = 2048;   // 24 hourly counts + a dozen scalars fits in a few hundred
const MAX_SCALE = 3;

let served = 0, failed = 0;
const startedAt = new Date().toISOString();

function send(res, code, type, body, extra) {
  res.writeHead(code, Object.assign({
    'Content-Type': type,
    'Content-Length': Buffer.byteLength(body),
    'X-Content-Type-Options': 'nosniff'
  }, extra || {}));
  res.end(body);
}

const server = http.createServer((req, res) => {
  let url;
  try { url = new URL(req.url, 'http://dial.local'); }
  catch (e) { return send(res, 400, 'text/plain; charset=utf-8', 'bad url\n'); }

  if (req.method !== 'GET' && req.method !== 'HEAD') {
    return send(res, 405, 'text/plain; charset=utf-8', 'GET only\n', { Allow: 'GET, HEAD' });
  }

  const p = url.pathname.replace(/\/+$/, '') || '/';

  if (p === '/healthz' || p === '/dial/healthz') {
    // ⚠ The literal below is what the deploy verify greps for. Changing it means changing
    //   the verify in the same commit — a health check nobody can assert is decoration.
    const body = JSON.stringify({
      ok: true, service: 'hq-dial', size: WIDTH + 'x' + HEIGHT,
      faces: Object.keys(FACES),
      verdicts: VERDICTS, served: served, failed: failed,
      fontProblems: assertFontWeights(), startedAt: startedAt
    }) + '\n';
    return send(res, 200, 'application/json; charset=utf-8', body,
                { 'Cache-Control': 'no-store' });
  }

  if (p !== '/dial') {
    return send(res, 404, 'text/plain; charset=utf-8', 'hq-dial: only /dial and /healthz\n');
  }

  if ((url.search || '').length > MAX_QUERY) {
    return send(res, 414, 'text/plain; charset=utf-8', 'query too long\n');
  }

  const q = {};
  for (const [k, v] of url.searchParams) if (!(k in q)) q[k] = v;   // first wins
  const scale = Math.min(MAX_SCALE, Math.max(1, Number(q.x) || 1));

  try {
    // ⚠ anim=1 is the FLOATING-IMAGE path only. =IMAGE() renders frame one of a GIF and
    //   never advances (settled 2026-08-30), so asking for a GIF through the cell formula
    //   would silently show a scrambled first frame — the worst possible failure. The
    //   board's =IMAGE() formula must never carry anim=1.
    const wantsGif = q.anim === '1' && q.face === 'board';
    if (wantsGif) {
      const gif = renderGif(q, scale);
      served++;
      if (req.method === 'HEAD') {
        res.writeHead(200, { 'Content-Type': 'image/gif', 'Content-Length': gif.length });
        return res.end();
      }
      return send(res, 200, 'image/gif', gif, {
        'Cache-Control': 'public, max-age=600',
        'X-Dial-Face': 'board',
        'X-Dial-Anim': 'settle-once'
      });
    }
    const png = renderPng(q, scale);
    served++;
    if (req.method === 'HEAD') {
      res.writeHead(200, { 'Content-Type': 'image/png', 'Content-Length': png.length });
      return res.end();
    }
    // Each minute is its own URL (t=HHmm), so a long max-age costs nothing and spares the
    // box a re-render for every viewer of the same minute.
    return send(res, 200, 'image/png', png, {
      'Cache-Control': 'public, max-age=600',
      'X-Dial-Verdict': VERDICTS.indexOf(q.s) >= 0 ? q.s : 'clear',
      'X-Dial-Face': (q.face && FACES[q.face]) ? q.face : 'dial'
    });
  } catch (e) {
    failed++;
    console.error('render failed:', e && e.stack ? e.stack : e);
    // ⚠ FAIL AS A NON-IMAGE, DELIBERATELY. =IMAGE() then errors, IFERROR catches it, and
    //   the banner falls back to its text chip — visible and obviously wrong. Returning a
    //   placeholder PNG would put a confident-looking dial on a broken renderer.
    return send(res, 500, 'text/plain; charset=utf-8', 'render failed\n',
                { 'Cache-Control': 'no-store' });
  }
});

registerFonts();
server.listen(PORT, '0.0.0.0', () => {
  console.log(`hq-dial listening on :${PORT} · ${WIDTH}x${HEIGHT} · verdicts ${VERDICTS.join(',')}`);
});

// ⚠ Compose sends SIGTERM on `down`/`restart`; without this the container waits out the
//   10s kill timeout on every deploy.
for (const sig of ['SIGTERM', 'SIGINT']) {
  process.on(sig, () => server.close(() => process.exit(0)));
}

module.exports = server;
