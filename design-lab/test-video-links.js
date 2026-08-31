/*
 * test-video-links.js — THE DRIFT TEST.
 * ---------------------------------------------------------------------------
 * `video.html` and `FloorBoard.html` are separate static files with no shared
 * module, so the link rule necessarily exists TWICE. Two copies of one rule is
 * exactly how `A-9` sorted after `A-50` in three files until August, and how
 * the board and the server had to be pinned together on noteHasHold.
 *
 * So this extracts BOTH copies out of the SHIPPED files and runs them over the
 * same corpus. It compares the DECISION — which branch fired — not the prose,
 * because a message may legitimately read differently on a board than on a
 * break-room page. The rule may not differ.
 *
 * ⚠ IT MUST FAIL LOUDLY IF EXTRACTION BREAKS. A harness that cannot see the
 * thing it names passes vacuously, which is how the command palette survived
 * two emoji sweeps. There are assertions on the extraction itself.
 *
 *   node test-video-links.js
 */
const fs = require('fs');
const path = require('path');

const ROOT = process.env.SRC || path.join(__dirname, '..');
let pass = 0, fail = 0;
function t(name, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (ok) { pass++; console.log('  ✓ ' + name); }
  else { fail++; console.log('  ✗ ' + name + '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)); }
}

/* Pull a named function out of a file by brace-matching from its opening `{`.
   Everything in both copies balances (the only braces inside are the regex
   quantifier {11} and the comment text), so a counter is sufficient — but the
   result is asserted to parse, so a future edit that breaks this fails here
   rather than silently extracting half a function. */
function extract(file, fnName) {
  const src = fs.readFileSync(path.join(ROOT, file), 'utf8');
  const at = src.indexOf('function ' + fnName + '(');
  if (at < 0) throw new Error('NOT FOUND: ' + fnName + ' in ' + file);
  let i = src.indexOf('{', at), depth = 0, end = -1;
  for (let j = i; j < src.length; j++) {
    if (src[j] === '{') depth++;
    else if (src[j] === '}') { depth--; if (depth === 0) { end = j + 1; break; } }
  }
  if (end < 0) throw new Error('UNBALANCED: ' + fnName + ' in ' + file);
  const code = src.slice(at, end);
  return { code, fn: new Function(code + '; return ' + fnName + ';')() };
}

console.log('EXTRACTION — this test is worthless if it cannot see both copies');
let radio, video;
{
  radio = extract('FloorBoard.html', '_radioParse');
  video = extract('video.html', '_vidParse');
  t('found _radioParse in FloorBoard.html', typeof radio.fn, 'function');
  t('found _vidParse in video.html',        typeof video.fn, 'function');
  t('  … and neither is a stub',
    radio.code.length > 400 && video.code.length > 400, true);
}

/* The DECISION, stripped of wording. */
const verdict = r => ({
  yt:   r.yt  || null,
  url:  r.url || null,
  err:  !!r.err,
  warn: !!r.warn
});

const CORPUS = [
  // --- the five YouTube shapes that all carry the 11-char id ---
  'https://www.youtube.com/watch?v=dQw4w9WgXcQ',
  'https://youtube.com/watch?v=dQw4w9WgXcQ&t=42s',
  'https://www.youtube.com/watch?list=PLxyz&v=dQw4w9WgXcQ',
  'https://youtu.be/dQw4w9WgXcQ',
  'https://youtu.be/dQw4w9WgXcQ?t=90',
  'https://www.youtube.com/live/dQw4w9WgXcQ',
  'https://www.youtube.com/shorts/dQw4w9WgXcQ',
  'https://www.youtube.com/embed/dQw4w9WgXcQ',
  // --- YouTube, but nothing playable in it ---
  'https://www.youtube.com/@somechannel',
  'https://www.youtube.com/feed/subscriptions',
  'https://youtube.com/',
  // --- the refusals ---
  'http://stream.example.com/live.mp3',
  'not a link at all',
  '',
  '   ',
  'ftp://example.com/a.mp4',
  // --- plain https media ---
  'https://ice1.somafm.com/groovesalad-128-mp3',
  'https://example.com/clip.mp4',
  'https://www.example.com/a/b/c.webm?x=1',
  // --- HLS, allowed but warned ---
  'https://example.com/live/index.m3u8',
  'https://example.com/live/index.m3u8?token=abc',
  // --- near-misses on the id length ---
  'https://youtu.be/short',
  'https://www.youtube.com/watch?v=TOOLONGIDENTIFIER123',
  // --- whitespace tolerance ---
  '   https://youtu.be/dQw4w9WgXcQ   '
];

console.log('\n⭐ THE TWO COPIES MUST AGREE ON EVERY CASE');
{
  let disagreements = [];
  CORPUS.forEach(s => {
    const a = verdict(radio.fn(s));
    const b = verdict(video.fn(s));
    if (JSON.stringify(a) !== JSON.stringify(b)) disagreements.push({ s, board: a, video: b });
  });
  t('all ' + CORPUS.length + ' inputs decided identically', disagreements, []);
}

console.log('\nAND THE RULE ITSELF IS RIGHT (pinned on the video copy)');
{
  const p = video.fn;
  t('watch link → id',        p('https://www.youtube.com/watch?v=dQw4w9WgXcQ').yt, 'dQw4w9WgXcQ');
  t('id survives a leading list param',
    p('https://www.youtube.com/watch?list=PLxyz&v=dQw4w9WgXcQ').yt, 'dQw4w9WgXcQ');
  t('youtu.be → id',          p('https://youtu.be/dQw4w9WgXcQ').yt, 'dQw4w9WgXcQ');
  t('live → id',              p('https://www.youtube.com/live/dQw4w9WgXcQ').yt, 'dQw4w9WgXcQ');
  t('shorts → id',            p('https://www.youtube.com/shorts/dQw4w9WgXcQ').yt, 'dQw4w9WgXcQ');
  t('embed → id',             p('https://www.youtube.com/embed/dQw4w9WgXcQ').yt, 'dQw4w9WgXcQ');

  t('a channel link is REFUSED with a reason',
    !!p('https://www.youtube.com/@somechannel').err, true);
  t('  … and the reason names the problem',
    /no video id/i.test(p('https://www.youtube.com/@somechannel').err), true);

  // ⚠ THE COMMONEST SILENT FAILURE. Many radio-directory entries are http-only,
  // and on an https page the browser blocks them with nothing in the UI.
  t('http:// is REFUSED', !!p('http://stream.example.com/live.mp3').err, true);
  t('  … and says WHY (mixed content), not just "failed"',
    /https/i.test(p('http://stream.example.com/live.mp3').err), true);

  t('plain text is refused',   !!p('hello').err, true);
  t('empty is refused',        !!p('').err, true);
  t('whitespace is refused',   !!p('   ').err, true);
  t('ftp is refused',          !!p('ftp://example.com/a.mp4').err, true);

  t('an https media url is accepted',
    p('https://example.com/clip.mp4').url, 'https://example.com/clip.mp4');
  t('  … and is named by its host',
    p('https://www.example.com/a/b/c.webm?x=1').name, 'example.com');

  t('m3u8 is ALLOWED', !!p('https://example.com/live/index.m3u8').url, true);
  t('  … but warned about', !!p('https://example.com/live/index.m3u8').warn, true);

  t('a 5-char id is not an id',  p('https://youtu.be/short').yt, undefined);
  t('surrounding whitespace is tolerated',
    p('   https://youtu.be/dQw4w9WgXcQ   ').yt, 'dQw4w9WgXcQ');
}

console.log('\n⚠ THE PAGE MUST NOT REGRESS ON ITS OWN PROMISES');
{
  const html = fs.readFileSync(path.join(ROOT, 'video.html'), 'utf8');
  const noComments = html.replace(/<!--[\s\S]*?-->/g, '').replace(/\/\*[\s\S]*?\*\//g, '');

  t('the YouTube API script is LAZY (only fetched on play)',
    /createElement\('script'\)/.test(noComments) &&
    !/<script[^>]+iframe_api/.test(noComments), true);
  t('embedding-disabled (101/150) is explained, not swallowed',
    /101[\s\S]{0,40}150/.test(noComments), true);
  t('the dim state never reaches zero opacity',
    /hqv-dim[\s\S]{0,120}opacity:\.18/.test(noComments), true);
  t('remove is bound before item in the chip delegate',
    noComments.indexOf('data-x') < noComments.lastIndexOf('data-i'), true);
  t('dvh is used, not bare vh, for the page height',
    /height:100dvh/.test(noComments), true);
}

console.log('\n' + (fail ? '❌ ' : '✅ ') + pass + ' passed · ' + fail + ' failed');
process.exit(fail ? 1 : 0);
