// test-telegram-finder.js — Telegram: plain messages in a PRIVATE chat run the Parts
// Finder; /pull carries a note; the Pull note can never become a formula. (2026-10-01)
// Loads the REAL TelegramCommands.js + MpnSearch.js + ZohoPull.js with Apps Script stubbed.
//
//   node test-telegram-finder.js        SRC=/other/dir overrides where the files are read
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');

let pass = 0, fail = 0;
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want);
  if (g === w) pass++; else { fail++; console.log('✗ ' + name + '\n    got  ' + g + '\n    want ' + w); }
}

const SENT = [], EDITS = [], APPLIED = [], CACHE = {};
const ctx = {
  console, Map, Date, Math, JSON, String, Number, Array, Object, RegExp, isFinite, isNaN, parseFloat, parseInt,
  TELEGRAM_ADMIN_CHAT_ID: '-100',
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k === 'TELEGRAM_COMMAND_CHATS' ? '-200' : '') }) },
  CacheService: { getScriptCache: () => ({ put: (k, v) => { CACHE[k] = v; }, get: k => (k in CACHE ? CACHE[k] : null), remove: k => { delete CACHE[k]; } }) },
  Utilities: { getUuid: () => '1234abcd-5678-90ef-aaaa-bbbbccccdddd' },
  Logger: { log() {} }
};
vm.createContext(ctx);
for (const f of ['MpnSearch.js', 'TelegramCommands.js', 'ZohoPull.js']) vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), ctx);

// stubs AFTER loading (the files define some of these names themselves)
// Stub ONLY the wire (_tgApi) — the real _tgSend / _tgEdit / _tgKeyboard run under test.
ctx._tgApi = (method, payload) => {
  const kb = payload.reply_markup && payload.reply_markup.inline_keyboard;
  if (method === 'sendMessage') SENT.push({ chatId: payload.chat_id, text: payload.text, buttons: kb });
  if (method === 'editMessageText') EDITS.push({ text: payload.text, buttons: kb });
  return { ok: true };
};
ctx.listTelegramWebAppUsers = () => ['777'];
ctx.getPartBasics = q => ({ ok: true, basics: { found: q === '166527', zohoAvailable: q === '300001' ? 2 : null } });
ctx._tgFormatPart = q => 'PART ' + q;
ctx.searchMpns = (t, o) => ({ ok: true, results: [{ query: t, matches: [] }], found: 0, missing: 1, who: o.who });
ctx.searchKeywords = (t) => ({ ok: true, query: t, total: 0, matches: [] });
ctx.computeZohoSoDiff = q => ({ ok: true, soNumber: 'SO-24609', customerName: 'Representaciones Dtz', totalFormatted: '$895.98', isFirstPull: true,
  summary: { totalLines: 2, new: 2 }, lines: [{ sku: '166527', zohoQty: 1, location: 'E-84', name: 'Piston' }, { sku: '173817', zohoQty: 2, location: 'E-54', name: 'Piston' }] });
ctx.applyZohoPullSelection = (q, sel, note) => { APPLIED.push({ q, n: sel.length, note }); return { ok: true, soNumber: 'SO-24609', applied: { inserted: sel.length }, skipped: [] }; };

const msg = (text, chat, from) => ({ message: { text, chat, from: from || { id: 777, first_name: 'Yassin' } } });
const PRIV = { id: 777, type: 'private' }, GROUP = { id: -100, type: 'group' };
const send = u => { SENT.length = 0; const r = ctx.handleTelegramCommand(u); return { r, sent: SENT.slice() }; };

// --- A · plain messages: private only, allowlisted only ------------------------------
let o = send(msg('04270701 04179234', GROUP));
eq('A1 a plain message in a GROUP is ignored (no reply)', [o.r.handled, o.sent.length], [false, 0]);
o = send(msg('04270701 04179234', PRIV));
eq('A2 private + allowlisted user → answered', [o.r.handled, o.r.command, o.sent.length], [true, '(auto)', 1]);
eq('A3 part numbers → the /find answer, then says how it searched',
   [/not on any listing/.test(o.sent[0].text), /searched as part numbers — start with \/search/.test(o.sent[0].text)], [true, true]);
o = send(msg('04270701', { id: 999, type: 'private' }, { id: 999, first_name: 'Stranger' }));
eq('A4 a stranger in a private chat gets nothing', [o.r.handled, o.sent.length], [false, 0]);
o = send(msg('04270701', { id: -200, type: 'private' }, { id: -200 }));
eq('A5 ...but a chat on the command allowlist is admitted', o.r.handled, true);
o = send(msg('166527', PRIV));
eq('A6 a SKU of ours → the /part card', o.sent[0].text, 'PART 166527');
o = send(msg('300001', PRIV));
eq('A7 a Zoho-only SKU still counts as ours', o.sent[0].text, 'PART 300001');
o = send(msg('412345', PRIV));
eq('A8 a 6-digit number that is not ours → searched as a part number, and says so',
   [/^↪ 412345 is not one of our SKUs/.test(o.sent[0].text), /not on any listing/.test(o.sent[0].text)], [true, true]);
o = send(msg('v2203 piston', PRIV));
eq('A9 words → the /search answer + the tip', [/Nothing on any listing contains all of those words/.test(o.sent[0].text),
   /searched as words — start with \/find/.test(o.sent[0].text)], [true, true]);
o = send(msg('', PRIV));
eq('A10 an empty message (a sticker, a photo with no caption) is ignored', [o.r.handled, o.sent.length], [false, 0]);
o = send(msg('/part 166527', GROUP));
eq('A11 commands in the group still work as before', [o.r.handled, o.sent[0].text], [true, 'PART 166527']);
eq('A12 /help mentions plain messages', /PRIVATE chat with me, just send a SKU/.test(ctx.TG_ROUTES['/help'].run('')), true);

// --- B · /pull with a note ------------------------------------------------------------
eq('B1 parse: order + note', ctx._tgParsePullArgs('SO-24609 note hold for payment'), { query: 'SO-24609', note: 'hold for payment' });
eq('B2 parse: no note', ctx._tgParsePullArgs('INV-022496'), { query: 'INV-022496', note: '' });
eq('B3 parse: NOTE in any case', ctx._tgParsePullArgs('SO-1 NOTE Call first').note, 'Call first');
eq('B4 parse: a note longer than 200 is cut', ctx._tgParsePullArgs('SO-1 note ' + 'x'.repeat(300)).note.length, 200);

o = send(msg('/pull SO-24609 note hold for payment', GROUP));
const card = o.sent[0];
eq('B5 the card shows the note', /📝 Note on every row: hold for payment/.test(card.text), true);
const data = card.buttons[0][0].callback_data;
eq('B6 the button carries the order + a short token', data, 'HQ:pull:SO-24609:1234abcd56');
eq('B7 ...well inside Telegram\'s 64-byte limit', Buffer.byteLength(data) <= 64, true);
eq('B8 the note waits in the cache under that token', CACHE['pn:1234abcd56'], 'hold for payment');

EDITS.length = 0; APPLIED.length = 0;
ctx.handleTelegramCommand({ callback_query: { id: 'c1', data, message: { chat: { id: -100 }, message_id: 5 } } });
eq('B9 the tap pulls every line WITH the note', APPLIED, [{ q: 'SO-24609', n: 2, note: 'hold for payment' }]);
eq('B10 the result says so', /📝 Note: hold for payment/.test(EDITS[0].text), true);

delete CACHE['pn:1234abcd56']; EDITS.length = 0; APPLIED.length = 0;
ctx.handleTelegramCommand({ callback_query: { id: 'c2', data, message: { chat: { id: -100 }, message_id: 5 } } });
eq('B11 an expired note → refuses, nothing pulled', [APPLIED.length, /note on this card has expired/.test(EDITS[0].text)], [0, true]);

o = send(msg('/pull SO-24609', GROUP));
eq('B12 no note → the button is exactly as before', o.sent[0].buttons[0][0].callback_data, 'HQ:pull:SO-24609');
EDITS.length = 0; APPLIED.length = 0;
ctx.handleTelegramCommand({ callback_query: { id: 'c3', data: 'HQ:pull:SO-24609', message: { chat: { id: -100 }, message_id: 6 } } });
eq('B13 ...and pulls with an empty note', APPLIED, [{ q: 'SO-24609', n: 2, note: '' }]);
eq('B14 /help shows the note form', /\/pull <SO or INV> \[note <text>\]/.test(ctx.TG_ROUTES['/help'].run('')), true);

// --- C · the Pull note can never become a formula -------------------------------------
eq('C1 plain note unchanged', ctx._pullNoteCell('hold for payment'), 'hold for payment');
eq('C2 = gets the text marker', ctx._pullNoteCell('=HYPERLINK("x")'), "'=HYPERLINK(\"x\")");
eq('C3 + - @ too', ['+1 call', '-2 short', '@john'].map(ctx._pullNoteCell), ["'+1 call", "'-2 short", "'@john"]);
eq('C4 blank stays blank', ctx._pullNoteCell('  '), '');

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
