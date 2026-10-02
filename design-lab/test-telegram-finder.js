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
eq('B4 parse: a whole order note longer than 200 is cut', ctx._tgParsePullNotes('x'.repeat(300), []).all.length, 200);
eq('B4b parse: newlines survive into the note', ctx._tgParsePullArgs('SO-1 note hold\n166527: this one').note, 'hold\n166527: this one');

const SKUS = ['166527', '173817', '000000'];
const pn = t => ctx._tgParsePullNotes(t, SKUS);
eq('N1 plain line → every row', [pn('call before shipping').all, Object.keys(pn('call before shipping').sku).length], ['call before shipping', 0]);
eq('N2 SKU: text → that line', pn('166527: hold this one').sku, { '166527': 'hold this one' });
eq('N3 both, several lines', [pn('call first\n173817: fragile').all, pn('call first\n173817: fragile').sku], ['call first', { '173817': 'fragile' }]);
eq('N4 a time is NOT a SKU', pn('Pickup 3:30 PM').all, 'Pickup 3:30 PM');
eq('N5 a word: is NOT a SKU', [pn('Urgent: call the customer').all, pn('Urgent: call the customer').errors.length], ['Urgent: call the customer', 0]);
eq('N6 a SKU not on the order → error, never a guess', pn('999999: hold').errors, ['999999 is not one of the lines being pulled']);
eq('N7 leading zeros / case still match', [pn('0: x').errors.length, Object.keys(pn('000000: special pkg').sku)], [0, ['000000']]);
eq('N8 clear', pn('clear').clear, true);
eq('N9 "SKU:" with nothing → removes that line\'s note', pn('166527:').sku, { '166527': '' });
eq('N10 merge keeps the order note when only SKU lines arrive',
   ctx._tgPullMergeNotes({ all: 'A', sku: { '166527': 'x' } }, pn('173817: y')), { all: 'A', sku: { '166527': 'x', '173817': 'y' } });

o = send(msg('/pull SO-24609 note hold for payment', GROUP));
const card = o.sent[0];
eq('B5 the card shows the note', /📝 Every row: hold for payment/.test(card.text), true);
eq('B5b the card has a 📝 Note button', card.buttons[0][0], { text: '📝 Note', callback_data: 'HQ:pnote:1234abcd56' });
const data = card.buttons[1][0].callback_data;
eq('B6 the button carries the order + a short token', data, 'HQ:pull:SO-24609:1234abcd56');
eq('B7 ...well inside Telegram\'s 64-byte limit', Buffer.byteLength(data) <= 64, true);
eq('B8 the note waits in the cache under that token', JSON.parse(CACHE['pn:1234abcd56']).all, 'hold for payment');

EDITS.length = 0; APPLIED.length = 0;
ctx.handleTelegramCommand({ callback_query: { id: 'c1', data, message: { chat: { id: -100 }, message_id: 5 } } });
eq('B9 the tap pulls every line WITH the note', APPLIED, [{ q: 'SO-24609', n: 2, note: 'hold for payment' }]);
eq('B10 the result says so', /📝 Every row: hold for payment/.test(EDITS[0].text), true);

delete CACHE['pn:1234abcd56']; EDITS.length = 0; APPLIED.length = 0;
ctx.handleTelegramCommand({ callback_query: { id: 'c2', data, message: { chat: { id: -100 }, message_id: 5 } } });
eq('B11 an expired card → refuses, nothing pulled', [APPLIED.length, /card has expired/.test(EDITS[0].text)], [0, true]);

CACHE['pn:oldoldold0'] = 'legacy note';
EDITS.length = 0; APPLIED.length = 0;
ctx.handleTelegramCommand({ callback_query: { id: 'c2b', data: 'HQ:pull:SO-24609:oldoldold0', message: { chat: { id: -100 }, message_id: 5 } } });
eq('B11b a card drawn before today (bare note string) still pulls with its note', APPLIED, [{ q: 'SO-24609', n: 2, note: 'legacy note' }]);

EDITS.length = 0; APPLIED.length = 0;
ctx.handleTelegramCommand({ callback_query: { id: 'c3', data: 'HQ:pull:SO-24609', message: { chat: { id: -100 }, message_id: 6 } } });
eq('B13 an old token-less button still pulls with an empty note', APPLIED, [{ q: 'SO-24609', n: 2, note: '' }]);
eq('B14 /help shows the note form', /\/pull <SO or INV> \[note <text>\]/.test(ctx.TG_ROUTES['/help'].run('')), true);

// --- B · typed per-SKU note + the wrong SKU --------------------------------------------
o = send(msg('/pull SO-24609 note call first\n173817: fragile', GROUP));
eq('B15 typed: order note + SKU note on the card', [/📝 Every row: call first/.test(o.sent[0].text), /📝 173817: fragile/.test(o.sent[0].text)], [true, true]);
o = send(msg('/pull SO-24609 note 999999: hold', GROUP));
eq('B16 typed: SKU not on the order → no card, says why', [!!o.sent[0].buttons, /999999 is not one of the lines being pulled/.test(o.sent[0].text)], [false, true]);

// --- R · the 📝 Note button → reply → card redrawn → pull -----------------------------
for (const k of Object.keys(CACHE)) delete CACHE[k];
o = send(msg('/pull SO-24609', GROUP));
eq('R1 no note → still a Note button', o.sent[0].buttons[0][0].callback_data, 'HQ:pnote:1234abcd56');
const RAW = [];
const realApi = ctx._tgApi;
ctx._tgApi = (m, p) => { RAW.push({ m, p }); return realApi(m, p); };
EDITS.length = 0; SENT.length = 0;
ctx.handleTelegramCommand({ callback_query: { id: 'c4', data: 'HQ:pnote:1234abcd56', message: { chat: { id: -100 }, message_id: 50 } } });
const prompt = RAW.find(x => x.m === 'sendMessage');
eq('R2 the tap posts a force-reply prompt, replying to the card', [!!prompt, prompt.p.reply_markup.force_reply, prompt.p.reply_to_message_id], [true, true, 50]);
eq('R3 ...ending with the ref token', /ref pn1234abcd56$/.test(prompt.p.text), true);
eq('R4 ...and the card is left alone (buttons kept)', EDITS.length, 0);
eq('R5 placeholder fits Telegram\'s 64 chars', prompt.p.reply_markup.input_field_placeholder.length <= 64, true);

const reply = (text, chat) => ({ message: { text, chat: chat || GROUP, from: { id: 777 },
  reply_to_message: { from: { is_bot: true }, text: prompt.p.text } } });
EDITS.length = 0;
o = send(reply('call before shipping\n166527: hold this one'));
eq('R6 the reply is handled (in a GROUP, though it has no /)', [o.r.handled, o.r.command], [true, '(pull note)']);
eq('R7 the card is redrawn WITH its buttons and the notes',
   [EDITS.length, /📝 Every row: call before shipping/.test(EDITS[0].text), /📝 166527: hold this one/.test(EDITS[0].text), EDITS[0].buttons.length], [1, true, true, 2]);
eq('R8 the chat gets a one-line ack', /Notes set on the SO-24609 card/.test(o.sent[0].text), true);

EDITS.length = 0;
o = send(reply('999999: nope'));
eq('R9 a SKU not on the order → nothing changed, card untouched', [/Nothing changed/.test(o.sent[0].text), EDITS.length], [true, 0]);
o = send(reply('hello', { id: -999, type: 'group' }));
eq('R10 a reply from a chat not allowlisted is ignored', [o.r.handled, o.sent.length], [false, 0]);
o = send({ message: { text: 'call me', chat: GROUP, from: { id: 777 }, reply_to_message: { from: { is_bot: false }, text: prompt.p.text } } });
eq('R11 a reply to a HUMAN quoting the ref is ignored', o.r.handled, false);

APPLIED.length = 0;
let SELS = null;
ctx.applyZohoPullSelection = (q, sel, note) => { SELS = sel; APPLIED.push({ q, n: sel.length, note }); return { ok: true, soNumber: 'SO-24609', applied: { inserted: sel.length }, skipped: [] }; };
ctx.handleTelegramCommand({ callback_query: { id: 'c5', data: 'HQ:pull:SO-24609:1234abcd56', message: { chat: { id: -100 }, message_id: 50 } } });
eq('R12 pull: order note to all, SKU note to its line only', [APPLIED[0].note, SELS], ['call before shipping',
   [{ sku: '166527', action: 'insert', note: 'hold this one' }, { sku: '173817', action: 'insert' }]]);

o = send(reply('clear'));
eq('R13 clear removes every note', [JSON.parse(CACHE['pn:1234abcd56']).all, JSON.parse(CACHE['pn:1234abcd56']).sku], ['', {}]);
delete CACHE['pn:1234abcd56'];
o = send(reply('late note'));
eq('R14 a reply after the card expired says so', /expired/.test(o.sent[0].text), true);
ctx._tgApi = realApi;


// --- A2 · an order already on the sheet: NEW LINES ONLY → Telegram may add them -------
for (const k of Object.keys(CACHE)) delete CACHE[k];
const L3 = [{ sku: '166527', zohoQty: 1, location: 'E-84', name: 'Piston', status: 'unchanged' },
            { sku: '173817', zohoQty: 2, location: 'E-54', name: 'Piston', status: 'unchanged' },
            { sku: '200001', zohoQty: 3, location: 'B-7', name: 'Gasket', status: 'new' }];
const repull = (lines, sum) => () => ({ ok: true, soNumber: 'SO-26018', customerName: 'Miguel', totalFormatted: '$1,044.68', isFirstPull: false,
  summary: Object.assign({ totalLines: lines.length, unchanged: 0, new: 0, qtyChanged: 0, removed: 0, anyChanges: true }, sum), lines });
ctx.computeZohoSoDiff = repull(L3, { unchanged: 2, new: 1 });
eq('M1 mode: first / add / modal / nothing', [
  ctx._tgPullMode({ isFirstPull: true, summary: { totalLines: 2, new: 2 } }),
  ctx._tgPullMode({ isFirstPull: false, summary: { new: 1, unchanged: 2 } }),
  ctx._tgPullMode({ isFirstPull: false, summary: { new: 1, qtyChanged: 1 } }),
  ctx._tgPullMode({ isFirstPull: false, summary: { new: 0, unchanged: 3 } })], ['first', 'add', null, null]);
o = send(msg('/pull SO-26018 note 200001: rush', GROUP));
eq('M2 card lists ONLY the new line + says the rest stay', [/200001/.test(o.sent[0].text), /166527/.test(o.sent[0].text), /1 new line to add · 2 already on the sheet/.test(o.sent[0].text)], [true, false, true]);
eq('M3 button reads Add 1 new', o.sent[0].buttons[1][0].text, '✅ Add 1 new');
o = send(msg('/pull SO-26018 note 166527: x', GROUP));
eq('M4 a note on a line already there is refused', /166527 is not one of the lines being pulled/.test(o.sent[0].text), true);
o = send(msg('/pull SO-26018 note 200001: rush', GROUP));
APPLIED.length = 0; EDITS.length = 0; SELS = null;
ctx.handleTelegramCommand({ callback_query: { id: 'm5', data: o.sent[0].buttons[1][0].callback_data, message: { chat: { id: -100 }, message_id: 70 } } });
eq('M5 tap adds ONLY the new line, with its note', SELS, [{ sku: '200001', action: 'insert', note: 'rush' }]);
eq('M6 result says ADDED and that the rest were untouched', [/✅ ADDED · SO-24609/.test(EDITS[0].text), /2 already there, untouched/.test(EDITS[0].text)], [true, true]);

o = send(msg('/pull SO-26018', GROUP));
const addData = o.sent[0].buttons[1][0].callback_data;
ctx.computeZohoSoDiff = repull(L3.concat([{ sku: '200002', zohoQty: 1, location: 'C-1', name: 'Seal', status: 'new' }]), { unchanged: 2, new: 2 });
APPLIED.length = 0; EDITS.length = 0;
ctx.handleTelegramCommand({ callback_query: { id: 'm7', data: addData, message: { chat: { id: -100 }, message_id: 71 } } });
eq('M7 a line that appeared in Zoho after the card → nothing pulled unseen', [APPLIED.length, /order changed since this card/.test(EDITS[0].text)], [0, true]);

ctx.computeZohoSoDiff = repull(L3.map(l => Object.assign({}, l, { status: l.status === 'new' ? 'qty_changed' : l.status })), { unchanged: 2, qtyChanged: 1 });
o = send(msg('/pull SO-26018', GROUP));
eq('M8 a qty change still needs the modal', [/Needs the Pull modal — 1 qty change/.test(o.sent[0].text), !!o.sent[0].buttons], [true, false]);
ctx.computeZohoSoDiff = repull(L3.slice(0, 2), { unchanged: 2, anyChanges: false });
o = send(msg('/pull SO-26018', GROUP));
eq('M9 nothing new → says so', /Nothing new — all 2 lines are already on the sheet/.test(o.sent[0].text), true);

// --- P · the row note in ZohoPull ------------------------------------------------------
eq('P1 order note + line note', ctx._pullRowNote('call first', 'fragile'), 'call first · fragile');
eq('P2 only one of them', [ctx._pullRowNote('', 'fragile'), ctx._pullRowNote('call first', undefined)], ['fragile', 'call first']);
eq('P3 neither', ctx._pullRowNote(null, ''), '');

// --- C · the Pull note can never become a formula -------------------------------------
eq('C1 plain note unchanged', ctx._pullNoteCell('hold for payment'), 'hold for payment');
eq('C2 = gets the text marker', ctx._pullNoteCell('=HYPERLINK("x")'), "'=HYPERLINK(\"x\")");
eq('C3 + - @ too', ['+1 call', '-2 short', '@john'].map(ctx._pullNoteCell), ["'+1 call", "'-2 short", "'@john"]);
eq('C4 blank stays blank', ctx._pullNoteCell('  '), '');

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
