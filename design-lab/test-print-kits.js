/**
 * test-print-kits.js — K chips on the printed pick list (OrderLook part 3, 2026-10-01).
 *
 * K1 · the REAL preparePrintSheet hands each printed line its K tag, numbered over the
 *      WHOLE sheet — so a part is tagged even when its parent line is still PENDING and
 *      therefore not on the paper. Switch off → no tags at all.
 * K2 · the REAL PrintFulfillment.html draws them: outlined on parts, solid on a parent,
 *      in all three tables; no tag → no chip markup.
 * K3 · rendered in Chromium: rows with chips are EXACTLY as tall as rows without, so the
 *      hand-tuned page capacities (16 / 21 rows) stay true. Writes a PNG to renders/.
 *
 * Usage: node test-print-kits.js    (SRC=/old/tree to run against the previous code)
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');
let pass = 0, fail = 0;
const ok = (l, c, d) => { c ? pass++ : fail++; console.log((c ? '  ✓ ' : '  ✗ ') + l + (c ? '' : '  → ' + JSON.stringify(d))); };

// ---- K1: the real server code ----------------------------------------------------------
console.log('K1 · preparePrintSheet');
const rows = [
  ['HQ'], [''], ['◈ SKU'],
  ['111111', 1, 'B-1', '05-1', '', 'PREPARING', 3],
  ['DIRECT'], ['◈ SKU'],
  ['157644', 1, 'D-41', 'SO-25792', 'Miguel', 'PENDING', 2],                      // parent K1 — NOT printed
  ['164979', 1, 'C-42', 'SO-25792', '↳ from KIT-157644 · Miguel', 'PREPARING', 237],
  ['164979', 2, 'C-42', 'SO-25792', 'Miguel', 'PREPARING', 237],                   // loose, same SKU
  ['217205', 1, 'D-42', 'SO-25792', '', 'PREPARING', 20],                          // parent K2 — printed
  ['181138', 3, 'E-23', 'SO-25792', '↳ from KIT-217205', 'PREPARING', 20],
  ['163962', 4, 'E-15', 'SO-25792', '↳ from KIT-157644 · Miguel', 'PREPARING', 23]
].map(r => { const o = r.slice(); while (o.length < 10) o.push(''); return o; });
const props = {};
function runPrint() {
  let captured = null;
  const sheet = {
    getRange: a => ({ getValue: () => (a === 'F2' ? 'Shipping - Yassin 1' : ''), getNumberFormats: () => rows.map(() => ['@']) }),
    getDataRange: () => ({ getValues: () => rows })
  };
  const sb = {
    console: { log() {} }, SPREADSHEET_ID: 'x', MAIN_SHEET_NAME: 'All orders', Set, Map,
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => sheet }), getUi: () => ({ showModalDialog() {} }) },
    Utilities: { formatDate: () => '1' },
    HtmlService: { createTemplateFromFile: () => { const t = {}; t.evaluate = () => { captured = t;
      return { setTitle() { return this; }, setWidth() { return this; }, setHeight() { return this; }, getContent: () => '' }; }; return t; } },
    _buildDirectCustomerMap: () => ({}), isPrintPaidShippingAlertsEnabled: () => false,
    _extractPickIdData: v => v, logActivity: () => {},
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null }) },
    buildKitMap: () => new Map([['157644', {}], ['217205', {}]])
  };
  vm.createContext(sb);
  vm.runInContext(fs.readFileSync(path.join(SRC, 'Schema.js'), 'utf8'), sb);
  sb.Schema.pickIdA1 = () => 'F2';
  // Helpers.js is large; take only the shared note parser from it, verbatim.
  const helpers = fs.readFileSync(path.join(SRC, 'Helpers.js'), 'utf8');
  const kt = helpers.slice(helpers.indexOf('function kitComponentTag'));
  vm.runInContext(kt.slice(0, kt.indexOf('\n}\n') + 3), sb);
  if (fs.existsSync(path.join(SRC, 'OrderLook.js'))) vm.runInContext(fs.readFileSync(path.join(SRC, 'OrderLook.js'), 'utf8'), sb);
  vm.runInContext(fs.readFileSync(path.join(SRC, 'FulfillmentService.js'), 'utf8'), sb);
  sb.preparePrintSheet({ forBoard: true });
  return captured;
}
props.ORDER_LOOK_ON = 'on';
let cap = runPrint();
const tags = (cap.directItems || []).map(x => x[0] + ':' + (x[10] || '-'));
ok('parts of a PENDING (unprinted) parent still carry its number', tags.indexOf('164979:C1') !== -1 && tags.indexOf('163962:C1') !== -1, tags);
ok('a printed parent is solid K2, its part K2', tags.indexOf('217205:P2') !== -1 && tags.indexOf('181138:C2') !== -1, tags);
ok('the loose line with the same SKU has no tag', tags.filter(t => t === '164979:-').length === 1, tags);
ok('the eBay line has no tag', (cap.ebayItems || []).every(x => !x[10]), (cap.ebayItems || []).map(x => x[10]));
delete props.ORDER_LOOK_ON;
cap = runPrint();
ok('switch OFF → no tags on paper', (cap.directItems || []).every(x => !x[10]), (cap.directItems || []).map(x => x[10]));

// ---- K2: the real template -----------------------------------------------------------
console.log('K2 · PrintFulfillment.html');
function compile(tpl) {
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  let js = 'var __o = [];\n', i = 0; const re = /<\?(!=|=)?([\s\S]*?)\?>/g; let m;
  while ((m = re.exec(tpl))) {
    js += '__o.push(' + JSON.stringify(tpl.slice(i, m.index)) + ');\n';
    if (m[1] === '=') js += '__o.push(__esc(' + m[2] + '));\n'; else if (m[1] === '!=') js += '__o.push(String(' + m[2] + '));\n'; else js += m[2] + '\n';
    i = re.lastIndex;
  }
  js += '__o.push(' + JSON.stringify(tpl.slice(i)) + ');\nreturn __o.join("");';
  return vars => { const n = Object.keys(vars); return new Function('__esc', ...n, js)(esc, ...n.map(k => vars[k])); };
}
const render = compile(fs.readFileSync(path.join(SRC, 'PrintFulfillment.html'), 'utf8'));
const it = (sku, tag, n) => [sku, 1, 'C-' + n, 'SO-25792', '', 30, '', '', '', '', tag];
const directItems = [];
for (let n = 0; n < 12; n++) directItems.push(it(String(160000 + n), n % 3 === 0 ? '' : (n % 4 === 0 ? 'P' : 'C') + (1 + n % 3), n));
const vars = { ebayItems: [it('111111', 'C1', 1)], directItems, amazonItems: [it('222222', 'P3', 2)],
  directCustomers: { 'SO-25792': 'Miguel' }, employeeId: 'Y', pickIdShipping: 'Y', pickIdAdjustment: '',
  printDate: '10/1/2026', printTime: '1:00 AM', printDay: 'Thursday', printDateLong: 'Thursday, October 1, 2026',
  docNumber: 'FUL · 10/01 · 01:00', estimatedPages: 2, showPaidShippingAlerts: false,
  metrics: { totalItems: 14, totalQty: 14, ebayCount: 1, directCount: 12, amazonCount: 1, distinctSkus: 14, lowStock: 0, paidShippingCount: 0, paidShippingTotal: '0.00' } };
let html = '';
try { html = render(vars); ok('template renders', true); } catch (e) { ok('template renders', false, e.message); }
const chips = (html.match(/class="kchip( par)?"/g) || []);
ok('a chip on every tagged line in all three tables', chips.length === 10, chips.length);
ok('parent chips are solid and read "▣ K"', (html.match(/class="kchip par">▣ K\d/g) || []).length === 3);
ok('part chips read "K"', (html.match(/class="kchip">K\d/g) || []).length === 7);
const plain = render(Object.assign({}, vars, { directItems: directItems.map(x => x.slice(0, 10)), ebayItems: [it('111111', '', 1)], amazonItems: [it('222222', '', 2)] }));
ok('no tags → no chip markup at all', plain.indexOf('class="kchip') === -1);
fs.mkdirSync(path.join(__dirname, 'renders'), { recursive: true });
fs.writeFileSync(path.join(__dirname, 'renders', 'print-kits.html'), html);

// ---- K3: rendered row heights ---------------------------------------------------------
(async () => {
  console.log('K3 · rendered in Chromium');
  let chromium; try { ({ chromium } = require('playwright')); } catch (e) { console.log('  (playwright missing — skipped)'); return done(); }
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 900, height: 1200 } });
  await p.setContent(html, { waitUntil: 'load' });
  await p.emulateMedia({ media: 'print' });
  const h = await p.$$eval('#direct-table tbody tr:not(.order-head)', trs => trs.filter(t => t.querySelector('.c-sku') && t.querySelector('.c-sku').textContent.trim())
    .map(t => ({ chip: !!t.querySelector('.kchip'), h: Math.round(t.getBoundingClientRect().height) })));
  const withC = [...new Set(h.filter(x => x.chip).map(x => x.h))], without = [...new Set(h.filter(x => !x.chip).map(x => x.h))];
  ok('rows with a chip are exactly as tall as rows without', withC.length === 1 && without.length === 1 && withC[0] === without[0], { withC, without });
  const el = await p.$('#direct-table');
  if (el) await el.screenshot({ path: path.join(__dirname, 'renders', 'print-kits.png') });
  await b.close(); done();
})();
function done() { console.log('\n' + (fail ? '✗ ' + fail + ' FAILED · ' : '✓ all · ') + pass + ' passed'); process.exit(fail ? 1 : 0); }
