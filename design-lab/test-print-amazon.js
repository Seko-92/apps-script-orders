/**
 * test-print-amazon.js — the printed pick list gains an AMAZON section (2026-09-26).
 *
 * Renders the REAL PrintFulfillment.html through a small emulation of Apps Script's
 * template engine (<? code ?> · <?= escaped ?> · <?!= raw ?>), then checks the sections,
 * the page breaks between them, and the counts. And the section SPLIT itself is checked
 * against the REAL preparePrintSheet loop, including the substring bug it replaced: an
 * eBay row whose NOTE mentions "DIRECT" used to push every row below it into Direct.
 *
 * Writes the rendered page to design-lab/renders/print-amazon.html for a visual look.
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');

let pass = 0, fail = 0;
const ok = (label, cond, detail) => {
  cond ? pass++ : fail++;
  console.log((cond ? '  ✓ ' : '  ✗ ') + label + (cond ? '' : '  → ' + JSON.stringify(detail)));
};

// ---- a minimal HtmlService template compiler -------------------------------------------
function compile(tpl) {
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  let js = 'var __o = [];\n', i = 0;
  const re = /<\?(!=|=)?([\s\S]*?)\?>/g; let m;
  while ((m = re.exec(tpl))) {
    js += '__o.push(' + JSON.stringify(tpl.slice(i, m.index)) + ');\n';
    if (m[1] === '=') js += '__o.push(__esc(' + m[2] + '));\n';
    else if (m[1] === '!=') js += '__o.push(String(' + m[2] + '));\n';
    else js += m[2] + '\n';
    i = re.lastIndex;
  }
  js += '__o.push(' + JSON.stringify(tpl.slice(i)) + ');\nreturn __o.join("");';
  return (vars) => {
    const names = Object.keys(vars);
    // eslint-disable-next-line no-new-func
    return new Function('__esc', ...names, js)(esc, ...names.map(n => vars[n]));
  };
}

const tpl = fs.readFileSync(path.join(SRC, 'PrintFulfillment.html'), 'utf8');
let render;
try { render = compile(tpl); ok('P0 the template compiles', true); }
catch (e) { ok('P0 the template compiles', false, e.message); process.exit(1); }

const item = (sku, so, note) => [sku, 1, 'A-1', so, note || '', 5, '', '', '', ''];
const base = {
  ebayItems: [item('111111', '05-11111-11111')],
  directItems: [item('444444', 'SO-100'), item('555555', 'SO-100')],
  amazonItems: [item('167517', 'AMZ-114-3941689-8772232', 'ship by 9/29')],
  directCustomers: { 'SO-100': 'Acme' },
  employeeId: 'Yassin · 1', pickIdShipping: 'Yassin · 1', pickIdAdjustment: '',
  printDate: '9/26/2026', printTime: '9:00 PM', printDay: 'Saturday',
  printDateLong: 'Saturday, September 26, 2026', docNumber: 'FUL · 09/26 · 21:00',
  estimatedPages: 3, showPaidShippingAlerts: false,
  metrics: { totalItems: 4, totalQty: 4, ebayCount: 1, directCount: 2, amazonCount: 1,
             distinctSkus: 4, lowStock: 4, paidShippingCount: 0, paidShippingTotal: '0.00' }
};

let html;
try { html = render(base); ok('P1 renders with all three sections', true); }
catch (e) { ok('P1 renders with all three sections', false, e.message); process.exit(1); }
fs.mkdirSync(path.join(__dirname, 'renders'), { recursive: true });
fs.writeFileSync(path.join(__dirname, 'renders', 'print-amazon.html'), html);

ok('P2 an Amazon table exists', /id="amazon-table"/.test(html));
ok('P3 its section head says Amazon and names the Seller Central step',
   /Amazon Orders<\/span>[\s\S]{0,200}Confirm shipment in Seller Central/.test(html));
ok('P4 the Amazon order gets its own order band', /▌ AMZ-114-3941689-8772232/.test(html));
ok('P5 Amazon rows are not in the Direct table',
   (html.split('id="direct-table"')[1] || '').split('id="amazon-table"')[0].indexOf('167517') === -1);
ok('P6 section order eBay → Direct → Amazon',
   html.indexOf('id="ebay-table"') < html.indexOf('id="direct-table"') &&
   html.indexOf('id="direct-table"') < html.indexOf('id="amazon-table"'));
const breaks = (html.match(/page-break-after:always/g) || []).length;
ok('P7 a page break between each pair of sections (2 for three sections)', breaks === 2, breaks);
ok('P8 the header total counts Amazon', /4 items<\/span>/.test(html));
ok('P9 the closing summary has an Amazon metric', /metric-num">1<\/div>\s*<div class="metric-lbl">Amazon/.test(html));
ok('P10 the page-fill script knows the Amazon table', /fillTableToPage\('amazon-table'/.test(html));

// Amazon only → no stray break, fills as the first section
const onlyAmz = render(Object.assign({}, base, { ebayItems: [], directItems: [],
  metrics: Object.assign({}, base.metrics, { ebayCount: 0, directCount: 0 }) }));
ok('P11 Amazon-only print has no page breaks', !/page-break-after:always/.test(onlyAmz));
// no Amazon → exactly today's output shape (no amazon table, one break)
const noAmz = render(Object.assign({}, base, { amazonItems: [],
  metrics: Object.assign({}, base.metrics, { amazonCount: 0 }) }));
ok('P12 without Amazon rows there is no Amazon table', !/id="amazon-table"/.test(noAmz));
ok('P13 …and no Amazon metric card', !/metric-lbl">Amazon/.test(noAmz));
ok('P14 …and still exactly one break (eBay → Direct), as before', (noAmz.match(/page-break-after:always/g) || []).length === 1);

// ---- the section SPLIT in preparePrintSheet, run for real ----------------------------------
{
  const code = fs.readFileSync(path.join(SRC, 'FulfillmentService.js'), 'utf8');
  const rows = [
    ['HQ'], [''], ['◈ SKU'],
    ['111111', 1, 'B-1', '05-1', 'customer asked for DIRECT shipping', 'PREPARING', 3, '', '', ''],
    ['222222', 1, 'B-2', '05-2', '', 'PREPARING', 3, '', '', ''],
    ['DIRECT'], ['◈ SKU'],
    ['444444', 1, 'D-4', 'SO-100', '', 'PREPARING', 5, '', '', ''],
    ['AMAZON'], ['◈ SKU'],
    ['167517', 1, 'K-7', 'AMZ-114-3941689-8772232', 'AMAZON', 'PREPARING', 12, '', '', '']
  ].map(r => { const o = r.slice(); while (o.length < 10) o.push(''); return o; });
  let captured = null;
  const sheet = {
    getRange: (a) => ({
      getValue: () => (a === 'F2' ? 'Shipping - Yassin 1' : ''),
      getNumberFormats: () => rows.map(() => ['@'])
    }),
    getDataRange: () => ({ getValues: () => rows })
  };
  const sb = {
    console: { log() {} }, SPREADSHEET_ID: 'x', MAIN_SHEET_NAME: 'All orders',
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => sheet }),
                      getUi: () => ({ showModalDialog() {} }) },
    Utilities: { formatDate: () => '1' },
    HtmlService: { createTemplateFromFile: () => {
      const t = {}; t.evaluate = () => { captured = t;
        return { setTitle() { return this; }, setWidth() { return this; }, setHeight() { return this; },
                 getContent: () => '' }; }; return t; } },
    _buildDirectCustomerMap: () => ({}), isPrintPaidShippingAlertsEnabled: () => false,
    _extractPickIdData: v => v, logActivity: () => {},
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) }
  };
  vm.createContext(sb);
  vm.runInContext(fs.readFileSync(path.join(SRC, 'Schema.js'), 'utf8'), sb);
  sb.Schema.pickIdA1 = () => 'F2';
  vm.runInContext(code, sb);
  sb.preparePrintSheet({ forBoard: true });
  const skus = a => (a || []).map(x => x[0]);
  ok('P15 ⭐ an eBay row whose NOTE says "DIRECT" stays in eBay (the old substring bug)',
     JSON.stringify(skus(captured.ebayItems)) === '["111111","222222"]', skus(captured.ebayItems));
  ok('P16 Direct gets only its own rows', JSON.stringify(skus(captured.directItems)) === '["444444"]', skus(captured.directItems));
  ok('P17 Amazon gets its own rows', JSON.stringify(skus(captured.amazonItems)) === '["167517"]', skus(captured.amazonItems));
  ok('P18 metrics count Amazon', captured.metrics.amazonCount === 1 && captured.metrics.totalItems === 4, captured.metrics);
}

console.log('\n' + (fail ? '✗ ' + fail + ' FAILED · ' : '✓ all · ') + pass + ' passed');
process.exit(fail ? 1 : 0);
