// ============================================================================
// MOCK TICK for the Floor Board render loop.
// Deliberately exercises EVERY visual feature the board can draw:
//   eBay order bands (2+ lines) · split "+N more" caps · DIRECT accordion
//   (one open, one shut-with-aisles, one single-line) · unexpanded KIT row ·
//   kit threads (2 hues × dash) + strip · shelf-count lane (hand ≤ 25) ·
//   one DEVIANCE row · floor note (**) · NOT FOUND shelf · dual shelf
//   L-208/C-51 · pack-size G-35 * 2 · ×N qty chips · PREP rows · paid strip ·
//   live feed with grouped multi-line orders.
// Shape derived from FloorBoard.html onTick/paint* on 2026-08-11 (v2.1+).
// ============================================================================
'use strict';

const rows = [
  // ── EBAY section (server order: orders anchored at earliest shelf) ──
  { channel: 'EBAY', orderId: '24-15021-77421', sku: '194244', qty: 1, location: 'A-14',      status: 'PENDING',   note: '', isKit: false, hand: 9 },
  // five-line order — the band case, anchored at A-31
  { channel: 'EBAY', orderId: '08-15017-44806', sku: '167517', qty: 1, location: 'A-31',      status: 'PENDING',   note: '', isKit: false },
  { channel: 'EBAY', orderId: '08-15017-44806', sku: '155394', qty: 2, location: 'B-12',      status: 'PENDING',   note: '', isKit: false },
  { channel: 'EBAY', orderId: '08-15017-44806', sku: '168138', qty: 1, location: 'C-44',      status: 'PREPARING', note: '', isKit: false },
  { channel: 'EBAY', orderId: '08-15017-44806', sku: '176701', qty: 4, location: 'F-3',       status: 'PENDING',   note: '', isKit: false },
  { channel: 'EBAY', orderId: '08-15017-44806', sku: '190455', qty: 1, location: 'L-77',      status: 'PENDING',   note: '', isKit: false },
  { channel: 'EBAY', orderId: '24-15008-33107', sku: '165447', qty: 2, location: 'B-30',      status: 'PENDING',   note: '** call before shipping', isKit: false, hand: 4, left: 1 },
  { channel: 'EBAY', orderId: '24-15003-91882', sku: '176688', qty: 8, location: 'C-18',      status: 'PENDING',   note: '', isKit: false },
  { channel: 'EBAY', orderId: '24-15015-60233', sku: '183112', qty: 1, location: 'G-35 * 2',  status: 'PENDING',   note: '', isKit: false, hand: 12 },
  { channel: 'EBAY', orderId: '24-15019-08841', sku: '157023', qty: 1, location: 'L-208/C-51',status: 'PENDING',   note: '', isKit: false },
  { channel: 'EBAY', orderId: '24-15004-11290', sku: '172764', qty: 1, location: 'NOT FOUND', status: 'PENDING',   note: '', isKit: false },

  // ── DIRECT section (SO-grouped, status-ranked) ──
  // SO-24612 — TWO kits in one order → per-row spines (hue 0 solid + hue 1 solid)
  { channel: 'DIRECT', orderId: 'SO-24612', sku: '162198', qty: 4, location: 'A-9',   status: 'PENDING',   note: '↳ from KIT-160029', isKit: false, kit: '160029', hand: 6 },
  { channel: 'DIRECT', orderId: 'SO-24612', sku: '155457', qty: 1, location: 'C-2',   status: 'PENDING',   note: '↳ from KIT-160029', isKit: false, kit: '160029' },
  { channel: 'DIRECT', orderId: 'SO-24612', sku: '158652', qty: 1, location: 'E-17',  status: 'PENDING',   note: '↳ from KIT-158652', isKit: false, kit: '158652' },
  { channel: 'DIRECT', orderId: 'SO-24612', sku: '167409', qty: 2, location: 'E-37',  status: 'PENDING',   note: '↳ from KIT-158652', isKit: false, kit: '158652', hand: 3 },
  { channel: 'DIRECT', orderId: 'SO-24612', sku: '167517', qty: 2, location: 'B-12',  status: 'PREPARING', note: '↳ from KIT-160029', isKit: false, kit: '160029' },
  { channel: 'DIRECT', orderId: 'SO-24612', sku: '159579', qty: 1, location: 'F-20',  status: 'PREPARING', note: '↳ from KIT-158652', isKit: false, kit: '158652' },
  // SO-24618 — will render SHUT (not seeded open; picker hasn't touched it)
  { channel: 'DIRECT', orderId: 'SO-24618', sku: '171003', qty: 1, location: 'A-51',  status: 'PENDING',   note: '↳ from KIT-217475', isKit: false, kit: '217475' },
  { channel: 'DIRECT', orderId: 'SO-24618', sku: '168138', qty: 1, location: 'B-44',  status: 'PENDING',   note: '↳ from KIT-217475', isKit: false, kit: '217475' },
  { channel: 'DIRECT', orderId: 'SO-24618', sku: '196080', qty: 2, location: 'D-7',   status: 'PENDING',   note: '', isKit: false },
  // SO-24631 — plain three-line order, no kit
  { channel: 'DIRECT', orderId: 'SO-24631', sku: '183220', qty: 2, location: 'A-22',  status: 'PENDING',   note: '', isKit: false, hand: 15 },
  { channel: 'DIRECT', orderId: 'SO-24631', sku: '190455', qty: 1, location: 'C-30',  status: 'PENDING',   note: '', isKit: false },
  { channel: 'DIRECT', orderId: 'SO-24631', sku: '196530', qty: 1, location: 'G-10',  status: 'PENDING',   note: '', isKit: false },
  // SO-24620 — unexpanded KIT decision row (READY box at K-55)
  { channel: 'DIRECT', orderId: 'SO-24620', sku: '217871', qty: 1, location: 'K-55',  status: 'PENDING',   note: '', isKit: true }
];

const orderAgeMin = {
  '24-15021-77421': 310, '08-15017-44806': 95, '24-15008-33107': 224,
  '24-15003-91882': 45,  '24-15015-60233': 30, '24-15019-08841': 25,
  '24-15004-11290': 12,  'SO-24612': 2900, 'SO-24618': 400,
  'SO-24631': 130, 'SO-24620': 65
};

const timeline = [
  { event: 'RECEIVED',  orderId: '24-15021-77421', hourFraction: 0.372 },
  { event: 'RECEIVED',  orderId: 'SO-24631',       hourFraction: 0.381 }, // 3 lines grouped
  { event: 'RECEIVED',  orderId: 'SO-24631',       hourFraction: 0.381 },
  { event: 'RECEIVED',  orderId: 'SO-24631',       hourFraction: 0.381 },
  { event: 'SHIPPED',   orderId: '23-14988-51002', hourFraction: 0.402 },
  { event: 'SHIPPED',   orderId: '23-14990-11873', hourFraction: 0.410 },
  { event: 'PREPARING', orderId: '08-15017-44806', hourFraction: 0.428 },
  { event: 'PRINTED',   orderId: '',               hourFraction: 0.433 },
  { event: 'SHIPPED',   orderId: 'SO-24580',       hourFraction: 0.451 },
  { event: 'SHIPPED',   orderId: 'SO-24580',       hourFraction: 0.451 },
  { event: 'RECEIVED',  orderId: '24-15015-60233', hourFraction: 0.472 },
  { event: 'RECEIVED',  orderId: '24-15019-08841', hourFraction: 0.489 },
  { event: 'CANCELED',  orderId: '24-14971-00518', hourFraction: 0.495 },
  { event: 'RECEIVED',  orderId: '24-15004-11290', hourFraction: 0.512 },
  { event: 'SHIPPED',   orderId: '24-14979-87359', hourFraction: 0.520 },
  { event: 'RECEIVED',  orderId: 'SO-24620',       hourFraction: 0.531 }
];

module.exports = {
  cockpit: {
    ebayGrab: 9, directGrab: 10,
    ebayPending: 11, directPending: 13,
    shippedToday: 41, receivedToday: 57,
    receivedEbay: 38, receivedDirect: 19,
    oldestPendingMinutes: 2900, pastRedlineCount: 3,
    prepQueueCount: 7, zohoPending: 4,
    lastSyncMinutes: 2,
    timeline: timeline,
    orderAgeMin: orderAgeMin
  },
  alerts: { paidShipping: { count: 2, rows: [] } },
  api: null,
  picker: 'Shipping - Hatem 7788',
  lastSync: '2m', lastSyncMinutes: 2,
  paceCar: null,
  openOrders: rows,
  openOrdersTotal: 27,                       // 24 shown + 3 past the cap
  openOrdersBy: { EBAY: 13, DIRECT: 14 },    // open on the floor (pre-cap truth)
  kits: [
    { key: 'SO-24612|160029', parent: '160029', order: 'SO-24612', total: 6, done: 2, hue: 0, dash: 0, left: ['A-9', 'C-2'] },
    { key: 'SO-24612|158652', parent: '158652', order: 'SO-24612', total: 5, done: 1, hue: 1, dash: 0, left: ['E-17', 'E-37'] },
    { key: 'SO-24618|217475', parent: '217475', order: 'SO-24618', total: 5, done: 2, hue: 0, dash: 1, left: ['A-51', 'B-44'] }
  ],
  serverTime: new Date().toISOString()
};
