// node run1.js <catalogue dir> <pdf path> <pages csv> <out json>
const { ocrBook } = require(process.argv[2] + "/lib/scan-book");
const r = ocrBook(process.argv[3], { pages: process.argv[4].split(",").map(Number) });
const rows = [];
r.sections.forEach(s => s.parts.forEach(p => rows.push({ page: p.page, pn: p.pn, ref: p.ref, qty: p.qty })));
require("fs").writeFileSync(process.argv[5], JSON.stringify(rows));
