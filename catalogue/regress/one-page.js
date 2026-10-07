const { ocrBook } = require(process.argv[2] + "/lib/scan-book");
const r = ocrBook(process.argv[3], { pages: process.argv[4].split(",").map(Number) });
console.log("models", JSON.stringify(r.models));
r.sections.forEach(s => { console.log(s.code, s.name); s.parts.forEach(p => console.log(" ", p.ref, p.pn, (p.name||"").padEnd(20), JSON.stringify(p.qty), p.qtyInk ? "ink" : "")); });
