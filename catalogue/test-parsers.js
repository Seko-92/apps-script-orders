// node catalogue/test-parsers.js — every fixture is a real layout seen in the manuals (2026-10-03).
"use strict";
const { parseKubota } = require("./lib/kubota");
const { parseKpad } = require("./lib/kpad");

let pass = 0, fail = 0;
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want);
  if (g === w) { pass++; } else { fail++; console.log("FAIL " + name + "\n   got  " + g + "\n   want " + w); }
}
const parts = r => r.sections.flatMap(s => s.parts);

// ---------- Kubota book ----------
const HDR = [
  "                                                                                                                     A:V2203-M-E2B-EU-X3",
  "                                                                                                    Q'TY/S.No.",
  "REF.No.     PART No.                                                                                Q'TE/No.S.                        REMARKS",
  "POS.No.    REFERENCE         PART NAME            DESIGNATION             BEZEICHNUNG              STUECK/S.Nr.        I.C.          REMARQUES",
  "BILD-Nr.   BESELL-Nr.                                                                                                               BEMERKUNGEN"
].join("\n");
const COVER = [
  "Spare Parts List", "V2203-M-E2B", "(19.06.2006 --> 31.03.2007)",
  "                                 V2203-M-E2B-EU-X3                        1G475-00000",
  "                            CONTENTS", "0102    PISTON AND CRANKSHAFT ･････････････････････････････ 11",
  "0504    FAN ･･････････････････････････････････････････････････････ 29"
].join("\n");
const PISTON = [
  "                 PISTON AND CRANKSHAFT", "0102             PISTON ET VILEBREQUIN", "                 KOLBEN UND KURBELWELLE", "",
  HDR,
  "                                                                                                  4            -",
  "010 1G780-2111-2        PISTON                PISTON               KOLBEN                                                     STD",
  "                                                                                                  4            -",
  "080 17331-2297-0        METAL,CRANKPIN        COUSSINET DE BIELLE METALLTEIL                                                  -0.20mm SET",
  "                                                                                                  8            -",
  "040 14109-2133-0        CIR CLIP,INTERNAL     CIRCLIP              SPRENGRING",
  "160     ----            BLANK                 BLANC                LEERE SPALTE",
  "                                                                     Interchangeable;   not interchangeable;"
].join("\n");
const OILF = [
  "             OIL FILTER", "0006         FILTRE D'HUILE", "",
  "REF.No.PART No.                                                                            Q'TE/No.S.                  REMARKS",
  "POS.No.                PART NAME           DESIGNATION           BEZEICHNUNG              STUECK/S.Nr.        I.C.    REMARQUES",
  "                                                                                         1            -",
  "010 336.021.004",
  "    16414-3243-0 CARTRIDGE,OIL FILTER CARTOUCHE FILTRANTE    OELFILTERPATRONE"
].join("\n");
const r1 = parseKubota([COVER, PISTON, OILF, "   ", "NUMERICAL INDEX"].join("\f"), "x.pdf");
eq("cover model", [r1.model, r1.codeNo, r1.validity], ["V2203-M-E2B-EU-X3", "1G475-00000", "(19.06.2006 --> 31.03.2007)"]);
eq("piston line", parts(r1)[0], { ref: "010", pn: "1G780-2111-2", name: "PISTON", remark: "STD", qty: [4, null], page: 2 });
eq("size remark kept", parts(r1)[1].remark, "-0.20mm SET");
eq("no remark", [parts(r1)[2].name, parts(r1)[2].remark, parts(r1)[2].qty], ["CIR CLIP,INTERNAL", "", [8, null]]);
eq("BLANK slot skipped", parts(r1).some(p => p.pn === "----"), false);
eq("xref first, PN below", [parts(r1)[3].pn, parts(r1)[3].xref, parts(r1)[3].name], ["16414-3243-0", "336.021.004", "CARTRIDGE,OIL FILTER"]);
eq("section from contents not found → flagged", r1.flags.filter(f => /section/.test(f.kind)).map(f => f.line), ["0504 FAN"]);
eq("sections", r1.sections.map(s => s.code + " " + s.name), ["0102 PISTON AND CRANKSHAFT", "0006 OIL FILTER"]);

// ---------- KPAD printout ----------
const KP = [
  "kpadweb.kubota.co.jp/kpad2/PartsInfoPrintUnite.do",
  "              D902-E4B-AVN-1 -> ENGINE -> 010000 MAIN BEARING CASE ## D902-E4B-AVN-1",
  "   No       Part Number            Part Name              Qty    RoundUp         IC        S/N              Remarks                  Kg",
  "   080 15861-22973 METAL,CRANKPIN                            3                                            -0.20mm/SET               0.03",
  "                           ASSY GUIDE,OIL",
  "   020      1G471-36502                                  1                                                                          0.14",
  "                           GAUGE",
  "            1G460-      METAL,ASSY(3-",
  "   040                                                          2                                          -0.20mm/SET              0.022",
  "            23943        02,CRANKSHAFT)",
  "        D902-E4B-AVN-1 -> ENGINE -> 010100 CAMSHAFT AND IDLE GEAR SHAFT ##",
  "                                   D902-E4B-AVN-1",
  "   No       Part Number            Part Name             Qty    RoundUp          IC        S/N              Remarks                  Kg",
  "   010      16851-15552    TAPPET                          6                                                                        0.02"
].join("\n");
const r2 = parseKpad(KP, "d902.pdf");
eq("kpad model", r2.model, "D902-E4B-AVN-1");
eq("kpad qty + remark, kg ignored", [parts(r2)[0].qty, parts(r2)[0].remark], [[3], "-0.20mm/SET"]);
eq("kpad name wrapped around line", parts(r2)[1].name, "ASSY GUIDE,OIL GAUGE");
eq("kpad PN wrapped above+below", [parts(r2)[2].pn, parts(r2)[2].name, parts(r2)[2].qty], ["1G460-23943", "METAL,ASSY(3-02,CRANKSHAFT)", [2]]);
eq("kpad section switches mid-page (wrapped ## header)", r2.sections.map(s => s.code + ":" + s.parts.length), ["010000:3", "010100:1"]);

// ---------- drawings ----------
const { drawingBand, mergeCallouts } = require("./lib/drawings");
const box = { w: 595, h: 841, words: [
  { x0: 35, y0: 36.5, x1: 71, y1: 54.5, t: "0102" },
  { x0: 36, y0: 405.4, x1: 53, y1: 412.4, t: "REF.No." } ] };
eq("band sits between the title block and REF.No.", drawingBand(box, "0102"), { x: 0, y: 70.5, w: 595, h: 306.9 });
eq("table-only page has no band", drawingBand({ w: 595, h: 841, words: [{ x0: 0, y0: 36, x1: 0, y1: 54, t: "0102" }, { x0: 0, y0: 120, x1: 0, y1: 127, t: "REF.No." }] }, "0102"), null);
const a1 = [{ ref: "010", x: 0.39, y: 0.62, w: 0.02, h: 0.02 }];
eq("same callout from two passes counts once", mergeCallouts(a1, [{ ref: "010", x: 0.395, y: 0.625, w: 0.02, h: 0.02 }]).length, 1);
eq("same ref at another spot is kept (a part drawn twice)", mergeCallouts(a1, [{ ref: "010", x: 0.7, y: 0.2, w: 0.02, h: 0.02 }]).length, 2);

// ---------- OCR of scanned books ----------
const K = require("./lib/kubota-scan");
eq("OCR: letter Z in a digit position", K.normalisePn("15471-3501-Z"), "15471-3501-2");
eq("OCR: lower-case l read for 1", K.normalisePn("1G780-2lll-2"), "1G780-2111-2");
eq("OCR: brackets from table rules stripped", K.normalisePn("|1A021-3515-0]"), "1A021-3515-0");
eq("OCR: number glued to the next word", K.normalisePn("17331-2105-0_[ASSY"), "17331-2105-0");
eq("OCR: no dashes, part number (starts 1) → 5-4-1", K.normalisePn("1547135012"), "15471-3501-2");
eq("OCR: no dashes, hardware (starts 0) → 5-5", K.normalisePn("0771500401"), "07715-00401");
eq("OCR: a word is not a number", K.normalisePn("ASSY"), null);
eq("OCR: remark spacing", K.fixRemark("-0.20MMSET"), "-0.20mm SET");
eq("OCR: remark O for 0", K.fixRemark("+O.50MM"), "+0.50mm");
// sizes sit at the BOTTOM of a cell: banding, not nearest-row, decides the owner
// (the old code accepted a read only within 0.7× the row height of the part number's centre,
//  so a size printed at the bottom of a tall cell was silently dropped)
const rows = [{ y: 100, h: 30, remark: "", qty: null }, { y: 166, h: 30, remark: "", qty: null }];
K.attachColumn(rows, [{ block: 1, par: 1, line: 1, x: 0, y: 118, w: 50, h: 20, t: "STD" }], "remark");
eq("OCR: a size at the bottom of its cell is kept, on its own row", rows.map(r => r.remark), ["STD", ""]);
// the column re-read fuses REF and part number; it still confirms and rescues
const L = { nameEnd: 900 };
const pageRows = [{ ref: "010", pn: "16423-2111-0", y: 100, h: 50, name: "PISTON" }];
const col = [{ t: "16423-2111-0", x: 380, y: 85, w: 300, h: 30, block: 1, par: 1, line: 1 },
             { t: "10007715-00401", x: 380, y: 285, w: 320, h: 30, block: 1, par: 1, line: 2 }];
const out = K.rescueRows(pageRows, [], col, L);
eq("OCR: two reads agreeing confirm a number", out[0].agree, true);
eq("OCR: a row only the column read found is rescued, REF split off", [out[1].pn, out[1].ref, out[1].rescued], ["07715-00401", "100", true]);

// section codes as OCR really reads them on the newer bilingual books (2026-10-04)
eq("OCR: section codes, every shape seen", ["0102", "E03.", "EO2-1.", "ES1.", "£11-1.", "E11-1", "BOS", "OIL", "C11-L", "1234-5"].map(K.sectionCodeOf),
   ["0102", "E03", "E02-1", "E51", "E11-1", "E11-1", "", "", "", ""]);

// qty vote (2026-10-04): majority of the cell/column reads; else light read → column → wiped read
const { voteQty } = require("./lib/scan-book");
eq("OCR: qty vote", [
  voteQty(3, null, 3, 3),      // D905: wiped + column + word agree
  voteQty(7, 2, 4, 2),         // V1505: the light reads agree on 2, the line-wiped read lost the base
  voteQty(5, 2, 4, null),      // no majority → the light read
  voteQty(null, null, 3, null),// only the column read
  voteQty(1, 1, 3, 3),         // tie 2–2 → the light read's value
], [3, 2, 2, 3, 1]);

// model columns (2026-10-04): the header line → models; the strict vote for variant cells
const { pageModels, voteStrict } = require("./lib/scan-book");
const mw = (t, x) => ({ t, x, y: 100, h: 20 });
eq("OCR: model line with a two-model column", pageModels([mw("A:D1703-BB-EC-1,", 10), mw("B:D1703-BB-EC-STD-2-SI,D1703-BB-EC-STD-3,", 200), mw("C:D1703-BBS-EC-1,", 600), mw("D:D1703-BG-SAE-1-S3", 800)]).map(m => m.col + "=" + m.name),
   ["A=D1703-BB-EC-1", "B=D1703-BB-EC-STD-2-SI, D1703-BB-EC-STD-3", "C=D1703-BBS-EC-1", "D=D1703-BG-SAE-1-S3"]);
eq("OCR: variant vote — a lone wiped read is not trusted (a dash read as 1), two agree win",
   [voteStrict(1, null, null), voteStrict(3, 3, null), voteStrict(null, 2, null), voteStrict(7, 2, null)], [null, 3, 2, null]);

// speed (2026-10-05): the 3rd cell read is skipped when needThird says it can't change the vote —
// prove that over EVERY combination, so a future vote change can't silently break the shortcut
const { needThird } = require("./lib/scan-book");
const Vs = [null, 1, 2, 3]; let skipBad = 0;
for (const w of Vs) for (const r of Vs) for (const c of Vs) for (const x of Vs) {
  if (needThird(w, r)) continue;
  if (voteQty(w, r, c, x) !== voteQty(w, r, c, null) || voteStrict(w, r, x) !== voteStrict(w, r, null)) skipBad++;
}
eq("OCR: skipping the 3rd qty read never changes a vote", skipBad, 0);

// prefix letter misread (2026-10-05): "1C010" read as "10010" / "16010"
const { fixMisreadPrefixes } = require("./lib/pn-fix");
const pfKnown = new Set(["1C010-5675-0", "10244-4232-0", "16241-6401-2"]);
const pfBooks = [
  { parts: [{ pn: "10010-5675-0" }, { pn: "10244-4232-0" }, { pn: "16241-6401-2" }, { pn: "16010-6303-0" },
            { pn: "1C020-1111-0" }, { pn: "10020-2222-0" }, { pn: "10030-3333-0" }, { pn: "16099-4444-0" }] },
  { parts: [{ pn: "1C010-6303-0" }, { pn: "1C030-9999-0" }, { pn: "1G030-9999-1" }] }];
const pfr = fixMisreadPrefixes(pfBooks, pn => pfKnown.has(pn));
eq("OCR prefix fix: known · real 10… untouched · real 16… untouched · read elsewhere · stem · no evidence",
   pfBooks[0].parts.map(p => p.pn),
   ["1C010-5675-0", "10244-4232-0", "16241-6401-2", "1C010-6303-0", "1C020-1111-0", "1C020-2222-0", "10030-3333-0", "16099-4444-0"]);
eq("OCR prefix fix: an inferred fix is not 'agreed'; tiers counted",
   [pfBooks[0].parts[3].ocr.agree, pfBooks[0].parts[3].ocr.prefixFix, pfr.byTier.known, pfr.byTier.read, pfr.byTier.stem, pfr.left], [false, "read", 1, 1, 1, 1]);


// ---------- variant names: a serial range is not a variant (user test 2026-10-06) ----------
const { isSerialRange, dropSerialRangeModels } = require("./lib/models");
eq("serial ranges recognised, real variant names kept",
   ["<=15000", "14000 to 15000 |", "=<<=15000", "<{=15000", "14000 to 15000 :", "D1703-BB-EC-1", "XXXX |", "ECHANGE", "V2203-M-E2B-EU-Schäffer"].map(isSerialRange),
   [true, true, true, true, true, false, false, false, false]);
const sr = { models: [{ col: "A", name: "<=15000" }, { col: "B", name: "14000 to 15000 |" }], sections: [{ parts: [{ qty: [4, 1] }, { qty: [null, null] }] }] };
dropSerialRangeModels(sr);
eq("all-range header → no variants, one qty column", [sr.models, sr.sections[0].parts.map(p => p.qty)], [[], [[4], [null]]]);
const mx = { models: [{ col: "A", name: "D1105-BB-EC-1" }, { col: "B", name: "<=15000" }, { col: "C", name: "XXXX |" }], sections: [{ parts: [{ qty: [1, 9, 2] }] }] };
dropSerialRangeModels(mx);
eq("mixed header → the range column alone goes", [mx.models.map(m => m.col), mx.sections[0].parts[0].qty], [["A", "C"], [1, 2]]);

// ---------- KPAD drawing band: between the two copies of the section title (D902, 2026-10-06) ----------
const { kpadDrawingBand } = require("./lib/drawings");
const kw = (y0, t) => ({ x0: 50, y0, x1: 100, y1: y0 + 10, t });
const kbox = { w: 596, h: 842, words: [kw(41, "000300"), kw(65, "Update"), kw(453, "000300"), kw(477, "Update"), kw(524, "No")] };
eq("KPAD band: from under the first title block to above the second title", kpadDrawingBand(kbox, "000300"), { x: 0, y: 81, w: 596, h: 364 });
eq("KPAD band: one title only (table page) → none", kpadDrawingBand({ w: 596, h: 842, words: [kw(41, "000300")] }, "000300"), null);

console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
