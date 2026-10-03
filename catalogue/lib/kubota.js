// Kubota "Spare Parts List" parser (text PDFs — pdftotext -layout output).
//
// Layout, measured on V2203-M-E2B (2026-10-03):
//   * One PDF page per \f. Each parts page opens with the section title in English, then a
//     line "0102   <French title>", then "A:<model>" (more columns B:, C: when a manual covers
//     variants), then the column header block starting "REF.No.".
//   * A part line:  "010 1G780-2111-2   PISTON   PISTON   KOLBEN        STD"
//     REF · PART No. · English name · French · German · remark (sizes live here).
//   * The quantity sits on the line ABOVE, one number per model column ("4   -").
//   * The NUMERICAL INDEX at the end repeats every number — parsing stops there.
//
// Pure: takes text, returns data. No file system, so it is Node-testable.

"use strict";

// Kubota numbers: 5 alnum + 4/5 digits + optional check digit (1G780-2111-2, 17331-2297-0),
// or the 5-5 hardware form (07715-00401, 04811-10300).
const PN_SHAPE = /^[0-9A-Z]{5}-\d{4,5}(-\d)?$/;
const PART_LINE = /^\s*(\d{3})\s+(\S+)(?:\s+(.*))?$/;
const QTY_LINE = /^\s+((?:\d+|-)(?:\s+(?:\d+|-))*)\s*$/;
const SECTION_LINE = /^\s*(\d{4})\s{2,}\S/;

function parseKubota(text, file) {
  const pages = String(text).split("\f");
  const out = {
    brand: "Kubota",
    file: file || "",
    model: "", codeNo: "", validity: "",
    models: [],          // model column letters → names, e.g. [{col:"A", name:"V2203-M-E2B-EU-X3"}]
    sections: [],
    flags: [],
    pageCount: pages.length
  };

  // ---- cover: model, code number, validity window -------------------------------------
  const head = pages.slice(0, 4).join("\n");
  const mv = head.match(/\(\s*\d{2}\.\d{2}\.\d{4}\s*-->\s*[\d.]*\s*\)/);
  if (mv) out.validity = mv[0].replace(/\s+/g, " ");
  const mc = head.match(/^\s*([A-Z0-9][A-Z0-9-]{3,})\s{4,}([0-9A-Z]{5}-\d{5})\s*$/m);
  if (mc) { out.model = mc[1]; out.codeNo = mc[2]; }

  // Contents page: the sections the manual promises. Used to prove nothing was missed.
  const contents = [];
  for (const p of pages.slice(0, 14)) {
    if (!/CONTENTS/.test(p)) continue;
    for (const l of p.split("\n")) {
      const m = l.match(/^(\d{4})\s+(.+?)\s*[･.·…]{3,}\s*(\d+)\s*$/);
      if (m && !contents.some(c => c.code === m[1])) contents.push({ code: m[1], name: m[2].trim(), printedPage: Number(m[3]) });
    }
    if (contents.length) break;
  }
  out.contents = contents;

  let current = null;     // the section being filled

  for (let p = 0; p < pages.length; p++) {
    const lines = pages[p].split("\n");
    const pageNo = p + 1;                     // PDF page index — what a viewer opens

    // Stop at the index: it repeats every number with no section context.
    const top = lines.slice(0, 8).join(" ");
    if (/NUMERICAL INDEX|INDEX NUMERIQUE|NUMERISCHEN INDEX/.test(top) && !/REF\.No\./.test(pages[p])) break;

    const hdrIdx = lines.findIndex(l => /^\s*REF\.No\./.test(l));
    if (hdrIdx < 0) continue;                 // cover, notice, contents, illustrations

    // ---- section (English title is the non-empty line above the "0102  French" line)
    let code = "", name = "";
    for (let i = 0; i < hdrIdx; i++) {
      const m = lines[i].match(SECTION_LINE);
      if (m) {
        code = m[1];
        for (let j = i - 1; j >= 0; j--) { if (lines[j].trim()) { name = lines[j].trim(); break; } }
        break;
      }
    }
    // Model columns for this page ("A:V2203-M-E2B-EU-X3   B:...")
    for (let i = 0; i < hdrIdx; i++) {
      const re = /\b([A-Z]):([A-Z0-9]\S+)/g; let m;
      while ((m = re.exec(lines[i]))) {
        if (!out.models.some(x => x.col === m[1])) out.models.push({ col: m[1], name: m[2] });
      }
    }
    if (!out.model && out.models.length) out.model = out.models[0].name;

    if (code && (!current || current.code !== code)) {
      current = { code, name, page: pageNo, parts: [] };
      out.sections.push(current);
    } else if (!current) {
      current = { code: code || "????", name: name || "(untitled)", page: pageNo, parts: [] };
      out.sections.push(current);
      out.flags.push({ page: pageNo, kind: "no-section", line: "" });
    }

    // Remarks column: where the remarks headers start on THIS page.
    let remarksCol = Infinity;
    for (let i = hdrIdx; i < Math.min(lines.length, hdrIdx + 4); i++) {
      ["REMARKS", "REMARQUES", "BEMERKUNGEN"].forEach(w => {
        const k = lines[i].indexOf(w);
        if (k >= 0) remarksCol = Math.min(remarksCol, k);
      });
    }

    let enWidth = 0;
    for (let i = hdrIdx; i < Math.min(lines.length, hdrIdx + 4); i++) {
      const a = lines[i].indexOf("PART NAME"), b = lines[i].indexOf("DESIGNATION");
      if (a >= 0 && b > a) { enWidth = b - a; break; }
    }

    let lastQty = null;
    for (let i = hdrIdx + 1; i < lines.length; i++) {
      const line = lines[i];
      if (/Interchangeable;/.test(line)) break;              // page footer legend
      const q = line.match(QTY_LINE);
      if (q && !PART_LINE.test(line)) {
        lastQty = q[1].split(/\s+/).map(x => (x === "-" ? null : Number(x)));
        continue;
      }
      const m = line.match(PART_LINE);
      if (!m) continue;
      const ref = m[1], pn = m[2];
      if (/^-+$/.test(pn)) { lastQty = null; continue; }   // "---- BLANK" placeholder slot
      let rest = m[3] || "";
      let restLine = line;
      let xref = "";
      // Name pushed to the next line, led by a customer cross-reference ("336.021.004 CARTRIDGE…")
      let pnUse = pn;
      if (!rest.trim() && i + 1 < lines.length) {
        const nx = lines[i + 1];
        const mx = nx.match(/^\s*(\S+)\s+(.*)$/);
        if (mx && !PART_LINE.test(nx)) {
          if (!PN_SHAPE.test(pn) && PN_SHAPE.test(mx[1])) { xref = pn; pnUse = mx[1]; }   // "010 336.021.004" / "16414-3243-0 NAME"
          else xref = mx[1];                                                                  // "010 16414-3243-0" / "336.021.004 NAME"
          rest = mx[2]; restLine = nx; i++;
        }
      }
      const restStart = restLine.length - rest.length;

      let remark = "";
      let body = rest;
      if (isFinite(remarksCol) && restLine.length > remarksCol - 4) {
        // Remark = the last 2+-space-separated chunk, if it starts near the remarks column.
        const chunks = [...rest.matchAll(/\S(?:.*?\S)??(?=\s{2,}|$)/g)];
        const lastC = chunks[chunks.length - 1];
        if (lastC && restStart + lastC.index >= remarksCol - 12) {
          remark = lastC[0].trim();
          body = rest.slice(0, lastC.index);
        }
      }
      let partName = (body.split(/\s{2,}/)[0] || "").trim();
      // English ran into French with a single space ("CARTRIDGE,OIL FILTER CARTOUCHE …"):
      // cut at the last space that fits the English column.
      if (enWidth && partName.length > enWidth + 1) {
        const cut = partName.lastIndexOf(" ", enWidth + 1);
        if (cut > 3) partName = partName.slice(0, cut);
      }

      const part = { ref, pn: pnUse, name: partName, remark, qty: lastQty, page: pageNo };
      if (xref) part.xref = xref;
      if (!PN_SHAPE.test(pnUse)) out.flags.push({ page: pageNo, kind: "pn-shape", line: line.trim() });
      if (!partName) out.flags.push({ page: pageNo, kind: "no-name", line: line.trim() });
      current.parts.push(part);
      lastQty = null;
    }
  }

  out.sections = out.sections.filter(s => s.parts.length);

  // Printed page → PDF page offset, learned from sections we did find.
  const offsets = {};
  out.sections.forEach(s => {
    const c = contents.find(c => c.code === s.code);
    if (c) { const o = s.page - c.printedPage; offsets[o] = (offsets[o] || 0) + 1; }
  });
  const offset = Number(Object.keys(offsets).sort((a, b) => offsets[b] - offsets[a])[0]);
  out.imagePages = pages.map((t, i) => (t.replace(/\s/g, "").length < 20 ? i + 1 : 0)).filter(Boolean);
  contents.forEach(c => {
    if (out.sections.some(s => s.code === c.code)) return;
    const pg = isFinite(offset) ? c.printedPage + offset : null;
    const isImage = pg && out.imagePages.indexOf(pg) >= 0;
    out.flags.push({ page: pg, kind: isImage ? "section-image-only" : "section-missing", line: c.code + " " + c.name });
  });
  out.models.sort((a, b) => a.col.localeCompare(b.col));
  return out;
}

module.exports = { parseKubota, PN_SHAPE };
