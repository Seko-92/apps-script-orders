// Kubota KPAD web printout parser ("kpadweb.kubota.co.jp … PartsInfoPrintUnite").
//
// Layout, measured on D902.pdf (2026-10-03):
//   * Section header: "D902-E4B-AVN-1 -> ENGINE -> 000600 OIL FILTER ## D902-E4B-AVN-1"
//     (printed twice per page — dedupe).
//   * Column header: "No  Part Number  Part Name  Qty  RoundUp  IC  S/N  Remarks  Kg"
//   * Part line: "010 HH150-32430 FILTER(CARTRIDGE,OIL)   1   …   0.2"
//   * A long name WRAPS around the part line: first half on the line above, the rest below
//     ("ASSY GUIDE,OIL" / "020 1G471-36502   1" / "GAUGE").
// Same output shape as kubota.js.

"use strict";

const { PN_SHAPE } = require("./kubota");

const SECTION = /^\s*(\S+)\s*->\s*\S+\s*->\s*(\d{6})\s+(.+?)\s*(?:##.*)?$/;
const PART = /^\s*(\d{3})\s+([0-9A-Z]{5}-[0-9A-Z-]{4,})\s*(.*)$/;

function parseKpad(text, file) {
  const pages = String(text).split("\f");
  const out = { brand: "Kubota", source: "kpad", file: file || "", model: "", codeNo: "", validity: "",
                models: [], sections: [], flags: [], contents: [], imagePages: [], pageCount: pages.length };
  let current = null;

  for (let p = 0; p < pages.length; p++) {
    const lines = pages[p].split("\n");
    const pageNo = p + 1;
    let pendingName = "", inTable = false;
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      // A page can close one section and open the next — switch wherever the header appears.
      const sm = line.match(SECTION);
      if (sm) {
        if (!out.model) { out.model = sm[1]; out.models.push({ col: "A", name: sm[1] }); }
        if (!current || current.code !== sm[2]) {
          current = out.sections.find(s => s.code === sm[2]);
          if (!current) { current = { code: sm[2], name: sm[3].trim(), page: pageNo, parts: [] }; out.sections.push(current); }
        }
        inTable = false; pendingName = "";
        continue;
      }
      if (/\bPart Number\b/.test(line) && /\bPart Name\b/.test(line)) { inTable = true; pendingName = ""; continue; }
      if (/kpadweb\.kubota/.test(line)) { inTable = false; continue; }
      if (!inTable || !current) continue;
      const m = line.match(PART);
      let mm = m;
      if (!mm && /^\s*\d{3}\s{6,}/.test(line)) {
        // Part number wrapped around the ref line: "1G460-  METAL,ASSY(3-" / "040 … 2 … -0.20mm" / "23943  02,CRANKSHAFT)"
        const up = (lines[i - 1] || "").match(/^\s+([0-9A-Z]{5}-)\s*(.*)$/);
        const dn = (lines[i + 1] || "").match(/^\s+([0-9A-Z]{4,6})\b\s*(.*)$/);
        if (up && dn) {
          const pnJoined = up[1] + dn[1];
          const nameUp = up[2].trim(), nameDn = dn[2].trim();
          const wrapName = nameUp && nameDn ? (/-$/.test(nameUp) ? nameUp + nameDn : nameUp + " " + nameDn) : (nameUp || nameDn);
          const refM = line.match(/^\s*(\d{3})\s+(.*)$/);
          mm = [line, refM[1], pnJoined, refM[2]];
          mm.wrapName = wrapName;
          i++;                                   // the line below is consumed
          pendingName = "";
        } else {
          out.flags.push({ page: pageNo, kind: "pn-split", line: line.trim() + " / " + (lines[i + 1] || "").trim() });
          continue;
        }
      }
      if (!mm) {
        const t = line.trim();
        if (/^[0-9A-Z]{5}-(\s|$)/.test(t)) continue;      // top half of a wrapped part number
        // a lone name fragment (wrapped name) — remember it for the next part line
        if (t && !/^\d/.test(t) && !/Interchangeable|Update Date/.test(t)) pendingName = t;
        continue;
      }
      const [, ref, pn, restRaw] = mm;
      const restStart = line.length - restRaw.length;
      const chunks = [...restRaw.matchAll(/\S(?:.*?\S)??(?=\s{2,}|$)/g)].map(c => ({ t: c[0], col: restStart + c.index }));
      // Order is name · qty · (remark) · kg. Columns drift from the header, so go by kind:
      // first text = name, first integer after it = qty, text after qty = remark, decimal = kg.
      let name = "", qty = null, remark = "";
      chunks.forEach(c => {
        const isInt = /^\d+$/.test(c.t), isDec = /^\d+\.\d+$/.test(c.t);
        if (isDec) return;
        if (qty == null && isInt) { qty = Number(c.t); return; }
        if (qty == null && !name) { name = c.t; return; }
        if (qty != null && !isInt) remark = remark ? remark + " " + c.t : c.t;
      });
      const unbalanced = n => (n.match(/\(/g) || []).length !== (n.match(/\)/g) || []).length;
      if (mm.wrapName && (!name || /\($|-$/.test(name) || unbalanced(name))) name = mm.wrapName;
      if (!name) {
        // wrapped: fragment above + fragment below
        const below = (lines[i + 1] || "").trim();
        const belowIsName = below && !PART.test(lines[i + 1]) && !/^\d/.test(below) && !/kpadweb/.test(below);
        name = [pendingName, belowIsName ? below : ""].filter(Boolean).join(" ");
        if (belowIsName) i++;
      }
      pendingName = "";
      const part = { ref, pn, name, remark, qty: qty == null ? null : [qty], page: pageNo };
      if (!PN_SHAPE.test(pn)) out.flags.push({ page: pageNo, kind: "pn-shape", line: line.trim() });
      if (!name) out.flags.push({ page: pageNo, kind: "no-name", line: line.trim() });
      else if (unbalanced(name)) out.flags.push({ page: pageNo, kind: "name-check", line: pn + " " + name });
      current.parts.push(part);
    }
  }
  out.sections = out.sections.filter(s => s.parts.length);
  return out;
}

module.exports = { parseKpad };
