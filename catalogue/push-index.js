// Send the part-number → engines index to Apps Script, so the Parts Finder can say
// "in N engine catalogues ↗".   (2026-10-09)
//
//   node catalogue/bundle.js        → writes catalogue/out/pn-index.json
//   node catalogue/push-index.js    → sends it (replaces the whole index)
//
// Run it after every publish. Goes through /api/board like import.js's MI evidence call;
// the write is allowed only with CATALOGUE_PUBLISH_KEY, read from the (gitignored) Secrets.js.
"use strict";
const fs = require("fs"), path = require("path"), os = require("os");
const { execFileSync } = require("child_process");

const API = "https://hq.yassinqurabi.com/api/board";
const file = path.join(__dirname, "out", "pn-index.json");
if (!fs.existsSync(file)) { console.error("No " + file + " — run node catalogue/bundle.js first."); process.exit(1); }

const secrets = fs.readFileSync(path.join(__dirname, "..", "Secrets.js"), "utf8");
const m = /var\s+CATALOGUE_PUBLISH_KEY\s*=\s*"([^"]+)"/.exec(secrets);
if (!m) { console.error("CATALOGUE_PUBLISH_KEY is not in Secrets.js."); process.exit(1); }

const index = JSON.parse(fs.readFileSync(file, "utf8"));
const engines = new Set(); Object.values(index).forEach(l => l.forEach(e => engines.add(e)));
const body = JSON.stringify({ action: "catalogueIndexPut", publishKey: m[1], index, meta: { engines: engines.size } });

// the body is ~260 KB — too long for one command-line argument, so it goes through a file
const tmp = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "hq-pnindex-")), "body.json");
fs.writeFileSync(tmp, body);
try {
  const out = execFileSync("curl", ["-s", "--max-time", "180", "-X", "POST", API,
    "-H", "Content-Type: application/json", "--data-binary", "@" + tmp]).toString();
  let j; try { j = JSON.parse(out); } catch (e) { console.error("Not JSON back:\n" + out.slice(0, 400)); process.exit(1); }
  if (!j.ok) { console.error("Refused: " + (j.reason || j.message || JSON.stringify(j).slice(0, 300))); process.exit(1); }
  console.log(`✓ index sent: ${j.keys} part numbers · ${engines.size} engines`);
} finally { fs.rmSync(path.dirname(tmp), { recursive: true, force: true }); }
