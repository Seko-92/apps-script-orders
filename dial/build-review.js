#!/usr/bin/env node
/**
 * build-review.js — renders the review page with every PNG base64-embedded.
 *
 * ⚠ The page is the ONLY preview the sheet gets. Same pattern the repo already uses for
 *   assessment-*.template.html: the template is the source, the rendered copy is disposable.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const OUT = process.argv[2] || path.join(__dirname, 'renders', 'review.html');
const R = path.join(__dirname, 'renders');
const tpl = fs.readFileSync(path.join(__dirname, 'review.template.html'), 'utf8');
const html = tpl.replace(/\{\{IMG:([a-z0-9@.\-]+)\}\}/gi, (_, f) => {
  const p = path.join(R, f);
  if (!fs.existsSync(p)) { console.error('MISSING RENDER: ' + f); return ''; }
  return 'data:image/png;base64,' + fs.readFileSync(p).toString('base64');
});
const left = (html.match(/\{\{IMG:/g) || []).length;
if (left) { console.error(`✗ ${left} placeholder(s) unresolved`); process.exit(1); }
fs.writeFileSync(OUT, html);
console.log(`${(html.length / 1024 / 1024).toFixed(2)} MB -> ${OUT}`);
