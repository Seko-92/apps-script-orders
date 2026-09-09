/** build-review.js — inline every 1:1 render into one page. The renders MUST ship at their
 *  natural width (1136px, 799px, 579px): the whole argument is that row one is judged at the
 *  size it lives at, and a page that scales them down repeats the 09-03 review mistake. */
'use strict';
const fs=require('fs'),path=require('path');
const R=(f)=>{const b=fs.readFileSync(path.join(__dirname,'renders',f));
  const m=f.endsWith('.gif')?'image/gif':'image/png';
  return `data:${m};base64,${b.toString('base64')}`;};
const OUT=process.argv[2];
let html=fs.readFileSync(path.join(__dirname,'review.tpl.html'),'utf8');
html=html.replace(/\{\{IMG:([a-zA-Z0-9._-]+)\}\}/g,(_,f)=>R(f));
fs.writeFileSync(OUT,html);
console.log('  '+OUT+'  '+(fs.statSync(OUT).size/1048576).toFixed(2)+' MB');
