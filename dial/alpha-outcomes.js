// Renders the three possible outcomes side by side so the user knows what to look for.
const { createCanvas, GlobalFonts } = require('@napi-rs/canvas');
const fs = require('fs');
GlobalFonts.registerFromPath(__dirname + '/src/fonts/Oswald-400.ttf', 'Oswald');

const W = 232, H = 56, PAD = 16, SCALE = 2;
const cv = createCanvas((W * 3 + PAD * 4) * SCALE, (H + 70) * SCALE);
const g = cv.getContext('2d');
g.scale(SCALE, SCALE);
g.fillStyle = '#2b2b2b'; g.fillRect(0, 0, cv.width, cv.height);

function bandWithHeadline(x, y) {              // what D1 actually looks like today
  g.fillStyle = '#1a1a1a'; g.fillRect(x, y, W, H);
  g.fillStyle = '#f0ece0'; g.font = '15px Oswald';
  g.fillText('12 to grab · eBay 8', x + 12, y + 26);
  g.fillStyle = '#8a8a86'; g.font = '11px Oswald';
  g.fillText('Direct 4', x + 12, y + 43);
}
function frame(x, y) {                          // the probe overlay itself
  g.strokeStyle = '#ff00ff'; g.lineWidth = 3;
  g.strokeRect(x + 1.5, y + 1.5, W - 3, H - 3);
  g.fillStyle = '#ffffff'; g.fillRect(x + W - 40, y + 6, 14, 14);
  g.fillStyle = '#000000'; g.fillRect(x + W - 22, y + 6, 14, 14);
  g.fillStyle = 'rgba(128,128,128,0.5)'; g.fillRect(x + W - 40, y + 26, 32, 22);
}
function label(x, y, t, c) {
  g.fillStyle = c; g.font = 'bold 12px Oswald';
  g.fillText(t, x, y);
}

let x = PAD, y = 40;
// A — alpha passes through: headline readable inside the frame
bandWithHeadline(x, y); frame(x, y);
label(x, y - 10, 'A · PASSES  -> text readable', '#7ed321');
label(x, y + H + 18, 'static frame is LIVE', '#7ed321');

// B — composited on white
x += W + PAD;
g.fillStyle = '#1a1a1a'; g.fillRect(x, y, W, H);
g.fillStyle = '#ffffff'; g.fillRect(x + 3, y + 3, W - 6, H - 6);
frame(x, y);
label(x, y - 10, 'B · WHITE  -> hole = swatch', '#ff6b6b');
label(x, y + H + 18, 'static frame is DEAD', '#ff6b6b');

// C — composited on black
x += W + PAD;
g.fillStyle = '#1a1a1a'; g.fillRect(x, y, W, H);
g.fillStyle = '#000000'; g.fillRect(x + 3, y + 3, W - 6, H - 6);
frame(x, y);
label(x, y - 10, 'C · BLACK  -> flat, no text', '#ff6b6b');
label(x, y + H + 18, 'static frame is DEAD', '#ff6b6b');

fs.writeFileSync(__dirname + '/renders/alpha-outcomes.png', cv.toBuffer('image/png'));
console.log('wrote renders/alpha-outcomes.png');
