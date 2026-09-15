"""Round 2 of the movie-limit probe (2026-09-16). Round 1 showed BYTES are not the limit (8 MB noise
   accepted) but could not separate FRAME COUNT (refused somewhere in 1,000–1,257) from PLAY TIME
   (refused somewhere in 59–86 s), because every bigger rung was also longer. This separates them.
     D = 100 frames, play time grows (70 s · 5 min · 11.7 min)
     G = frames grow, play stays under 50 s (1,100 · 1,200)
     R = the real movie opening, 832 frames (accepted count) WITH its real 45 s pauses (5.3 min)

   python3 probe-ladder2.py <movie.gif> [outdir]   → probe-d*/g*/r1.gif + ladder2.json
   Pairs with probeMovieLimits2() in BrandTheme.js."""
import os, sys, json
from PIL import Image, ImageSequence
MOVIE = sys.argv[1]
HERE = sys.argv[2] if len(sys.argv) > 2 else os.getcwd()
os.makedirs(HERE, exist_ok=True)
W, H = 799, 133
out = []

def note(name, path):
    im = Image.open(path); n = 0; ms = 0
    for f in ImageSequence.Iterator(im):
        n += 1; ms += f.info.get('duration', 0)
    assert im.size == (W, H), im.size
    out.append({'id': name, 'bytes': os.path.getsize(path), 'frames': n, 'secs': round(ms / 1000)})
    print(out[-1])

def dots(n, ms, path):
    frames = []
    for i in range(n):
        im = Image.new('P', (W, H), 0)
        im.putpalette([26, 26, 26, 255, 212, 0] + [0] * 762)
        x = (i * 7) % (W - 12); y = 20 + int(45 * (1 + ((i // 25) % 2)))
        im.paste(1, (x, y, x + 12, y + 12))
        frames.append(im)
    frames[0].save(path, save_all=True, append_images=frames[1:], duration=ms, loop=0)

for name, n, ms in [('d1', 100, 700), ('d2', 100, 3000), ('d3', 100, 7000), ('g1', 1100, 40), ('g2', 1200, 40)]:
    p = os.path.join(HERE, 'probe-%s.gif' % name); dots(n, ms, p); note(name, p)

# R1: exact byte-prefix of the movie ending after stored frame 832, delays untouched
data = open(MOVIE, 'rb').read()
flags = data[10]; pos = 13 + ((3 << ((flags & 7) + 1)) if flags & 0x80 else 0)
ends = []
while pos < len(data):
    b = data[pos]
    if b == 0x21:
        q = pos + 2
        while data[q]: q += data[q] + 1
        pos = q + 1
    elif b == 0x2C:
        packed = data[pos + 9]; q = pos + 10
        if packed & 0x80: q += 3 << ((packed & 7) + 1)
        q += 1
        while data[q]: q += data[q] + 1
        pos = q + 1; ends.append(pos)
    elif b == 0x3B: break
    else: raise SystemExit('bad block 0x%02x at %d' % (b, pos))
p = os.path.join(HERE, 'probe-r1.gif')
open(p, 'wb').write(data[:ends[831]] + b'\x3B')
note('r1', p)
json.dump(out, open(os.path.join(HERE, 'ladder2.json'), 'w'), indent=1)
