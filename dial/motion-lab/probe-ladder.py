"""Build the movie-limit probe ladder. All rungs are 799x133 (A1:E2).
   B = bytes with few frames (noise) · F = frames with few bytes (moving dot) ·
   M = exact byte-prefixes of the real movie, pauses shortened to 1.5 s.

   python3 probe-ladder.py <movie.gif> [outdir]   → probe-*.gif + ladder.json
   Pairs with probeMovieLimits() in BrandTheme.js (upload the files to /opt/hq-app/mast/ first)."""
import os, random, json, sys
from PIL import Image, ImageSequence
MOVIE = sys.argv[1]
HERE = sys.argv[2] if len(sys.argv) > 2 else os.getcwd()
os.makedirs(HERE, exist_ok=True)
W, H = 799, 133
random.seed(7)
out = []

def note(name, path):
    im = Image.open(path); n = 0; ms = 0
    for f in ImageSequence.Iterator(im):
        n += 1; ms += f.info.get('duration', 0)
    assert im.size == (W, H), im.size
    out.append({'id': name, 'bytes': os.path.getsize(path), 'frames': n, 'secs': round(ms / 1000)})
    print(out[-1])

# ── B: bytes, few frames ──
def noise_frame():
    im = Image.frombytes('P', (W, H), os.urandom(W * H))
    return im
pal = [random.randrange(256) for _ in range(768)]
for name, target in [('b1', 1.0), ('b2', 2.5), ('b3', 5.0), ('b4', 8.0)]:
    n = 3
    while True:
        frames = [noise_frame() for _ in range(n)]
        for f in frames: f.putpalette(pal)
        p = os.path.join(HERE, 'probe-%s.gif' % name)
        frames[0].save(p, save_all=True, append_images=frames[1:], duration=400, loop=0, optimize=False)
        mb = os.path.getsize(p) / 1048576
        if mb >= target: break
        n = max(n + 1, int(n * target / mb + 0.999))
    note(name, p)

# ── F: frames, few bytes ──
for name, n in [('f1', 300), ('f2', 1000), ('f3', 2500)]:
    frames = []
    for i in range(n):
        im = Image.new('P', (W, H), 0)
        im.putpalette([26, 26, 26, 255, 212, 0] + [0] * 762)
        x = (i * 5) % (W - 12); y = 20 + int(45 * (1 + ((i // 25) % 2)))
        im.paste(1, (x, y, x + 12, y + 12))
        frames.append(im)
    p = os.path.join(HERE, 'probe-%s.gif' % name)
    frames[0].save(p, save_all=True, append_images=frames[1:], duration=40, loop=0)
    note(name, p)

# ── M: exact prefixes of the real movie ──
data = open(MOVIE, 'rb').read()
flags = data[10]; pos = 13 + ((3 << ((flags & 7) + 1)) if flags & 0x80 else 0)
ends, gce = [], []          # byte offset after each image; offsets of GCE delay fields
while pos < len(data):
    b = data[pos]
    if b == 0x21:
        label = data[pos + 1]; q = pos + 2
        if label == 0xF9: gce.append(q + 2)       # size(1) packed(1) delay(2)
        while data[q]: q += data[q] + 1
        pos = q + 1
    elif b == 0x2C:
        packed = data[pos + 9]; q = pos + 10
        if packed & 0x80: q += 3 << ((packed & 7) + 1)
        q += 1                                    # LZW min code size
        while data[q]: q += data[q] + 1
        pos = q + 1; ends.append(pos)
    elif b == 0x3B: break
    else: raise SystemExit('bad block 0x%02x at %d' % (b, pos))
print('movie: %d stored frames, %d bytes' % (len(ends), len(data)))
for name, target in [('m1', 1.5), ('m2', 3.0), ('m3', 6.0)]:
    k = next(i for i, e in enumerate(ends) if e >= target * 1048576)
    buf = bytearray(data[:ends[k]]) + b'\x3B'
    for g in gce:
        if g < ends[k]:
            d = buf[g] | (buf[g + 1] << 8)
            if d > 150: buf[g] = 150; buf[g + 1] = 0     # pauses → 1.5 s
    p = os.path.join(HERE, 'probe-%s.gif' % name)
    open(p, 'wb').write(buf)
    note(name, p)

json.dump(out, open(os.path.join(HERE, 'ladder.json'), 'w'), indent=1)
