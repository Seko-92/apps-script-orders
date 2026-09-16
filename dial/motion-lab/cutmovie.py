"""cutmovie.py — re-cut the 24 fps movie from its existing frames, under Google's frame ceiling.

   Google refuses an inserted GIF past 1,000–1,099 STORED frames (measured 2026-09-16 with
   probeMovieLimits / probeMovieLimits2). Bytes and play time do not matter, so pauses are free
   (one held frame) and smoothness is kept by choosing fewer moves rather than dropping frames.

   python3 cutmovie.py <manifest.json> <out.gif> <spec> [dry]
     spec  = comma list of scene[+] in play order; '+' = the eBay light pass in the pause before it
             e.g.  scanplus,firing+,torque,kit+
     dry   = count and seam-check only, no GIF written

   ⚠ The first pause is FIRST_MS (8 s) so the banner moves soon after a reload (a GIF restarts on
     every load); the remaining 45 - 8 s is held at the END, so every gap — the loop wrap included —
     is still 45 s.
"""
import hashlib, json, os, sys, time
from PIL import Image, ImageChops, ImageDraw, ImageSequence
MAN, OUT, SPEC = sys.argv[1], sys.argv[2], sys.argv[3]
DRY = len(sys.argv) > 4
PAUSE, FIRST_MS = 45000, 8000
CEILING = 1000            # measured: 1,000 frames accepted, 1,100 refused (probeMovieLimits2)
BG = (26, 26, 26); EB = (478, 79, 582, 123)
items = json.load(open(MAN))

# ── split the long cut into blocks: [pause items] + [scene items] ──
blocks, cur = {}, None
for it in items:
    if it['tag'] in ('rest', 'rest-a'):
        cur = {'pause': [], 'scene': []}
    if it['tag'] in ('rest', 'rest-a', 'sheen', 'rest-b'):
        cur['pause'].append(it)
    else:
        cur['scene'].append(it)
        name = it['tag'].replace('-back', '')
        if name != 'hold': blocks.setdefault(name, cur)
plain = [b['pause'] for b in blocks.values() if len(b['pause']) == 1][0]
sheen = [b['pause'] for b in blocks.values() if len(b['pause']) > 1][0]

def px(path): return Image.open(path).convert('RGB')
still = px(plain[0]['f'])
for p in (sheen[0], sheen[-1]):
    assert px(p['f']).tobytes() == still.tobytes(), 'a light-pass pause does not start/end on the still frame'

seq = []
for i, tok in enumerate(SPEC.split(',')):
    name, lit = tok.rstrip('+'), tok.endswith('+')
    if name not in blocks: raise SystemExit('unknown scene ' + name + ' · have: ' + ', '.join(blocks))
    pause = [dict(x) for x in (sheen if lit else plain)]
    if i == 0:                                   # shorten the opening pause, keep the light pass intact
        cut = PAUSE - FIRST_MS
        for x in pause:                          # rest / rest-a first, then rest-b; each keeps >= 1 s
            if x['tag'] in ('rest', 'rest-a', 'rest-b') and cut > 0:
                take = min(cut, x['ms'] - 1000); x['ms'] -= take; cut -= take
        assert cut == 0 and sum(x['ms'] for x in pause) == FIRST_MS, 'could not shorten the opening pause'
    seq += pause + [dict(x) for x in blocks[name]['scene']]
seq.append({'f': plain[0]['f'], 'ms': PAUSE - FIRST_MS, 'tag': 'tail'})   # opening 8 s + tail 37 s = a normal 45 s gap at the wrap

# ── palette over THIS cut, eBay box quantised separately (same recipe as encmovie.py) ──
moving = [it for it in seq if it['tag'] not in ('rest', 'rest-a', 'rest-b', 'hold', 'tail')]
keys = [px(it['f']) for it in moving[::max(1, len(moving) // 48)]][:48] + [still]
strip = Image.new('RGB', (799, 133 * len(keys))); ebs = Image.new('RGB', (104 * len(keys), 44))
for i, k in enumerate(keys):
    im = k.copy(); ImageDraw.Draw(im).rectangle(EB, fill=BG); strip.paste(im, (0, 133 * i)); ebs.paste(k.crop(EB), (104 * i, 0))
pb = strip.quantize(colors=200, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE).getpalette()[:600]
pe = ebs.quantize(colors=48, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE).getpalette()[:144]
pal = [26, 26, 26] + pb + pe; pal += [0, 0, 0] * (256 - len(pal) // 3)
PIM = Image.new('P', (1, 1)); PIM.putpalette(pal[:768])

t0 = time.time()
q = [px(it['f']).quantize(palette=PIM, dither=Image.Dither.NONE) for it in seq]
rgb = [x.convert('RGB') for x in q]
hashes = [hashlib.md5(x.tobytes()).hexdigest() for x in rgb]
stored = 1 + sum(1 for a, b in zip(hashes, hashes[1:]) if a != b)

# ── seams: how big is the jump INTO and OUT OF every pause, and at the loop wrap? ──
def diff(a, b):
    d = ImageChops.difference(a, b).convert('L')
    return sum(i * n for i, n in enumerate(d.histogram())) / (799 * 133)
inside = sorted(diff(rgb[i], rgb[i + 1]) for i in range(len(seq) - 1)
                if seq[i]['tag'] == seq[i + 1]['tag'] and seq[i]['tag'] not in ('sheen',))
typical = inside[len(inside) // 2] if inside else 0
seams = []
for i in range(len(seq) - 1):
    if seq[i]['tag'] != seq[i + 1]['tag'] and 'hold' not in (seq[i]['tag'], seq[i + 1]['tag']):
        seams.append((seq[i]['tag'] + '→' + seq[i + 1]['tag'], diff(rgb[i], rgb[i + 1])))
seams.append(('LOOP WRAP', diff(rgb[-1], rgb[0])))
worst = max(seams, key=lambda s: s[1])
play = sum(it['ms'] for it in seq) / 60000
print(f'{SPEC}: {len(seq)} items · {stored} stored frames · plays {play:.2f} min · '
      f'worst seam {worst[0]} {worst[1]:.3f} (median in-move step {typical:.3f})')
if stored > CEILING:
    raise SystemExit(f'REFUSED: {stored} stored frames > {CEILING} — Google refuses past 1,000–1,099. Cut a move.')
if DRY: sys.exit(0)

first = q[0]
first.save(OUT, save_all=True, append_images=q[1:], duration=[it['ms'] for it in seq], loop=0, disposal=1, optimize=False)
expect = [hashes[0]] + [h for a, h in zip(hashes, hashes[1:]) if h != a]
got, total = [], 0
for f in ImageSequence.Iterator(Image.open(OUT)):
    got.append(hashlib.md5(f.convert('RGB').tobytes()).hexdigest()); total += f.info.get('duration', 0)
print(f'{OUT}: {len(got)} frames stored · {os.path.getsize(OUT)/1048576:.2f} MB · plays {total/60000:.2f} min · '
      f'decoded-match={got == expect} · encode {time.time()-t0:.0f}s')
for s in seams: print('   seam', s[0], round(s[1], 3))
