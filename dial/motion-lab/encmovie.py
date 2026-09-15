"""encmovie.py — stream the movie frames into one GIF, then prove it decodes to exactly those frames.

   python3 encmovie.py <out.gif> [quick]      quick = pauses shortened for previewing
"""
import hashlib, json, os, sys, time
from PIL import Image, ImageDraw, ImageSequence
os.chdir(os.path.dirname(os.path.abspath(__file__)))
OUT = sys.argv[1]; QUICK = len(sys.argv) > 2
BG = (26, 26, 26); EB = (478, 79, 582, 123)
items = json.load(open('movie/manifest.json'))

def dur(it):
    if not QUICK: return it['ms']
    return {'rest': 3500, 'rest-a': 1500, 'rest-b': 2000}.get(it['tag'], it['ms'])

# palette: every 40th frame across the whole film, body and eBay box quantised separately
t0 = time.time()
moving = [it for it in items if not it['tag'].startswith(('rest', 'hold'))]
keys = [Image.open(it['f']).convert('RGB') for it in moving[::max(1, len(moving) // 48)]][:48] + [Image.open(items[0]['f']).convert('RGB')]
strip = Image.new('RGB', (799, 133 * len(keys))); ebs = Image.new('RGB', (104 * len(keys), 44))
for i, k in enumerate(keys):
    im = k.copy(); ImageDraw.Draw(im).rectangle(EB, fill=BG); strip.paste(im, (0, 133 * i)); ebs.paste(k.crop(EB), (104 * i, 0))
pb = strip.quantize(colors=200, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE).getpalette()[:600]
pe = ebs.quantize(colors=48, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE).getpalette()[:144]
pal = [26, 26, 26] + pb + pe; pal += [0, 0, 0] * (256 - len(pal) // 3)
PIM = Image.new('P', (1, 1)); PIM.putpalette(pal[:768])

hashes = []
def frames():
    for it in items:
        q = Image.open(it['f']).convert('RGB').quantize(palette=PIM, dither=Image.Dither.NONE)
        hashes.append(hashlib.md5(q.convert('RGB').tobytes()).hexdigest())
        yield q

g = frames(); first = next(g)
first.save(OUT, save_all=True, append_images=g, duration=[dur(it) for it in items], loop=0, disposal=1, optimize=False)
t1 = time.time()

# verify: decoded frames == quantised source frames (identical neighbours are merged by the encoder)
expect = [hashes[0]] + [h for a, h in zip(hashes, hashes[1:]) if h != a]
got, total_ms = [], 0
for f in ImageSequence.Iterator(Image.open(OUT)):
    got.append(hashlib.md5(f.convert('RGB').tobytes()).hexdigest()); total_ms += f.info.get('duration', 0)
print(f'{OUT}: {len(items)} frames in, {len(got)} stored · {os.path.getsize(OUT)/1048576:.2f} MB · '
      f'plays {total_ms/60000:.2f} min · decoded-match={got == expect} · encode {t1-t0:.0f}s')
