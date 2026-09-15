"""encode.py — turn a captured frame manifest into verified GIFs (row 1, 799x133).

   python3 encode.py <manifest.json> <name> <out.gif> [rest_ms]

   Palette: 200 body colours + 48 reserved for the eBay logo's box + exact #1a1a1a ground.
   Pillow crops each frame to what changed and merges identical frames. The result is decoded
   back and compared with the quantised source frames — never trust the encoder's word.
"""
import json, os, sys
from PIL import Image, ImageDraw, ImageSequence
BG = (26, 26, 26); EB = (478, 79, 582, 123)

def palette(frames):
    keys = frames[::max(1, len(frames) // 14)][:16]
    strip = Image.new('RGB', (799, 133 * len(keys)))
    for i, k in enumerate(keys):
        im = k.copy(); ImageDraw.Draw(im).rectangle(EB, fill=BG); strip.paste(im, (0, 133 * i))
    pb = strip.quantize(colors=200, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE).getpalette()[:600]
    pe = Image.new('RGB', (104 * len(keys), 44))
    for i, k in enumerate(keys): pe.paste(k.crop(EB), (104 * i, 0))
    pe = pe.quantize(colors=48, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE).getpalette()[:144]
    pal = [26, 26, 26] + pb + pe; pal += [0, 0, 0] * (256 - len(pal) // 3)
    pim = Image.new('P', (1, 1)); pim.putpalette(pal[:768]); return pim

def write(items, out, rest=3500):
    frames = [Image.open(it['f']).convert('RGB') for it in items]
    pim = palette(frames); q = [f.quantize(palette=pim, dither=Image.Dither.NONE) for f in frames]
    dur = [rest if it['ms'] == 'REST' else it['ms'] for it in items]
    q[0].save(out, save_all=True, append_images=q[1:], duration=dur, loop=0, disposal=1, optimize=False)
    dec = [f.convert('RGB').tobytes() for f in ImageSequence.Iterator(Image.open(out))]
    src = [x.convert('RGB').tobytes() for x in q]; dedup = [src[0]] + [s for a, s in zip(src, src[1:]) if s != a]
    return os.path.getsize(out) / 1024, sum(dur) / 1000, dec == dedup

if __name__ == '__main__':
    man = json.load(open(sys.argv[1])); entry = man[sys.argv[2]]
    items = entry['frames'] if isinstance(entry, dict) else entry
    kb, sec, ok = write(items, sys.argv[3], int(sys.argv[4]) if len(sys.argv) > 4 else 3500)
    print(f'{sys.argv[3]}  {kb:.0f} KB  loop {sec:.1f}s  decoded-match={ok}')
