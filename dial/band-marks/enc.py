"""enc.py — manifest.json → band-direct-v1.gif + band-parcel-v1.gif, verified.

   ⚠ The band's yellow #ffd400 is RESERVED at palette index 0, so every background pixel is
     exactly the band colour — an approximate yellow would show a seam around the GIF.
   ⚠ Both loops are the SAME length (LOOP_MS) and the parcel's action is placed ~22 s after the
     truck's, so the two bands never move at the same moment.
   Decodes the GIF back and compares it with the quantised frames — never trust the encoder.
"""
import json, os
from PIL import Image, ImageSequence
YEL = (255, 212, 0)
LOOP_MS = 51660
PLAN = {   # which: (rest before the action, out file)
  'direct': (6000,  'band-direct-v1.gif'),
  'parcel': (28000, 'band-parcel-v1.gif'),
}
man = json.load(open('manifest.json'))
for which, (lead, out) in PLAN.items():
    frames = [Image.open(f).convert('RGB') for f in man[which]]
    # ⚠ headless Chrome's canvas lands the ground one step off (254,211,0). Snap every pixel
    #   within 6 of the band yellow to it EXACTLY, or the GIF shows a box edge on the band.
    def snap(im):
        px = im.load()
        for y in range(im.height):
            for x in range(im.width):
                r, g, b_ = px[x, y]
                if abs(r - 255) + abs(g - 212) + abs(b_) <= 6: px[x, y] = YEL
        return im
    frames = [snap(f) for f in frames]
    strip = Image.new('RGB', (frames[0].width, frames[0].height * len(frames)))
    for i, f in enumerate(frames): strip.paste(f, (0, i * f.height))
    adapt = strip.quantize(colors=255, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE).getpalette()[:765]
    pal = list(YEL) + adapt; pal += [0] * (768 - len(pal))
    pim = Image.new('P', (1, 1)); pim.putpalette(pal)
    q = [f.quantize(palette=pim, dither=Image.Dither.NONE) for f in frames]
    action = 1000 / 24
    tail = LOOP_MS - lead - action * (len(frames) - 1)
    dur = [lead] + [round(action)] * (len(frames) - 2) + [round(tail)]
    q[0].save(out, save_all=True, append_images=q[1:], duration=dur, loop=0, disposal=1, optimize=False)
    dec = [f.convert('RGB').tobytes() for f in ImageSequence.Iterator(Image.open(out))]
    src = [x.convert('RGB').tobytes() for x in q]; dedup = [src[0]] + [s for a, s in zip(src, src[1:]) if s != a]
    corner = Image.open(out).convert('RGB').getpixel((0, 0))
    print(f'{out}: {os.path.getsize(out)/1024:.0f} KB · {len(dec)} stored frames · loop {sum(dur)/1000:.1f}s · '
          f'decoded-match={dec == dedup} · corner={corner} (band {YEL})')
