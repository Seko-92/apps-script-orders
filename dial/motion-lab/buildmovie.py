"""buildmovie.py — one renderer with every scene: round one (moves.html) merged into level.html."""
import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))
lv = open('level.html').read()
mv = open('moves.html').read()

def rep(s, a, b, n=1):
    assert s.count(a) == n, (a[:70], s.count(a))
    return s.replace(a, b)

# the coin flip needs a horizontal scale on the whole roundel, bar included
lv = rep(lv, "  const s = o.scale || 1;\n  ctx.save(); ctx.translate(130, 66.5); ctx.scale(s, s);",
            "  const s = o.scale || 1, sx = s * (o.scaleX === undefined ? 1 : o.scaleX);\n  ctx.save(); ctx.translate(130, 66.5); ctx.scale(sx, s);")
lv = rep(lv, "ctx.fillRect(130 + (BAR.x - 130) * s, 66.5 + (BAR.y - 66.5) * s, BAR.w * s, BAR.h * s);",
            "ctx.fillRect(130 + (BAR.x - 130) * sx, 66.5 + (BAR.y - 66.5) * s, BAR.w * sx, BAR.h * s);")

# ⚠ SEAMS. A scene must end exactly where the pause begins, or the ring snaps back every loop.
# conveyor: the pulley turns a whole number of turns, not whatever the belt distance happened to be
lv = rep(lv, "  base({ ringRot: -e * d / (2 * Math.PI * 58) * 360, arcRot: -e * d / (2 * Math.PI * 46) * 360 });",
            "  const tr = Math.max(1, Math.round(d / (2 * Math.PI * 58))), ta = Math.max(1, Math.round(d / (2 * Math.PI * 46)));\n"
            "  base({ ringRot: -e * tr * 360, arcRot: -e * ta * 360 });")

# round one's moves, minus conveyor (already in SIG, with the fix above)
i0 = mv.index('const MOVES = {'); i1 = mv.index('\nconst cache = {};')
block = mv[i0:i1].replace('const MOVES = {', 'const MOVES1 = {', 1)
block = rep(block, "base({ ringRot: e * 180, arcRot: -e * 180 });", "base({ ringRot: e * 360, arcRot: -e * 360 });")   # odometer seam
block = rep(block, "glyph(B, q, { dy: (1 - back(lu)) * 44 })", "glyph(B, q, { dy: (1 - back(lu, 1.5)) * 44 })")
extra = block + """
const stag = (i, n, u, spread) => cl((u - spread * (n > 1 ? i / (n - 1) : 0)) / (1 - spread));
['typewriter', 'odometer', 'decode', 'assembly', 'title'].forEach(k => { SIG[k] = MOVES1[k]; });
"""
lv = rep(lv, "\nconst cache = {};", "\n" + extra + "\nconst cache = {};")
open('movie.html', 'w').write(lv)
print('movie.html', len(lv))
