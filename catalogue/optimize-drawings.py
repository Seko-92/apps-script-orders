#!/usr/bin/env python3
"""Losslessly recompress the 1-bit drawing PNGs written by import.js --drawings.
pdftoppm writes them unoptimised; Pillow's optimize halves them (2026-10-03: 11.7 -> 5.7 MB).
Every file is checked pixel-identical before it replaces the original.
Needs Pillow:  python3 -m venv v && v/bin/pip install pillow && v/bin/python catalogue/optimize-drawings.py
"""
import glob, os, sys
from PIL import Image, ImageChops

root = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "out", "drawings")
before = after = 0
for f in glob.glob(os.path.join(root, "*", "*.png")):
    a = Image.open(f).convert("1")
    tmp = f + ".tmp.png"
    a.save(tmp, optimize=True)
    same = ImageChops.difference(a, Image.open(tmp).convert("1")).getbbox() is None
    s0, s1 = os.path.getsize(f), os.path.getsize(tmp)
    if same and s1 < s0:
        os.replace(tmp, f); after += s1
    else:
        os.remove(tmp); after += s0
    before += s0
print(f"{before/1e6:.1f} MB -> {after/1e6:.1f} MB (pixel-identical)")
