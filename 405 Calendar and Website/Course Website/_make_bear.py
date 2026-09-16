# -*- coding: utf-8 -*-
"""Turn Nico's Bruin Bear into a header mark, 2026-09-12.

The file he dropped is 237x129 with an OPAQUE WHITE background and the bear
in UCLA blue (44, 113, 208). Both are wrong for the slot it has to sit in:
the panel header is a navy bar, so a white rectangle would show as a white
rectangle, and a blue bear on dark blue would disappear. He asked for the
deck's dark yellow instead.

Recolouring a 2-colour silhouette properly means recovering the ALPHA, not
just swapping pixels. Every edge pixel is a blend of white and bear-blue, so
solve that blend per pixel and keep the result as transparency -- otherwise
the bear ships with a white halo, which on a navy bar is the one thing worse
than a white box.

    P = a*BEAR + (1-a)*WHITE   ->   a = (255 - P_red) / (255 - 44)

Red is the channel with the widest separation (211 of 255), so it gives the
steadiest estimate.

Written at 4x the display height so it stays sharp on a retina screen, and
the source file is kept: it is the build input.
"""

import os
from PIL import Image

ASSETS = os.path.join(
    r"C:\Users\nvoigtla\Claude Code\Teaching\405 Calendar and Website",
    "Course Website", "assets")
SRC = os.path.join(ASSETS, "Bruin Bear.png")
OUT = os.path.join(ASSETS, "bruin-bear.png")

DARK_YELLOW = (184, 134, 11)        # #B8860B -- the deck's dark yellow
BEAR_R = 44                          # red channel of the source blue
DISPLAY_H = 16                       # the header mark's height in CSS px
SCALE = 4

im = Image.open(SRC).convert("RGBA")
w, h = im.size
src = im.load()

out = Image.new("RGBA", (w, h), (0, 0, 0, 0))
dst = out.load()
for y in range(h):
    for x in range(w):
        r, g, b, a0 = src[x, y]
        alpha = (255 - r) * 255 // (255 - BEAR_R)
        alpha = max(0, min(255, alpha))
        if a0 == 0:
            alpha = 0
        dst[x, y] = DARK_YELLOW + (alpha,)

# crop to the bear itself, so the mark is not mostly padding
box = out.getbbox()
out = out.crop(box)
print("  cropped to %dx%d (was %dx%d)" % (out.size + (w, h)))

target_h = DISPLAY_H * SCALE
target_w = max(1, round(out.width * target_h / out.height))
out = out.resize((target_w, target_h), Image.LANCZOS)
out.save(OUT, optimize=True)

print("  wrote %s  %dx%d, %.1f KB"
      % (os.path.basename(OUT), out.width, out.height,
         os.path.getsize(OUT) / 1024.0))
print("  display size: %d x %d css px" % (round(target_w / SCALE), DISPLAY_H))
