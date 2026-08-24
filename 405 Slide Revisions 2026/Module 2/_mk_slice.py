# -*- coding: utf-8 -*-
"""Cut a single wedge (one slice) out of the whole-pizza photo image14.jpeg
for slide 10 (Nico, 2026-08-23: "only show one slice of pizza").
Output: _source_images/pizza_slice.png — RGBA, tight crop, transparent
background, upsampled 2.5x + mild unsharp so a 4"-wide placement holds up.
BUILD INPUT generator — rerunnable, deterministic."""
import os
import sys
from PIL import Image, ImageDraw, ImageFilter

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "_source_images", "image14.jpeg")
DST = os.path.join(HERE, "_source_images", "pizza_slice.png")

# pizza disc in the source photo, measured off a grid overlay (700x525)
CX, CY = 362.0, 257.0
RX, RY = 238.0, 203.0
SCALE = 2.5          # upsample before masking (source is only 700x525)
SUP = 4              # mask supersampling, for a clean wedge edge
HALF = 28.0          # half-angle of the wedge, degrees
A0 = 300.0           # wedge bisector, degrees cw from 3 o'clock (y down)


def build(out=DST, preview=None):
    im = Image.open(SRC).convert("RGBA")
    im = im.resize((int(im.width * SCALE), int(im.height * SCALE)),
                   Image.LANCZOS)
    im = im.filter(ImageFilter.UnsharpMask(radius=2.0, percent=60, threshold=3))
    W, H = im.size
    mask = Image.new("L", (W * SUP, H * SUP), 0)
    d = ImageDraw.Draw(mask)
    s = SCALE * SUP
    d.pieslice([(CX - RX) * s, (CY - RY) * s, (CX + RX) * s, (CY + RY) * s],
               A0 - HALF, A0 + HALF, fill=255)
    mask = mask.resize((W, H), Image.LANCZOS)
    mask = mask.filter(ImageFilter.GaussianBlur(1.2))
    im.putalpha(mask)
    im = im.crop(im.getbbox())
    im.save(out)
    print("wrote", out, im.size)
    if preview:
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        bg.alpha_composite(im)
        bg.convert("RGB").save(preview)
    return im


if __name__ == "__main__":
    build(preview=sys.argv[1] if len(sys.argv) > 1 else None)
