# -*- coding: utf-8 -*-
"""Zoomed crops of the two spots that need checking: the inline button's
vertical alignment on slide 2, and whether slide 12's jump pill collides
with the podcast link label."""
import os

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
PD = os.path.join(HERE, "_probe")
SLIDE_W_IN = 13.3333

CROPS = [
    ("Module 1 - Revised_test_s02.png", "s02: navy jump symbol (30% smaller)",
     7.9, 4.25, 11.2, 5.00),
    ("Module 1 - Revised_test_s09.png", "s09: navy jump symbol",
     9.3, 3.65, 12.4, 4.40),
    ("Module 1 - Revised_test_s12.png", "s12: navy jump pill + gold podcast marker",
     0.20, 6.38, 10.10, 7.10),
    ("Module 1 - Revised_test_s17.png", "s17: navy jump pill",
     0.20, 6.38, 5.20, 7.10),
]

tiles = []
for fn, label, x0, y0, x1, y1 in CROPS:
    im = Image.open(os.path.join(PD, fn)).convert("RGB")
    ppi = im.width / SLIDE_W_IN
    box = (int(x0 * ppi), int(y0 * ppi), int(x1 * ppi), int(y1 * ppi))
    crop = im.crop(box)
    scale = 1000.0 / crop.width
    crop = crop.resize((1000, max(1, int(crop.height * scale))),
                       Image.LANCZOS)
    tiles.append((label, crop))

H = sum(t.height + 26 for _, t in tiles) + 10
sheet = Image.new("RGB", (1010, H), "white")
d = ImageDraw.Draw(sheet)
y = 6
for label, t in tiles:
    d.text((6, y), label, fill="black")
    sheet.paste(t, (6, y + 18))
    d.rectangle([6, y + 18, 6 + t.width, y + 18 + t.height],
                outline=(180, 180, 180))
    y += t.height + 26
out = os.path.join(PD, "_probe_crops.png")
sheet.save(out)
print("saved", out)
