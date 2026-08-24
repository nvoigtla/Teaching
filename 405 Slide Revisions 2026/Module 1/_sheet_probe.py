# -*- coding: utf-8 -*-
"""Contact sheet of the _probe renders (one image, low context cost)."""
import glob
import os
import sys

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
PAT = sys.argv[1] if len(sys.argv) > 1 else "Module 1 - Revised_test_s*.png"
OUT = sys.argv[2] if len(sys.argv) > 2 else "_probe_sheet_links.png"
COLS = int(sys.argv[3]) if len(sys.argv) > 3 else 2
CELL_W = int(sys.argv[4]) if len(sys.argv) > 4 else 760

files = sorted(glob.glob(os.path.join(HERE, "_probe", PAT)))
if not files:
    raise SystemExit("no renders match %s" % PAT)
ims = []
for f in files:
    im = Image.open(f).convert("RGB")
    h = int(im.height * CELL_W / im.width)
    ims.append((os.path.basename(f), im.resize((CELL_W, h), Image.LANCZOS)))
cell_h = max(im.height for _, im in ims) + 22
rows = (len(ims) + COLS - 1) // COLS
sheet = Image.new("RGB", (COLS * (CELL_W + 8) + 8,
                          rows * (cell_h + 8) + 8), "white")
d = ImageDraw.Draw(sheet)
for i, (name, im) in enumerate(ims):
    cx = 8 + (i % COLS) * (CELL_W + 8)
    cy = 8 + (i // COLS) * (cell_h + 8)
    d.text((cx, cy + 4), name.replace("Module 1 - Revised_test_", ""),
           fill="black")
    sheet.paste(im, (cx, cy + 20))
    d.rectangle([cx, cy + 20, cx + CELL_W, cy + 20 + im.height],
                outline=(190, 190, 190))
sheet.save(os.path.join(HERE, "_probe", OUT))
print("saved _probe/%s  (%d renders)" % (OUT, len(ims)))
