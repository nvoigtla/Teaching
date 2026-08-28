"""Montage the PNGs in _probe/ into _probe/_sheet.png for a single read.

Usage: python _sheet.py [cols] [cell_width]
"""
import glob
import math
import sys
from pathlib import Path

from PIL import Image

D = Path(__file__).parent / "_probe"


def main():
    cols = int(sys.argv[1]) if len(sys.argv) > 1 else 2
    cw = int(sys.argv[2]) if len(sys.argv) > 2 else 900
    files = sorted(p for p in D.glob("*.png") if p.name != "_sheet.png")
    if not files:
        print("no PNGs in _probe/")
        return
    ch = int(cw * 9 / 16)
    rows = math.ceil(len(files) / cols)
    sheet = Image.new("RGB", (cols * cw, rows * ch), "#888888")
    for i, f in enumerate(files):
        im = Image.open(f).convert("RGB").resize((cw, ch))
        sheet.paste(im, ((i % cols) * cw, (i // cols) * ch))
    out = D / "_sheet.png"
    sheet.save(out)
    print(f"{len(files)} slides -> {out} {sheet.size}")


if __name__ == "__main__":
    main()
