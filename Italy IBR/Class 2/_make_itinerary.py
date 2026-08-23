"""
Build `_source_images/itinerary_2026.png` from the Legacy Ventures PDF.

Two edits are applied to the rendered grid, both by moving pixels that are
already there -- no text is re-typed, so the table's own font, colour and
weight survive exactly:

  1. **Wednesday 11.00 - 12.30 resolved to Eataly** (confirmed 2026-08-22).
     The cell read "11.00 - 12.30 / Dolce&Gabbana / Luxury Fashion / or /
     Eataly / Food & Beverage".  The time line and the "Eataly / Food &
     Beverage" block are lifted out as image patches, the cell interior is
     cleared, and the two patches are pasted back as a vertically centred
     stack -- so the cell now reads "11.00 - 12.30 / Eataly / Food & Beverage".
  2. **Saturday column dropped** -- the students fly home that day, so the
     schedule runs Sep 6 - 11.  The crop stops on the Friday|Saturday rule,
     which becomes the table's new right border.

Nothing is guessed: the grid rules are found by scanning for dark rows and
columns, the cell is located by the rules that bracket the "Dolce&Gabbana"
span, and every text position comes from the PDF's own text spans (PyMuPDF)
scaled by RENDER.

    python _make_itinerary.py
"""

import os

import fitz
import numpy as np
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
PDF = os.path.join(HERE, "UCLA Italy 2026 - Itinerary Latest.pdf")
OUT = os.path.join(HERE, "_source_images", "itinerary_2026.png")
RENDER = 3.2                      # PDF points -> pixels
CROP_L, CROP_T, CROP_B = 126, 168, 1714     # table edges in the full render


def pt(v):
    return int(round(v * RENDER))


def group(vals, gap=2):
    out, cur = [], [vals[0]]
    for v in vals[1:]:
        if v - cur[-1] <= gap:
            cur.append(v)
        else:
            out.append(cur)
            cur = [v]
    out.append(cur)
    return [sum(g) // len(g) for g in out]


def main():
    doc = fitz.open(PDF)
    page = doc[0]
    pm = page.get_pixmap(matrix=fitz.Matrix(RENDER, RENDER))
    im = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)
    a = np.array(im).astype(int)
    dark = a.sum(axis=2) < 560

    # ---- vertical rules -------------------------------------------------
    top, bot = pt(58), pt(535)
    h = bot - top
    cols = group([x for x in range(a.shape[1])
                  if dark[top:bot, x].sum() > 0.55 * h])
    assert len(cols) == 9, "expected 9 vertical rules, found %d" % len(cols)
    wed_l, wed_r = cols[4], cols[5]          # Wednesday column
    fri_sat = cols[7]                        # Friday | Saturday

    # ---- horizontal rules, below the blue header band -------------------
    w = wed_r - wed_l
    rows = [y for y in group([y for y in range(a.shape[0])
                              if dark[y, wed_l + 3:wed_r - 3].sum()
                              > 0.85 * (w - 6)]) if y >= pt(80)]

    # ---- every text span in the Wednesday column -----------------------
    wed_spans = []
    for blk in page.get_text("dict")["blocks"]:
        for line in blk.get("lines", []):
            for sp in line["spans"]:
                x0, y0, x1, y1 = sp["bbox"]
                if wed_l / RENDER < x0 < wed_r / RENDER:
                    wed_spans.append((sp["text"].strip(), y0, y1))

    # locate the cell by the rules bracketing the text we are replacing
    anchor = next(y0 for (t, y0, _) in wed_spans if t.startswith("Dolce"))
    cell_top = max(y for y in rows if y < pt(anchor))
    cell_bot = min(y for y in rows if y > pt(anchor))

    spans = {t: (y0, y1) for (t, y0, y1) in wed_spans
             if cell_top / RENDER < y0 < cell_bot / RENDER}
    time_key = next(k for k in spans if k.startswith("11.00"))
    for need in ("Dolce&Gabbana", "or", "Eataly", "Food & Beverage"):
        assert need in spans, "span %r not found in the Wednesday cell" % need

    time_y0, time_y1 = spans[time_key]
    or_y1 = spans["or"][1]
    eat_y0 = spans["Eataly"][0]
    food_y1 = spans["Food & Beverage"][1]
    gap = spans["Dolce&Gabbana"][0] - time_y1        # the cell's own line gap

    # ---- lift the two blocks we keep, clear the cell, paste them back ---
    t_box = (wed_l + 3, pt(time_y0) - 12, wed_r - 3, pt(time_y1) + 8)
    e_box = (wed_l + 3, max(pt(or_y1) + 2, pt(eat_y0) - 12),
             wed_r - 3, pt(food_y1) + 6)
    patch_time = im.crop(t_box)
    patch_eat = im.crop(e_box)

    ink = (pt(time_y1) - pt(time_y0)) + pt(gap) + (pt(food_y1) - pt(eat_y0))
    ink_top = (cell_top + cell_bot) // 2 - ink // 2
    time_dst = ink_top - (pt(time_y0) - t_box[1])
    eat_dst = (ink_top + (pt(time_y1) - pt(time_y0)) + pt(gap)
               - (pt(eat_y0) - e_box[1]))

    bg = tuple(im.getpixel((wed_l + 8, cell_bot - 8)))   # the cell's own white
    im.paste(bg, (wed_l + 3, cell_top + 3, wed_r - 3, cell_bot - 3))
    im.paste(patch_time, (t_box[0], time_dst))
    im.paste(patch_eat, (e_box[0], eat_dst))

    # ---- crop: table only, Saturday dropped ----------------------------
    out = im.crop((CROP_L, CROP_T, fri_sat + 3, CROP_B))
    out.save(OUT)
    print("wrote %s  %dx%d  (aspect %.4f)"
          % (OUT, out.width, out.height, out.width / out.height))
    print("  Wednesday cell rows %d-%d: time -> y %d, Eataly block -> y %d"
          % (cell_top, cell_bot, time_dst, eat_dst))
    print("  Saturday dropped at the Friday|Saturday rule, x=%d" % fri_sat)


if __name__ == "__main__":
    main()
