# -*- coding: utf-8 -*-
"""Crop the camelcamelcamel price-history screenshot down to its PLOT
AREA, so video slide 33 can carry the axis labels as native text.

The screenshot's own labels render at about 4.6 pt on the slide, which is
unreadable in a taped video (Nico, 2026-08-24). Cropping the label
margins off lets the build draw them at 17 pt instead.

Geometry measured off the source PNG (1038 x 471):
    y-axis line          x = 41
    plot right edge      x = 983   (the $83.99 / $62.98 / $35.00 labels
                                    sit outside it, from x 988)
    x-axis line          y = 440
    y tick label centres y = 21 ($90), 160 ($70), 300 ($50), 440 ($30)
    x tick marks         x = 180, 343, 500, 663, 826  (May .. Sep)

    price -> row:  y = 440 - 6.9835 * (price - 30)

BUILD INPUT - regenerates `_source_images_video/camel_plot.png`.
"""
from pathlib import Path

from PIL import Image

HERE = Path(__file__).parent
SRC = HERE / "_source_images_video" / "ct_amazonrecent_image11.png"
DST = HERE / "_source_images_video" / "camel_plot.png"

# a little headroom above the highest price, down to just past the axis
CROP = (41, 12, 983, 443)


def main():
    im = Image.open(SRC).convert("RGB")
    out = im.crop(CROP)
    out.save(DST)
    print("%s  %dx%d" % (DST.name, out.size[0], out.size[1]))


if __name__ == "__main__":
    main()
