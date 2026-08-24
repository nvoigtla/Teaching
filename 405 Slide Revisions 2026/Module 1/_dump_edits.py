# -*- coding: utf-8 -*-
"""Side-by-side dump of the hand-edited slides: canonical (Nico's edits)
vs the fresh build, every shape with rendered geometry in inches."""
import os
import sys

import _diff_slides as D

HERE = os.path.dirname(os.path.abspath(__file__))
CAN = os.path.join(HERE, "Module 1 - Revised.pptx")
TEST = os.path.join(HERE, "Module 1 - Revised_test.pptx")


def show(deck, disp, label):
    shapes, _ = D.dump(deck, disp)
    print("--- %s" % label)
    for i, s in enumerate(shapes):
        mark = ""
        print("  [%2d] %s%-6s (%7.3f,%7.3f) %7.3f x %7.3f | %s%s"
              % (i, "  " * s[1], s[0], s[2], s[3], s[4], s[5],
                 s[6][:52], mark))


for disp in [int(a) for a in sys.argv[1:]] or [11, 23, 26]:
    print("=" * 34, "DISPLAY", disp, "=" * 34)
    show(CAN, disp, "canonical (hand-edited)")
    show(TEST, disp, "fresh build")
