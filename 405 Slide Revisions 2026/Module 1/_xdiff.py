# -*- coding: utf-8 -*-
"""Cross-deck member-level dump: compare a slide in one deck against a
slide in another. Usage: _xdiff.py deckA dispA deckB dispB"""
import sys

import _diff_slides as D


def show(deck, disp, label):
    shapes, notes = D.dump(deck, disp)
    print("--- %s : %s display %d  (%d shapes)"
          % (label, deck, disp, len(shapes)))
    for i, s in enumerate(shapes):
        print("  [%2d] %s%-6s (%7.3f,%7.3f) %7.3f x %7.3f | %s | %s"
              % (i, "  " * s[1], s[0], s[2], s[3], s[4], s[5],
                 s[6][:56], s[7][:34]))
    return notes


nA = show(sys.argv[1], int(sys.argv[2]), "A")
print()
nB = show(sys.argv[3], int(sys.argv[4]), "B")
print()
if D.norm(nA) != D.norm(nB):
    print("NOTES DIFFER")
    print("  A:", D.norm(nA)[:400])
    print("  B:", D.norm(nB)[:400])
else:
    print("notes identical")
