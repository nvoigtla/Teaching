# -*- coding: utf-8 -*-
"""Nico's round-2 fixes on `Module 3 - Revised.pptx` (2026-08-27).

  1. Slide 71 title: "we will often usually use" -> "we will often use".
  2. Slide 22 bullet: strip the trailing period after "(labor)".
  3. Slides 49 and 54: the LARGE filled cards get slightly rounded
     corners and the deck's soft drop shadow.  The small navy cost /
     table cells stay flat, as Nico asked - rounding a 2.5 x 0.4"
     grid cell reads as a mistake, not as a lifted card.

Boxes are matched on GEOMETRY, not shape name (python-pptx names
content shapes "Rectangle N" too).  Corner radius is set to a constant
RENDERED 0.08" on every box, so the `adj` value is computed per shape
from its own short side - a single shared adj would give a big panel a
much rounder corner than a thin header bar.

Direct zip + lxml surgery; the deck is never round-tripped through
python-pptx.

Usage:  python _round_and_text_fixes.py [--dry-run]
"""
import shutil
import sys

from lxml import etree as ET

import _poll_chrome_pass as base

A = base.A
P = base.P
EMU = base.EMU
DECK = base.DECK
DRY = "--dry-run" in sys.argv

RADIUS_IN = 0.08          # rendered corner radius, inches

SHADOW = (
    '<a:effectLst xmlns:a="%s">'
    '<a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl"'
    ' rotWithShape="0"><a:srgbClr val="000000">'
    '<a:alpha val="45000"/></a:srgbClr></a:outerShdw>'
    "</a:effectLst>" % A
)

# (slide, left, top, w, h) of every box to round + shade, in inches.
ROUND = [
    (49, 0.47, 2.70, 6.00, 0.70),     # navy header, left column
    (49, 6.87, 2.70, 6.00, 0.70),     # navy header, right column
    (49, 0.47, 3.45, 6.00, 2.65),     # cream body panel, left
    (49, 6.87, 3.45, 6.00, 2.65),     # cream body panel, right
    (54, 0.87, 2.45, 5.50, 1.40),     # "Your own car" card
    (54, 6.97, 2.45, 5.50, 1.40),     # "Company car" card
]

# already rounded, only the missing shade:
SHADE_ONLY = [
    (54, 0.87, 4.00, 5.50, 1.50),     # "Costs associated with your car"
]


def q(ns, t):
    return base.q(ns, t)


def match(el, box):
    g = base.geom(el)
    if g is None:
        return False
    want = [v * EMU for v in box[1:]]
    return all(abs(a - b) < 0.02 * EMU for a, b in zip(g, want))


def set_shadow(spPr):
    """Replace the (usually empty) effectLst - never append a second one,
    a duplicate <a:effectLst> makes PowerPoint reject the file."""
    old = spPr.find(q(A, "effectLst"))
    new = ET.fromstring(SHADOW)
    if old is not None:
        spPr.replace(old, new)
        return
    ln = spPr.find(q(A, "ln"))
    if ln is not None:
        ln.addnext(new)
    else:
        spPr.append(new)


def round_corners(spPr, cx, cy):
    adj = int(round(RADIUS_IN * EMU / float(min(cx, cy)) * 100000))
    adj = max(1000, min(adj, 50000))
    pg = spPr.find(q(A, "prstGeom"))
    if pg is None:
        return None
    pg.set("prst", "roundRect")
    for ch in list(pg):
        pg.remove(ch)
    av = ET.SubElement(pg, q(A, "avLst"))
    gd = ET.SubElement(av, q(A, "gd"))
    gd.set("name", "adj")
    gd.set("fmla", "val %d" % adj)
    return adj


def main():
    pkg = base.Pkg(DECK)
    slides = pkg.slides()
    log = []

    # ---------------- 1 + 2: text fixes ----------------
    for n, old, new, label in (
        (71, "usually use a na", "use a na", "slide 71 title"),
        (22, None, None, "slide 22 trailing period"),
    ):
        tree = pkg.xml(slides[n - 1])
        spTree = tree.find(q(P, "cSld")).find(q(P, "spTree"))
        hit = False
        for c in base.shape_kids(spTree):
            for t in c.iter(q(A, "t")):
                s = t.text or ""
                if n == 71 and s.startswith("usually use a na"):
                    t.text = s.replace("usually use a na", "use a na", 1)
                    hit = True
                elif n == 22 and s.endswith("(labor)."):
                    t.text = s[:-1]
                    hit = True
        if hit:
            log.append("  %-26s fixed" % label)
            if not DRY:
                pkg.set_xml(slides[n - 1], tree)
        else:
            log.append("  %-26s NOT FOUND" % label)

    # ---------------- 3: rounding + shade ----------------
    todo = {}
    for box in ROUND:
        todo.setdefault(box[0], []).append(("round", box))
    for box in SHADE_ONLY:
        todo.setdefault(box[0], []).append(("shade", box))

    for n in sorted(todo):
        tree = pkg.xml(slides[n - 1])
        spTree = tree.find(q(P, "cSld")).find(q(P, "spTree"))
        done = 0
        for kind, box in todo[n]:
            found = False
            for c in base.shape_kids(spTree):
                if ET.QName(c).localname != "sp" or not match(c, box):
                    continue
                spPr = c.find(q(P, "spPr"))
                if spPr is None:
                    continue
                g = base.geom(c)
                if kind == "round":
                    adj = round_corners(spPr, g[2], g[3])
                    set_shadow(spPr)
                    log.append("  slide %-3d %.2fx%.2f  rounded (adj %d) + shade"
                               % (n, box[3], box[4], adj))
                else:
                    set_shadow(spPr)
                    log.append("  slide %-3d %.2fx%.2f  shade added"
                               % (n, box[3], box[4]))
                found = True
                done += 1
                break
            if not found:
                log.append("  slide %-3d %.2fx%.2f  NOT FOUND" % (n, box[3], box[4]))
        if done and not DRY:
            pkg.set_xml(slides[n - 1], tree)

    print("\n".join(log))
    if DRY:
        print("")
        print("(dry run - nothing written)")
        return
    tmp = DECK + ".tmp"
    pkg.write(tmp)
    shutil.move(tmp, DECK)
    print("")
    print("wrote %s" % DECK)


if __name__ == "__main__":
    main()
