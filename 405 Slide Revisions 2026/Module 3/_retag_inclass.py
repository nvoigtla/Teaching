# -*- coding: utf-8 -*-
"""Re-tag the in-class example slides of `Module 3 - Revised.pptx`.

2026-08-27, Nico: slides 95 - 110 (the block after the concept map in the
"SLIDES NOT USED IN THE VIDEOS" appendix) will be shown as examples in the
on-campus class, so their top-bar tag gains an "In Class * Examples" level
between the module and the topic:

    Module 3 * Wage Searchers
    -> Module 3 * In Class * Examples * Wage Searchers

The topic level is whatever the slide already said, so nothing is retyped
per slide - the new level is spliced into the existing string.

Slides 93 (the appendix divider, no top bar) and 94 (concept map) keep
their tags, as Nico specified "from slide 95 onward".

The label is measured in the real font (PIL ImageFont, Calibri Bold at
the tag's own size) against the tag box width, so a tag that would run
past the box fails loudly instead of shipping.

Direct zip + lxml surgery; no python-pptx round-trip.

Usage:  python _retag_inclass.py [--dry-run]
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

FIRST, LAST = 95, 110
SEP = "\u00b7"                       # the deck's middle dot
PREFIX = "Module 3 " + SEP + " "
INSERT = "In Class " + SEP + " Examples " + SEP + " "

try:
    from PIL import ImageFont
except ImportError:
    ImageFont = None

FONTS = [r"C:\Windows\Fonts\calibrib.ttf", r"C:\Windows\Fonts\Carlito-Bold.ttf"]


def q(ns, t):
    return base.q(ns, t)


def top_bar_tag(spTree):
    """The top-bar tag text box: a text shape in the navy band."""
    for c in base.shape_kids(spTree):
        if ET.QName(c).localname != "sp":
            continue
        g = base.geom(c)
        if g is None or not base.txt(c):
            continue
        if g[1] / EMU < 0.55 and g[2] / EMU > 4.0:
            return c
    return None


def measure(text, pt):
    if ImageFont is None:
        return None
    for path in FONTS:
        try:
            f = ImageFont.truetype(path, int(round(pt * 4)))
        except OSError:
            continue
        return f.getlength(text) / 4.0 / 72.0     # inches
    return None


def main():
    pkg = base.Pkg(DECK)
    slides = pkg.slides()
    log = []
    changed = 0
    problems = []

    for n in range(FIRST, LAST + 1):
        pn = slides[n - 1]
        tree = pkg.xml(pn)
        spTree = tree.find(q(P, "cSld")).find(q(P, "spTree"))
        tag = top_bar_tag(spTree)
        if tag is None:
            problems.append("slide %d: no top-bar tag found" % n)
            continue
        runs = list(tag.iter(q(A, "r")))
        if len(runs) != 1:
            problems.append("slide %d: tag has %d runs, expected 1"
                            % (n, len(runs)))
            continue
        t = runs[0].find(q(A, "t"))
        old = t.text or ""
        if INSERT in old:
            log.append("  slide %-4d already tagged" % n)
            continue
        if not old.startswith(PREFIX):
            problems.append("slide %d: tag %r does not start with %r"
                            % (n, old, PREFIX))
            continue
        new = PREFIX + INSERT + old[len(PREFIX):]

        # does it still fit the tag box?
        sz = runs[0].find(q(A, "rPr"))
        pt = int(sz.get("sz")) / 100.0 if (sz is not None and sz.get("sz")) else 16.0
        g = base.geom(tag)
        box_w = g[2] / EMU
        w = measure(new, pt)
        if w is not None and w > box_w - 0.1:
            problems.append("slide %d: tag %.2f\" wide at %g pt, box is %.2f\""
                            % (n, w, pt, box_w))
            continue

        t.text = new
        changed += 1
        log.append("  slide %-4d %s" % (n, new))
        if not DRY:
            pkg.set_xml(pn, tree)

    print("\n".join(log))
    print("")
    print("tags rewritten: %d" % changed)
    if problems:
        print("")
        print("PROBLEMS:")
        for p in problems:
            print("  " + p)
        sys.exit(1)
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
