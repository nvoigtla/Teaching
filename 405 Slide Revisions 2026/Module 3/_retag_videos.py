# -*- coding: utf-8 -*-
"""Splice the video number into every taped slide's top-bar tag.

2026-08-27, Nico: in a taped module the tag reads
`Module N * Video k * <topic>`, so a student sees which video a slide
came from as well as which topic it teaches:

    Module 3 * Course Roadmap        -> Module 3 * Video 1 * Course Roadmap
    Module 3 * The Production Fn     -> Module 3 * Video 2 * The Production Fn
    Module 3 * Agenda                -> Module 3 * Video 2 * Agenda

The video BLOCKS are read off the deck itself - a video title card is a
slide carrying the gold "Module 3  *  Video k" line - so the boundaries
cannot drift out of sync with the cards.  The `<topic>` level is
whatever the slide already said; only the video level is spliced in.

Exempt, per Nico:
  * the CONCEPT MAP, in every copy (`Module 3 * Concept Map`) - a
    reference slide belonging to no single video;
  * the SUMMARY closer (`Module 3 * Summary`);
  * anything already tagged `In Class * Examples` (the appendix block);
  * every slide with no top bar - the deck title slide, the video title
    cards, the section / appendix dividers, the PollEverywhere slides
    and the BACKUP slides (full-bleed figures with a caption).

The label is measured in the real font against the tag box width, so a
tag that would run past the box fails loudly instead of shipping.

Usage:  python _retag_videos.py [--dry-run]
"""
import re
import shutil
import sys

from lxml import etree as ET

import _poll_chrome_pass as base
from _retag_inclass import measure

A = base.A
P = base.P
EMU = base.EMU
DECK = base.DECK
DRY = "--dry-run" in sys.argv

from _deck_guard import require_committed

MODULE = "Module 3"
SEP = "·"
PREFIX = MODULE + " " + SEP + " "

CARD_RE = re.compile(re.escape(MODULE) + r"\s+" + SEP + r"\s+Video\s+(\d+)")

# tags that keep their two-level form wherever they appear
EXEMPT_TAILS = ("Concept Map", "Summary")


def q(ns, t):
    return base.q(ns, t)


def top_bar_tag(spTree):
    """The tag text box sitting ON the navy top bar.

    A backup slide is a full-bleed figure with a caption near the top,
    which a geometry-only test mistakes for a tag - so require the navy
    bar itself to be present before looking for the label.
    """
    has_bar = False
    for c in base.shape_kids(spTree):
        if ET.QName(c).localname != "sp":
            continue
        g = base.geom(c)
        if g is None:
            continue
        spPr = c.find(q(P, "spPr"))
        if spPr is None:
            continue
        fill = spPr.find(q(A, "solidFill"))
        clr = fill.find(q(A, "srgbClr")) if fill is not None else None
        if (clr is not None and clr.get("val") == "0B2B4E"
                and g[1] / EMU < 0.05 and g[2] / EMU > 12.0):
            has_bar = True
            break
    if not has_bar:
        return None
    for c in base.shape_kids(spTree):
        if ET.QName(c).localname != "sp":
            continue
        g = base.geom(c)
        if g is None or not base.txt(c):
            continue
        if g[1] / EMU < 0.55 and g[2] / EMU > 4.0:
            return c
    return None


def card_video(spTree):
    """The video number if this slide is a video title card, else None."""
    for c in base.shape_kids(spTree):
        m = CARD_RE.search(base.txt(c))
        if m:
            return int(m.group(1))
    return None


def main():
    if not DRY:
        require_committed(DECK)
    pkg = base.Pkg(DECK)
    slides = pkg.slides()
    log = []
    skipped = []
    problems = []
    changed = 0
    video = None

    for n, pn in enumerate(slides, 1):
        tree = pkg.xml(pn)
        spTree = tree.find(q(P, "cSld")).find(q(P, "spTree"))

        k = card_video(spTree)
        if k is not None:
            video = k
            skipped.append("%d: video %d title card" % (n, k))
            continue

        tag = top_bar_tag(spTree)
        if tag is None:
            skipped.append("%d: no top bar" % n)
            continue

        runs = list(tag.iter(q(A, "r")))
        if len(runs) != 1:
            problems.append("slide %d: tag has %d runs, expected 1"
                            % (n, len(runs)))
            continue
        t = runs[0].find(q(A, "t"))
        old = t.text or ""

        if " " + SEP + " Video " in old or " " + SEP + " In Class " + SEP in old:
            skipped.append("%d: already %s" % (n, old))
            continue
        if any(old.endswith(tail) for tail in EXEMPT_TAILS):
            skipped.append("%d: exempt (%s)" % (n, old))
            continue
        if not old.startswith(PREFIX):
            problems.append("slide %d: tag %r does not start with %r"
                            % (n, old, PREFIX))
            continue
        if video is None:
            problems.append("slide %d: tagged %r but no video card seen yet"
                            % (n, old))
            continue

        new = PREFIX + "Video %d " % video + SEP + " " + old[len(PREFIX):]

        rPr = runs[0].find(q(A, "rPr"))
        pt = int(rPr.get("sz")) / 100.0 if (rPr is not None and rPr.get("sz")) else 16.0
        box_w = base.geom(tag)[2] / EMU
        w = measure(new, pt)
        if w is not None and w > box_w - 0.1:
            problems.append("slide %d: tag %.2f\" at %g pt, box is %.2f\""
                            % (n, w, pt, box_w))
            continue

        t.text = new
        changed += 1
        log.append("  %3d  %s" % (n, new))
        if not DRY:
            pkg.set_xml(pn, tree)

    print("\n".join(log))
    print("")
    print("tags rewritten: %d" % changed)
    print("untouched     : %d" % len(skipped))
    for s in skipped:
        print("    " + s)
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
