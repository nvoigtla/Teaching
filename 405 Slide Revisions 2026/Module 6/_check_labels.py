# -*- coding: utf-8 -*-
"""Find labels that a line crosses, or that a later shape covers.

2026-09-09 (Nico, on Module 6 slide 8): "the 'MR' label can't be clicked
... they should also not overlap with other lines, so that they're not
covered.  Check similar cases where labels overlap with lines everywhere."

Two failure modes, both invisible in a PNG render when the covering shape
is opaque, and both fatal to clicking the label in PowerPoint:

  COVERED  a label sits inside a filled shape that is emitted LATER, so
           the shape is painted over it and takes the mouse;
  ONLINE   a connector's segment passes through the label's box, so the
           line is drawn across the words.

Both are reported with the offending shape, so the fix is obvious.  A
label is any text shape under ~2.6 x 0.7" -- big text blocks and bullet
boxes are not labels and are skipped.

Usage:  python _check_labels.py ["Module 6 - Revised.pptx"]
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).parent
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0

SHAPE_RE = re.compile(r'<p:(sp|cxnSp|pic|graphicFrame)\b.*?</p:\1>', re.S)
NAME_RE = re.compile(r'<p:cNvPr[^>]*name="([^"]*)"')
XFRM_RE = re.compile(r'<a:xfrm([^>]*)>\s*<a:off x="(-?\d+)" y="(-?\d+)"/>'
                     r'\s*<a:ext cx="(\d+)" cy="(\d+)"/>')
T_RE = re.compile(r'<a:t(?:\s[^>]*)?>(.*?)</a:t>', re.S)
FILL_RE = re.compile(r'<a:solidFill>')
ALPHA_RE = re.compile(r'<a:alpha val="(\d+)"/>')

LABEL_MAX_W, LABEL_MAX_H = 2.60, 0.72


def parts(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = {r.get("Id"): r.get("Target") for r in
            ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    return ["ppt/" + rels[s.get("{%s}id" % R)].replace("../", "")
            for s in pres.find("{%s}sldIdLst" % P)]


def unesc(s):
    for a, b in (("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
                 ("&quot;", '"'), ("&apos;", "'")):
        s = s.replace(a, b)
    return s


def seg_hits_rect(x1, y1, x2, y2, rx, ry, rw, rh):
    """Does the segment (x1,y1)-(x2,y2) touch the rectangle?"""
    rx2, ry2 = rx + rw, ry + rh
    if max(x1, x2) < rx or min(x1, x2) > rx2:
        return False
    if max(y1, y2) < ry or min(y1, y2) > ry2:
        return False
    if abs(x2 - x1) < 1e-9 or abs(y2 - y1) < 1e-9:
        return True                       # axis-parallel, boxes overlap
    # clip the segment to the rect (Liang-Barsky)
    dx, dy = x2 - x1, y2 - y1
    t0, t1 = 0.0, 1.0
    for pp, qq in ((-dx, x1 - rx), (dx, rx2 - x1),
                   (-dy, y1 - ry), (dy, ry2 - y1)):
        if abs(pp) < 1e-12:
            if qq < 0:
                return False
            continue
        r = qq / pp
        if pp < 0:
            if r > t1:
                return False
            t0 = max(t0, r)
        else:
            if r < t0:
                return False
            t1 = min(t1, r)
    return t0 <= t1


def main():
    deck = Path(sys.argv[1]) if len(sys.argv) > 1 else \
        HERE / "Module 6 - Revised.pptx"
    z = zipfile.ZipFile(str(deck))
    covered, online, wrongend = [], [], []
    for n, part in enumerate(parts(z), 1):
        xml = z.read(part).decode("utf-8")
        body = xml[xml.find("<p:spTree"):xml.find("</p:spTree>")]
        shapes = []
        for z_i, m in enumerate(SHAPE_RE.finditer(body)):
            blk, kind = m.group(), m.group(1)
            xf = XFRM_RE.search(blk)
            if not xf:
                continue
            attrs = xf.group(1)
            x, y = int(xf.group(2)) / EMU, int(xf.group(3)) / EMU
            w, h = int(xf.group(4)) / EMU, int(xf.group(5)) / EMU
            txt = " ".join(unesc(s) for s in T_RE.findall(blk)).strip()
            alpha = ALPHA_RE.search(blk)
            shapes.append({
                "z": z_i, "kind": kind, "x": x, "y": y, "w": w, "h": h,
                "text": txt, "name": (NAME_RE.search(blk).group(1)
                                      if NAME_RE.search(blk) else "?"),
                "filled": bool(FILL_RE.search(blk)),
                "alpha": int(alpha.group(1)) if alpha else 100000,
                "flipH": 'flipH="1"' in attrs, "flipV": 'flipV="1"' in attrs,
                "head": '<a:headEnd' in blk and 'type="none"' not in
                        blk.split('<a:headEnd')[1][:40],
                "tail": '<a:tailEnd' in blk and 'type="none"' not in
                        blk.split('<a:tailEnd')[1][:40],
            })
        # An AXIS TITLE is not a candidate label here: it is anchored to
        # the arrow TIP by its own rule (Teaching CLAUDE.md), so ARROWEND
        # must never measure an annotation leader against it.  Slide 40's
        # kink callout was reported against the joint panel's "P" once the
        # callout itself moved inside a group, which is exactly the wrong
        # shape to compare.  The titles carry a name as of 2026-09-15.
        labels = [s for s in shapes
                  if s["kind"] == "sp" and s["text"]
                  and not s["name"].startswith("sdaxistitle:")
                  and s["w"] <= LABEL_MAX_W and s["h"] <= LABEL_MAX_H]
        for lb in labels:
            # COVERED: a filled shape emitted later, containing the label
            for s in shapes:
                if s["z"] <= lb["z"] or not s["filled"] or s is lb:
                    continue
                if s["kind"] not in ("sp", "pic"):
                    continue
                if s["text"]:
                    continue           # a box WITH text is its own label
                if (s["x"] <= lb["x"] and s["y"] <= lb["y"]
                        and s["x"] + s["w"] >= lb["x"] + lb["w"]
                        and s["y"] + s["h"] >= lb["y"] + lb["h"]):
                    covered.append((n, lb["text"][:26], lb["name"],
                                    s["name"], s["alpha"] // 1000))
            # ONLINE: a connector crossing the label's box
            for s in shapes:
                if s["kind"] != "cxnSp":
                    continue
                x1 = s["x"] + (s["w"] if s["flipH"] else 0)
                x2 = s["x"] + (0 if s["flipH"] else s["w"])
                y1 = s["y"] + (s["h"] if s["flipV"] else 0)
                y2 = s["y"] + (0 if s["flipV"] else s["h"])
                # the INNER area: a label box carries slack around the
                # glyphs, so a steep line clipping a corner is not a
                # strike-through.  Test the middle 60%.
                bw, bh = lb["w"] * 0.60, lb["h"] * 0.60
                bx = lb["x"] + (lb["w"] - bw) / 2.0
                by = lb["y"] + (lb["h"] - bh) / 2.0

                def _inside(px, py, m=0.10):
                    return (bx - m <= px <= bx + bw + m
                            and by - m <= py <= by + bh + m)

                # a leader line ANCHORED to this label is the house
                # pattern for a callout, not a strike-through: it is only
                # a defect when the connector passes right through, i.e.
                # neither endpoint is at the label.
                if _inside(x1, y1) or _inside(x2, y2):
                    continue
                if seg_hits_rect(x1, y1, x2, y2, bx, by, bw, bh):
                    online.append((n, lb["text"][:26], lb["name"],
                                   s["name"]))
        # ARROWEND: a label belongs at its arrow's ORIGIN, not its head
        for s in shapes:
            if s["kind"] != "cxnSp" or not (s["head"] or s["tail"]):
                continue
            # AXES are exempt: an axis title is anchored to the arrow TIP
            # by design (Teaching CLAUDE.md).  An axis is a long,
            # perfectly horizontal or vertical arrow; the decision tree's
            # Yes/No arrows are axis-parallel too but far shorter.
            axis_parallel = s["w"] < 0.02 or s["h"] < 0.02
            if axis_parallel and max(s["w"], s["h"]) > 2.5:
                continue
            # A DOUBLE-headed arrow is exempt: it measures a span rather
            # than pointing at anything, so it has no origin for a label
            # to sit at.  The practice deck's "Mark-up" arrow spans MC to
            # P* and its label belongs beside the span, either end.
            if s["head"] and s["tail"]:
                continue
            x1 = s["x"] + (s["w"] if s["flipH"] else 0)
            x2 = s["x"] + (0 if s["flipH"] else s["w"])
            y1 = s["y"] + (s["h"] if s["flipV"] else 0)
            y2 = s["y"] + (0 if s["flipV"] else s["h"])
            # OOXML draws head at the START and tail at the END, so the
            # pointed end is whichever carries a marker
            if s["head"] and not s["tail"]:
                tail, headp = (x2, y2), (x1, y1)
            else:
                tail, headp = (x1, y1), (x2, y2)
            best, bestd = None, 1e9
            for lb in labels:
                lx = lb["x"] + lb["w"] / 2.0
                ly = lb["y"] + lb["h"] / 2.0
                d = min(((lx - tail[0]) ** 2 + (ly - tail[1]) ** 2) ** 0.5,
                        ((lx - headp[0]) ** 2 + (ly - headp[1]) ** 2) ** 0.5)
                if d < bestd:
                    best, bestd = lb, d
            if best is None or bestd > 1.20:
                continue                      # no label belongs to it
            lx = best["x"] + best["w"] / 2.0
            ly = best["y"] + best["h"] / 2.0
            dt = ((lx - tail[0]) ** 2 + (ly - tail[1]) ** 2) ** 0.5
            dh = ((lx - headp[0]) ** 2 + (ly - headp[1]) ** 2) ** 0.5
            if dh < dt - 0.15:
                wrongend.append((n, best["text"][:26], s["name"],
                                 dt, dh))

    print("== ARROWEND (label nearer the arrow's HEAD than its origin) ==")
    for n, txt, sname, dt, dh in wrongend:
        print("   s%-4d %-26s  %s  origin %.2f\" vs head %.2f\""
              % (n, txt, sname, dt, dh))
    print("   none" if not wrongend else "   -> %d" % len(wrongend))
    print()
    print("== COVERED (a later filled shape sits on top of the label) ==")
    for n, txt, lname, sname, a in covered:
        print("   s%-4d %-26s  under %s  (alpha %d%%)" % (n, txt, sname, a))
    print("   none" if not covered else "   -> %d" % len(covered))
    print("\n== ONLINE (a connector is drawn through the label) ==")
    for n, txt, lname, sname in online:
        print("   s%-4d %-26s  crossed by %s" % (n, txt, sname))
    print("   none" if not online else "   -> %d" % len(online))


if __name__ == "__main__":
    main()
