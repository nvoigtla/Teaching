"""Audit every figure in a .docx against the label-placement rules.

Reports, per figure:
  CURVE-THROUGH-LABEL   a curve segment crosses a label's box   (rule 1)
  LABEL-OVERLAP         two label boxes overlap each other
Curves are read as prstGeom="line" (endpoints reconstructed from the
bbox + flipH/flipV) and as custGeom point paths.  Guides (dashed gray)
and axes are excluded -- a dashed guide is meant to reach its label.

Usage: python ruleaudit.py <docx> [<docx> ...]
"""
import sys
import zipfile
from xml.etree import ElementTree as ET

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WPG = "{http://schemas.microsoft.com/office/word/2010/wordprocessingGroup}"
WPS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"
E = 914400.0

SKIP = ("guide", "axis", "tick mark")


def shape_text(sp):
    t = sp.find(".//" + W + "txbxContent")
    if t is None:
        return None
    return "".join(n.text or "" for n in t.iter(W + "t")).strip()


def segs_of(sp, off, scale):
    """Every line segment of this shape, in rendered inches."""
    (ox, oy), (chx, chy), (sx, sy) = off
    x = sp.find(".//" + A + "xfrm")
    if x is None:
        return []
    o, e = x.find(A + "off"), x.find(A + "ext")
    if o is None or e is None:
        return []
    bx, by = int(o.get("x")), int(o.get("y"))
    cx, cy = int(e.get("cx")), int(e.get("cy"))
    flipH = x.get("flipH") == "1"
    flipV = x.get("flipV") == "1"

    def R(px, py):
        return ((ox + (px - chx) * sx) / E, (oy + (py - chy) * sy) / E)

    prst = sp.find(".//" + A + "prstGeom")
    if prst is not None and prst.get("prst") == "line":
        if flipH == flipV:
            return [(R(bx, by), R(bx + cx, by + cy))]
        return [(R(bx + cx, by), R(bx, by + cy))]

    cust = sp.find(".//" + A + "custGeom")
    if cust is None:
        return []
    out = []
    for path in cust.iter(A + "path"):
        pw = int(path.get("w") or cx or 1)
        ph = int(path.get("h") or cy or 1)
        pts = []
        for node in path:
            pt = node.find(A + "pt")
            if pt is None:
                continue
            px = bx + int(pt.get("x")) * cx / float(pw)
            py = by + int(pt.get("y")) * cy / float(ph)
            tag = node.tag.split("}")[1]
            if tag == "moveTo":
                pts = [R(px, py)]
            else:
                pts.append(R(px, py))
        out += list(zip(pts, pts[1:]))
    return out


# A label box is bigger than its glyphs: `label()` gives the text
# 0.08" of width slack (0.04 a side) and sets a 0.22" box around ~0.152"
# of 11 pt cap height (0.034 a side).  Nico's own accepted positions let
# a curve cross that padding, so the rule is about the GLYPHS.
PAD_X, PAD_Y = 0.040, 0.034


def seg_hits_box(p, q, box, pad=None):
    """True if segment p-q enters the label's glyph area."""
    x0, y0, x1, y1 = box
    x0, y0 = x0 + PAD_X, y0 + PAD_Y
    x1, y1 = x1 - PAD_X, y1 - PAD_Y
    if x1 <= x0 or y1 <= y0:
        return False
    (ax, ay), (bx, by) = p, q
    dx, dy = bx - ax, by - ay
    t0, t1 = 0.0, 1.0
    for lo, hi, a, d in ((x0, x1, ax, dx), (y0, y1, ay, dy)):
        if abs(d) < 1e-12:
            if a < lo or a > hi:
                return False
            continue
        r0, r1 = (lo - a) / d, (hi - a) / d
        if r0 > r1:
            r0, r1 = r1, r0
        t0, t1 = max(t0, r0), min(t1, r1)
        if t0 > t1:
            return False
    return True


def audit(path):
    root = ET.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    print("=========", path)
    bad = 0
    for fi, wgp in enumerate(root.iter(WPG + "wgp")):
        gx = wgp.find(WPG + "grpSpPr").find(".//" + A + "xfrm")
        o, e = gx.find(A + "off"), gx.find(A + "ext")
        co, ce = gx.find(A + "chOff"), gx.find(A + "chExt")
        ox, oy = int(o.get("x")), int(o.get("y"))
        ex, ey = int(e.get("cx")), int(e.get("cy"))
        chx, chy = (int(co.get("x")), int(co.get("y"))) if co is not None else (0, 0)
        cex, cey = (int(ce.get("cx")), int(ce.get("cy"))) if ce is not None else (ex, ey)
        off = ((ox, oy), (chx, chy), (ex / float(cex), ey / float(cey)))

        labels, curves = [], []
        for sp in wgp.iter(WPS + "wsp"):
            nv = sp.find(WPS + "cNvPr")
            name = (nv.get("name") if nv is not None else "") or ""
            t = shape_text(sp)
            if t:
                x = sp.find(".//" + A + "xfrm")
                b, ee = x.find(A + "off"), x.find(A + "ext")
                bx, by = int(b.get("x")), int(b.get("y"))
                cx, cy = int(ee.get("cx")), int(ee.get("cy"))
                sx, sy = off[2]
                X = (ox + (bx - chx) * sx) / E
                Y = (oy + (by - chy) * sy) / E
                labels.append((name, t, (X, Y, X + cx * sx / E, Y + cy * sy / E)))
            elif not any(s in name.lower() for s in SKIP):
                curves.append((name, segs_of(sp, off, None)))

        for lname, t, box in labels:
            # A legend label sits on an opaque white badge that is drawn
            # after the gridlines, so anything behind it is hidden.
            if ("axis title" in lname or "tick" in lname
                    or "legend" in lname):
                continue
            for ci, (cname, segs) in enumerate(curves):
                for p, q in segs:
                    if seg_hits_box(p, q, box):
                        print("   fig%d  CURVE-THROUGH-LABEL  %-14s <- %s#%d"
                              % (fi, repr(t)[:14], cname, ci))
                        print("          label box (%.3f,%.3f)-(%.3f,%.3f)"
                              " seg (%.3f,%.3f)-(%.3f,%.3f)"
                              % (box[0], box[1], box[2], box[3],
                                 p[0], p[1], q[0], q[1]))
                        bad += 1
                        break
                else:
                    continue
                break
        for i in range(len(labels)):
            for j in range(i + 1, len(labels)):
                a, b = labels[i][2], labels[j][2]
                ow = min(a[2], b[2]) - max(a[0], b[0])
                oh = min(a[3], b[3]) - max(a[1], b[1])
                ow -= 2 * PAD_X
                oh -= 2 * PAD_Y
                if ow > 0.01 and oh > 0.01:
                    print("   fig%d  LABEL-OVERLAP  %-12s / %-12s  %.3f x %.3f"
                          % (fi, repr(labels[i][1])[:12],
                             repr(labels[j][1])[:12], ow, oh))
                    bad += 1
    print("   -> %d finding(s)" % bad)
    return bad


if __name__ == "__main__":
    total = sum(audit(p) for p in sys.argv[1:])
    print("TOTAL findings:", total)
