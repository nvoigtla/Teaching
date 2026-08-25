# -*- coding: utf-8 -*-
"""Audit a built deck for FILLED boxes that are missing the deck's soft
drop shadow.  Every cream / gold / navy card is supposed to read as a
lifted card (Teaching CLAUDE.md); the flat chrome is excluded by
geometry, and table-cell fills live inside graphicFrames, not shapes.

    python _shade_audit.py "Module 2 - Video Part Revised.pptx"
"""
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
EMU = 914400.0

# fills that mark a CONTENT card (chrome uses navy / gold / rule gray too,
# so chrome is excluded by geometry below, not by color)
CARD_FILLS = {"FDF6E6", "F6E8C9", "E09F3E", "0B2B4E", "B0B5BC"}


def q(ns, t):
    return "{%s}%s" % (ns, t)


def order(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    m = {r.get("Id"): r.get("Target").split("/")[-1] for r in rels}
    return [m[s.get(R)] for s in pres.find(q(P, "sldIdLst"))]


def is_chrome(x, y, w, h):
    """Top bar, title rule, gold strips, footer rule, page number."""
    if h < 0.09:                       # the thin rules and accent strips
        return True
    if y < 0.5 and w > 12.0:           # navy top bar
        return True
    if y > 7.1:                        # anything in the footer band
        return True
    return False


def walk(el, out, off=(0.0, 0.0), scale=(1.0, 1.0), choff=(0.0, 0.0)):
    for ch in el:
        tag = ET.QName(ch).localname
        if tag == "grpSp":
            xf = ch.find(q(P, "grpSpPr") + "/" + q(A, "xfrm"))
            if xf is None:
                continue
            o, e = xf.find(q(A, "off")), xf.find(q(A, "ext"))
            co, ce = xf.find(q(A, "chOff")), xf.find(q(A, "chExt"))
            gx, gy = int(o.get("x")) / EMU, int(o.get("y")) / EMU
            gw, gh = int(e.get("cx")) / EMU, int(e.get("cy")) / EMU
            cx, cy = int(co.get("x")) / EMU, int(co.get("y")) / EMU
            cw, chh = int(ce.get("cx")) / EMU, int(ce.get("cy")) / EMU
            sx = gw / cw if cw else 1.0
            sy = gh / chh if chh else 1.0
            walk(ch.find(q(P, "grpSp")) if False else ch, out,
                 (off[0] + (gx - cx * sx), off[1] + (gy - cy * sy)),
                 (scale[0] * sx, scale[1] * sy), (cx, cy))
            continue
        if tag != "sp":
            continue
        spPr = ch.find(q(P, "spPr"))
        if spPr is None:
            continue
        fill = spPr.find(q(A, "solidFill") + "/" + q(A, "srgbClr"))
        if fill is None or fill.get("val", "").upper() not in CARD_FILLS:
            continue
        xf = spPr.find(q(A, "xfrm"))
        if xf is None:
            continue
        o, e = xf.find(q(A, "off")), xf.find(q(A, "ext"))
        x = off[0] + int(o.get("x")) / EMU * scale[0]
        y = off[1] + int(o.get("y")) / EMU * scale[1]
        w = int(e.get("cx")) / EMU * scale[0]
        h = int(e.get("cy")) / EMU * scale[1]
        if is_chrome(x, y, w, h):
            continue
        shadow = spPr.find(q(A, "effectLst") + "/" + q(A, "outerShdw"))
        txt = "".join(t.text or "" for t in ch.iter(q(A, "t")))[:46]
        out.append((fill.get("val").upper(), shadow is not None,
                    x, y, w, h, txt))


def main():
    deck = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
        "Module 2 - Video Part Revised.pptx")
    z = zipfile.ZipFile(deck)
    bad = 0
    for disp, pn in enumerate(order(z), 1):
        tree = ET.fromstring(z.read("ppt/slides/" + pn))
        spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
        found = []
        walk(spTree, found)
        for col, has, x, y, w, h, txt in found:
            if has:
                continue
            bad += 1
            print("s%02d  #%s  (%.2f, %.2f) %.2f x %.2f  %r"
                  % (disp, col, x, y, w, h, txt))
    z.close()
    print("\n%s: %d filled box(es) without a shade" % (deck.name, bad))


if __name__ == "__main__":
    main()
