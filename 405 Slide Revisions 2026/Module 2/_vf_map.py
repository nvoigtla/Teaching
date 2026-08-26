# -*- coding: utf-8 -*-
"""List the slides of the three FINAL video decks (Videos Final/) beside
the slides of "Module 2 - Video Part Revised.pptx", so each of ours can
be matched to its final counterpart.

Prints, per deck, the display number, the action title, the top-bar tag,
the shape count and the number of animation clicks.

    python _vf_map.py
"""
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
VF = HERE / "Videos Final"
FINALS = [
    "Module 2 - Video 1 - Elasticity and Revenue.pptx",
    "Module 2 - Video 2 - Marginal Revenue.pptx",
    "Module 2 - Video 3 - Demand Estimation.pptx",
]
import os
OURS = os.environ.get("VF_OURS", "Module 2 - Video Part Revised.pptx")

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
EMU = 914400.0


def q(ns, t):
    return "{%s}%s" % (ns, t)


def order(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    m = {r.get("Id"): r.get("Target").split("/")[-1] for r in rels}
    return [m[s.get(R)] for s in pres.find(q(P, "sldIdLst"))]


def shape_texts(spTree):
    """(y, x, text) for every top-level shape that carries text."""
    out = []
    for c in spTree:
        tag = ET.QName(c).localname
        if tag == "AlternateContent":
            ch = c.find("{%s}Choice" % MC)
            c = ch[0] if (ch is not None and len(ch)) else None
            if c is None:
                continue
            tag = ET.QName(c).localname
        if tag not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
            continue
        pr = c.find(q(P, "spPr"))
        if pr is None:
            pr = c.find(q(P, "grpSpPr"))
        xf = c.find(q(P, "xfrm")) if tag == "graphicFrame" else (
            pr.find(q(A, "xfrm")) if pr is not None else None)
        x = y = 0.0
        if xf is not None:
            o = xf.find(q(A, "off"))
            if o is not None:
                x, y = int(o.get("x")) / EMU, int(o.get("y")) / EMU
        txt = "".join(t.text or "" for t in c.iter(q(A, "t")))
        out.append((y, x, tag, txt))
    return out


def slide_row(z, part):
    tree = ET.fromstring(z.read("ppt/slides/" + part))
    spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
    items = shape_texts(spTree)
    tag = ""
    title = ""
    for y, x, t, txt in items:
        if not txt.strip():
            continue
        if y < 0.45 and not tag:
            tag = txt.strip()
        elif 0.45 <= y < 1.25 and not title:
            title = txt.strip()
    raw = z.read("ppt/slides/" + part).decode("utf-8", "ignore")
    clicks = raw.count('nodeType="clickEffect"')
    return tag, title, len(items), clicks


def dump(path, label):
    z = zipfile.ZipFile(path)
    parts = order(z)
    print("\n" + "=" * 78)
    print("%s  (%d slides)" % (label, len(parts)))
    print("=" * 78)
    for i, pn in enumerate(parts, 1):
        tag, title, n, clicks = slide_row(z, pn)
        print("%3d | %-34s | %-46s | %2d sh | %2d clk"
              % (i, tag[:34], title[:46], n, clicks))
    z.close()


def main():
    dump(HERE / OURS, "OURS: " + OURS)
    for f in FINALS:
        dump(VF / f, "FINAL: " + f)


if __name__ == "__main__":
    main()
