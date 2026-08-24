# -*- coding: utf-8 -*-
"""Deck-wide check of the new handout rule: wherever a build paints one
picture over another, the later picture (later in document order = painted
on top) must fully contain the earlier one. Reports every violation with
the shortfall on each edge."""
import os
import sys
import zipfile

from lxml import etree as ET

import _diff_slides as D

A, P, R = D.A, D.P, D.R
EMU = 914400.0
HERE = os.path.dirname(os.path.abspath(__file__))
DECK = os.path.join(HERE, sys.argv[1] if len(sys.argv) > 1
                    else "Module 1 - Revised.pptx")
TOL = 0.005          # inches; ignore sub-hairline shortfalls
MIN_OVERLAP = 0.35   # fraction of the smaller area that must overlap


def boxes(tree):
    spTree = tree.find(".//" + D.q(P, "cSld") + "/" + D.q(P, "spTree"))
    out = []
    for c in spTree:
        if ET.QName(c).localname != "pic":
            continue
        sppr = c.find(D.q(P, "spPr"))
        xf = sppr.find(D.q(A, "xfrm")) if sppr is not None else None
        if xf is None:
            continue
        off, ext = xf.find(D.q(A, "off")), xf.find(D.q(A, "ext"))
        if off is None or ext is None:
            continue
        out.append((int(off.get("x")) / EMU, int(off.get("y")) / EMU,
                    int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU))
    return out


z = zipfile.ZipFile(DECK)
pres = ET.fromstring(z.read("ppt/presentation.xml"))
prels = {r.get("Id"): r.get("Target") for r in
         ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
order = ["ppt/" + prels[s.get(D.q(R, "id"))].lstrip("/")
         for s in pres.find(D.q(P, "sldIdLst"))]

bad = 0
checked = 0
for disp, part in enumerate(order, 1):
    bs = boxes(ET.fromstring(z.read(part)))
    if len(bs) < 2:
        continue
    for i in range(len(bs)):
        for j in range(i + 1, len(bs)):
            ax, ay, aw, ah = bs[i]      # painted first (underneath)
            bx, by, bw, bh = bs[j]      # painted later (on top)
            ox = min(ax + aw, bx + bw) - max(ax, bx)
            oy = min(ay + ah, by + bh) - max(ay, by)
            if ox <= 0 or oy <= 0:
                continue
            small = min(aw * ah, bw * bh)
            if (ox * oy) / small < MIN_OVERLAP:
                continue
            checked += 1
            short = []
            if bx - ax > TOL:
                short.append("left %.3f\"" % (bx - ax))
            if by - ay > TOL:
                short.append("top %.3f\"" % (by - ay))
            if (ax + aw) - (bx + bw) > TOL:
                short.append("right %.3f\"" % ((ax + aw) - (bx + bw)))
            if (ay + ah) - (by + bh) > TOL:
                short.append("bottom %.3f\"" % ((ay + ah) - (by + bh)))
            if short:
                bad += 1
                print("display %-3d pic %d under pic %d — NOT covered: %s"
                      % (disp, i, j, ", ".join(short)))
                print("            under: (%.3f,%.3f) %.3f x %.3f"
                      % (ax, ay, aw, ah))
                print("            over : (%.3f,%.3f) %.3f x %.3f"
                      % (bx, by, bw, bh))
print("\n%d overlapping picture pair(s) checked, %d violation(s)"
      % (checked, bad))
z.close()
