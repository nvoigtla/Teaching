# -*- coding: utf-8 -*-
"""Map each Videos-Final slide to its best-matching display in the main deck."""
import os, sys, zipfile, difflib
from lxml import etree as ET
import _diff_slides as D

HERE = os.path.dirname(os.path.abspath(__file__))
MAIN = os.path.join(HERE, "Module 1 - Revised.pptx")
VIDS = [
    "Module 1 - Video 1 - Introduction.pptx",
    "Module 1 - Video 2 - Markets.pptx",
    "Module 1 - Video 3 - Demand and Supply.pptx",
    "Module 1 - Video 4 - Equilibrium.pptx",
]

def parts(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rmap = {r.get("Id"): r.get("Target") for r in
            ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    return ["ppt/" + rmap[s.get(D.q(D.R, "id"))].lstrip("/").replace("../", "")
            for s in pres.find(D.q(D.P, "sldIdLst"))]

def sig(z, part):
    tree = ET.fromstring(z.read(part))
    sp = tree.find(".//" + D.q(D.P, "cSld") + "/" + D.q(D.P, "spTree"))
    return D.norm("".join(t.text or "" for t in sp.iter(D.q(D.A, "t"))))

zm = zipfile.ZipFile(MAIN)
mparts = parts(zm)
msigs = [sig(zm, p) for p in mparts]

for v in VIDS:
    zv = zipfile.ZipFile(os.path.join(HERE, "Videos Final", v))
    vparts = parts(zv)
    print("="*80); print(v)
    for i, p in enumerate(vparts, 1):
        s = sig(zv, p)
        best = sorted(((difflib.SequenceMatcher(None, s, m).ratio(), j+1)
                       for j, m in enumerate(msigs)), reverse=True)[:3]
        print("  v%-3d -> " % i + "  ".join("main %3d (%.3f)" % (j, r) for r, j in best)
              + "   | " + s[:70])
    zv.close()
zm.close()
