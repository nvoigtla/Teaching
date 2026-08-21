# -*- coding: utf-8 -*-
"""Locate the agenda/divider slides in Module 3 - Revised.pptx and dump
their chrome (tag, title, footer, timing presence). Read-only."""
import os
import sys
import zipfile
from pathlib import Path
from lxml import etree as ET

sys.stdout.reconfigure(encoding='utf-8')
HERE = Path(__file__).parent
DECK = HERE / "Module 3 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def q(ns, t):
    return "{%s}%s" % (ns, t)


z = zipfile.ZipFile(DECK)
data = {n: z.read(n) for n in z.namelist()}
z.close()
pres = ET.fromstring(data["ppt/presentation.xml"])
rid2t = {r.get("Id"): r.get("Target") for r in
         ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
order = [os.path.basename(rid2t[s.get(q(R, "id"))])
         for s in pres.find(q(P, "sldIdLst"))]
print("total slides:", len(order))

for disp in range(1, len(order) + 1):
    part = "ppt/slides/" + order[disp - 1]
    tree = ET.fromstring(data[part])
    texts = [t.text for t in tree.iter(q(A, "t")) if t.text]
    joined = " | ".join(texts)
    if "Picking the Right Inputs" in joined or "Lowest Price" in joined:
        has_timing = tree.find(q(P, "timing")) is not None
        n_shapes = len([c for c in tree.find(
            ".//" + q(P, "cSld") + "/" + q(P, "spTree"))
            if ET.QName(c).localname in ("sp", "pic", "graphicFrame",
                                         "grpSp", "cxnSp")])
        print("\n=== display %d  (%s)  shapes=%d  timing=%s" %
              (disp, part, n_shapes, has_timing))
        for t in texts:
            print("   ", repr(t))
