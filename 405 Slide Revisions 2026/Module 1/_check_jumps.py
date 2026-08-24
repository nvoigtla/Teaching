# -*- coding: utf-8 -*-
"""Resolve every slide-jump action to its destination display number."""
import os
import sys
import zipfile

from lxml import etree as ET

import _diff_slides as D

A, P, R, REL = D.A, D.P, D.R, D.REL
HERE = os.path.dirname(os.path.abspath(__file__))
DECK = os.path.join(HERE, sys.argv[1] if len(sys.argv) > 1
                    else "Module 1 - Revised.pptx")

z = zipfile.ZipFile(DECK)
pres = ET.fromstring(z.read("ppt/presentation.xml"))
prels = {r.get("Id"): r.get("Target")
         for r in ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
order = ["ppt/" + prels[s.get(D.q(R, "id"))].lstrip("/").replace("../", "")
         for s in pres.find(D.q(P, "sldIdLst"))]
disp_of = {p: i + 1 for i, p in enumerate(order)}

total = 0
for disp, part in enumerate(order, 1):
    tree = ET.fromstring(z.read(part))
    relp = part.replace("slides/", "slides/_rels/") + ".rels"
    try:
        srels = {r.get("Id"): (r.get("Type"), r.get("Target"))
                 for r in ET.fromstring(z.read(relp))}
    except KeyError:
        srels = {}
    for sp in tree.iter(D.q(P, "sp")):
        geom = sp.find(".//" + D.q(A, "prstGeom"))
        prst = geom.get("prst") if geom is not None else "-"
        txt = D.norm("".join(t.text or "" for t in sp.iter(D.q(A, "t"))))
        for hl in sp.iter(D.q(A, "hlinkClick")):
            if "sldjump" not in (hl.get("action") or ""):
                continue
            rid = hl.get(D.q(R, "id"))
            ty, tgt = srels.get(rid, ("?", "?"))
            # slide rels are relative to ppt/slides/
            tp = ("ppt/slides/" + tgt.split("/")[-1]
                  if tgt != "?" else "?")
            total += 1
            print("display %-3d %-22s %-38s -> display %s"
                  % (disp, prst, txt[:38] or "(no text)",
                     disp_of.get(tp, "UNRESOLVED %s" % tgt)))
print("\n%d slide-jump actions" % total)
z.close()
