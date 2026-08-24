# -*- coding: utf-8 -*-
"""Which slides have speaker notes, and how substantial are they?
Poll slides are flagged: their notes carry the PollEverywhere payload and
must never be rewritten."""
import os
import sys
import zipfile

from lxml import etree as ET

import _diff_slides as D

A, P, R, REL = D.A, D.P, D.R, D.REL
HERE = os.path.dirname(os.path.abspath(__file__))
DECK = os.path.join(HERE, sys.argv[1] if len(sys.argv) > 1
                    else "Module 1 - Revised.pptx")
SPLICED = {7, 8, 24, 25, 28, 29, 49, 50}

z = zipfile.ZipFile(DECK)
pres = ET.fromstring(z.read("ppt/presentation.xml"))
prels = {r.get("Id"): r.get("Target") for r in
         ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
order = ["ppt/" + prels[s.get(D.q(R, "id"))].lstrip("/")
         for s in pres.find(D.q(P, "sldIdLst"))]

missing, short, ok = [], [], []
for disp, part in enumerate(order, 1):
    tree = ET.fromstring(z.read(part))
    title = ""
    for sp in tree.iter(D.q(P, "sp")):
        t = D.norm("".join(x.text or "" for x in sp.iter(D.q(A, "t"))))
        sppr = sp.find(D.q(P, "spPr"))
        xf = sppr.find(D.q(A, "xfrm")) if sppr is not None else None
        if xf is None:
            continue
        off = xf.find(D.q(A, "off"))
        if off is None or off.get("y") is None:
            continue
        y = int(off.get("y")) / 914400.0
        if 0.4 < y < 0.8 and t:
            title = t
            break
    notes = D.norm(D.notes_text(z, part))
    words = len(notes.split())
    tag = "POLL" if disp in SPLICED else ""
    row = (disp, words, title[:46], tag)
    if words == 0:
        missing.append(row)
    elif words < 25:
        short.append(row)
    else:
        ok.append(row)

print("=== NO NOTES (%d) ===" % len(missing))
for d, w, t, g in missing:
    print("  %3d  %-46s %s" % (d, t, g))
print("\n=== THIN, under 25 words (%d) ===" % len(short))
for d, w, t, g in short:
    print("  %3d  %3dw  %-46s %s" % (d, w, t, g))
print("\n=== substantive (%d) ===" % len(ok))
print("  " + ", ".join(str(d) for d, w, t, g in ok))
z.close()
