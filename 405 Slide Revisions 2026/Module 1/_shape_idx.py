# -*- coding: utf-8 -*-
"""Print the selector index (_animate.py's own numbering) for every
top-level shape on a slide, so per-slide plans can be written against the
real indices instead of a guess."""
import os
import sys
import zipfile

from lxml import etree as ET

sys.argv = [sys.argv[0]] + sys.argv[1:]          # keep DECK arg for _animate
import _animate as AN

DECK = AN.DECK
disp = int(sys.argv[2]) if len(sys.argv) > 2 else 36

z = zipfile.ZipFile(DECK)
data = {n: z.read(n) for n in z.namelist()}
z.close()
pres = ET.fromstring(data["ppt/presentation.xml"])
rid2t = {r.get("Id"): r.get("Target") for r in
         ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
order = [os.path.basename(rid2t[s.get(AN.q(AN.R, "id"))])
         for s in pres.find(AN.q(AN.P, "sldIdLst"))]
tree = ET.fromstring(data["ppt/slides/" + order[disp - 1]])
spTree = tree.find(".//" + AN.q(AN.P, "cSld") + "/" + AN.q(AN.P, "spTree"))
shapes = AN.collect_shapes(spTree)
print("display %d — %d top-level shapes" % (disp, len(shapes)))
for s in shapes:
    print("  idx=%-8s id=%-4s %-6s (%6.2f,%5.2f) %5.2f x %4.2f  chrome=%-5s %r"
          % (s["idx"], s["id"], s["tag"], s["x"], s["y"], s["w"], s["h"],
             AN.is_chrome(s), s["text"][:38]))
