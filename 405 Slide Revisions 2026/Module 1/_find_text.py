# -*- coding: utf-8 -*-
import sys, zipfile
from lxml import etree as ET
import _diff_slides as D, _vdiff as V
deck=sys.argv[1]; needle=sys.argv[2]
z=zipfile.ZipFile(deck)
pres=ET.fromstring(z.read("ppt/presentation.xml"))
rmap={r.get("Id"):r.get("Target") for r in ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
parts=["ppt/"+rmap[s.get(D.q(D.R,"id"))].lstrip("/").replace("../","") for s in pres.find(D.q(D.P,"sldIdLst"))]
for i,p in enumerate(parts,1):
    t=ET.fromstring(z.read(p))
    txt=D.norm("".join(x.text or "" for x in t.iter(D.q(D.A,"t"))))
    if needle in txt: print("display %d (%s)"%(i,p.split('/')[-1]))
