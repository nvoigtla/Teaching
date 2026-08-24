# -*- coding: utf-8 -*-
import sys, re
from lxml import etree as ET
import _diff_slides as D, _vdiff as V
tree,_,_ = V.load(sys.argv[1], int(sys.argv[2]))
sp = tree.find(".//"+D.q(D.P,"cSld")+"/"+D.q(D.P,"spTree"))
want=set(sys.argv[3:])
for c in sp.iter():
    if ET.QName(c).localname not in ("sp","pic","cxnSp","graphicFrame","grpSp"): continue
    cnv=c.find(".//"+D.q(D.P,"cNvPr"))
    if cnv is None or cnv.get("id") not in want: continue
    s=re.sub(r' xmlns:[a-z0-9]+="[^"]*"','',ET.tostring(c,pretty_print=True).decode())
    print(s[:2500]); print("~"*60)
