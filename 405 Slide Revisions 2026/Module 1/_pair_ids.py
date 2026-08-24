# -*- coding: utf-8 -*-
"""Check whether shape ids/names are preserved between a video slide and its main-deck twin."""
import os, sys, zipfile
from lxml import etree as ET
import _diff_slides as D

def spTree(deck, disp):
    z = zipfile.ZipFile(deck)
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rmap = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    parts = ["ppt/" + rmap[s.get(D.q(D.R,"id"))].lstrip("/").replace("../","") for s in pres.find(D.q(D.P,"sldIdLst"))]
    tree = ET.fromstring(z.read(parts[disp-1]))
    z.close()
    return tree.find(".//"+D.q(D.P,"cSld")+"/"+D.q(D.P,"spTree"))

def flat(el, path="", out=None):
    if out is None: out=[]
    for c in el:
        tag = ET.QName(c).localname
        if tag == "AlternateContent":
            ch = c.find("{http://schemas.openxmlformats.org/markup-compatibility/2006}Choice")
            if ch is not None: flat(ch, path, out)
            continue
        if tag not in ("sp","pic","graphicFrame","cxnSp","grpSp"): continue
        cnv = c.find(".//"+D.q(D.P,"cNvPr"))
        nm = cnv.get("name") if cnv is not None else ""
        sid = cnv.get("id") if cnv is not None else "?"
        txt = D.norm("".join(t.text or "" for t in c.iter(D.q(D.A,"t"))))
        out.append((path, tag, sid, nm, txt[:60]))
        if tag=="grpSp": flat(c, path+nm+"/", out)
    return out

a = flat(spTree(sys.argv[1], int(sys.argv[2])))
b = flat(spTree(sys.argv[3], int(sys.argv[4])))
print("--- VIDEO"); [print("  %-24s %-12s id=%-4s %-28s %r" % x) for x in a]
print("--- MAIN");  [print("  %-24s %-12s id=%-4s %-28s %r" % x) for x in b]
