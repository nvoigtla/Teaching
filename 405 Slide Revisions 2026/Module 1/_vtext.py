# -*- coding: utf-8 -*-
"""Paragraph/run-level dump of one shape in a slide (or all text shapes)."""
import sys
from lxml import etree as ET
import _diff_slides as D
import _vdiff as V

def dump(deck, disp, want=None):
    tree,_,_ = V.load(deck, disp)
    spTree = tree.find(".//"+D.q(D.P,"cSld")+"/"+D.q(D.P,"spTree"))
    for sp in spTree.iter(D.q(D.P,"sp")):
        cnv = sp.find(".//"+D.q(D.P,"cNvPr")); tx = sp.find(D.q(D.P,"txBody"))
        if cnv is None or tx is None: continue
        sid = cnv.get("id")
        if want and sid != want: continue
        print("  id=%s name=%r" % (sid, cnv.get("name")))
        for i,p in enumerate(tx.findall(D.q(D.A,"p"))):
            ppr = p.find(D.q(D.A,"pPr"))
            lvl = ppr.get("lvl") if ppr is not None else None
            mar = ppr.get("marL") if ppr is not None else None
            sb = ppr.find(D.q(D.A,"spcBef")) if ppr is not None else None
            sbv = sb.find(D.q(D.A,"spcPts")).get("val") if sb is not None and sb.find(D.q(D.A,"spcPts")) is not None else "-"
            runs=[]
            for r in p.findall(D.q(D.A,"r")):
                rp = r.find(D.q(D.A,"rPr"))
                sz = rp.get("sz") if rp is not None else "-"
                b = "b" if rp is not None and rp.get("b")=="1" else ""
                it = "i" if rp is not None and rp.get("i")=="1" else ""
                clr = rp.find(".//"+D.q(D.A,"srgbClr")) if rp is not None else None
                runs.append("[%s%s%s%s]%r" % (sz,b,it,("#"+clr.get("val")) if clr is not None else "", r.find(D.q(D.A,"t")).text))
            print("    p%-2d lvl=%-4s marL=%-6s spcBef=%-5s %s" % (i, lvl, mar, sbv, " ".join(runs)))

if __name__=="__main__":
    dump(sys.argv[1], int(sys.argv[2]), sys.argv[3] if len(sys.argv)>3 else None)
