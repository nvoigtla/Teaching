# -*- coding: utf-8 -*-
import sys, re
from lxml import etree as ET
import _diff_slides as D, _vdiff as V
tree,_,_ = V.load(sys.argv[1], int(sys.argv[2]))
sp = tree.find(".//"+D.q(D.P,"cSld")+"/"+D.q(D.P,"spTree"))
def col(el):
    if el is None: return "-"
    s=el.find(D.q(D.A,"srgbClr"))
    if s is not None: return "#"+s.get("val")
    sc=el.find(D.q(D.A,"schemeClr"))
    if sc is not None:
        mods="".join("%s=%s"%(ET.QName(m).localname,m.get("val")) for m in sc)
        return "scheme:%s %s"%(sc.get("val"),mods)
    return "?"
for c in sp:
    tag=ET.QName(c).localname
    if tag not in ("sp","grpSp","pic"): continue
    cnv=c.find(".//"+D.q(D.P,"cNvPr")); sid=cnv.get("id"); nm=cnv.get("name")
    sppr=c.find(D.q(D.P,"spPr"))
    fill = col(sppr.find(D.q(D.A,"solidFill"))) if sppr is not None else "-"
    if sppr is not None and sppr.find(D.q(D.A,"noFill")) is not None: fill="noFill"
    alpha=""
    if sppr is not None:
        sf=sppr.find(D.q(D.A,"solidFill"))
        if sf is not None:
            a=sf.find(".//"+D.q(D.A,"alpha"))
            if a is not None: alpha=" alpha=%s"%a.get("val")
    tx=c.find(D.q(D.P,"txBody"))
    runs=[]
    if tx is not None:
        for p in tx.findall(D.q(D.A,"p")):
            for r in p.findall(D.q(D.A,"r")):
                rp=r.find(D.q(D.A,"rPr"))
                runs.append("[%s]%r"%(col(rp.find(D.q(D.A,"solidFill"))) if rp is not None else "-", r.find(D.q(D.A,"t")).text[:44]))
    print("id=%-4s %-20s fill=%-22s%s %s"%(sid,nm[:20],fill,alpha," ".join(runs)[:150]))
