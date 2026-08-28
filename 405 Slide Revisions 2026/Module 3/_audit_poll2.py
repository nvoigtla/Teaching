# -*- coding: utf-8 -*-
"""Detailed poll-chrome audit: badge members, z-order, animation targeting."""
import zipfile, sys
from lxml import etree as ET
A="http://schemas.openxmlformats.org/drawingml/2006/main"
P="http://schemas.openxmlformats.org/presentationml/2006/main"
R="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU=914400.0
def q(n,t): return "{%s}%s"%(n,t)
DECK=sys.argv[1] if len(sys.argv)>1 else "Module 3 - Revised.pptx"
z=zipfile.ZipFile(DECK); parts={n:z.read(n) for n in z.namelist()}
pres=ET.fromstring(parts["ppt/presentation.xml"]); prels=ET.fromstring(parts["ppt/_rels/presentation.xml.rels"])
rid2t={r.get("Id"):r.get("Target") for r in prels}
order=["ppt/slides/"+rid2t[s.get(q(R,"id"))].split("/")[-1] for s in pres.find(q(P,"sldIdLst"))]
def txt(e): return "".join(t.text or "" for t in e.iter(q(A,"t")))
GOLD="E09F3E"
for i,pn in enumerate(order,1):
    tree=ET.fromstring(parts[pn]); sp=tree.find(q(P,"cSld")).find(q(P,"spTree"))
    kids=[c for c in sp if ET.QName(c).localname in ("sp","pic","graphicFrame","cxnSp","grpSp")]
    # animation targets
    anim=set()
    tm=tree.find(q(P,"timing"))
    if tm is not None:
        for e in tm.iter(q(P,"spTgt")): anim.add(e.get("spid"))
    found=[]
    for idx,c in enumerate(kids):
        t=txt(c).strip()
        nv=c.find(".//"+q(P,"cNvPr"))
        nm=nv.get("name"); sid=nv.get("id")
        if t in ("Poll Break","POLL","Discussion Break","Group Discussion"):
            found.append((idx,c,nm,sid,t))
    # also loose gold parallelograms at the bottom
    for idx,c in enumerate(kids):
        if ET.QName(c).localname!="sp": continue
        if any(f[0]==idx for f in found): continue
        cg=c.find(".//"+q(A,"custGeom"))
        sf=c.find(".//"+q(A,"solidFill")+"/"+q(A,"srgbClr"))
        o=c.find(".//"+q(A,"off"))
        if cg is not None and sf is not None and sf.get("val")==GOLD and o is not None and o.get("y") and int(o.get("y"))/EMU>5.8:
            nv=c.find(".//"+q(P,"cNvPr"))
            found.append((idx,c,nv.get("name"),nv.get("id"),"<parallelogram>"))
    if not found: continue
    found.sort()
    print("\n### slide %d  (%d shapes, %s)  timing=%s" % (i,len(kids),pn.split('/')[-1], "yes" if tm is not None else "no"))
    for idx,c,nm,sid,t in found:
        o=c.find(".//"+q(A,"off")); e=c.find(".//"+q(A,"ext"))
        g=""
        if o is not None and o.get("x") and e is not None and e.get("cx"):
            g="L%.3f T%.3f W%.3f H%.3f"%(int(o.get("x"))/EMU,int(o.get("y"))/EMU,int(e.get("cx"))/EMU,int(e.get("cy"))/EMU)
        # animated?
        ids={sid}
        if ET.QName(c).localname=="grpSp":
            for nvv in c.iter(q(P,"cNvPr")): ids.add(nvv.get("id"))
        an="ANIMATED" if (ids & anim) else "static"
        print("   z=%-3d %-8s id=%-6s %-20s %-36s %-9s %s"%(idx,ET.QName(c).localname,sid,nm[:20],g,an,t))
