# -*- coding: utf-8 -*-
"""Paragraph/run-level dump of a shape (by index from _dump1) on a display slide."""
import sys, zipfile
from lxml import etree as ET
import _diff_slides as D
A=D.A; P=D.P
MC="http://schemas.openxmlformats.org/markup-compatibility/2006"

def shapes(deck, disp):
    z=zipfile.ZipFile(deck); part=D.slide_part(z, disp)
    tree=ET.fromstring(z.read(part))
    spTree=tree.find(".//"+D.q(P,"cSld")+"/"+D.q(P,"spTree"))
    out=[]
    def walk(el):
        for c in el:
            tag=ET.QName(c).localname
            if tag=="AlternateContent":
                ch=c.find("{%s}Choice"%MC)
                if ch is not None and len(ch): walk(ch)
                continue
            if tag not in ("sp","pic","graphicFrame","cxnSp","grpSp"): continue
            out.append((tag,c))
            if tag=="grpSp": walk(c)
    walk(spTree); z.close(); return out

def show(deck, disp, idxs, label):
    sh=shapes(deck, disp)
    print("### %s display %d"%(label, disp))
    for i in idxs:
        tag,c=sh[i]
        print("  --- [%d] %s"%(i,tag))
        for pi,para in enumerate(c.iter(D.q(A,"p"))):
            pPr=para.find(D.q(A,"pPr"))
            info=""
            if pPr is not None:
                info=" lvl=%s marL=%s indent=%s algn=%s buNone=%s"%(
                    pPr.get("lvl"),pPr.get("marL"),pPr.get("indent"),pPr.get("algn"),
                    pPr.find(D.q(A,"buNone")) is not None)
                sb=pPr.find(D.q(A,"spcBef"))
                if sb is not None:
                    pts=sb.find(".//"+D.q(A,"spcPts"))
                    if pts is not None: info+=" spcBef=%s"%pts.get("val")
            print("    p%d%s"%(pi,info))
            for r in para.findall(D.q(A,"r")):
                rPr=r.find(D.q(A,"rPr"))
                a=""
                if rPr is not None:
                    clr=rPr.find(".//"+D.q(A,"srgbClr"))
                    a="sz=%s b=%s i=%s u=%s clr=%s"%(rPr.get("sz"),rPr.get("b"),
                        rPr.get("i"),rPr.get("u"),clr.get("val") if clr is not None else None)
                t=r.find(D.q(A,"t"))
                print("       run[%s] %r"%(a, (t.text or "") if t is not None else ""))
if __name__=="__main__":
    deck=sys.argv[1]; disp=int(sys.argv[2]); idxs=[int(x) for x in sys.argv[3:]]
    show(deck, disp, idxs, deck)
