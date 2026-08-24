# -*- coding: utf-8 -*-
"""Print connector endpoints (rendered inches, flip-aware) for a display
slide, descending into groups via the chOff/chExt transform."""
import sys, zipfile
from lxml import etree as ET
import _diff_slides as D
A=D.A; P=D.P; EMU=914400.0
MC="http://schemas.openxmlformats.org/markup-compatibility/2006"

def walk(el, ox=0.0, oy=0.0, sx=1.0, sy=1.0, depth=0):
    for c in el:
        tag=ET.QName(c).localname
        if tag=="AlternateContent":
            ch=c.find("{%s}Choice"%MC)
            if ch is not None and len(ch): walk(ch,ox,oy,sx,sy,depth)
            continue
        if tag not in ("sp","pic","graphicFrame","cxnSp","grpSp"): continue
        sppr=c.find(D.q(P,"spPr"))
        if sppr is None: sppr=c.find(D.q(P,"grpSpPr"))
        xf=sppr.find(D.q(A,"xfrm")) if sppr is not None else None
        if tag=="graphicFrame": xf=c.find(D.q(P,"xfrm"))
        if xf is None: continue
        off=xf.find(D.q(A,"off")); ext=xf.find(D.q(A,"ext"))
        x=ox+int(off.get("x"))*sx; y=oy+int(off.get("y"))*sy
        w=int(ext.get("cx"))*sx;  h=int(ext.get("cy"))*sy
        fh=xf.get("flipH")=="1"; fv=xf.get("flipV")=="1"
        txt=D.norm("".join(t.text or "" for t in c.iter(D.q(A,"t"))))[:38]
        pre="  "*depth
        if tag=="cxnSp":
            x0,x1=(x+w,x) if fh else (x,x+w)
            y0,y1=(y+h,y) if fv else (y,y+h)
            ln=c.find(".//"+D.q(A,"ln"))
            wt=ln.get("w") if ln is not None else None
            clr=c.find(".//"+D.q(A,"srgbClr"))
            dash=c.find(".//"+D.q(A,"prstDash"))
            head=c.find(".//"+D.q(A,"tailEnd"))
            print("%s cxn  (%7.3f,%7.3f) -> (%7.3f,%7.3f)  w=%s clr=%s dash=%s head=%s"
                  % (pre, x0/EMU, y0/EMU, x1/EMU, y1/EMU, wt,
                     clr.get("val") if clr is not None else None,
                     dash.get("val") if dash is not None else None,
                     head.get("type") if head is not None else None))
        else:
            print("%s %-4s (%7.3f,%7.3f) %7.3fx%6.3f fh=%d fv=%d | %s"
                  % (pre, tag, x/EMU, y/EMU, w/EMU, h/EMU, fh, fv, txt))
        if tag=="grpSp":
            cho=xf.find(D.q(A,"chOff")); che=xf.find(D.q(A,"chExt"))
            csx=w/int(che.get("cx")) if che is not None and int(che.get("cx")) else sx
            csy=h/int(che.get("cy")) if che is not None and int(che.get("cy")) else sy
            cox=x-int(cho.get("x"))*csx if cho is not None else x
            coy=y-int(cho.get("y"))*csy if cho is not None else y
            print("%s   [scale %.4f x %.4f]"%(pre,csx,csy))
            walk(c,cox,coy,csx,csy,depth+1)

deck=sys.argv[1]; disp=int(sys.argv[2])
z=zipfile.ZipFile(deck); t=ET.fromstring(z.read(D.slide_part(z,disp)))
walk(t.find(".//"+D.q(P,"cSld")+"/"+D.q(P,"spTree")))
