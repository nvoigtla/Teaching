# -*- coding: utf-8 -*-
"""Full shape dump (geometry + text + run formatting + hyperlinks) of a
slide in an arbitrary deck."""
import sys, zipfile
from lxml import etree as ET
import _diff_slides as D
A=D.A; P=D.P; R=D.R; REL=D.REL
EMU=914400.0
MC="http://schemas.openxmlformats.org/markup-compatibility/2006"

deck=sys.argv[1]; disps=[int(x) for x in sys.argv[2:]]
z=zipfile.ZipFile(deck)
for disp in disps:
    part=D.slide_part(z, disp)
    try:
        rels=ET.fromstring(z.read(part.replace("slides/","slides/_rels/")+".rels"))
        rmap={r.get("Id"):(r.get("Target"), r.get("TargetMode")) for r in rels}
    except KeyError:
        rmap={}
    t=ET.fromstring(z.read(part))
    sp=t.find(".//"+D.q(P,"cSld")+"/"+D.q(P,"spTree"))
    print("="*70); print("DECK %s  SLIDE %d" % (deck, disp))
    def walk(el, ox=0.0, oy=0.0, sx=1.0, sy=1.0, depth=0):
        for c in el:
            tag=ET.QName(c).localname
            if tag=="AlternateContent":
                ch=c.find("{%s}Choice"%MC)
                if ch is not None and len(ch): walk(ch,ox,oy,sx,sy,depth)
                continue
            if tag not in ("sp","pic","graphicFrame","cxnSp","grpSp"): continue
            pr=c.find(D.q(P,"spPr")) or c.find(D.q(P,"grpSpPr"))
            xf=c.find(D.q(P,"xfrm")) if tag=="graphicFrame" else (pr.find(D.q(A,"xfrm")) if pr is not None else None)
            x=y=w=h=0.0; fh=fv=False
            if xf is not None:
                o=xf.find(D.q(A,"off")); e=xf.find(D.q(A,"ext"))
                if o is not None and e is not None:
                    x=ox+int(o.get("x"))*sx; y=oy+int(o.get("y"))*sy
                    w=int(e.get("cx"))*sx; h=int(e.get("cy"))*sy
                fh=xf.get("flipH")=="1"; fv=xf.get("flipV")=="1"
            pre="  "*depth
            geom=c.find(".//"+D.q(A,"prstGeom"))
            g=geom.get("prst") if geom is not None else ("custGeom" if c.find(".//"+D.q(A,"custGeom")) is not None else "")
            fill=c.find(D.q(P,"spPr")+"/"+D.q(A,"solidFill")+"/"+D.q(A,"srgbClr")) if pr is not None else None
            ln=c.find(".//"+D.q(A,"ln"))
            lnc=ln.find(".//"+D.q(A,"srgbClr")) if ln is not None else None
            print("%s%-5s (%6.3f,%6.3f) %6.3fx%6.3f fh=%d fv=%d prst=%-12s fill=%s ln=%s/%s"
                  % (pre, tag, x/EMU, y/EMU, w/EMU, h/EMU, fh, fv, g,
                     fill.get("val") if fill is not None else None,
                     lnc.get("val") if lnc is not None else None,
                     ln.get("w") if ln is not None else None))
            for para in c.iter(D.q(A,"p")):
                runs=[]
                for r in para.findall(D.q(A,"r")):
                    rPr=r.find(D.q(A,"rPr"))
                    clr=rPr.find(".//"+D.q(A,"srgbClr")) if rPr is not None else None
                    hl=rPr.find(D.q(A,"hlinkClick")) if rPr is not None else None
                    tgt=""
                    if hl is not None:
                        rid=hl.get(D.q(R,"id"))
                        tgt=" LINK->%s" % (rmap.get(rid,("?",""))[0] if rid else "(none)")
                    tt=r.find(D.q(A,"t"))
                    runs.append("[sz=%s b=%s i=%s u=%s c=%s%s] %r" % (
                        rPr.get("sz") if rPr is not None else None,
                        rPr.get("b") if rPr is not None else None,
                        rPr.get("i") if rPr is not None else None,
                        rPr.get("u") if rPr is not None else None,
                        clr.get("val") if clr is not None else None,
                        tgt, (tt.text or "") if tt is not None else ""))
                if runs:
                    print("%s    p: %s" % (pre, " | ".join(runs)))
            if tag=="grpSp":
                cho=xf.find(D.q(A,"chOff")); che=xf.find(D.q(A,"chExt"))
                csx=w/int(che.get("cx")) if che is not None and int(che.get("cx")) else sx
                csy=h/int(che.get("cy")) if che is not None and int(che.get("cy")) else sy
                walk(c, x-int(cho.get("x"))*csx if cho is not None else x,
                     y-int(cho.get("y"))*csy if cho is not None else y, csx, csy, depth+1)
    walk(sp)
z.close()
