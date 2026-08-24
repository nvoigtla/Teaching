# -*- coding: utf-8 -*-
"""Cross-deck member-level diff: (deckA, dispA) vs (deckB, dispB).
Reports shape count, per-member signature/geometry deltas, notes, and the
click structure of both slides."""
import os, sys, zipfile
from lxml import etree as ET
import _diff_slides as D

HERE = os.path.dirname(os.path.abspath(__file__))
TOL = 0.011

def clicks(deck, disp):
    z = zipfile.ZipFile(deck)
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rmap = {r.get("Id"): r.get("Target") for r in
            ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    parts = ["ppt/" + rmap[s.get(D.q(D.R, "id"))].lstrip("/").replace("../", "")
             for s in pres.find(D.q(D.P, "sldIdLst"))]
    tree = ET.fromstring(z.read(parts[disp - 1]))
    spTree = tree.find(".//" + D.q(D.P, "cSld") + "/" + D.q(D.P, "spTree"))
    # spid -> signature
    idx = {}
    def walk(el):
        for c in el:
            tag = ET.QName(c).localname
            if tag == "AlternateContent":
                ch = c.find("{http://schemas.openxmlformats.org/markup-compatibility/2006}Choice")
                if ch is not None: walk(ch)
                continue
            if tag not in ("sp","pic","graphicFrame","cxnSp","grpSp"): continue
            cnv = c.find(".//" + D.q(D.P, "cNvPr"))
            if cnv is None: continue
            xf = c.find(D.q(D.P,"xfrm")) if tag=="graphicFrame" else None
            if xf is None:
                sppr = c.find(D.q(D.P,"spPr"))
                if sppr is None: sppr = c.find(D.q(D.P,"grpSpPr"))
                xf = sppr.find(D.q(D.A,"xfrm")) if sppr is not None else None
            pos=""
            if xf is not None:
                off,ext = xf.find(D.q(D.A,"off")), xf.find(D.q(D.A,"ext"))
                if off is not None and ext is not None:
                    pos = "(%.2f,%.2f)%.2fx%.2f" % (int(off.get("x"))/D.EMU,int(off.get("y"))/D.EMU,
                                                     int(ext.get("cx"))/D.EMU,int(ext.get("cy"))/D.EMU)
            txt = D.norm("".join(t.text or "" for t in c.iter(D.q(D.A,"t"))))
            nm = cnv.get("name") or ""
            idx[cnv.get("id")] = "%-11s %-22s %-18s %r" % (tag, pos, nm[:18], txt[:44])
            if tag == "grpSp": walk(c)
    walk(spTree)
    paras = {}
    for c in spTree.iter(D.q(D.P,"sp")):
        cnv = c.find(".//"+D.q(D.P,"cNvPr"))
        tx = c.find(D.q(D.P,"txBody"))
        if cnv is not None and tx is not None:
            paras[cnv.get("id")] = [D.norm("".join(t.text or "" for t in p.iter(D.q(D.A,"t"))))
                                    for p in tx.findall(D.q(D.A,"p"))]
    timing = tree.find(D.q(D.P,"timing"))
    out = []
    if timing is not None:
        ctn = None
        for c in timing.iter(D.q(D.P,"cTn")):
            if c.get("nodeType") == "mainSeq": ctn = c; break
        if ctn is not None:
            lvl1 = ctn.find(D.q(D.P,"childTnLst"))
            if lvl1 is not None:
                for par in lvl1.findall(D.q(D.P,"par")):
                    eff=[]; seen=set()
                    for tgt in par.iter(D.q(D.P,"spTgt")):
                        sid = tgt.get("spid")
                        rg = tgt.find(".//"+D.q(D.P,"pRg"))
                        sub=""
                        if rg is not None:
                            st,end = int(rg.get("st")), int(rg.get("end"))
                            pt = paras.get(sid,[])
                            sub = " para%d-%d %r" % (st,end,[pt[i][:30] if i<len(pt) else "?" for i in range(st,end+1)])
                        k=(sid,sub)
                        if k in seen: continue
                        seen.add(k); eff.append(k)
                    out.append([(sid, idx.get(sid,"?"), sub) for sid,sub in eff])
    z.close()
    return out

def report(deckA, dispA, deckB, dispB, verbose=False):
    s1,n1 = D.dump(deckA, dispA)
    s2,n2 = D.dump(deckB, dispB)
    msgs=[]
    if len(s1)!=len(s2): msgs.append("SHAPE COUNT %d (video) vs %d (main)" % (len(s1),len(s2)))
    for i in range(min(len(s1),len(s2))):
        a,b = s1[i],s2[i]
        if (a[0],a[1],a[6],a[7])!=(b[0],b[1],b[6],b[7]):
            d=[]
            if a[0]!=b[0] or a[1]!=b[1]: d.append("type/depth %s@%d vs %s@%d"%(a[0],a[1],b[0],b[1]))
            if a[6]!=b[6]: d.append("TEXT %r vs %r"%(a[6],b[6]))
            if a[7]!=b[7]: d.append("runs %s vs %s"%(a[7],b[7]))
            msgs.append("[%d] %s"%(i,"; ".join(d)))
            continue
        for j,lbl in ((2,"x"),(3,"y"),(4,"w"),(5,"h")):
            if abs(a[j]-b[j])>TOL:
                msgs.append("[%d] %-24s %s: %.3f (video) vs %.3f (main)"%(i,(a[6][:24] or a[0]),lbl,a[j],b[j]))
    if D.norm(n1)!=D.norm(n2): msgs.append("NOTES differ")
    c1,c2 = clicks(deckA,dispA), clicks(deckB,dispB)
    if len(c1)!=len(c2): msgs.append("CLICK COUNT %d (video) vs %d (main)"%(len(c1),len(c2)))
    else:
        for k,(x,y) in enumerate(zip(c1,c2),1):
            sx=[(e[1],e[2]) for e in x]; sy=[(e[1],e[2]) for e in y]
            if sx!=sy: msgs.append("CLICK %d differs"%k)
    return msgs, (s1,s2,n1,n2,c1,c2)

if __name__=="__main__":
    a,da,b,db = sys.argv[1],int(sys.argv[2]),sys.argv[3],int(sys.argv[4])
    msgs,(s1,s2,n1,n2,c1,c2) = report(a,da,b,db)
    print("### %s d%d  VS  %s d%d" % (os.path.basename(a),da,os.path.basename(b),db))
    for m in msgs: print("   "+m)
    if not msgs: print("   IDENTICAL")
    if len(sys.argv)>5:
        print("--- VIDEO shapes"); [print("  %s%-4s (%6.2f,%5.2f) %5.2fx%4.2f | %s | %s"%("  "*s[1],s[0],s[2],s[3],s[4],s[5],s[6],s[7])) for s in s1]
        print("--- MAIN shapes");  [print("  %s%-4s (%6.2f,%5.2f) %5.2fx%4.2f | %s | %s"%("  "*s[1],s[0],s[2],s[3],s[4],s[5],s[6],s[7])) for s in s2]
        print("--- VIDEO clicks")
        for k,cl in enumerate(c1,1):
            print("  CLICK %d"%k); [print("     "+e[1]+e[2]) for e in cl]
        print("--- MAIN clicks")
        for k,cl in enumerate(c2,1):
            print("  CLICK %d"%k); [print("     "+e[1]+e[2]) for e in cl]
        print("--- VIDEO notes:"); print(n1[:2000])
        print("--- MAIN notes:");  print(n2[:2000])
