# -*- coding: utf-8 -*-
"""Raw-XML sanity diff: per shape id, compare a canonicalized dump of the
spPr (fill/line/effects/geom) and the rels (images, hyperlinks)."""
import os, re, sys, zipfile, hashlib, difflib
from lxml import etree as ET
import _diff_slides as D, _vdiff as V

MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"

def parts_of(deck):
    z=zipfile.ZipFile(deck)
    pres=ET.fromstring(z.read("ppt/presentation.xml"))
    rmap={r.get("Id"):r.get("Target") for r in ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    return z, ["ppt/"+rmap[s.get(D.q(D.R,"id"))].lstrip("/").replace("../","") for s in pres.find(D.q(D.P,"sldIdLst"))]

def props(deck, disp):
    z, parts = parts_of(deck); part=parts[disp-1]
    tree=ET.fromstring(z.read(part))
    spTree=tree.find(".//"+D.q(D.P,"cSld")+"/"+D.q(D.P,"spTree"))
    out={}
    def walk(el):
        for c in el:
            tag=ET.QName(c).localname
            if tag=="AlternateContent":
                ch=c.find(MC+"Choice")
                if ch is not None: walk(ch)
                continue
            if tag not in ("sp","pic","graphicFrame","cxnSp","grpSp"): continue
            cnv=c.find(".//"+D.q(D.P,"cNvPr")); sid=cnv.get("id") if cnv is not None else "?"
            sppr=c.find(D.q(D.P,"spPr")) or c.find(D.q(D.P,"grpSpPr"))
            bits=[]
            if sppr is not None:
                for e in sppr:
                    ln=ET.QName(e).localname
                    if ln=="xfrm": continue
                    bits.append(re.sub(r' xmlns:[a-z0-9]+="[^"]*"', "", ET.tostring(e).decode()))
            # hyperlinks + blip rels
            for h in c.iter(D.q(D.A,"hlinkClick")): bits.append("HLINK:"+(h.get(D.q(D.R,"embed")) or h.get(D.q(D.R,"id")) or ""))
            for b in c.iter(D.q(D.A,"blip")): bits.append("BLIP:"+(b.get(D.q(D.R,"embed")) or ""))
            for h in c.iter(D.q(D.P,"cNvPr")): pass
            out[sid]="\n".join(bits)
            if tag=="grpSp": walk(c)
    walk(spTree)
    # rel map for this slide
    rp=part.replace("slides/","slides/_rels/")+".rels"
    rels={}
    try:
        for r in ET.fromstring(z.read(rp)).iter("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            rels[r.get("Id")]=(r.get("Type").split("/")[-1], r.get("Target"), r.get("TargetMode") or "")
    except KeyError: pass
    # md5 of media
    media={}
    for rid,(ty,tg,tm) in rels.items():
        if ty=="image" and not tm:
            p="ppt/"+tg.replace("../","")
            try: media[rid]=hashlib.md5(z.read(p)).hexdigest()[:10]
            except KeyError: media[rid]="?"
    z.close()
    return out, rels, media

def cmp(vdeck, vd, mdeck, md):
    pv, rv, mv = props(vdeck, vd)
    pm, rm, mm = props(mdeck, md)
    msgs=[]
    for k in sorted(set(pv)&set(pm), key=int):
        if pv[k]!=pm[k]:
            dl=[l for l in difflib.unified_diff(pm[k].split("\n"), pv[k].split("\n"), lineterm="", n=0) if l[:1] in "+-" and l[:3] not in ("+++","---")]
            msgs.append("  ~ id=%s spPr/rels differ:\n      %s" % (k, "\n      ".join(dl[:12])))
    # media
    sv=sorted(mv.values()); sm=sorted(mm.values())
    if sv!=sm: msgs.append("  ~ IMAGE SET differs: main %s vs video %s"%(sm,sv))
    # external hyperlinks
    ev=sorted(t for ty,t,tm in rv.values() if tm=="External"); em=sorted(t for ty,t,tm in rm.values() if tm=="External")
    if ev!=em: msgs.append("  ~ EXTERNAL LINKS differ:\n      main : %s\n      video: %s"%(em,ev))
    return msgs

MAP=[("Module 1 - Video 1 - Introduction.pptx",[(1,67),(2,2),(3,9),(4,10),(5,11),(6,13),(7,17),(8,1),(9,69),(10,95),(11,100)]),
     ("Module 1 - Video 2 - Markets.pptx",[(1,70),(2,71),(3,72),(6,75),(7,76)]),
     ("Module 1 - Video 3 - Demand and Supply.pptx",[(i,i+76) for i in range(1,11)]),
     ("Module 1 - Video 4 - Equilibrium.pptx",[(i,i+86) for i in range(1,8)])]
HERE=os.path.dirname(os.path.abspath(__file__))
MAIN=os.path.join(HERE,"Module 1 - Revised.pptx")
for deck,pairs in MAP:
    for dv,dm in pairs:
        m=cmp(os.path.join(HERE,"Videos Final",deck), dv, MAIN, dm)
        if m:
            print("### %s d%d vs main d%d" % (deck.split(" - ")[1], dv, dm))
            print("\n".join(m))
