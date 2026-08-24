# -*- coding: utf-8 -*-
"""Id-keyed diff between a Videos-Final slide and its main-deck twin.
Shape ids survived Nico's extract+polish, so we can pair exactly."""
import os, sys, zipfile
from lxml import etree as ET
import _diff_slides as D

MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
EMU = 914400.0

def load(deck, disp):
    z = zipfile.ZipFile(deck)
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rmap = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    parts = ["ppt/" + rmap[s.get(D.q(D.R,"id"))].lstrip("/").replace("../","") for s in pres.find(D.q(D.P,"sldIdLst"))]
    part = parts[disp-1]
    tree = ET.fromstring(z.read(part))
    notes = D.notes_text(z, part)
    z.close()
    return tree, notes, part

def shapes(tree):
    """id -> dict(tag,name,path,x,y,w,h,text,runs,order)"""
    spTree = tree.find(".//"+D.q(D.P,"cSld")+"/"+D.q(D.P,"spTree"))
    out = {}
    ctr = [0]
    def walk(el, path, ox, oy, sx, sy):
        for c in el:
            tag = ET.QName(c).localname
            if tag == "AlternateContent":
                ch = c.find(MC+"Choice")
                if ch is not None: walk(ch, path, ox, oy, sx, sy)
                continue
            if tag not in ("sp","pic","graphicFrame","cxnSp","grpSp"): continue
            cnv = c.find(".//"+D.q(D.P,"cNvPr"))
            sid = cnv.get("id") if cnv is not None else "?"
            nm = cnv.get("name") if cnv is not None else ""
            if tag=="graphicFrame": xf = c.find(D.q(D.P,"xfrm"))
            else:
                sppr = c.find(D.q(D.P,"spPr"))
                if sppr is None: sppr = c.find(D.q(D.P,"grpSpPr"))
                xf = sppr.find(D.q(D.A,"xfrm")) if sppr is not None else None
            x=y=w=h=0.0
            if xf is not None:
                off,ext = xf.find(D.q(D.A,"off")), xf.find(D.q(D.A,"ext"))
                if off is not None and ext is not None:
                    x = ox + int(off.get("x"))*sx; y = oy + int(off.get("y"))*sy
                    w = int(ext.get("cx"))*sx;      h = int(ext.get("cy"))*sy
            txt = D.norm("".join(t.text or "" for t in c.iter(D.q(D.A,"t"))))
            runs=[]
            for r in c.iter(D.q(D.A,"rPr")):
                clr = r.find(".//"+D.q(D.A,"srgbClr"))
                runs.append("%s%s%s%s" % (r.get("sz") or "-", "b" if r.get("b")=="1" else "",
                            "i" if r.get("i")=="1" else "", ("#"+clr.get("val")) if clr is not None else ""))
            ctr[0]+=1
            out[sid] = dict(tag=tag, name=nm, path=path, x=x/EMU, y=y/EMU, w=w/EMU, h=h/EMU,
                            text=txt, runs=",".join(runs), order=ctr[0])
            if tag=="grpSp":
                cho, che = xf.find(D.q(D.A,"chOff")), xf.find(D.q(D.A,"chExt"))
                csx = w/int(che.get("cx")) if che is not None and int(che.get("cx")) else sx
                csy = h/int(che.get("cy")) if che is not None and int(che.get("cy")) else sy
                cox = x - int(cho.get("x"))*csx if cho is not None else x
                coy = y - int(cho.get("y"))*csy if cho is not None else y
                walk(c, path+("%s(%s)/"%(nm,sid)), cox, coy, csx, csy)
    walk(spTree, "", 0.0, 0.0, 1.0, 1.0)
    return out

def clicks(tree):
    spTree = tree.find(".//"+D.q(D.P,"cSld")+"/"+D.q(D.P,"spTree"))
    paras={}
    for c in spTree.iter(D.q(D.P,"sp")):
        cnv=c.find(".//"+D.q(D.P,"cNvPr")); tx=c.find(D.q(D.P,"txBody"))
        if cnv is not None and tx is not None:
            paras[cnv.get("id")]=[D.norm("".join(t.text or "" for t in p.iter(D.q(D.A,"t")))) for p in tx.findall(D.q(D.A,"p"))]
    timing = tree.find(D.q(D.P,"timing"))
    if timing is None: return []
    ctn=None
    for c in timing.iter(D.q(D.P,"cTn")):
        if c.get("nodeType")=="mainSeq": ctn=c; break
    if ctn is None: return []
    lvl1 = ctn.find(D.q(D.P,"childTnLst"))
    if lvl1 is None: return []
    res=[]
    for par in lvl1.findall(D.q(D.P,"par")):
        eff=[]; seen=set()
        for tgt in par.iter(D.q(D.P,"spTgt")):
            sid=tgt.get("spid"); rg=tgt.find(".//"+D.q(D.P,"pRg")); sub=""
            if rg is not None:
                st,end=int(rg.get("st")),int(rg.get("end")); pt=paras.get(sid,[])
                sub=" para%d-%d %r"%(st,end,[pt[i][:34] if i<len(pt) else "?" for i in range(st,end+1)])
            if (sid,sub) in seen: continue
            seen.add((sid,sub)); eff.append((sid,sub))
        res.append(eff)
    return res

IGNORE_TEXT = None  # set of texts to ignore (page numbers etc.)

def diff(vdeck, vd, mdeck, md, tol=0.011, show_clicks=True):
    tv, nv, _ = load(vdeck, vd); tm, nm_, _ = load(mdeck, md)
    sv, sm = shapes(tv), shapes(tm)
    print("### VIDEO %s d%d   vs   MAIN d%d" % (os.path.basename(vdeck), vd, md))
    only_v = [k for k in sv if k not in sm]; only_m=[k for k in sm if k not in sv]
    for k in sorted(only_v, key=lambda k: sv[k]['order']):
        s=sv[k]; print("  + ONLY IN VIDEO  id=%-4s %-11s %-22s (%6.2f,%5.2f) %5.2fx%4.2f path=%s | %r"%(k,s['tag'],s['name'][:22],s['x'],s['y'],s['w'],s['h'],s['path'],s['text'][:60]))
    for k in sorted(only_m, key=lambda k: sm[k]['order']):
        s=sm[k]; print("  - ONLY IN MAIN   id=%-4s %-11s %-22s (%6.2f,%5.2f) %5.2fx%4.2f path=%s | %r"%(k,s['tag'],s['name'][:22],s['x'],s['y'],s['w'],s['h'],s['path'],s['text'][:60]))
    for k in sorted(set(sv)&set(sm), key=lambda k: sv[k]['order']):
        a,b = sv[k], sm[k]; d=[]
        if a['path']!=b['path']: d.append("PATH %r -> %r" % (b['path'], a['path']))
        for f in ('x','y','w','h'):
            if abs(a[f]-b[f])>tol: d.append("%s %.3f(main) -> %.3f(video)"%(f,b[f],a[f]))
        if a['text']!=b['text']: d.append("TEXT %r -> %r"%(b['text'][:110], a['text'][:110]))
        if a['runs']!=b['runs']: d.append("RUNS %s -> %s"%(b['runs'][:120], a['runs'][:120]))
        if a['name']!=b['name']: d.append("NAME %r -> %r"%(b['name'],a['name']))
        if d: print("  ~ id=%-4s %-11s %-20s : %s"%(k,a['tag'],a['name'][:20],"; ".join(d)))
    if D.norm(nv)!=D.norm(nm_):
        print("  ~ NOTES DIFFER")
        print("     MAIN : %r"%nm_[:700]); print("     VIDEO: %r"%nv[:700])
    cv, cm = clicks(tv), clicks(tm)
    if show_clicks:
        def lab(sid, sub, S): 
            s=S.get(sid); 
            return "id=%-4s %-11s %-18s %r%s"%(sid, s['tag'] if s else '?', (s['name'][:18] if s else '?'), (s['text'][:40] if s else '?'), sub)
        same = len(cv)==len(cm) and all([ (x) == (y) for x,y in zip(cv,cm)])
        print("  CLICKS: main %d, video %d%s"%(len(cm),len(cv)," (IDENTICAL)" if same else " (DIFFER)"))
        if not same:
            print("   -- MAIN clicks")
            for i,cl in enumerate(cm,1):
                print("      %d) "%i + "\n         ".join(lab(s,u,sm) for s,u in cl))
            print("   -- VIDEO clicks")
            for i,cl in enumerate(cv,1):
                print("      %d) "%i + "\n         ".join(lab(s,u,sv) for s,u in cl))

if __name__=="__main__":
    diff(sys.argv[1], int(sys.argv[2]), sys.argv[3], int(sys.argv[4]))
