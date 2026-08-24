# -*- coding: utf-8 -*-
"""Dump click-by-click animation choreography for a display slide,
resolving spid -> shape signature (type + rendered position + text)."""
import sys, zipfile
from lxml import etree as ET
import _diff_slides as D

A = D.A; P = D.P; R = D.R
EMU = 914400.0
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def shape_index(deck, disp):
    """spid -> (tag, x, y, w, h, text)"""
    z = zipfile.ZipFile(deck)
    part = D.slide_part(z, disp)
    tree = ET.fromstring(z.read(part))
    spTree = tree.find(".//" + D.q(P, "cSld") + "/" + D.q(P, "spTree"))
    idx = {}
    def walk(el, ox=0.0, oy=0.0, sx=1.0, sy=1.0):
        for c in el:
            tag = ET.QName(c).localname
            if tag == "AlternateContent":
                ch = c.find("{%s}Choice" % MC)
                if ch is not None and len(ch):
                    walk(ch, ox, oy, sx, sy)
                continue
            if tag not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
                continue
            nv = None
            for cand in ("nvSpPr", "nvPicPr", "nvGraphicFramePr",
                         "nvCxnSpPr", "nvGrpSpPr"):
                nv = c.find(D.q(P, cand))
                if nv is not None:
                    break
            sid = nv.find(D.q(P, "cNvPr")).get("id") if nv is not None else None
            if tag == "graphicFrame":
                xf = c.find(D.q(P, "xfrm"))
            else:
                sppr = c.find(D.q(P, "spPr"))
                if sppr is None:
                    sppr = c.find(D.q(P, "grpSpPr"))
                xf = sppr.find(D.q(A, "xfrm")) if sppr is not None else None
            x = y = w = h = 0.0
            if xf is not None:
                off = xf.find(D.q(A, "off")); ext = xf.find(D.q(A, "ext"))
                if off is not None and ext is not None:
                    x = ox + int(off.get("x")) * sx
                    y = oy + int(off.get("y")) * sy
                    w = int(ext.get("cx")) * sx
                    h = int(ext.get("cy")) * sy
            txt = D.norm("".join(t.text or "" for t in c.iter(D.q(A, "t"))))
            if sid:
                idx[sid] = (tag, x/EMU, y/EMU, w/EMU, h/EMU, txt[:44])
            if tag == "grpSp":
                cho = xf.find(D.q(A, "chOff")); che = xf.find(D.q(A, "chExt"))
                csx = w/int(che.get("cx")) if che is not None and int(che.get("cx")) else sx
                csy = h/int(che.get("cy")) if che is not None and int(che.get("cy")) else sy
                cox = x - int(cho.get("x"))*csx if cho is not None else x
                coy = y - int(cho.get("y"))*csy if cho is not None else y
                walk(c, cox, coy, csx, csy)
    walk(spTree)
    z.close()
    return tree, idx


def dump(deck, disp, label):
    tree, idx = shape_index(deck, disp)
    timing = tree.find(D.q(P, "timing"))
    print("### %s display %d" % (label, disp))
    if timing is None:
        print("    (no timing)")
        return
    # find mainSeq: par with ctn nodeType mainSeq
    main = None
    for ctn in timing.iter(D.q(P, "cTn")):
        if ctn.get("nodeType") == "mainSeq":
            main = ctn
            break
    if main is None:
        print("    (no mainSeq)")
        return
    clicks = main.find(D.q(P, "childTnLst"))
    if clicks is None:
        print("    (empty mainSeq)")
        return
    for ci, clickpar in enumerate(clicks, 1):
        # each click = <p:par> containing cTn/childTnLst/par(group)/...
        effs = []
        for beh in clickpar.iter(D.q(P, "cTn")):
            if beh.get("nodeType") in ("clickEffect", "withEffect", "afterEffect"):
                trig = beh.get("nodeType")
                # find spid + prst + paragraph range
                par = beh.getparent()
                spid = None; prg = None; prst = None
                for tgt in par.iter(D.q(P, "spTgt")):
                    spid = tgt.get("spid")
                    tr = tgt.find(D.q(P, "txEl"))
                    if tr is not None:
                        pr = tr.find(D.q(P, "pRg"))
                        if pr is not None:
                            prg = "p%s-%s" % (pr.get("st"), pr.get("end"))
                    break
                for anim in par.iter(D.q(P, "animEffect")):
                    prst = anim.get("transition")
                for cbe in par.iter(D.q(P, "cBhvr")):
                    pass
                effs.append((trig, spid, prg))
        if not effs:
            continue
        print("  click %d:" % ci)
        for trig, spid, prg in effs:
            s = idx.get(spid, ("?", 0, 0, 0, 0, "??"))
            print("      %-12s spid=%-5s %-5s (%7.3f,%7.3f) %6.3fx%5.3f %-8s | %s"
                  % (trig, spid, s[0], s[1], s[2], s[3], s[4], prg or "", s[5]))


if __name__ == "__main__":
    CAN = "Module 2 - In Class Revised.pptx"
    TEST = "Module 2 - In Class Revised_test.pptx"
    a = int(sys.argv[1]); b = int(sys.argv[2]) if len(sys.argv) > 2 else a
    dump(CAN, a, "CANON")
    dump(TEST, b, "BUILD")
