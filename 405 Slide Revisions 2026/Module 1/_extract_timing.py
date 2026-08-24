# -*- coding: utf-8 -*-
"""Extract a slide's animation choreography: the main sequence broken into
CLICKS, each click listing the shapes revealed together, resolved from
spid to a signature of shape type + position + text.

Usage: _extract_timing.py "<deck>.pptx" <display> [<display> ...]
"""
import os
import sys
import zipfile

from lxml import etree as ET

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0
HERE = os.path.dirname(os.path.abspath(__file__))


def q(ns, t):
    return "{%s}%s" % (ns, t)


def norm(t):
    return " ".join((t or "").split())


def shape_index(spTree):
    """spid -> signature, walking groups too."""
    out = {}

    def walk(el, depth, prefix):
        for c in el:
            tag = ET.QName(c).localname
            if tag == "AlternateContent":
                ch = c.find("{http://schemas.openxmlformats.org/"
                            "markup-compatibility/2006}Choice")
                if ch is not None:
                    walk(ch, depth, prefix)
                continue
            if tag not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
                continue
            cnv = c.find(".//" + q(P, "cNvPr"))
            if cnv is None:
                continue
            sid = cnv.get("id")
            xf = (c.find(q(P, "xfrm")) if tag == "graphicFrame"
                  else None)
            if xf is None:
                sppr = c.find(q(P, "spPr"))
                if sppr is None:
                    sppr = c.find(q(P, "grpSpPr"))
                xf = sppr.find(q(A, "xfrm")) if sppr is not None else None
            pos = ""
            if xf is not None:
                off, ext = xf.find(q(A, "off")), xf.find(q(A, "ext"))
                if off is not None and ext is not None:
                    pos = ("(%.2f,%.2f) %.2fx%.2f"
                           % (int(off.get("x")) / EMU,
                              int(off.get("y")) / EMU,
                              int(ext.get("cx")) / EMU,
                              int(ext.get("cy")) / EMU))
            txt = norm("".join(t.text or "" for t in c.iter(q(A, "t"))))
            out[sid] = "%s%-12s %-22s %r" % (prefix, tag, pos, txt[:44])
            if tag == "grpSp":
                walk(c, depth + 1, prefix + "  ")
    walk(spTree, 0, "")
    return out


def paras_of(spTree, sid):
    """Paragraph texts of a shape, for pRg-targeted effects."""
    for c in spTree.iter(q(P, "sp")):
        cnv = c.find(".//" + q(P, "cNvPr"))
        if cnv is not None and cnv.get("id") == sid:
            tx = c.find(q(P, "txBody"))
            if tx is None:
                return []
            return [norm("".join(t.text or "" for t in p.iter(q(A, "t"))))
                    for p in tx.findall(q(A, "p"))]
    return []


def dump(deck, disp):
    z = zipfile.ZipFile(os.path.join(HERE, deck))
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rmap = {r.get("Id"): r.get("Target") for r in
            ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    parts = ["ppt/" + rmap[s.get(q(R, "id"))].lstrip("/").replace("../", "")
             for s in pres.find(q(P, "sldIdLst"))]
    part = parts[disp - 1]
    tree = ET.fromstring(z.read(part))
    spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
    idx = shape_index(spTree)
    print("=" * 72)
    print("%s  display %d  (%s)" % (deck, disp, part.split("/")[-1]))
    print("--- shapes")
    for sid, sig in sorted(idx.items(), key=lambda kv: int(kv[0])):
        print("   id=%-4s %s" % (sid, sig))
    timing = tree.find(q(P, "timing"))
    if timing is None:
        print("--- NO TIMING (static slide)")
        z.close()
        return
    # main sequence: the outermost par under the seq with nodeType mainSeq
    # nodeType="mainSeq" sits on the <p:cTn> INSIDE the <p:seq>
    ctn = None
    for c in timing.iter(q(P, "cTn")):
        if c.get("nodeType") == "mainSeq":
            ctn = c
            break
    if ctn is None:
        print("--- no mainSeq")
        z.close()
        return
    lvl1 = ctn.find(q(P, "childTnLst"))
    if lvl1 is None:
        print("--- mainSeq has no childTnLst")
        z.close()
        return
    print("--- clicks")
    click_no = 0
    for par in lvl1.findall(q(P, "par")):          # one per click group
        click_no += 1
        effects = []
        for beh in par.iter(q(P, "cTn")):
            pass
        for tgt in par.iter(q(P, "spTgt")):
            sid = tgt.get("spid")
            rg = tgt.find(".//" + q(P, "pRg"))
            sub = ""
            if rg is not None:
                st, end = rg.get("st"), rg.get("end")
                ptexts = paras_of(spTree, sid)
                names = []
                for i in range(int(st), int(end) + 1):
                    names.append(ptexts[i][:34] if i < len(ptexts) else "?")
                sub = "  para %s-%s %r" % (st, end, names)
            # trigger type of the enclosing par
            effects.append((sid, sub))
        # de-dup while keeping order
        seen = set()
        uniq = []
        for e in effects:
            if e not in seen:
                seen.add(e)
                uniq.append(e)
        print("  CLICK %d  (%d effect%s)"
              % (click_no, len(uniq), "" if len(uniq) == 1 else "s"))
        for sid, sub in uniq:
            print("      id=%-4s %s%s" % (sid, idx.get(sid, "?"), sub))
    z.close()


if __name__ == "__main__":
    deck = sys.argv[1]
    for d in sys.argv[2:]:
        dump(deck, int(d))
