# -*- coding: utf-8 -*-
"""Paragraph- and run-level text of one slide, in OURS and in its final
counterpart, so wording / emphasis / indent changes can be ported exactly.

    python _vf_text.py 10          # every text shape, both decks
    python _vf_text.py 10 --final  # only the final deck
"""
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

import _diff_slides as D
import _vf_diff as V

A, P = D.A, D.P
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
EMU = 914400.0


def shapes(deck, disp):
    z = zipfile.ZipFile(deck)
    tree = ET.fromstring(z.read(D.slide_part(z, disp)))
    z.close()
    spTree = tree.find(".//" + D.q(P, "cSld") + "/" + D.q(P, "spTree"))
    out = []

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
            pr = c.find(D.q(P, "spPr"))
            if pr is None:
                pr = c.find(D.q(P, "grpSpPr"))
            xf = (c.find(D.q(P, "xfrm")) if tag == "graphicFrame"
                  else (pr.find(D.q(A, "xfrm")) if pr is not None else None))
            x = y = w = h = 0.0
            if xf is not None:
                o, e = xf.find(D.q(A, "off")), xf.find(D.q(A, "ext"))
                if o is not None and e is not None:
                    x = ox + int(o.get("x")) * sx
                    y = oy + int(o.get("y")) * sy
                    w, h = int(e.get("cx")) * sx, int(e.get("cy")) * sy
            if tag == "grpSp" and xf is not None:
                cho, che = xf.find(D.q(A, "chOff")), xf.find(D.q(A, "chExt"))
                csx = w / int(che.get("cx")) if che is not None else sx
                csy = h / int(che.get("cy")) if che is not None else sy
                out.append((tag, x / EMU, y / EMU, w / EMU, h / EMU, None))
                walk(c, x - int(cho.get("x")) * csx,
                     y - int(cho.get("y")) * csy, csx, csy)
                continue
            paras = []
            for p in c.iter(D.q(A, "p")):
                ppr = p.find(D.q(A, "pPr"))
                lvl = ppr.get("lvl") if ppr is not None else None
                marl = ppr.get("marL") if ppr is not None else None
                indent = ppr.get("indent") if ppr is not None else None
                spc = None
                if ppr is not None:
                    sb = ppr.find(D.q(A, "spcBef") + "/" + D.q(A, "spcPts"))
                    if sb is not None:
                        spc = sb.get("val")
                runs = []
                for r in p.findall(D.q(A, "r")):
                    rpr = r.find(D.q(A, "rPr"))
                    t = r.find(D.q(A, "t"))
                    fmt = ""
                    if rpr is not None:
                        clr = rpr.find(".//" + D.q(A, "srgbClr"))
                        fmt = "%s%s%s%s" % (
                            rpr.get("sz") or "",
                            "B" if rpr.get("b") == "1" else "",
                            "I" if rpr.get("i") == "1" else "",
                            ("#" + clr.get("val")) if clr is not None else "")
                    runs.append((t.text if t is not None else "", fmt))
                # keep EMPTY paragraphs: PowerPoint counts them, so the
                # pRg indices in the timing block depend on them
                paras.append((lvl, marl, indent, spc, runs))
            out.append((tag, x / EMU, y / EMU, w / EMU, h / EMU, paras))
    walk(spTree)
    return out


def show(deck, disp, label):
    print("#" * 74)
    print("%s  -  %s slide %d" % (label, Path(deck).name, disp))
    print("#" * 74)
    for tag, x, y, w, h, paras in shapes(deck, disp):
        if not paras:
            print("  <%s> (%6.2f,%5.2f) %5.2f x %4.2f" % (tag, x, y, w, h))
            continue
        print("  <%s> (%6.2f,%5.2f) %5.2f x %4.2f" % (tag, x, y, w, h))
        for lvl, marl, indent, spc, runs in paras:
            meta = "lvl=%s marL=%s ind=%s spcB=%s" % (lvl, marl, indent, spc)
            print("      p[%s]" % meta)
            for txt, fmt in runs:
                print("        %-12s %r" % (fmt, txt))


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    disp = int(args[0])
    deck, fd = V.MAP[disp]
    if "--final" not in sys.argv:
        show(V.OURS, disp, "OURS")
    if "--ours" not in sys.argv:
        show(deck, fd, "FINAL")
