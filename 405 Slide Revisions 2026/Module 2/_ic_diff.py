# -*- coding: utf-8 -*-
"""Compare Nico's hand-edited In-Class deck with a fresh pipeline build.

He deleted six slides on 2026-08-25 (two duplicate PollEverywhere slides
and the four-slide Amazon e-book example), so his display numbers no
longer line up with the build's.  MAP translates.

    python _ic_diff.py --all       # one line per slide
    python _ic_diff.py 18          # full shape listing, both decks
    python _ic_diff.py 18 --brief  # only the shapes that differ
    python _ic_diff.py 18 --anim   # click structure, both decks
"""
import os
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

import _diff_slides as D

HERE = Path(__file__).parent
HIS = str(HERE / os.environ.get("IC_HIS", "Module 2 - In Class Revised.pptx"))
BUILD = str(HERE / os.environ.get("IC_BUILD",
                                  "Module 2 - In Class Revised_test.pptx"))

A, P = D.A, D.P
EMU = D.EMU


def MAP(his):
    """His display number -> the same slide in a fresh build.

    From 2026-08-25 the build itself drops the six slides he deleted, so
    the two decks line up one to one.
    """
    return his


def struct(deck, disp):
    return D.dump(deck, disp)


def tree_of(deck, disp):
    z = zipfile.ZipFile(deck)
    t = ET.fromstring(z.read(D.slide_part(z, disp)))
    z.close()
    return t


def beats(deck, disp):
    timing = tree_of(deck, disp).find(D.q(P, "timing"))
    if timing is None:
        return []
    seq = timing.find(".//" + D.q(P, "seq"))
    if seq is None:
        return []
    out = []
    child = seq.find(D.q(P, "cTn")).find(D.q(P, "childTnLst"))
    for click in (child if child is not None else []):
        ids, seen = [], []
        for tgt in click.iter(D.q(P, "spTgt")):
            prg = tgt.find(".//" + D.q(P, "pRg"))
            span = ("" if prg is None
                    else ":p%s-%s" % (prg.get("st"), prg.get("end")))
            ids.append(tgt.get("spid") + span)
        for i in ids:
            if i not in seen:
                seen.append(i)
        out.append(seen)
    return out


def sigmap(deck, disp):
    spTree = tree_of(deck, disp).find(".//" + D.q(P, "cSld") + "/"
                                      + D.q(P, "spTree"))
    out = {}

    def walk(el, ox=0.0, oy=0.0, sx=1.0, sy=1.0):
        for c in el:
            tag = ET.QName(c).localname
            if tag == "AlternateContent":
                ch = c.find("{http://schemas.openxmlformats.org/"
                            "markup-compatibility/2006}Choice")
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
            cnv = c.find(".//" + D.q(P, "cNvPr"))
            txt = D.norm("".join(t.text or "" for t in c.iter(D.q(A, "t"))))
            if cnv is not None:
                out[cnv.get("id")] = (tag, round(x / EMU, 2),
                                      round(y / EMU, 2), txt[:44])
            if tag == "grpSp" and xf is not None:
                cho, che = xf.find(D.q(A, "chOff")), xf.find(D.q(A, "chExt"))
                csx = w / int(che.get("cx")) if che is not None else sx
                csy = h / int(che.get("cy")) if che is not None else sy
                walk(c, x - int(cho.get("x")) * csx,
                     y - int(cho.get("y")) * csy, csx, csy)
    walk(spTree)
    return out


def moved_for_real(a, b):
    """Ignore PowerPoint's spAutoFit height recompute (see _vf_diff)."""
    if abs(a[2] - b[2]) > 0.011 or abs(a[4] - b[4]) > 0.011:
        return True
    dy = abs(a[3] - b[3])
    dc = abs((a[3] + a[5] / 2.0) - (b[3] + b[5] / 2.0))
    return dy > 0.011 and dc > 0.011


def pagenum(shape):
    return abs(shape[2] - 12.55) < 0.02 and abs(shape[3] - 7.20) < 0.02


def show_anim(deck, disp, label):
    sm, bs = sigmap(deck, disp), beats(deck, disp)
    print("  %s: %d click(s)" % (label, len(bs)))
    for i, ids in enumerate(bs, 1):
        parts = []
        for sid in ids:
            base = sid.split(":")[0]
            s = sm.get(base)
            parts.append("%s%s" % (
                ("%s(%.2f,%.2f)%r" % (s[0], s[1], s[2], s[3])) if s
                else "?" + base, sid[len(base):]))
        print("    %2d. %s" % (i, "  +  ".join(parts)))


def one(his, anim=False, brief=False):
    b = MAP(his)
    s1, n1 = struct(HIS, his)
    s2, n2 = struct(BUILD, b)
    print("=" * 78)
    print("HIS %d   <->   build %d      (%d vs %d shapes)"
          % (his, b, len(s1), len(s2)))
    print("=" * 78)
    if brief and len(s1) == len(s2):
        for i, (x, y) in enumerate(zip(s1, s2)):
            if pagenum(x) or (not moved_for_real(x, y) and x[6] == y[6]):
                continue
            print("  [%02d] his   (%6.2f,%5.2f) %5.2fx%4.2f | %s"
                  % (i, x[2], x[3], x[4], x[5], x[6][:220]))
            print("       build (%6.2f,%5.2f) %5.2fx%4.2f | %s"
                  % (y[2], y[3], y[4], y[5], y[6][:220]))
    else:
        for lbl, ss in (("HIS", s1), ("BUILD", s2)):
            print("--- %s: %d leaf shapes" % (lbl, len(ss)))
            for s in ss:
                print("   %s%-4s (%6.2f,%5.2f) %5.2fx%4.2f | %-44s | %s"
                      % ("  " * s[1], s[0], s[2], s[3], s[4], s[5],
                         s[6][:44], s[7]))
    if D.norm(n1) != D.norm(n2):
        print("--- NOTES DIFFER")
    if anim:
        print("--- ANIMATION")
        show_anim(HIS, his, "his  ")
        show_anim(BUILD, b, "build")


def summary():
    z = zipfile.ZipFile(HIS)
    n = len(ET.fromstring(z.read("ppt/presentation.xml"))
            .find(D.q(P, "sldIdLst")))
    z.close()
    print("%-5s %-6s %-9s %-9s %s"
          % ("his", "build", "shapes", "clicks", "flags"))
    for his in range(1, n + 1):
        b = MAP(his)
        s1, n1 = struct(HIS, his)
        s2, n2 = struct(BUILD, b)
        b1, b2 = beats(HIS, his), beats(BUILD, b)
        flags = []
        if len(s1) != len(s2):
            flags.append("SHAPES")
        else:
            mv = sum(1 for x, y in zip(s1, s2)
                     if not pagenum(x) and moved_for_real(x, y))
            if mv:
                flags.append("GEOM %d" % mv)
            tx = sum(1 for x, y in zip(s1, s2)
                     if x[6] != y[6] and not pagenum(x))
            if tx:
                flags.append("TEXT %d" % tx)
        if len(b1) != len(b2):
            flags.append("CLICKS")
        if D.norm(n1) != D.norm(n2):
            flags.append("NOTES")
        g1 = sum(1 for s in s1 if s[0] == "grp")
        g2 = sum(1 for s in s2 if s[0] == "grp")
        if g1 != g2:
            flags.append("GROUPS %d->%d" % (g2, g1))
        if flags:
            print("%-5d %-6d %3d->%-4d %3d->%-4d %s"
                  % (his, b, len(s2), len(s1), len(b2), len(b1),
                     " ".join(flags)))


if __name__ == "__main__":
    args = sys.argv[1:]
    if "--all" in args:
        summary()
    else:
        for a in args:
            if not a.startswith("--"):
                one(int(a), anim="--anim" in args, brief="--brief" in args)
