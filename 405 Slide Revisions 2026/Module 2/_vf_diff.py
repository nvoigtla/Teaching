# -*- coding: utf-8 -*-
"""Compare one of OUR video slides with its counterpart in the three
FINAL video decks (Videos Final/).

    python _vf_diff.py 22            # our display 22 vs its final twin
    python _vf_diff.py 22 --anim     # ... and the click structure
    python _vf_diff.py --all         # every mapped pair, summary only

Reports leaf-level geometry / text / run formatting, the grouping
structure, the notes, and (with --anim) the animation beats resolved to
shape signatures so the two choreographies can be compared beat for beat.
"""
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

import _diff_slides as D

HERE = Path(__file__).parent
import os
OURS = str(HERE / os.environ.get("VF_OURS",
                                 "Module 2 - Video Part Revised.pptx"))
VF = HERE / "Videos Final"
V1 = str(VF / "Module 2 - Video 1 - Elasticity and Revenue.pptx")
V2 = str(VF / "Module 2 - Video 2 - Marginal Revenue.pptx")
V3 = str(VF / "Module 2 - Video 3 - Demand Estimation.pptx")

A, P = D.A, D.P
EMU = D.EMU

# our display number -> (final deck, its display number)
MAP = {
    # our display -> its slide in one of the final decks.
    # (Since 2026-08-25 our order matches theirs, so the matched slides
    # run in the same sequence.)
    # ---- Video 1 -------------------------------------------------
    1: (V1, 1), 2: (V1, 2), 3: (V1, 3), 4: (V1, 4), 5: (V1, 5),
    6: (V1, 6), 7: (V1, 7), 8: (V1, 8), 9: (V1, 9), 10: (V1, 10),
    11: (V1, 11), 14: (V1, 12),
    # 12, 13 (Netflix pair) and 15 (Mega Millions) are not in the final
    # ---- Video 2 -------------------------------------------------
    16: (V2, 1), 17: (V2, 2), 18: (V2, 3), 19: (V2, 4), 20: (V2, 5),
    21: (V2, 6), 22: (V2, 7), 23: (V2, 8), 24: (V2, 9), 25: (V2, 10),
    # 26, 27 (Inside Out 2 pair) are not in the final
    # ---- Video 3 -------------------------------------------------
    28: (V3, 1), 29: (V3, 2), 30: (V3, 3), 31: (V3, 4), 32: (V3, 5),
    33: (V3, 6), 34: (V3, 7), 35: (V3, 8), 36: (V3, 9), 37: (V3, 10),
    38: (V3, 11), 39: (V3, 12), 40: (V3, 13), 43: (V3, 14),
    46: (V3, 15), 49: (V3, 16), 50: (V3, 17), 51: (V3, 18),
    52: (V3, 19), 53: (V3, 20), 54: (V3, 21), 55: (V3, 22),
    56: (V3, 23),
    # 41/42, 44/45, 47/48 (poll pairs) and 57 (summary) are not in the
    # final decks
}

# the order the final decks imply, with our unmatched slides kept in
# place relative to their neighbours
NEW_ORDER = (
    [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]
    + [16, 17, 18, 19, 20, 21, 24, 25, 22, 23, 26, 27]
    + [28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39]
    + [42, 43, 44, 45, 46, 47, 48, 49, 50, 51]
    + [40, 41]
    + [52, 53, 54, 55, 56, 57]
)


def tree_of(deck, disp):
    z = zipfile.ZipFile(deck)
    part = D.slide_part(z, disp)
    tree = ET.fromstring(z.read(part))
    z.close()
    return tree


def struct(deck, disp):
    """Nested shape listing (groups shown with their children indented)."""
    return D.dump(deck, disp)


def beats(deck, disp):
    """[(trigger, [spid, ...]), ...] in mainSeq order."""
    tree = tree_of(deck, disp)
    timing = tree.find(D.q(P, "timing"))
    if timing is None:
        return []
    seq = timing.find(".//" + D.q(P, "seq"))
    if seq is None:
        return []
    main = seq.find(D.q(P, "cTn"))
    out = []
    for click in main.find(D.q(P, "childTnLst")) or []:
        ids = []
        for tgt in click.iter(D.q(P, "spTgt")):
            prg = tgt.find(".//" + D.q(P, "pRg"))
            span = ("" if prg is None
                    else ":p%s-%s" % (prg.get("st"), prg.get("end")))
            ids.append(tgt.get("spid") + span)
        seen = []
        for i in ids:
            if i not in seen:
                seen.append(i)
        out.append(seen)
    return out


def sigmap(deck, disp):
    """spid -> a stable signature (kind, x, y, text)."""
    tree = tree_of(deck, disp)
    spTree = tree.find(".//" + D.q(P, "cSld") + "/" + D.q(P, "spTree"))
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


def show_anim(deck, disp, label):
    sm = sigmap(deck, disp)
    bs = beats(deck, disp)
    print("  %s: %d click(s)" % (label, len(bs)))
    for i, ids in enumerate(bs, 1):
        parts = []
        for sid in ids:
            base = sid.split(":")[0]
            span = sid[len(base):]
            s = sm.get(base)
            parts.append("%s%s" % (
                ("%s(%.2f,%.2f)%r" % (s[0], s[1], s[2], s[3])) if s
                else "?" + base, span))
        print("    %2d. %s" % (i, "  +  ".join(parts)))


def one(disp, anim=False):
    if disp not in MAP:
        print("display %d: no counterpart in the final decks - untouched"
              % disp)
        return
    deck, fd = MAP[disp]
    s1, n1 = struct(OURS, disp)
    s2, n2 = struct(deck, fd)
    print("=" * 78)
    print("OUR %d   <->   %s  slide %d" % (disp, Path(deck).name, fd))
    print("=" * 78)
    print("--- OURS: %d leaf shapes" % len(s1))
    for s in s1:
        print("   %s%-4s (%6.2f,%5.2f) %5.2fx%4.2f | %-46s | %s"
              % ("  " * s[1], s[0], s[2], s[3], s[4], s[5], s[6][:46], s[7]))
    print("--- FINAL: %d leaf shapes" % len(s2))
    for s in s2:
        print("   %s%-4s (%6.2f,%5.2f) %5.2fx%4.2f | %-46s | %s"
              % ("  " * s[1], s[0], s[2], s[3], s[4], s[5], s[6][:46], s[7]))
    if D.norm(n1) != D.norm(n2):
        print("--- NOTES DIFFER")
        print("  ours : %s" % n1[:400].replace("\n", " | "))
        print("  final: %s" % n2[:400].replace("\n", " | "))
    if anim:
        print("--- ANIMATION")
        show_anim(OURS, disp, "ours ")
        show_anim(deck, fd, "final")


def moved_for_real(a, b):
    """True if a shape really moved, ignoring PowerPoint's spAutoFit
    height recompute.  A top-anchored box keeps its y and changes h; a
    centred box keeps its vertical CENTRE and changes y.  Either way the
    shape did not move - only x, width, or BOTH y and centre matter."""
    if abs(a[2] - b[2]) > 0.011 or abs(a[4] - b[4]) > 0.011:
        return True
    dy = abs(a[3] - b[3])
    dc = abs((a[3] + a[5] / 2.0) - (b[3] + b[5] / 2.0))
    return dy > 0.011 and dc > 0.011


def pagenum(shape):
    """The footer slide number - different on every pair by design."""
    return abs(shape[2] - 12.55) < 0.02 and abs(shape[3] - 7.20) < 0.02


def brief(disp):
    """Only the leaf shapes that actually differ (index-aligned)."""
    if disp not in MAP:
        print("s%-3d no counterpart - untouched" % disp)
        return
    deck, fd = MAP[disp]
    s1, n1 = struct(OURS, disp)
    s2, n2 = struct(deck, fd)
    print("--- s%d  vs  %s #%d   (%d vs %d shapes)"
          % (disp, Path(deck).name[11:24], fd, len(s1), len(s2)))
    if len(s1) != len(s2):
        print("    shape counts differ - use the full view")
        return
    for i, (a, b) in enumerate(zip(s1, s2)):
        if pagenum(a):
            continue
        same_geom = not moved_for_real(a, b)
        if same_geom and a[6] == b[6]:
            continue
        print("  [%02d] ours  (%6.2f,%5.2f) %5.2fx%4.2f | %s"
              % (i, a[2], a[3], a[4], a[5], a[6][:300]))
        print("       final (%6.2f,%5.2f) %5.2fx%4.2f | %s"
              % (b[2], b[3], b[4], b[5], b[6][:300]))
    if D.norm(n1) != D.norm(n2):
        print("  NOTES DIFFER")


def summary():
    print("%-5s %-42s %-9s %-9s %s"
          % ("ours", "final", "shapes", "clicks", "flags"))
    for disp in sorted(MAP):
        deck, fd = MAP[disp]
        s1, n1 = struct(OURS, disp)
        s2, n2 = struct(deck, fd)
        b1, b2 = beats(OURS, disp), beats(deck, fd)
        flags = []
        if len(s1) != len(s2):
            flags.append("SHAPES")
        if len(b1) != len(b2):
            flags.append("CLICKS")
        if D.norm(n1) != D.norm(n2):
            flags.append("NOTES")
        g1 = sum(1 for s in s1 if s[0] == "grp")
        g2 = sum(1 for s in s2 if s[0] == "grp")
        if g1 != g2:
            flags.append("GROUPS %d->%d" % (g1, g2))
        if len(s1) == len(s2):
            # noise to ignore: the footer page number always differs, and
            # PowerPoint recomputes the HEIGHT of every spAutoFit text box
            # on save, so height alone is not a hand-edit
            moved = sum(1 for a, b in zip(s1, s2)
                        if not pagenum(a) and moved_for_real(a, b))
            if moved:
                flags.append("GEOM %d" % moved)
            txt = sum(1 for a, b in zip(s1, s2)
                      if a[6] != b[6] and not pagenum(a))
            if txt:
                flags.append("TEXT %d" % txt)
        t1 = [s[6] for s in s1 if 0.45 <= s[3] < 1.25 and s[6]]
        t2 = [s[6] for s in s2 if 0.45 <= s[3] < 1.25 and s[6]]
        if t1[:1] != t2[:1]:
            flags.append("TITLE")
        print("%-5d %-42s %3d->%-4d %3d->%-4d %s"
              % (disp, "%s #%d" % (Path(deck).name[11:24], fd),
                 len(s1), len(s2), len(b1), len(b2),
                 " ".join(flags) or "-"))


if __name__ == "__main__":
    args = sys.argv[1:]
    if "--all" in args:
        summary()
    elif "--brief" in args:
        for a in args:
            if not a.startswith("--"):
                brief(int(a))
    else:
        anim = "--anim" in args
        for a in args:
            if a.startswith("--"):
                continue
            one(int(a), anim)
