# -*- coding: utf-8 -*-
"""Whole-deck diff: canonical (Nico's hand-edited) vs. fresh side build.

Matches slides across the two decks by text signature (Jaccard on the set of
normalized text strings), then reports member-level geometry/text/format
differences, notes differences, and click-structure differences.

Usage:  python _diff2.py [--full] [display ...]
"""
import sys
import os
import zipfile
import difflib
from lxml import etree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
EMU = 914400.0

MATH_OFFSETS = [(0x1D434, 0x1D44D, ord('A')), (0x1D44E, 0x1D467, ord('a'))]


def norm(t):
    out = []
    for ch in t:
        o = ord(ch)
        for lo, hi, base in MATH_OFFSETS:
            if lo <= o <= hi:
                ch = chr(base + o - lo)
                break
        out.append(ch)
    return " ".join("".join(out).split())


def q(ns, t):
    return "{%s}%s" % (ns, t)


def slide_parts(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    rmap = {r.get("Id"): r.get("Target") for r in rels}
    out = []
    for sld in pres.find(q(P, "sldIdLst")):
        out.append("ppt/" + rmap[sld.get(q(R, "id"))].lstrip("/"))
    return out


def decode(el, ox=0.0, oy=0.0, sx=1.0, sy=1.0, out=None, depth=0):
    if out is None:
        out = []
    for c in el:
        tag = ET.QName(c).localname
        if tag == "AlternateContent":
            ch = c.find(q(MC, "Choice"))
            if ch is not None and len(ch):
                decode(ch, ox, oy, sx, sy, out, depth)
            continue
        if tag not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
            continue
        if tag == "graphicFrame":
            xf = c.find(q(P, "xfrm"))
        else:
            sppr = c.find(q(P, "spPr"))
            if sppr is None:
                sppr = c.find(q(P, "grpSpPr"))
            xf = sppr.find(q(A, "xfrm")) if sppr is not None else None
        x = y = w = h = 0.0
        if xf is not None:
            off = xf.find(q(A, "off"))
            ext = xf.find(q(A, "ext"))
            if off is not None and ext is not None:
                x = ox + int(off.get("x")) * sx
                y = oy + int(off.get("y")) * sy
                w = int(ext.get("cx")) * sx
                h = int(ext.get("cy")) * sy
        if tag == "grpSp":
            cho = xf.find(q(A, "chOff")) if xf is not None else None
            che = xf.find(q(A, "chExt")) if xf is not None else None
            csx = w / int(che.get("cx")) if che is not None and int(che.get("cx")) else sx
            csy = h / int(che.get("cy")) if che is not None and int(che.get("cy")) else sy
            cox = x - int(cho.get("x")) * csx if cho is not None else x
            coy = y - int(cho.get("y")) * csy if cho is not None else y
            out.append(("grp", depth, x / EMU, y / EMU, w / EMU, h / EMU, "", ""))
            decode(c, cox, coy, csx, csy, out, depth + 1)
            continue
        txt = norm("".join(t.text or "" for t in c.iter(q(A, "t"))))
        runs = []
        for r in c.iter(q(A, "rPr")):
            sz = r.get("sz")
            b = r.get("b")
            i = r.get("i")
            clr = r.find(".//" + q(A, "srgbClr"))
            runs.append("%s%s%s%s" % (
                sz or "-", "b" if b == "1" else "", "i" if i == "1" else "",
                ("#" + clr.get("val")) if clr is not None else ""))
        out.append((tag, depth, x / EMU, y / EMU, w / EMU, h / EMU,
                    txt[:120], ",".join(runs[:10])))
    return out


def notes_text(z, part):
    rp = part.replace("slides/", "slides/_rels/") + ".rels"
    try:
        rels = ET.fromstring(z.read(rp))
    except KeyError:
        return ""
    for r in rels.iter(q(REL, "Relationship")):
        if r.get("Type").endswith("/notesSlide"):
            np_ = "ppt/" + r.get("Target").replace("../", "")
            tree = ET.fromstring(z.read(np_))
            outp = []
            for sp in tree.iter(q(P, "sp")):
                ph = sp.find(".//" + q(P, "ph"))
                if ph is not None and ph.get("type") == "body":
                    for para in sp.iter(q(A, "p")):
                        t = "".join(n.text or "" for n in para.iter(q(A, "t")))
                        outp.append(t)
            return "\n".join(outp)
    return ""


def clicks(z, part):
    """Return list of click groups; each is a list of target shape ids."""
    tree = ET.fromstring(z.read(part))
    timing = tree.find(q(P, "timing"))
    if timing is None:
        return []
    # id -> signature
    spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
    sig = {}
    for c in spTree.iter():
        tag = ET.QName(c).localname
        if tag in ("cNvPr",):
            pass
    for c in spTree.iter(q(P, "cNvPr")):
        sig[c.get("id")] = c.get("name")
    groups = []
    for par in timing.iter(q(P, "par")):
        pass
    # walk the main sequence: seq > cTn > childTnLst > par (click groups)
    seq = None
    for s in timing.iter(q(P, "seq")):
        if s.get("concurrent") == "1" or seq is None:
            seq = s
            break
    if seq is None:
        return []
    ctn = seq.find(q(P, "cTn"))
    child = ctn.find(q(P, "childTnLst")) if ctn is not None else None
    if child is None:
        return []
    for clickpar in child:
        tgts = []
        for tgt in clickpar.iter(q(P, "spTgt")):
            tgts.append(sig.get(tgt.get("spid"), tgt.get("spid")))
        if tgts:
            groups.append(tgts)
    return groups


class Deck(object):
    def __init__(self, path):
        self.path = path
        self.z = zipfile.ZipFile(path)
        self.parts = slide_parts(self.z)
        self.shapes = []
        self.notes = []
        self.clicks = []
        for p in self.parts:
            tree = ET.fromstring(self.z.read(p))
            spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
            self.shapes.append(decode(spTree))
            self.notes.append(notes_text(self.z, p))
            self.clicks.append(clicks(self.z, p))

    def texts(self, i):
        return set(s[6] for s in self.shapes[i] if s[6])


def sim(a, b):
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return len(a & b) / float(len(a | b))


def match(can, bld):
    """Global greedy matching canonical slide -> build slide (no monotonicity)."""
    n, m = len(can.shapes), len(bld.shapes)
    ct = [can.texts(i) for i in range(n)]
    bt = [bld.texts(j) for j in range(m)]
    cand = []
    for i in range(n):
        for j in range(m):
            s = sim(ct[i], bt[j])
            if s >= 0.35:
                cand.append((s, i, j))
    cand.sort(reverse=True)
    ci, bj = {}, {}
    for s, i, j in cand:
        if i in ci or j in bj:
            continue
        ci[i] = (j, s)
        bj[j] = i
    return [(i, ci.get(i, (-1, 0.0))[0], ci.get(i, (-1, 0.0))[1]) for i in range(n)]


def shape_key(s):
    return (s[0], s[1], s[6])


def is_pagenum(s):
    return s[6].isdigit() and s[2] > 12.2 and s[3] > 6.9


def diff_slide(cs, bs):
    """Return list of difference strings."""
    cs = [s for s in cs if not is_pagenum(s)]
    bs = [s for s in bs if not is_pagenum(s)]
    out = []
    ck = [shape_key(s) for s in cs]
    bk = [shape_key(s) for s in bs]
    sm = difflib.SequenceMatcher(None, ck, bk)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for i, j in zip(range(i1, i2), range(j1, j2)):
                a, b = cs[i], bs[j]
                dp = max(abs(a[2] - b[2]), abs(a[3] - b[3]))
                dz = max(abs(a[4] - b[4]), abs(a[5] - b[5]))
                if dp > 0.012 or dz > 0.012:
                    out.append("  MOVED/RESIZED %-4s '%s' can(%.2f,%.2f %.2fx%.2f) bld(%.2f,%.2f %.2fx%.2f)"
                               % (a[0], a[6][:50], a[2], a[3], a[4], a[5], b[2], b[3], b[4], b[5]))
                if a[7] != b[7]:
                    out.append("  FORMAT   %-4s '%s' can[%s] bld[%s]"
                               % (a[0], a[6][:50], a[7][:70], b[7][:70]))
        elif tag == "delete":
            for i in range(i1, i2):
                a = cs[i]
                out.append("  ONLY-IN-CANONICAL %-4s (%.2f,%.2f %.2fx%.2f) '%s'"
                           % (a[0], a[2], a[3], a[4], a[5], a[6][:70]))
        elif tag == "insert":
            for j in range(j1, j2):
                b = bs[j]
                out.append("  ONLY-IN-BUILD     %-4s (%.2f,%.2f %.2fx%.2f) '%s'"
                           % (b[0], b[2], b[3], b[4], b[5], b[6][:70]))
        elif tag == "replace":
            for i in range(i1, i2):
                a = cs[i]
                out.append("  CHANGED-CAN %-4s (%.2f,%.2f %.2fx%.2f) '%s'"
                           % (a[0], a[2], a[3], a[4], a[5], a[6][:70]))
            for j in range(j1, j2):
                b = bs[j]
                out.append("  CHANGED-BLD %-4s (%.2f,%.2f %.2fx%.2f) '%s'"
                           % (b[0], b[2], b[3], b[4], b[5], b[6][:70]))
    return out


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    full = "--full" in sys.argv
    only = [int(x) for x in sys.argv[1:] if not x.startswith("-")]
    can = Deck(os.path.join(HERE, "Module 1 - Revised.pptx"))
    bld = Deck(os.path.join(HERE, "_sidebuild.pptx"))
    print("canonical %d slides / build %d slides" % (len(can.shapes), len(bld.shapes)))
    pairs = match(can, bld)
    matched_b = set(j for _, j, _ in pairs if j >= 0)
    print()
    print("=== SLIDE MAPPING (canonical -> build) ===")
    for i, j, s in pairs:
        print("  can %3d -> bld %s   sim=%.2f   %s"
              % (i + 1, ("%3d" % (j + 1)) if j >= 0 else " NONE", s,
                 (sorted(can.texts(i), key=len, reverse=True) or [""])[0][:60]))
    print()
    print("=== BUILD SLIDES WITH NO CANONICAL MATCH (deleted by hand) ===")
    for j in range(len(bld.shapes)):
        if j not in matched_b:
            t = sorted(bld.texts(j), key=len, reverse=True)
            print("  bld %3d  %s" % (j + 1, " | ".join(t[:3])[:110]))
    print()
    print("=== PER-SLIDE DIFFERENCES ===")
    for i, j, s in pairs:
        if j < 0:
            continue
        if only and (i + 1) not in only:
            continue
        d = diff_slide(can.shapes[i], bld.shapes[j])
        nd = can.notes[i].strip() != bld.notes[j].strip()
        cc = can.clicks[i]
        bc = bld.clicks[j]
        cd = len(cc) != len(bc)
        if d or nd or cd:
            print("--- can %d (bld %d) sim=%.2f" % (i + 1, j + 1, s))
            for line in d:
                print(line)
            if cd:
                print("  CLICKS canonical=%d build=%d" % (len(cc), len(bc)))
                if full:
                    for k, g in enumerate(cc):
                        print("     can click %d: %s" % (k + 1, g))
                    for k, g in enumerate(bc):
                        print("     bld click %d: %s" % (k + 1, g))
            elif full and cc != bc:
                print("  CLICK TARGETS DIFFER")
                for k, (g1, g2) in enumerate(zip(cc, bc)):
                    if g1 != g2:
                        print("     click %d can=%s bld=%s" % (k + 1, g1, g2))
            if nd:
                print("  NOTES DIFFER")
                if full:
                    for line in difflib.unified_diff(
                            bld.notes[j].splitlines(), can.notes[i].splitlines(),
                            "build", "canonical", lineterm="", n=0):
                        print("     " + line[:200])


if __name__ == "__main__":
    main()
