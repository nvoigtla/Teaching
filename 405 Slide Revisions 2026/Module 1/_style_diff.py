# -*- coding: utf-8 -*-
"""Read-only: STYLE diff of In-Class vs Revised, pair by pair.

`_diff2.py` compares position, size, text and run formats. It is blind to
geometry presets, fills, outlines and effects - which is how a picture that
gained rounded corners and a drop shadow in the in-class deck was reported
as "identical" and never ported (Revised slide 39, 2026-09-23).

This walks the same shape tree and compares, per shape:
    prstGeom prst + its adjust values
    fill      (solid / none / group / blip / gradient, and the colour)
    line      (colour, width, dash, or noFill)
    effects   (outerShdw and its blur / dist / dir / alpha)
and, for a native table, every cell's fill.

Usage:  python _style_diff.py
"""
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, ".")
import zipfile                                              # noqa: E402
import difflib                                              # noqa: E402
from lxml import etree as ET                                # noqa: E402
import _diff2 as D                                          # noqa: E402
from _ic_vs_rev import MAP                                  # noqa: E402

A = D.A
P = D.P
MC = D.MC


def q(ns, t):
    return "{%s}%s" % (ns, t)


def colour(el):
    if el is None:
        return "-"
    s = el.find(".//" + q(A, "srgbClr"))
    if s is not None:
        a = s.find(q(A, "alpha"))
        return s.get("val") + ("@%s" % a.get("val") if a is not None else "")
    c = el.find(".//" + q(A, "schemeClr"))
    if c is not None:
        return "scheme:" + c.get("val")
    return "?"


def fill_of(sp):
    for kind in ("solidFill", "noFill", "grpFill", "blipFill", "gradFill",
                 "pattFill"):
        e = sp.find(q(A, kind))
        if e is not None:
            return kind + ("=" + colour(e) if kind == "solidFill" else "")
    return "inherit"


def line_of(sp):
    ln = sp.find(q(A, "ln"))
    if ln is None:
        return "inherit"
    if ln.find(q(A, "noFill")) is not None:
        return "none"
    dash = ln.find(q(A, "prstDash"))
    return "w=%s,%s%s" % (ln.get("w") or "-", colour(ln),
                          ",dash=" + dash.get("val") if dash is not None else "")


def effect_of(sp):
    eff = sp.find(q(A, "effectLst"))
    if eff is None:
        return "inherit"
    if not len(eff):
        return "none"
    out = []
    for ch in eff:
        n = ET.QName(ch).localname
        if n == "outerShdw":
            out.append("shdw(blur=%s,dist=%s,dir=%s,%s)"
                       % (ch.get("blurRad"), ch.get("dist"), ch.get("dir"),
                          colour(ch)))
        else:
            out.append(n)
    return "+".join(out)


def geom_of(sp):
    g = sp.find(q(A, "prstGeom"))
    if g is None:
        return "custGeom" if sp.find(q(A, "custGeom")) is not None else "-"
    adj = [gd.get("fmla") for gd in g.iter(q(A, "gd"))]
    return g.get("prst") + ("[" + ",".join(adj) + "]" if adj else "")


def walk(el, out=None, depth=0):
    if out is None:
        out = []
    for c in el:
        tag = ET.QName(c).localname
        if tag == "AlternateContent":
            ch = c.find(q(MC, "Choice"))
            if ch is not None and len(ch):
                walk(ch, out, depth)
            continue
        if tag not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
            continue
        txt = D.norm("".join(t.text or "" for t in c.iter(q(A, "t"))))[:40]
        # the footer page number is a live field: its TEXT differs between
        # the two decks by design, so key it on a constant instead
        if txt.isdigit():
            txt = "#pagenum"
        if tag == "grpSp":
            sp = c.find(q(P, "grpSpPr"))
            out.append((tag, depth, "", "fill=%s" % (fill_of(sp) if sp is not None else "-")))
            walk(c, out, depth + 1)
            continue
        if tag == "graphicFrame":
            cells = []
            for tc in c.iter(q(A, "tc")):
                tcPr = tc.find(q(A, "tcPr"))
                cells.append(colour(tcPr.find(q(A, "solidFill")))
                             if tcPr is not None else "-")
            out.append((tag, depth, txt, "cells=" + ",".join(cells)))
            continue
        sp = c.find(q(P, "spPr"))
        if sp is None:
            out.append((tag, depth, txt, "-"))
            continue
        out.append((tag, depth, txt,
                    "geom=%s fill=%s ln=%s eff=%s"
                    % (geom_of(sp), fill_of(sp), line_of(sp), effect_of(sp))))
    return out


def styles(deck):
    z = zipfile.ZipFile(deck)
    parts = D.slide_parts(z)
    out = []
    for p in parts:
        tree = ET.fromstring(z.read(p))
        spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
        out.append(walk(spTree))
    z.close()
    return out


def main():
    ic = styles("Module 1 - In Class.pptx")
    rv = styles("Module 1 - Revised.pptx")
    print("STYLE differences (IC dominant)\n")
    bad = 0
    for i in sorted(MAP):
        j = MAP[i]
        if j is None:
            continue
        a, b = ic[i - 1], rv[j - 1]
        ka = [(x[0], x[1], x[2]) for x in a]
        kb = [(x[0], x[1], x[2]) for x in b]
        lines = []
        sm = difflib.SequenceMatcher(None, ka, kb)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                for m, n in zip(range(i1, i2), range(j1, j2)):
                    if a[m][3] != b[n][3]:
                        lines.append("   %-5s '%s'\n       IC: %s\n       RV: %s"
                                     % (a[m][0], a[m][2][:34], a[m][3], b[n][3]))
            elif tag in ("delete", "replace"):
                for m in range(i1, i2):
                    lines.append("   ONLY-IN-IC %-5s '%s'  %s"
                                 % (a[m][0], a[m][2][:30], a[m][3]))
                for n in range(j1, j2):
                    lines.append("   ONLY-IN-RV %-5s '%s'  %s"
                                 % (b[n][0], b[n][2][:30], b[n][3]))
            elif tag == "insert":
                for n in range(j1, j2):
                    lines.append("   ONLY-IN-RV %-5s '%s'  %s"
                                 % (b[n][0], b[n][2][:30], b[n][3]))
        if lines:
            bad += 1
            print("--- IC %2d -> RV %2d" % (i, j))
            print("\n".join(lines))
    print("\nslides with style differences: %d" % bad)


if __name__ == "__main__":
    main()
