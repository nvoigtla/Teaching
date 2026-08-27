# -*- coding: utf-8 -*-
"""Read-only formatting audit of Module 3 against the Teaching CLAUDE.md
conventions. Reports; changes nothing."""
import os
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

sys.stdout.reconfigure(encoding="utf-8")
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0

PALETTE = {"0B2B4E", "E09F3E", "555B66", "C8CDD3", "FDF6E6", "B0B5BC",
           "F6E8C9", "FFFFFF", "000000", "BFBFBF",
           "0070C0", "B8860B",            # reserved pedagogical colors
           "1B5E20", "C00000", "FF0000", "008000"}  # deck accents

SMALL = {"a", "an", "the", "and", "but", "or", "nor", "of", "in", "at",
         "to", "for", "with", "on", "as", "by", "from", "vs"}


def q(ns, t):
    return "{%s}%s" % (ns, t)


def order_of(data):
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target") for r in
             ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    return [os.path.basename(rid2t[s.get(q(R, "id"))])
            for s in pres.find(q(P, "sldIdLst"))]


def shape_text(sp):
    return "".join(t.text or "" for t in sp.iter(q(A, "t")))


def geom(sp):
    off = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "off"))
    ext = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "ext"))
    if off is None or ext is None:
        return None
    return (int(off.get("x")) / EMU, int(off.get("y")) / EMU,
            int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU)


def title_case_bad(title):
    """Words that should be capitalised but are not (never the reverse)."""
    words = re.findall(r"[A-Za-z][A-Za-z'’\-]*", title)
    bad = []
    for j, w in enumerate(words):
        if w[0].isupper():
            continue
        if w.lower() in SMALL and 0 < j < len(words) - 1:
            continue
        bad.append(w)
    return bad


deck = Path(__file__).parent / "Module 3 - Revised.pptx"
z = zipfile.ZipFile(deck)
data = {n: z.read(n) for n in z.namelist()}
z.close()
order = order_of(data)

colors, titles, periods, small_text, flat = [], [], [], [], []
for i, base in enumerate(order, 1):
    tree = ET.fromstring(data["ppt/slides/" + base])
    spTree = tree.find(q(P, "cSld") + "/" + q(P, "spTree"))
    for sp in spTree.iter(q(P, "sp")):
        g = geom(sp)
        txt = shape_text(sp).strip()
        # --- palette outliers (shape fills and run colors)
        for clr in sp.iter(q(A, "srgbClr")):
            v = clr.get("val").upper()
            if v not in PALETTE:
                colors.append((i, v, txt[:40]))
        # --- title band: title case
        if g and 0.40 < g[1] < 0.95 and g[0] < 1.0 and g[2] > 8 and txt:
            bad = title_case_bad(txt)
            if bad:
                titles.append((i, txt[:70], bad[:6]))
        # --- body text: trailing periods on short lines, tiny type
        if g and g[1] > 1.2:
            for para in sp.iter(q(A, "p")):
                line = "".join(t.text or "" for t in para.iter(q(A, "t")))
                line = line.strip()
                if (line and len(line) < 75 and line.endswith(".")
                        and not line.endswith("etc.")
                        and line.count(".") == 1):
                    periods.append((i, line[:60]))
                for rpr in para.iter(q(A, "rPr")):
                    sz = rpr.get("sz")
                    if sz and int(sz) < 1400 and line and len(line) > 12:
                        small_text.append((i, int(sz) / 100.0, line[:45]))
                        break
        # --- filled content boxes drawn flat (no shadow / square corners)
        pg = sp.find(q(P, "spPr") + "/" + q(A, "prstGeom"))
        fill = sp.find(q(P, "spPr") + "/" + q(A, "solidFill") + "/"
                       + q(A, "srgbClr"))
        eff = sp.find(q(P, "spPr") + "/" + q(A, "effectLst") + "/"
                      + q(A, "outerShdw"))
        if (pg is not None and fill is not None and eff is None and txt
                and g and g[1] > 1.0 and g[3] < 2.0
                and pg.get("prst") == "rect"):
            flat.append((i, fill.get("val"), g, txt[:45]))


def head(t):
    print("\n" + t)
    print("-" * len(t))


head("1. Colors outside the palette")
seen = {}
for i, v, t in colors:
    seen.setdefault(v, []).append((i, t))
for v, hits in sorted(seen.items(), key=lambda kv: -len(kv[1])):
    slides = sorted({h[0] for h in hits})
    print("  #%s  on slides %s   e.g. %r"
          % (v, slides[:12], hits[0][1]))

head("2. Slide titles not in title case")
for i, t, bad in titles:
    print("  %2d  %-68s  -> %s" % (i, t, bad))

head("3. Trailing periods on short lines")
for i, t in periods[:40]:
    print("  %2d  %s" % (i, t))
print("  (%d total)" % len(periods))

head("4. Body text below 14 pt")
for i, sz, t in small_text[:40]:
    print("  %2d  %4.1f pt  %s" % (i, sz, t))
print("  (%d total)" % len(small_text))

head("5. Flat filled content boxes (square corners, no shadow)")
for i, v, g, t in flat[:40]:
    print("  %2d  #%s  [%.2f,%.2f %.2fx%.2f]  %s"
          % (i, v, g[0], g[1], g[2], g[3], t))
print("  (%d total)" % len(flat))
