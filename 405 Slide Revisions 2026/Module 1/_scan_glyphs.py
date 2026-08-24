"""Inventory every link affordance in the canonical deck:
gold arrow glyphs, pointer pills, back pills, invisible click overlays,
and the slide-jump / external-hyperlink targets they carry."""
import os
import zipfile

from lxml import etree as ET

import _diff_slides as D

A, P, R = D.A, D.P, D.R
REL = D.REL
HERE = os.path.dirname(os.path.abspath(__file__))
import sys
DECK = os.path.join(HERE, sys.argv[1] if len(sys.argv) > 1
                    else "Module 1 - Revised.pptx")
GLYPHS = "▶►➤➡➜→←◀◄⏭⏮"
EMU = 914400.0

z = zipfile.ZipFile(DECK)
pres = ET.fromstring(z.read("ppt/presentation.xml"))
prels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
rmap = {r.get("Id"): r.get("Target") for r in prels}
order = []
for sid in pres.find(D.q(P, "sldIdLst")):
    order.append("ppt/" + rmap[sid.get(D.q(R, "id"))].lstrip("/"))

# map display index by slide part, for resolving jump targets
disp_of_part = {p: i + 1 for i, p in enumerate(order)}

for disp, part in enumerate(order, 1):
    tree = ET.fromstring(z.read(part))
    relp = part.replace("slides/", "slides/_rels/") + ".rels"
    try:
        srels = {r.get("Id"): (r.get("Type"), r.get("Target"))
                 for r in ET.fromstring(z.read(relp))}
    except KeyError:
        srels = {}
    rows = []
    for sp in tree.iter(D.q(P, "sp")):
        nv = sp.find(".//" + D.q(P, "cNvPr"))
        name = nv.get("name") if nv is not None else "?"
        txt = D.norm("".join(t.text or "" for t in sp.iter(D.q(A, "t"))))
        geom = sp.find(".//" + D.q(A, "prstGeom"))
        prst = geom.get("prst") if geom is not None else "-"
        off = sp.find(".//" + D.q(A, "off"))
        ext = sp.find(".//" + D.q(A, "ext"))
        pos = ""
        if off is not None and ext is not None:
            pos = "(%.2f,%.2f) %.2fx%.2f" % (
                int(off.get("x")) / EMU, int(off.get("y")) / EMU,
                int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU)
        # link info
        links = []
        for hl in sp.iter(D.q(A, "hlinkClick")):
            rid = hl.get(D.q(R, "id"))
            act = hl.get("action") or ""
            tgt = ""
            if rid and rid in srels:
                ty, target = srels[rid]
                if ty.endswith("/slide"):
                    tp = "ppt/" + target.replace("../", "")
                    tgt = "-> display %s" % disp_of_part.get(tp, "?")
                else:
                    tgt = target[:60]
            links.append(("JUMP " if "sldjump" in act else "URL  ") + tgt)
        has_glyph = any(c in GLYPHS for c in txt)
        transparent = False
        al = sp.find(".//" + D.q(A, "alpha"))
        if al is not None and al.get("val") == "0":
            transparent = True
        if has_glyph or links or transparent:
            kind = []
            if has_glyph:
                kind.append("GLYPH")
            if transparent:
                kind.append("OVERLAY")
            rows.append("  %-14s %-22s %-26s %-40s %s"
                        % ("/".join(kind) or "LINK", prst, pos,
                           txt[:40], "; ".join(links)))
    if rows:
        print("=== display %d" % disp)
        for r in rows:
            print(r)
z.close()
