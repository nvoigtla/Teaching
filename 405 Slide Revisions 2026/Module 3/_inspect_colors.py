# -*- coding: utf-8 -*-
"""Read-only: show every off-palette color in context (fill vs. text,
geometry, the shape's own words) so each can be mapped by hand."""
import os
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

sys.stdout.reconfigure(encoding="utf-8")
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0
STRAYS = {"FFF5E0", "F4F1EA", "8B1A1A", "FFF2CC", "C0C0C0", "4F6128",
          "2E7D32", "C8CED6", "EEECE1", "D9D9D9", "404040", "9EC5F7",
          "FDECDB"}


def q(ns, t):
    return "{%s}%s" % (ns, t)


deck = Path(__file__).parent / "Module 3 - Revised.pptx"
z = zipfile.ZipFile(deck)
data = {n: z.read(n) for n in z.namelist()}
z.close()
pres = ET.fromstring(data["ppt/presentation.xml"])
rid2t = {r.get("Id"): r.get("Target") for r in
         ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
order = [os.path.basename(rid2t[s.get(q(R, "id"))])
         for s in pres.find(q(P, "sldIdLst"))]

for i, base in enumerate(order, 1):
    tree = ET.fromstring(data["ppt/slides/" + base])
    spTree = tree.find(q(P, "cSld") + "/" + q(P, "spTree"))
    for sp in spTree.iter(q(P, "sp")):
        txt = "".join(t.text or "" for t in sp.iter(q(A, "t"))).strip()
        off = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "off"))
        ext = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "ext"))
        g = ("[%.2f,%.2f %.2fx%.2f]"
             % (int(off.get("x")) / EMU, int(off.get("y")) / EMU,
                int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU)
             ) if off is not None and ext is not None else "[-]"
        link = "LINK" if sp.find(".//" + q(A, "hlinkClick")) is not None \
            else ""
        for clr in sp.iter(q(A, "srgbClr")):
            v = clr.get("val").upper()
            if v not in STRAYS:
                continue
            role = {"solidFill": "fill", "ln": "line", "rPr": "text",
                    "outerShdw": "shadow"}.get(
                        clr.getparent().tag.split("}")[-1], "?")
            if role == "fill":
                gp = clr.getparent().getparent().tag.split("}")[-1]
                role = "line-fill" if gp == "ln" else (
                    "text" if gp == "rPr" else "fill")
            print("%2d  #%s  %-9s %-26s %-4s %s"
                  % (i, v, role, g, link, txt[:44].replace("\n", " / ")))
