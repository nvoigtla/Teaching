# -*- coding: utf-8 -*-
"""Give a hand-added picture the deck's standard treatment — rounded
corners and the soft drop shadow (Teaching CLAUDE.md, "Pictures").

2026-08-26 (Nico): the image dropped onto the deck title slide.

Scope is deliberately narrow: only pictures on the slides named in
TARGETS, and only those that are still square-cornered and unshaded.
Logos, book covers and screenshots elsewhere in the deck are the flat
exception and must not be swept up, which is why this does not run over
every picture in the file.

Geometry and shadow are _apply_picture_style from Module 2's helper
layer (corner 8 %, 4 pt blur, 3 pt offset, 45° down-right, 50 % black),
reproduced here because the deck is edited by OOXML surgery rather than
rebuilt.

Rerunnable: a picture that already has rounding and a shadow is left
alone.
"""
import os
import sys
import zipfile
import shutil
from pathlib import Path

from lxml import etree as ET

sys.stdout.reconfigure(encoding="utf-8")
HERE = Path(__file__).parent
DECK = HERE / "Module 3 - Revised.pptx"

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Slides to treat, each an ALL-OF list of strings the slide must carry.
# Pick a marker unique to that slide: "Production and Costs" alone also
# matches the FOOTER of every content slide, which swept 18 pictures
# deck-wide on the first run.
TARGETS = [["Production and Costs", "EMBA"]]      # the deck title slide
CORNER_PCT = 8
SHADOW = ('<a:effectLst xmlns:a="%s"><a:outerShdw blurRad="50800" '
          'dist="38100" dir="2700000" rotWithShape="0">'
          '<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>'
          '</a:outerShdw></a:effectLst>' % A)


def q(ns, t):
    return "{%s}%s" % (ns, t)


def order_of(data):
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target") for r in
             ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    return [os.path.basename(rid2t[s.get(q(R, "id"))])
            for s in pres.find(q(P, "sldIdLst"))]


def style_picture(pic):
    """roundRect geometry + soft shadow, in schema order (xfrm ->
    prstGeom -> ... -> effectLst)."""
    spPr = pic.find(q(P, "spPr"))
    if spPr is None:
        return False
    pg = spPr.find(q(A, "prstGeom"))
    shd = spPr.find(q(A, "effectLst") + "/" + q(A, "outerShdw"))
    if pg is not None and pg.get("prst") == "roundRect" and shd is not None:
        return False
    for old in spPr.findall(q(A, "prstGeom")):
        spPr.remove(old)
    pg = ET.Element(q(A, "prstGeom"))
    pg.set("prst", "roundRect")
    av = ET.SubElement(pg, q(A, "avLst"))
    gd = ET.SubElement(av, q(A, "gd"))
    gd.set("name", "adj")
    gd.set("fmla", "val %d" % (CORNER_PCT * 1000))
    xfrm = spPr.find(q(A, "xfrm"))
    if xfrm is not None:
        xfrm.addnext(pg)
    else:
        spPr.insert(0, pg)
    for old in spPr.findall(q(A, "effectLst")):
        spPr.remove(old)
    spPr.append(ET.fromstring(SHADOW))
    return True


z = zipfile.ZipFile(DECK)
data = {n: z.read(n) for n in z.namelist()}
z.close()

hits = 0
for i, base in enumerate(order_of(data), 1):
    part = "ppt/slides/" + base
    tree = ET.fromstring(data[part])
    text = " ".join(t.text or "" for t in tree.iter(q(A, "t")))
    if not any(all(m in text for m in group) for group in TARGETS):
        continue
    touched = False
    for pic in tree.iter(q(P, "pic")):
        xfrm = pic.find(q(P, "spPr") + "/" + q(A, "xfrm"))
        ext = xfrm.find(q(A, "ext")) if xfrm is not None else None
        if style_picture(pic):
            touched = True
            hits += 1
            print("  display %2d: rounded + shaded  %.2f x %.2f\""
                  % (i, int(ext.get("cx")) / 914400,
                     int(ext.get("cy")) / 914400))
    if touched:
        data[part] = ET.tostring(tree, xml_declaration=True,
                                 encoding="UTF-8", standalone=True)

if not hits:
    print("nothing to do — the targeted pictures are already styled")
    raise SystemExit(0)

out = DECK.with_suffix(".pic_tmp.pptx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
    for name, blob in data.items():
        zout.writestr(name, blob)
shutil.move(str(out), str(DECK))
print("saved %s" % DECK.name)
