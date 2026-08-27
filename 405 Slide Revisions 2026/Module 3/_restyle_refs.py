# -*- coding: utf-8 -*-
"""Bring Module 3's reference boxes onto the current deck standard.

The deck still carried the ORIGINAL deck's pointer to a teaching note —
a white "folded corner / PDF" icon plus a separate text box reading
"SEE TEACHING NOTE  →  Bang-for-the-Buck Rule". The current standard
(Module 2, 2026-08-24/25, and Teaching CLAUDE.md "Post-work reference
box") is ONE gold-bordered rounded box, white fill, soft drop shadow,
navy bold, with the fixed glyph vocabulary: ✎ = problem set,
▤ = teaching note, ▶ = video. Position: bottom-RIGHT corner overlaying
the footer, right edge on the PS_BOX_XY line so every pointer in the
course lands on the same spot.

Module 1's older "➜ Problem Set 1" pills predate that vocabulary
(2026-08-23) and are not the model here.

The box is generated with Module 2's own helper layer, so the styling
is the shared one rather than a re-implementation, and is then spliced
into the frozen deck by OOXML surgery. Rerunnable: it is a no-op once
the old shapes are gone.
"""
import copy
import os
import sys
import zipfile
import shutil
from pathlib import Path

from lxml import etree as ET
from PIL import ImageFont

sys.stdout.reconfigure(encoding="utf-8")
HERE = Path(__file__).parent
DECK = HERE / "Module 3 - Revised.pptx"
M2_DIR = HERE.parent / "Module 2"
sys.path.insert(0, str(M2_DIR))

import _build_Module2InClass as M2  # noqa: E402

from pptx.util import Inches, Pt  # noqa: E402

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# The old pointer was THREE shapes: a peach dashed callout, a white
# "folded corner / PDF" icon, and the label text. Identified by content
# and fill, never by shape name.
OLD_TEXTS = ("SEE TEACHING NOTE", "PDF")
OLD_FILL = "FAC090"          # peach — not in the deck palette
LABEL = "Teaching Note:  Bang-for-the-Buck Rule"
SIZE = 15
HEIGHT = Inches(0.50)
RIGHT_EDGE = M2.PS_BOX_XY[0] + Inches(3.00)     # shared right edge
TOP = M2.PS_BOX_XY[1]


def q(ns, t):
    return "{%s}%s" % (ns, t)


def order_of(data):
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target") for r in
             ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    return [os.path.basename(rid2t[s.get(q(R, "id"))])
            for s in pres.find(q(P, "sldIdLst"))]


def text_of(el):
    return "".join(t.text or "" for t in el.iter(q(A, "t")))


def measure(text, pts, bold=True):
    """Rendered width in inches, measured with the real font."""
    name = "calibrib.ttf" if bold else "calibri.ttf"
    try:
        font = ImageFont.truetype(name, int(pts * 4))
    except OSError:
        font = ImageFont.truetype("C:/Windows/Fonts/" + name, int(pts * 4))
    return font.getlength(text) / 4.0 / 72.0


# --- generate the replacement box ------------------------------------------
text = "%s  %s" % (M2.TN_GLYPH, LABEL)
width = int(Inches(measure(text, SIZE) + 0.34))   # 0.1 margins + slack
prs = M2.Presentation()
prs.slide_width = int(M2.SLIDE_W)
prs.slide_height = int(M2.SLIDE_H)
slide = M2._blank_slide(prs)
M2._add_reference_box(slide, int(RIGHT_EDGE - width), int(TOP), width,
                      int(HEIGHT), LABEL, kind="tn", size=SIZE)
tmp = HERE / "_refs_tmp.pptx"
prs.save(str(tmp))
z = zipfile.ZipFile(tmp)
tdata = {n: z.read(n) for n in z.namelist()}
z.close()
ttree = ET.fromstring(tdata["ppt/slides/" + order_of(tdata)[0]])
new_sp = ttree.find(q(P, "cSld") + "/" + q(P, "spTree") + "/" + q(P, "sp"))
tmp.unlink()
print("reference box: %.2f x %.2f\" at x %.2f\"  %r"
      % (width / 914400, HEIGHT / 914400,
         (RIGHT_EDGE - width) / 914400, text))

# --- splice it into every slide that still has the old pointer -------------
z = zipfile.ZipFile(DECK)
data = {n: z.read(n) for n in z.namelist()}
z.close()

def shape_id(sp):
    return sp.find(q(P, "nvSpPr") + "/" + q(P, "cNvPr")).get("id")


def drop_effects_for(tree, spid):
    """Remove the animation beat that revealed a deleted shape. An effect
    left pointing at a missing spid makes PowerPoint refuse to open the
    deck ("PowerPoint could not open the file") — the editing canvas and
    python-pptx never notice."""
    timing = tree.find(q(P, "timing"))
    if timing is None:
        return 0
    gone = 0
    for par in list(timing.iter(q(P, "par"))):
        cTn = par.find(q(P, "cTn"))
        if cTn is None or cTn.get("nodeType") is None:
            continue
        spids = {t.get("spid") for t in par.iter(q(P, "spTgt"))}
        if spids == {spid}:
            par.getparent().remove(par)
            gone += 1
    return gone


def click_carrier(tree, ids):
    """Of *ids*, the one whose effect starts the beat (nodeType
    clickEffect). The new box inherits that id so the slide keeps its
    click count; the others are deleted with their with-previous
    effects."""
    timing = tree.find(q(P, "timing"))
    if timing is not None:
        for par in timing.iter(q(P, "par")):
            cTn = par.find(q(P, "cTn"))
            if cTn is None or cTn.get("nodeType") != "clickEffect":
                continue
            spids = {t.get("spid") for t in par.iter(q(P, "spTgt"))}
            hit = spids & set(ids)
            if len(hit) == 1:
                return hit.pop()
    return ids[0]


hits = 0
for i, base in enumerate(order_of(data), 1):
    part = "ppt/slides/" + base
    tree = ET.fromstring(data[part])
    spTree = tree.find(q(P, "cSld") + "/" + q(P, "spTree"))
    old = []
    for sp in spTree.findall(q(P, "sp")):
        t = text_of(sp).strip()
        fill = sp.find(q(P, "spPr") + "/" + q(A, "solidFill") + "/"
                       + q(A, "srgbClr"))
        if (t.startswith(OLD_TEXTS[0]) or t == OLD_TEXTS[1]
                or (fill is not None and fill.get("val") == OLD_FILL)):
            old.append(sp)
    if not any(text_of(sp).strip().startswith(OLD_TEXTS[0]) for sp in old):
        continue
    by_id = {shape_id(sp): sp for sp in old}
    keep_id = click_carrier(tree, list(by_id))
    box = copy.deepcopy(new_sp)
    cNvPr = box.find(q(P, "nvSpPr") + "/" + q(P, "cNvPr"))
    cNvPr.set("id", keep_id)
    cNvPr.set("name", "Reference Box")
    spTree.replace(by_id[keep_id], box)
    print("  display %d: pointer restyled onto shape id %s" % (i, keep_id))
    for sid, sp in by_id.items():
        if sid == keep_id:
            continue
        spTree.remove(sp)
        n = drop_effects_for(tree, sid)
        print("  display %d: dropped old shape id %s (%d effect(s))"
              % (i, sid, n))
    data[part] = ET.tostring(tree, xml_declaration=True, encoding="UTF-8",
                             standalone=True)
    hits += 1

if not hits:
    print("nothing to do — no old-style pointers left")
    raise SystemExit(0)

out = DECK.with_suffix(".refs_tmp.pptx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
    for name, blob in data.items():
        zout.writestr(name, blob)
shutil.move(str(out), str(DECK))
print("saved %s" % DECK.name)
