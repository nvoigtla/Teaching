# -*- coding: utf-8 -*-
"""Retrofit Module 3's five agenda/divider slides (displays 8, 13, 31,
41, 62) to the numbered-circle outline format (Teaching CLAUDE.md
"Module-Outline / Agenda Slides"), approved by Nico 2026-08-20.

Module 3's build script is FROZEN (deck = source of truth), so this is
in-place OOXML surgery: the replacement slides are generated with
Module 2's reference maker (make_m2_outline) in a temp deck, then each
target slide's <p:spTree> is swapped for the generated one. Chrome,
footer (live slidenum field), bands, circles all come from the maker;
speaker notes and slide rels stay untouched (the new shapes reference
no rels). None of the five slides carries animations (verified).

Rerunnable: each run regenerates the same five spTrees from scratch.
"""
import os
import sys
import zipfile
import shutil
from pathlib import Path

from lxml import etree as ET

sys.stdout.reconfigure(encoding='utf-8')
HERE = Path(__file__).parent
DECK = HERE / "Module 3 - Revised.pptx"
M2_DIR = HERE.parent / "Module 2"
sys.path.insert(0, str(M2_DIR))

import _build_Module2InClass as M2  # noqa: E402  (helper layer + maker)

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def q(ns, t):
    return "{%s}%s" % (ns, t)


# --- the approved Module 3 flat outline (5 items) -------------------------
M3_OUTLINE = [
    ("The Production Function",
     "How inputs (capital and labor) are turned into output"),
    ("Short Run: Hiring Decisions",
     "How many workers to hire when capital is fixed"),
    ("Long Run: The Optimal Input Mix",
     "Choosing between capital and labor when both can be adjusted"),
    ("Cost Concepts",
     "Fixed, variable, marginal, and sunk costs, and which ones matter "
     "for decisions"),
    ("Economies of Scale and Scope",
     "When bigger or broader production lowers cost per unit"),
]

# display -> (highlight_set, page_num)
TARGETS = {
    8:  {0, 1, 2},   # Part 1 divider -> items 1-3
    13: {1},         # Short Run
    31: {2},         # Long Run
    41: {3, 4},      # Part 2 divider -> items 4-5
    62: {4},         # Economies of Scale and Scope
}
TAG = "Module 3 · Outline"
TITLE = "Outline of Module 3"

# --- generate replacement slides with the M2 maker ------------------------
M2.M2_OUTLINE = M3_OUTLINE
M2.FOOTER_TEXT = "Management 405  ·  Module 3  ·  Production and Costs"

prs = M2.Presentation()
prs.slide_width = int(M2.SLIDE_W)
prs.slide_height = int(M2.SLIDE_H)
for disp in sorted(TARGETS):
    M2.make_m2_outline(prs, disp, section_tag=TAG, title=TITLE,
                       highlight_set=TARGETS[disp])
tmp = HERE / "_agenda_tmp.pptx"
prs.save(str(tmp))
print("generated %d replacement slides -> %s" % (len(TARGETS), tmp.name))

# --- read both decks, map display -> slide part ---------------------------
def slide_order(data):
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target") for r in
             ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    return [os.path.basename(rid2t[s.get(q(R, "id"))])
            for s in pres.find(q(P, "sldIdLst"))]


z = zipfile.ZipFile(tmp)
tdata = {n: z.read(n) for n in z.namelist()}
z.close()
torder = slide_order(tdata)          # temp slide i = sorted(TARGETS)[i]

z = zipfile.ZipFile(DECK)
data = {n: z.read(n) for n in z.namelist()}
z.close()
order = slide_order(data)

# --- transplant spTrees ----------------------------------------------------
for i, disp in enumerate(sorted(TARGETS)):
    tpart = "ppt/slides/" + torder[i]
    ttree = ET.fromstring(tdata[tpart])
    new_sptree = ttree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))

    part = "ppt/slides/" + order[disp - 1]
    tree = ET.fromstring(data[part])
    csld = tree.find(q(P, "cSld"))
    old_sptree = csld.find(q(P, "spTree"))
    csld.replace(old_sptree, new_sptree)
    timing = tree.find(q(P, "timing"))
    if timing is not None:
        tree.remove(timing)
    data[part] = ET.tostring(tree, xml_declaration=True,
                             encoding="UTF-8", standalone=True)
    print("display %2d (%s) <- %s  [items %s]"
          % (disp, part, tpart, sorted(TARGETS[disp])))

out = DECK.with_suffix(".retrofit_tmp.pptx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
    for name, blob in data.items():
        zout.writestr(name, blob)
shutil.move(str(out), str(DECK))
tmp.unlink()
print("saved %s" % DECK)
