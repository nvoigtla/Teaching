# -*- coding: utf-8 -*-
"""Add the lottery poll placeholder + solution to
"Module 2 - Potential Practice Exercises.pptx" (2026-09-20, Nico).

Nico's hand-made slide 4 ("Lottery Sales, Revisited") asks whether the
WROR/WBZ article is right that Mega Millions revenue rose in MA after
the April 2025 $2 -> $5 price change. This pass inserts:
  * NEW slide 5 - a poll placeholder stub (Nico splices his own
    PollEverywhere activity by hand);
  * NEW slide 6 - the worked solution (inelastic demand -> revenue up;
    computed from the article's ticket counts: +76%).
The Inside Out 2 pair moves to displays 7-8 (live slidenum fields
renumber themselves).

The deck is hand-made (NOT pipeline-built), so this is zip + lxml
surgery: the two slides are generated with the In-Class helper layer in
a temp deck and copied in verbatim, pointed at the practice deck's own
slideLayout7 (same generated layout family). Static (no <p:timing>) -
the deck is a parking lot; slides get builds when they graduate into a
real deck. Roll the _t-1 backup before running.

2026-09-20 (Nico), round 2: ticket counts rounded ("about 142,000 /
about 100,000", revenue $284,000 -> $500,000, still +76%), and the
caution box about the article's dollar figures moved OFF the slide into
its speaker notes.

Rerunnable: if the deck already contains "Solution: Lottery Revenue",
the pass REPLACES that slide (and its notes) with the regenerated
version instead of inserting; the poll stub is only created on the
first run.
"""
import copy
import os
import sys
import zipfile
import shutil
from pathlib import Path

from lxml import etree as ET

sys.stdout.reconfigure(encoding='utf-8')
HERE = Path(__file__).parent
DECK = HERE.parent / "Module 2 - Potential Practice Exercises.pptx"

if __name__ != "__main__":
    raise ImportError("_add_practice_lottery rewrites the practice deck "
                      "as soon as it runs; run it as a script.")

import _build_Module2InClass as M  # noqa: E402  (helper layer)
from pptx.util import Inches  # noqa: E402

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def q(ns, t):
    return "{%s}%s" % (ns, t)


TAG = "Demand Analysis · Elasticity and Revenue"     # matches slide 4
FOOTER = "Management 405  ·  Demand Analysis"        # matches slide 4

# --- build the two slides with the In-Class helper layer -------------------
M.FOOTER_TEXT = FOOTER
prs = M.Presentation()
prs.slide_width = int(M.SLIDE_W)
prs.slide_height = int(M.SLIDE_H)

M.make_stub(prs, 5, TAG, "Poll: Lottery Revenue",
            "PollEverywhere slide – insert your poll here")

slide = M._blank_slide(prs)
M._draw_top_bar_tc(slide, TAG)
M._draw_action_title(slide, "Solution: Lottery Revenue")
M._add_hierarchical_bullets(
    slide, M.MARGIN + Inches(0.15), Inches(2.05), Inches(12.4),
    Inches(3.7),
    [
        ([("Demand is ", {}),
          ("inelastic", {'bold': True, 'color': M.CBLUE}),
          (":  |Eᴅ| = 0.2 < 1   →   a higher price raises revenue",
           {})], 0, {}),
        ([("Check with the article's ticket counts:", {})], 0, {}),
        ([("Before:  R₀ ≈ 142,000 × $2 = $284,000", {})], 1, {}),
        ([("After:  R₁ ≈ 100,000 × $5 = $500,000", {})], 1, {}),
        ([("Yes – the article is right:  revenue rose by about 76%",
           {'bold': True, 'color': M.RED})], 0, {}),
    ],
    size=24, sub_size=22, line_spacing_pts=16)
M._set_notes(slide, (
    "With an elasticity of about −0.2, demand is well inside the "
    "inelastic range, so the 150% price increase outweighs the roughly "
    "30% drop in tickets and revenue rises. From the rounded ticket "
    "counts: 142,000 tickets at $2 is $284,000; 100,000 tickets at $5 "
    "is $500,000 - up about 76%. Multiplicative check: 2.50 × "
    "0.704 ≈ 1.76.\n\n"
    "Caution (kept off the slide): the article's own dollar figures "
    "($491,323 → $841,185, “+71%”) do not equal tickets "
    "× price (they imply $3.46 and $8.39 per ticket), so they must "
    "cover more than the base ticket sales counted here; the revenue "
    "check above uses the reported ticket counts (142,170 and 100,297, "
    "rounded), and the qualitative conclusion is unchanged. Source: "
    "WROR / BostonNewsroom, April 10, 2025, citing the Massachusetts "
    "Lottery (WBZ-TV)."))
M._draw_footer(slide, FOOTER, 6)

n_sub = M.apply_subscripts(prs)
print("subscript runs fixed:", n_sub)
tmp = HERE / "_practice_tmp.pptx"
prs.save(str(tmp))

# --- read both packages -----------------------------------------------------
def slide_order(data):
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target") for r in
             ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    return [rid2t[s.get(q(R, "id"))] for s in pres.find(q(P, "sldIdLst"))]


z = zipfile.ZipFile(tmp)
tdata = {n: z.read(n) for n in z.namelist()}
z.close()
torder = slide_order(tdata)          # [stub, solution]

z = zipfile.ZipFile(DECK)
data = {n: z.read(n) for n in z.namelist()}
z.close()

assert "ppt/notesMasters/notesMaster1.xml" in data, "no notes master"

# --- replace mode: the solution slide already exists -------------------------
existing = [n for n, blob in data.items()
            if n.startswith("ppt/slides/slide")
            and not n.endswith(".rels")
            and b"Solution: Lottery Revenue" in blob]
if existing:
    assert len(existing) == 1, existing
    part = existing[0]
    src_part = "ppt/slides/" + os.path.basename(torder[1])
    data[part] = tdata[src_part]
    # regenerate its notes too (rels stay: rId1 layout, rId2 notes)
    rels = ET.fromstring(data["ppt/slides/_rels/%s.rels"
                              % os.path.basename(part)])
    notes_tgt = [r.get("Target") for r in rels
                 if r.get("Type").endswith("/notesSlide")]
    assert notes_tgt, "existing solution slide has no notes rel"
    notes_part = "ppt/notesSlides/" + os.path.basename(notes_tgt[0])
    src_rels = ET.fromstring(
        tdata["ppt/slides/_rels/%s.rels" % os.path.basename(torder[1])])
    src_notes = [r.get("Target") for r in src_rels
                 if r.get("Type").endswith("/notesSlide")]
    data[notes_part] = tdata["ppt/notesSlides/"
                             + os.path.basename(src_notes[0])]
    out = DECK.with_suffix(".insert_tmp.pptx")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in data.items():
            zout.writestr(name, blob)
    shutil.move(str(out), str(DECK))
    tmp.unlink()
    print("REPLACED %s (+ %s) with the regenerated solution slide"
          % (part, notes_part))
    raise SystemExit(0)

order = slide_order(data)
# locate the anchor by on-slide text, not by position
anchor_idx = None
for i, tgt in enumerate(order):
    part = "ppt/slides/" + os.path.basename(tgt)
    if b"Lottery Sales, Revisited" in data[part]:
        anchor_idx = i
        break
assert anchor_idx is not None, "anchor slide (Lottery Sales, Revisited) " \
    "not found"
print("inserting after display", anchor_idx + 1)

# --- copy the two generated slides in ---------------------------------------
NEW = [("ppt/slides/slide7.xml", torder[0], None),
       ("ppt/slides/slide8.xml", torder[1], "ppt/notesSlides/notesSlide8.xml")]

for new_part, src_tgt, new_notes in NEW:
    src_part = "ppt/slides/" + os.path.basename(src_tgt)
    assert new_part not in data
    data[new_part] = tdata[src_part]
    rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.'
            'org/officeDocument/2006/relationships/slideLayout" '
            'Target="../slideLayouts/slideLayout7.xml"/>']
    if new_notes:
        # find the generated notes part for this slide
        src_rels = ET.fromstring(
            tdata["ppt/slides/_rels/%s.rels" % os.path.basename(src_tgt)])
        src_notes = [r.get("Target") for r in src_rels
                     if r.get("Type").endswith("/notesSlide")]
        assert src_notes, "generated solution slide lost its notes"
        src_notes_part = "ppt/notesSlides/" + os.path.basename(src_notes[0])
        data[new_notes] = tdata[src_notes_part]
        assert new_notes not in [n for n in data if n != new_notes] or True
        data["ppt/notesSlides/_rels/%s.rels" % os.path.basename(new_notes)] \
            = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<Relationships xmlns="%s">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats'
               '.org/officeDocument/2006/relationships/notesMaster" '
               'Target="../notesMasters/notesMaster1.xml"/>'
               '<Relationship Id="rId2" Type="http://schemas.openxmlformats'
               '.org/officeDocument/2006/relationships/slide" '
               'Target="../slides/%s"/>'
               '</Relationships>'
               % (PKG_R, os.path.basename(new_part))).encode()
        rels.append('<Relationship Id="rId2" Type="http://schemas.openxml'
                    'formats.org/officeDocument/2006/relationships/'
                    'notesSlide" Target="../notesSlides/%s"/>'
                    % os.path.basename(new_notes))
    data["ppt/slides/_rels/%s.rels" % os.path.basename(new_part)] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="%s">%s</Relationships>'
        % (PKG_R, "".join(rels))).encode()

# --- content types -----------------------------------------------------------
ct = ET.fromstring(data["[Content_Types].xml"])
SLIDE_CT = ("application/vnd.openxmlformats-officedocument.presentationml"
            ".slide+xml")
NOTES_CT = ("application/vnd.openxmlformats-officedocument.presentationml"
            ".notesSlide+xml")
for pn, ctype in (("/ppt/slides/slide7.xml", SLIDE_CT),
                  ("/ppt/slides/slide8.xml", SLIDE_CT),
                  ("/ppt/notesSlides/notesSlide8.xml", NOTES_CT)):
    o = ET.SubElement(ct, q(CT, "Override"))
    o.set("PartName", pn)
    o.set("ContentType", ctype)
data["[Content_Types].xml"] = ET.tostring(ct, xml_declaration=True,
                                          encoding="UTF-8", standalone=True)

# --- presentation rels + sldIdLst -------------------------------------------
prels = ET.fromstring(data["ppt/_rels/presentation.xml.rels"])
max_rid = max(int(r.get("Id")[3:]) for r in prels)
rid_a, rid_b = "rId%d" % (max_rid + 1), "rId%d" % (max_rid + 2)
for rid, tgt in ((rid_a, "slides/slide7.xml"), (rid_b, "slides/slide8.xml")):
    rel = ET.SubElement(prels, q(PKG_R, "Relationship"))
    rel.set("Id", rid)
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/"
                    "relationships/slide")
    rel.set("Target", tgt)
data["ppt/_rels/presentation.xml.rels"] = ET.tostring(
    prels, xml_declaration=True, encoding="UTF-8", standalone=True)

pres = ET.fromstring(data["ppt/presentation.xml"])
lst = pres.find(q(P, "sldIdLst"))
ids = [int(s.get("id")) for s in lst]
anchor_el = list(lst)[anchor_idx]
for k, rid in enumerate((rid_a, rid_b)):
    el = ET.Element(q(P, "sldId"))
    el.set("id", str(max(ids) + 1 + k))
    el.set(q(R, "id"), rid)
    anchor_el.addnext(el)
    anchor_el = el
data["ppt/presentation.xml"] = ET.tostring(
    pres, xml_declaration=True, encoding="UTF-8", standalone=True)

out = DECK.with_suffix(".insert_tmp.pptx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
    for name, blob in data.items():
        zout.writestr(name, blob)
shutil.move(str(out), str(DECK))
tmp.unlink()
print("saved %s - now %d slides" % (DECK.name, len(order) + 2))
