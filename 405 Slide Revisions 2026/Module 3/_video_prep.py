# -*- coding: utf-8 -*-
"""Video conversion, step 1 of 2 — a TITLE SLIDE for every taped section
(Nico, 2026-08-26).

Module 3 is taped in full, so all seven outline sections become videos.
Each gets the deck-standard video title card, placed immediately BEFORE
its anchor - normally that section's agenda slide, which is the order
the Module 2 video deck uses (title card -> outline -> content). The
Introduction's card opens the deck: a video-mode deck starts by naming
the video the viewer is about to watch, so it sits ahead of the deck
title slide. Anchors and
video numbers come from VIDEO_ANCHOR / COVERAGE_LABEL in _m3_outline.py.

The card is the Module 2 layout, verbatim (_video_title_slide in
_build_Module2Video.py): section name navy 60 pt bold at y 2.10", the
"Module N  ·  Video k" line gold 40 pt bold at 3.25", a 4" gold strip,
the course line gray bold 26 pt, my name gray 22 pt, and the footer rule
with no footer text and no page number. No speaker notes — the video
decks' title cards carry none.

Step 2 (the coverage pills on the agenda slides) lives in
_retrofit_agenda.py; run this FIRST, then that, so the pills and the
cached page numbers are computed on the final slide order.

Rerunnable: a card already sitting at its anchor is RESTAMPED with the
current title and video number rather than duplicated, which is how a
renamed section or a renumbered video (the Introduction pushing the
teaching sections from Videos 1-6 to 2-7) is picked up.
"""
import copy
import os
import sys
import zipfile
import shutil
from pathlib import Path

from lxml import etree as ET

sys.stdout.reconfigure(encoding="utf-8")
HERE = Path(__file__).parent
DECK = HERE / "Module 3 - Revised.pptx"
M2_DIR = HERE.parent / "Module 2"
sys.path.insert(0, str(M2_DIR))

import _build_Module2InClass as M2  # noqa: E402

from pptx.enum.text import PP_ALIGN  # noqa: E402
from pptx.util import Inches  # noqa: E402

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

AGENDA_TITLE = "Outline of Module 3"
SUBTITLE = "Module 3  ·  Video %d"
CARD_MARK = "Module 3  ·  Video"        # any card, whatever its number

# outline item -> card title, video number and anchor, all from the
# shared outline so a renamed item renames its card too
from _m3_outline import M3_OUTLINE, COVERAGE_LABEL, VIDEO_ANCHOR  # noqa

VIDEOS = {i: (int(COVERAGE_LABEL[i].split()[-1]), M3_OUTLINE[i][0],
              VIDEO_ANCHOR[i])
          for i in sorted(VIDEO_ANCHOR)}


def q(ns, t):
    return "{%s}%s" % (ns, t)


def order_of(data):
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target") for r in
             ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    return [os.path.basename(rid2t[s.get(q(R, "id"))])
            for s in pres.find(q(P, "sldIdLst"))]


def slide_text(blob):
    return " ".join(t.text or "" for t in ET.fromstring(blob).iter(q(A, "t")))


def make_video_title(prs, main, video_line):
    """Copied from _video_title_slide in _build_Module2Video.py so the two
    decks' title cards stay identical to the EMU."""
    slide = M2._blank_slide(prs)
    M2._add_text(slide, 0, Inches(2.10), M2.SLIDE_W, Inches(1.1), main,
                 size=60, bold=True, color=M2.NAVY, font="Calibri",
                 align=PP_ALIGN.CENTER)
    M2._add_text(slide, 0, Inches(3.25), M2.SLIDE_W, Inches(0.75),
                 video_line, size=40, bold=True, color=M2.GOLD,
                 font="Calibri", align=PP_ALIGN.CENTER)
    M2._add_rect(slide, int((M2.SLIDE_W - Inches(4.0)) / 2), Inches(4.28),
                 Inches(4.0), 54864, M2.GOLD)
    M2._add_text(slide, 0, Inches(4.62), M2.SLIDE_W, Inches(0.55),
                 "Management 405", size=26, bold=True, color=M2.GRAY,
                 font="Calibri", align=PP_ALIGN.CENTER)
    M2._add_text(slide, 0, Inches(5.32), M2.SLIDE_W, Inches(0.5),
                 "Prof. Nico Voigtländer  ·  UCLA Anderson",
                 size=22, color=M2.GRAY, font="Calibri",
                 align=PP_ALIGN.CENTER)
    M2._add_rect(slide, 0, Inches(7.15), M2.SLIDE_W, Inches(0.02), M2.RULE)
    M2._add_rect(slide, M2.MARGIN, Inches(7.135), M2.GOLD_W, Inches(0.05),
                 M2.GOLD)
    return slide


def insert_slide(data, after_display, sld_xml):
    """Add a notes-free slide right after *after_display* (1-based, 0 =
    front of the deck)."""
    nums = [int(n.rsplit("slide", 1)[1].split(".")[0])
            for n in data if n.startswith("ppt/slides/slide")
            and n.endswith(".xml")]
    n = max(nums) + 1
    part = "ppt/slides/slide%d.xml" % n
    body = sld_xml.decode("utf-8")
    assert "r:id=" not in body and "r:embed=" not in body, \
        "generated card references a relationship — rels needed"
    data[part] = sld_xml

    root = ET.Element("{%s}Relationships" % PKG_R, nsmap={None: PKG_R})
    ET.SubElement(root, "{%s}Relationship" % PKG_R, Id="rId1",
                  Type="http://schemas.openxmlformats.org/officeDocument/"
                       "2006/relationships/slideLayout",
                  Target="../slideLayouts/slideLayout1.xml")
    data["ppt/slides/_rels/slide%d.xml.rels" % n] = ET.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)

    ct = ET.fromstring(data["[Content_Types].xml"])
    ET.SubElement(ct, "{%s}Override" % CT, PartName="/" + part,
                  ContentType="application/vnd.openxmlformats-officedocument"
                              ".presentationml.slide+xml")
    data["[Content_Types].xml"] = ET.tostring(
        ct, xml_declaration=True, encoding="UTF-8", standalone=True)

    prels = ET.fromstring(data["ppt/_rels/presentation.xml.rels"])
    rid = "rId%d" % (max(int(r.get("Id")[3:]) for r in prels) + 1)
    ET.SubElement(prels, "{%s}Relationship" % PKG_R, Id=rid,
                  Type="http://schemas.openxmlformats.org/officeDocument/"
                       "2006/relationships/slide",
                  Target="slides/slide%d.xml" % n)
    data["ppt/_rels/presentation.xml.rels"] = ET.tostring(
        prels, xml_declaration=True, encoding="UTF-8", standalone=True)

    pres = ET.fromstring(data["ppt/presentation.xml"])
    lst = pres.find(q(P, "sldIdLst"))
    el = ET.Element(q(P, "sldId"))
    el.set("id", str(max(int(s.get("id")) for s in lst) + 1))
    el.set(q(R, "id"), rid)
    lst.insert(after_display, el)
    data["ppt/presentation.xml"] = ET.tostring(
        pres, xml_declaration=True, encoding="UTF-8", standalone=True)
    return "slide%d.xml" % n


# --- generate the six cards ------------------------------------------------
prs = M2.Presentation()
prs.slide_width = int(M2.SLIDE_W)
prs.slide_height = int(M2.SLIDE_H)
for i in sorted(VIDEOS):
    vid, main, _ = VIDEOS[i]
    make_video_title(prs, main, SUBTITLE % vid)
tmp = HERE / "_vtitle_tmp.pptx"
prs.save(str(tmp))
z = zipfile.ZipFile(tmp)
tdata = {n: z.read(n) for n in z.namelist()}
z.close()
torder = order_of(tdata)
cards = {i: tdata["ppt/slides/" + torder[k]]
         for k, i in enumerate(sorted(VIDEOS))}
tmp.unlink()

# --- place each card before its section's agenda slide ---------------------
z = zipfile.ZipFile(DECK)
data = {n: z.read(n) for n in z.namelist()}
z.close()

agenda = [i + 1 for i, base in enumerate(order_of(data))
          if AGENDA_TITLE in slide_text(data["ppt/slides/" + base])]
print("agenda slides at %s  (index 0 = the module overview)" % agenda)


def anchor_display(kind, ref):
    if kind == "top":
        return 1
    if kind == "agenda":
        if ref >= len(agenda):
            raise SystemExit("no agenda slide with index %d" % ref)
        return agenda[ref]
    hits = [i + 1 for i, base in enumerate(order_of(data))
            if ref in slide_text(data["ppt/slides/" + base])]
    if len(hits) != 1:
        raise SystemExit("anchor %r matches %d slides" % (ref, len(hits)))
    return hits[0]


targets = sorted(((anchor_display(*VIDEOS[i][2]), i) for i in VIDEOS),
                 reverse=True)                 # back to front, so earlier
added = swapped = 0                            # positions stay valid
for at, i in targets:
    vid, main, anchor = VIDEOS[i]
    # a "top" card IS the first slide; every other card sits just before
    # its anchor
    if anchor[0] == "top":
        prev, pos, where = order_of(data)[0], 0, "at the top of the deck"
    else:
        prev = order_of(data)[at - 2] if at > 1 else None
        pos, where = at - 1, "before its anchor"
    if prev is not None and CARD_MARK in slide_text(data["ppt/slides/" + prev]):
        # a card is already there — restamp it, which is how a renumbered
        # or renamed video gets picked up
        part = "ppt/slides/" + prev
        tree = ET.fromstring(data[part])
        csld = tree.find(q(P, "cSld"))
        card = ET.fromstring(cards[i])
        csld.replace(csld.find(q(P, "spTree")),
                     card.find(q(P, "cSld") + "/" + q(P, "spTree")))
        data[part] = ET.tostring(tree, xml_declaration=True,
                                 encoding="UTF-8", standalone=True)
        print("  Video %d: %-32s restamped at display %d"
              % (vid, main, pos + 1))
        swapped += 1
        continue
    insert_slide(data, pos, cards[i])
    print("  Video %d: %-32s inserted at display %d (%s)"
          % (vid, main, pos + 1, where))
    added += 1

if not added and not swapped:
    print("nothing to do — every title card is already in place")
    raise SystemExit(0)

out = DECK.with_suffix(".video_tmp.pptx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
    for name, blob in data.items():
        zout.writestr(name, blob)
shutil.move(str(out), str(DECK))
print("saved %s (%d slides)" % (DECK.name, len(order_of(data))))
