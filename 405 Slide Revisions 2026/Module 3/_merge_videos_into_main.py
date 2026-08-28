# -*- coding: utf-8 -*-
"""Merge the hand-revised video decks back into `Module 3 - Revised.pptx`.

2026-08-27, at Nico's request.  For every slide of Videos 1 - 7 that has a
counterpart in the main deck, the video's version REPLACES the main deck's
(so his hand edits come across in full: text, pictures, grouping and the
`<p:timing>` animation build).  Video slides with no counterpart are
inserted at the position they occupy in the video.  Slides that were cut
from the videos stay where they are, and a copy of each is appended under
a "SLIDES NOT USED IN THE VIDEOS" divider.

Decisions Nico made when this was planned (2026-08-27):
  * the AI-researcher wage-searcher pair is KEPT as well as the new Rivian
    designer pair, so both examples run in the in-class deck;
  * PollEverywhere slides are NOT copied into the appendix (a duplicate
    __PE_POLL_EMBED_ID would have two slides claiming one poll);
  * the appendix carries the cut TEACHING slides only, not the class-only
    front matter (Zoom logistics, announcements, recap) or the summary
    closer.

Two exceptions to "adopt everything", both his:
  * the BACKUP block stays at the end of the main deck, and the two
    "Very high / Very low MPL image" links are re-pointed at the main
    deck's backup slides, not the video's copies;
  * the production-function-table links are re-attached after adoption
    (the videos have them stripped, since they would be dead there).

Everything is done with direct zip + lxml surgery.  The deck is never
round-tripped through python-pptx, which would drop the NULL video
relationship on the Amazon slide and the PollEverywhere tag parts.

Usage:  python _merge_videos_into_main.py [--dry-run]
"""
import copy
import os
import shutil
import sys
import zipfile

from lxml import etree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)

MAIN = "Module 3 - Revised.pptx"

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

RT_SLIDE = R + "/slide"
RT_LAYOUT = R + "/slideLayout"
RT_NOTES = R + "/notesSlide"
RT_NOTESMASTER = R + "/notesMaster"


def q(ns, t):
    return "{%s}%s" % (ns, t)


VIDEOS = {
    "V1": "Module 3 - Video 1 - Introduction.pptx",
    "V2": "Module 3 - Video 2 - Production Function.pptx",
    "V3": "Module 3 - Video 3 - Short Run Hiring.pptx",
    "V4": "Module 3 - Video 4 - Wage Searchers.pptx",
    "V5": "Module 3 - Video 5 - Long Run.pptx",
    "V6": "Module 3 - Video 6 - Cost Concepts.pptx",
    "V7": "Module 3 - Video 7 - Economies of Scale and Scope.pptx",
}

# (video, video slide) -> main display index in the ORIGINAL 88-slide deck.
REPLACE = [
    ("V1", 3, 6), ("V1", 4, 7), ("V1", 5, 8),
    ("V2", 2, 11), ("V2", 3, 12), ("V2", 4, 13), ("V2", 5, 14),
    ("V2", 6, 15),
    ("V3", 2, 17), ("V3", 3, 18), ("V3", 4, 19), ("V3", 5, 20),
    ("V3", 7, 21), ("V3", 8, 22), ("V3", 9, 23), ("V3", 10, 24),
    ("V3", 11, 26), ("V3", 12, 27), ("V3", 13, 28),
    ("V4", 2, 30), ("V4", 3, 31), ("V4", 4, 32),
    ("V5", 2, 38), ("V5", 3, 39), ("V5", 4, 40), ("V5", 5, 41),
    ("V5", 6, 42), ("V5", 7, 43), ("V5", 9, 45), ("V5", 10, 46),
    ("V6", 2, 49), ("V6", 3, 50), ("V6", 4, 51), ("V6", 5, 54),
    ("V6", 6, 55), ("V6", 8, 56), ("V6", 9, 62), ("V6", 10, 63),
    ("V6", 11, 64), ("V6", 12, 67), ("V6", 13, 68), ("V6", 14, 69),
    ("V7", 2, 71), ("V7", 3, 72), ("V7", 4, 73), ("V7", 5, 74),
    ("V7", 6, 75), ("V7", 7, 76), ("V7", 8, 77), ("V7", 9, 78),
]

# (video, video slide, "after" main display index) — new in the videos.
INSERT = [
    ("V3", 6, 20),      # "Do You See Diminishing MPL... and MPK?"
    ("V4", 5, 32),      # "Example: The Full Cost of a New Rivian Designer"
    ("V4", 6, 32),      # "Solution: Marginal Cost of the 3rd Designer = $3M"
    ("V6", 7, 55),      # "Accounting for Opportunity Costs"
]

# Cut from the videos; a copy of each goes in the closing appendix, in
# main-deck order.  Poll slides and the class-only front matter are out by
# Nico's decision; the backup block is its own thing and stays put.
APPENDIX = [9, 33, 35, 36, 47, 52, 53, 57, 58, 60, 61, 65, 66, 79, 80,
            81, 84]

# Hand-edited by Nico in PowerPoint on 2026-08-27: the divider gained a
# third line naming what the appendix is FOR.  Three paragraphs, 90 pt.
APPENDIX_TITLE = ("SLIDES NOT USED\n"
                  "IN THE VIDEOS\n"
                  "-- FOR In-Class APPLICATIONS")
DIVIDER_SRC = 86          # the BACKUP divider, cloned for the appendix


# ==========================================================================
# package helpers
# ==========================================================================

class Pkg(object):
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.parts = {n: z.read(n) for n in z.namelist()}
        self.order = list(z.namelist())
        z.close()

    def xml(self, name):
        return ET.fromstring(self.parts[name])

    def set_xml(self, name, tree):
        self.parts[name] = ET.tostring(tree, xml_declaration=True,
                                       encoding="UTF-8", standalone=True)

    def rels_name(self, part):
        d, b = part.rsplit("/", 1)
        return "%s/_rels/%s.rels" % (d, b)

    def rels(self, part):
        rn = self.rels_name(part)
        if rn not in self.parts:
            return None
        return self.xml(rn)

    def slide_order(self):
        pres = self.xml("ppt/presentation.xml")
        prels = self.xml("ppt/_rels/presentation.xml.rels")
        rid2t = {r.get("Id"): r.get("Target") for r in prels}
        out = []
        for s in pres.find(q(P, "sldIdLst")):
            out.append("ppt/slides/"
                       + rid2t[s.get(q(R, "id"))].split("/")[-1])
        return out


def ct_map(pkg):
    """part name -> content type, from Overrides plus extension Defaults."""
    ct = pkg.xml("[Content_Types].xml")
    over, default = {}, {}
    for el in ct:
        tag = ET.QName(el).localname
        if tag == "Override":
            over[el.get("PartName").lstrip("/")] = el.get("ContentType")
        elif tag == "Default":
            default[el.get("Extension").lower()] = el.get("ContentType")
    return over, default


def add_ct_override(pkg, part, content_type):
    ct = pkg.xml("[Content_Types].xml")
    for el in ct:
        if (ET.QName(el).localname == "Override"
                and el.get("PartName") == "/" + part):
            return
    el = ET.SubElement(ct, q(CT, "Override"))
    el.set("PartName", "/" + part)
    el.set("ContentType", content_type)
    pkg.set_xml("[Content_Types].xml", ct)


def add_ct_default(pkg, ext, content_type):
    ct = pkg.xml("[Content_Types].xml")
    for el in ct:
        if (ET.QName(el).localname == "Default"
                and (el.get("Extension") or "").lower() == ext.lower()):
            return
    el = ET.Element(q(CT, "Default"))
    el.set("Extension", ext)
    el.set("ContentType", content_type)
    ct.insert(0, el)
    pkg.set_xml("[Content_Types].xml", ct)


def free_name(pkg, folder, stem, ext):
    i = 1
    while True:
        cand = "%s/%s%d.%s" % (folder, stem, i, ext)
        if cand not in pkg.parts:
            return cand
        i += 1


def new_rels_doc():
    return ET.Element(q(RELS, "Relationships"))


def put_rel(rels_el, rid, rtype, target, mode=None):
    el = ET.SubElement(rels_el, q(RELS, "Relationship"))
    el.set("Id", rid)
    el.set("Type", rtype)
    el.set("Target", target)
    if mode:
        el.set("TargetMode", mode)
    return el


# ==========================================================================
# importing one slide from a video package
# ==========================================================================

def import_media(dst, src, src_target, src_part_dir, cache,
                 src_over, src_default, dst_default):
    """Copy a non-slide dependency (image / chart / xlsx / tags) across."""
    src_part = os.path.normpath(
        os.path.join(src_part_dir, src_target)).replace("\\", "/")
    if src_part in cache:
        return cache[src_part]
    blob = src.parts[src_part]
    folder, base = src_part.rsplit("/", 1)
    stem = "".join(c for c in base.rsplit(".", 1)[0] if not c.isdigit())
    ext = base.rsplit(".", 1)[1]
    new_part = free_name(dst, folder, stem or "part", ext)
    dst.parts[new_part] = blob
    cache[src_part] = new_part

    if src_part in src_over:
        add_ct_override(dst, new_part, src_over[src_part])
    elif ext.lower() in src_default and ext.lower() not in dst_default:
        add_ct_default(dst, ext, src_default[ext.lower()])

    # A chart carries its own rels (the embedded workbook); recurse.
    src_rels = src.rels(src_part)
    if src_rels is not None:
        new_rels = new_rels_doc()
        for rel in src_rels:
            rtype, tgt = rel.get("Type"), rel.get("Target")
            if rel.get("TargetMode") == "External":
                put_rel(new_rels, rel.get("Id"), rtype, tgt, "External")
                continue
            sub = import_media(dst, src, tgt, folder, cache,
                               src_over, src_default, dst_default)
            put_rel(new_rels, rel.get("Id"), rtype,
                    os.path.relpath(sub, folder).replace("\\", "/"))
        dst.set_xml(dst.rels_name(new_part), new_rels)
    return new_part


def import_slide(dst, src, src_slide, dst_slide, cache,
                 src_over, src_default, dst_default, slide_link_map):
    """Put the video's slide XML + dependency closure into `dst_slide`.

    `dst_slide` must already exist in dst (replacement) or be a fresh part
    name (insertion).  Returns nothing; dst is mutated.
    """
    dst.parts[dst_slide] = src.parts[src_slide]

    src_rels = src.rels(src_slide)
    new_rels = new_rels_doc()
    notes_part = None
    for rel in src_rels:
        rid, rtype, tgt = rel.get("Id"), rel.get("Type"), rel.get("Target")
        if rel.get("TargetMode") == "External":
            put_rel(new_rels, rid, rtype, tgt, "External")
            continue
        if rtype == RT_LAYOUT:
            put_rel(new_rels, rid, rtype, "../slideLayouts/slideLayout1.xml")
            continue
        if rtype == RT_SLIDE:
            # a slide-jump hyperlink: remap to the main deck's own slide
            vpart = tgt.split("/")[-1]
            mapped = slide_link_map.get(vpart)
            if mapped is None:
                raise KeyError("no link remap for %s (from %s)"
                               % (vpart, src_slide))
            put_rel(new_rels, rid, rtype, "../slides/" + mapped)
            continue
        if rtype == RT_NOTES:
            notes_part = (rid, tgt)
            continue
        sub = import_media(dst, src, tgt, "ppt/slides", cache,
                           src_over, src_default, dst_default)
        put_rel(new_rels, rid, rtype,
                os.path.relpath(sub, "ppt/slides").replace("\\", "/"))

    # --- notes: copy the part and re-point its own rels
    if notes_part is not None:
        rid, tgt = notes_part
        sp = os.path.normpath(
            os.path.join("ppt/slides", tgt)).replace("\\", "/")
        new_notes = free_name(dst, "ppt/notesSlides", "notesSlide", "xml")
        dst.parts[new_notes] = src.parts[sp]
        if sp in src_over:
            add_ct_override(dst, new_notes, src_over[sp])
        nrels = src.rels(sp)
        nnew = new_rels_doc()
        if nrels is not None:
            for r2 in nrels:
                t2, g2 = r2.get("Type"), r2.get("Target")
                if t2 == RT_SLIDE:
                    put_rel(nnew, r2.get("Id"), t2,
                            "../slides/" + dst_slide.split("/")[-1])
                elif t2 == RT_NOTESMASTER:
                    put_rel(nnew, r2.get("Id"), t2,
                            "../notesMasters/notesMaster1.xml")
                elif r2.get("TargetMode") == "External":
                    put_rel(nnew, r2.get("Id"), t2, g2, "External")
                else:
                    sub = import_media(dst, src, g2, "ppt/notesSlides",
                                       cache, src_over, src_default,
                                       dst_default)
                    put_rel(nnew, r2.get("Id"), t2,
                            os.path.relpath(sub, "ppt/notesSlides")
                            .replace("\\", "/"))
        dst.set_xml(dst.rels_name(new_notes), nnew)
        put_rel(new_rels, rid, RT_NOTES,
                "../notesSlides/" + new_notes.split("/")[-1])

    dst.set_xml(dst.rels_name(dst_slide), new_rels)


# ==========================================================================
# duplicating a slide that is already in the main deck (the appendix)
# ==========================================================================

def clone_slide(pkg, src_slide, dup_charts=True):
    """Copy a slide inside the same package.  Media is shared (that is
    normal and keeps the 11 MB video from being duplicated); chart parts
    are duplicated so two frames never own one chart."""
    new_slide = free_name(pkg, "ppt/slides", "slide", "xml")
    pkg.parts[new_slide] = pkg.parts[src_slide]
    add_ct_override(pkg, new_slide,
                    "application/vnd.openxmlformats-officedocument."
                    "presentationml.slide+xml")

    old = pkg.rels(src_slide)
    new = new_rels_doc()
    for rel in old:
        rid, rtype, tgt = rel.get("Id"), rel.get("Type"), rel.get("Target")
        mode = rel.get("TargetMode")
        if mode == "External":
            put_rel(new, rid, rtype, tgt, "External")
            continue
        if rtype == RT_NOTES:
            sp = os.path.normpath(
                os.path.join("ppt/slides", tgt)).replace("\\", "/")
            nn = free_name(pkg, "ppt/notesSlides", "notesSlide", "xml")
            pkg.parts[nn] = pkg.parts[sp]
            add_ct_override(pkg, nn,
                            "application/vnd.openxmlformats-officedocument."
                            "presentationml.notesSlide+xml")
            nrels = pkg.rels(sp)
            nnew = new_rels_doc()
            if nrels is not None:
                for r2 in nrels:
                    if r2.get("Type") == RT_SLIDE:
                        put_rel(nnew, r2.get("Id"), r2.get("Type"),
                                "../slides/" + new_slide.split("/")[-1])
                    else:
                        put_rel(nnew, r2.get("Id"), r2.get("Type"),
                                r2.get("Target"), r2.get("TargetMode"))
            pkg.set_xml(pkg.rels_name(nn), nnew)
            put_rel(new, rid, rtype,
                    "../notesSlides/" + nn.split("/")[-1])
            continue
        if dup_charts and "/charts/" in tgt:
            sp = os.path.normpath(
                os.path.join("ppt/slides", tgt)).replace("\\", "/")
            nc = free_name(pkg, "ppt/charts", "chart", "xml")
            pkg.parts[nc] = pkg.parts[sp]
            over, _ = ct_map(pkg)
            add_ct_override(pkg, nc, over.get(
                sp, "application/vnd.openxmlformats-officedocument."
                    "drawingml.chart+xml"))
            crels = pkg.rels(sp)
            if crels is not None:
                cnew = new_rels_doc()
                for r2 in crels:
                    put_rel(cnew, r2.get("Id"), r2.get("Type"),
                            r2.get("Target"), r2.get("TargetMode"))
                pkg.set_xml(pkg.rels_name(nc), cnew)
            put_rel(new, rid, rtype, "../charts/" + nc.split("/")[-1])
            continue
        put_rel(new, rid, rtype, tgt, mode)
    pkg.set_xml(pkg.rels_name(new_slide), new)
    return new_slide


# ==========================================================================
# re-attaching the production-function-table links
# ==========================================================================

# Nico kept these "back to the table for revision" links in the main deck.
# The videos have them stripped (they would be dead across decks), so they
# are put back after adoption.  V3.10 still carries its "link" run; on the
# other two he deleted the "(link)" wording by hand, so the link goes onto
# the caption run that IS there and that run is underlined, the deck's own
# cue for a jump link.  No wording is changed.
TABLE_LINKS = [
    (24, "link"),
    (26, "From the production-function table"),
    (43, "Production-function table"),
]


def attach_table_link(pkg, slide_part, run_text, target_part):
    tree = pkg.xml(slide_part)
    rels = pkg.rels(slide_part)
    used = {r.get("Id") for r in rels}
    n = 1
    while ("rId%d" % n) in used:
        n += 1
    rid = "rId%d" % n

    hit = None
    for r_el in tree.iter(q(A, "r")):
        t_el = r_el.find(q(A, "t"))
        if t_el is not None and (t_el.text or "").strip() == run_text:
            hit = r_el
            break
    if hit is None:
        return False

    rPr = hit.find(q(A, "rPr"))
    if rPr is None:
        rPr = ET.Element(q(A, "rPr"))
        hit.insert(0, rPr)
    rPr.set("u", "sng")
    for old in rPr.findall(q(A, "hlinkClick")):
        rPr.remove(old)
    hl = ET.SubElement(rPr, q(A, "hlinkClick"))
    hl.set(q(R, "id"), rid)
    hl.set("action", "ppaction://hlinksldjump")

    put_rel(rels, rid, RT_SLIDE,
            "../slides/" + target_part.split("/")[-1])
    pkg.set_xml(pkg.rels_name(slide_part), rels)
    pkg.set_xml(slide_part, tree)
    return True


# ==========================================================================
# slide-list surgery
# ==========================================================================

def register_slide(pkg, part):
    """Give a slide part a presentation-level rId, return it."""
    prels = pkg.xml("ppt/_rels/presentation.xml.rels")
    used = {r.get("Id") for r in prels}
    for r in prels:
        if r.get("Target").split("/")[-1] == part.split("/")[-1]:
            return r.get("Id")
    n = 1
    while ("rId%d" % n) in used:
        n += 1
    rid = "rId%d" % n
    put_rel(prels, rid, RT_SLIDE, "slides/" + part.split("/")[-1])
    pkg.set_xml("ppt/_rels/presentation.xml.rels", prels)
    return rid


def rebuild_sldIdLst(pkg, parts_in_order):
    pres = pkg.xml("ppt/presentation.xml")
    lst = pres.find(q(P, "sldIdLst"))
    old = {}
    prels = pkg.xml("ppt/_rels/presentation.xml.rels")
    rid2t = {r.get("Id"): r.get("Target") for r in prels}
    for s in lst:
        old[rid2t[s.get(q(R, "id"))].split("/")[-1]] = s.get("id")
    for s in list(lst):
        lst.remove(s)
    nid = 256
    used = set()
    for part in parts_in_order:
        base = part.split("/")[-1]
        rid = register_slide(pkg, part)
        sid = old.get(base)
        if sid is None or sid in used:
            while str(nid) in used:
                nid += 1
            sid = str(nid)
        used.add(sid)
        el = ET.SubElement(lst, q(P, "sldId"))
        el.set("id", sid)
        el.set(q(R, "id"), rid)
    pkg.set_xml("ppt/presentation.xml", pres)


def prune_presentation_rels(pkg, parts_in_order):
    """Drop slide rels that no longer appear in the slide list."""
    keep = {p.split("/")[-1] for p in parts_in_order}
    prels = pkg.xml("ppt/_rels/presentation.xml.rels")
    for r in list(prels):
        if r.get("Type") == RT_SLIDE:
            if r.get("Target").split("/")[-1] not in keep:
                prels.remove(r)
    pkg.set_xml("ppt/_rels/presentation.xml.rels", prels)


# ==========================================================================
# the appendix divider
# ==========================================================================

# "BACKUP" fits on one line at the divider's 140 pt; the appendix title
# does not (26.5" against a 12.78" box), so it is set on two lines at
# 90 pt, where the longer line measures 9.09".
APPENDIX_LINES = ["SLIDES NOT USED", "IN THE VIDEOS"]
APPENDIX_PT = 90


def set_divider_title(pkg, part):
    tree = pkg.xml(part)
    for sp in tree.iter(q(P, "sp")):
        txt = "".join(x.text or "" for x in sp.iter(q(A, "t")))
        if txt.strip() != "BACKUP":
            continue
        txBody = sp.find(q(P, "txBody"))
        template = txBody.find(q(A, "p"))
        rPr = template.find(q(A, "r") + "/" + q(A, "rPr"))
        for old_p in txBody.findall(q(A, "p")):
            txBody.remove(old_p)
        for line in APPENDIX_LINES:
            p_el = ET.SubElement(txBody, q(A, "p"))
            pPr = ET.SubElement(p_el, q(A, "pPr"))
            pPr.set("algn", "ctr")
            r_el = ET.SubElement(p_el, q(A, "r"))
            new_rPr = copy.deepcopy(rPr)
            new_rPr.set("sz", str(APPENDIX_PT * 100))
            r_el.append(new_rPr)
            t_el = ET.SubElement(r_el, q(A, "t"))
            t_el.text = line
        # the box is spAutoFit; give it a sane cached height for 2 lines
        ext = sp.find(".//" + q(A, "ext"))
        ext.set("cy", str(int(2.62 * 914400)))
        break
    pkg.set_xml(part, tree)


# ==========================================================================
# orphan sweep
# ==========================================================================

def prune_orphans(pkg):
    """Drop parts no longer reachable from the package root.

    Replacing a slide leaves its old pictures, charts, workbooks and notes
    behind; without this the merged deck carries both the main deck's old
    media and the videos' copies.
    """
    reachable = set()

    def visit(part):
        if part in reachable or part not in pkg.parts:
            return
        reachable.add(part)
        rn = pkg.rels_name(part)
        if rn not in pkg.parts:
            return
        reachable.add(rn)
        base = part.rsplit("/", 1)[0]
        for rel in pkg.xml(rn):
            if rel.get("TargetMode") == "External":
                continue
            tgt = os.path.normpath(
                os.path.join(base, rel.get("Target"))).replace("\\", "/")
            visit(tgt)

    reachable.add("[Content_Types].xml")
    reachable.add("_rels/.rels")
    for rel in pkg.xml("_rels/.rels"):
        if rel.get("TargetMode") == "External":
            continue
        visit(os.path.normpath(rel.get("Target")).replace("\\", "/"))

    dead = [n for n in pkg.parts
            if n not in reachable and not n.startswith("docProps")]
    freed = sum(len(pkg.parts[n]) for n in dead)
    for n in dead:
        del pkg.parts[n]

    # drop the Content_Types overrides for the parts that just went
    ct = pkg.xml("[Content_Types].xml")
    for el in list(ct):
        if ET.QName(el).localname == "Override":
            if el.get("PartName").lstrip("/") not in pkg.parts:
                ct.remove(el)
    pkg.set_xml("[Content_Types].xml", ct)
    print("pruned %d orphaned parts (%.1f MB)" % (len(dead), freed / 1e6))
    return dead


# ==========================================================================
# main
# ==========================================================================

def main(dry=False):
    dst = Pkg(MAIN)
    order = dst.slide_order()
    assert len(order) == 88, "expected the 88-slide main deck, got %d" % len(order)
    dst_over, dst_default = ct_map(dst)

    srcs, src_orders, src_cts = {}, {}, {}
    for tag, path in VIDEOS.items():
        srcs[tag] = Pkg(path)
        src_orders[tag] = srcs[tag].slide_order()
        src_cts[tag] = ct_map(srcs[tag])

    # --- link remaps: video slide part -> main deck slide part ------------
    # Nico's exception 1: the backup links point at the MAIN deck's backup
    # slides (display 87 / 88), never at the video's own copies.
    link_map = {
        "V3": {
            src_orders["V3"][14].split("/")[-1]: order[86].split("/")[-1],
            src_orders["V3"][15].split("/")[-1]: order[87].split("/")[-1],
            src_orders["V3"][4].split("/")[-1]: order[19].split("/")[-1],
        },
    }

    cache = {tag: {} for tag in VIDEOS}
    new_parts = {}

    # --- replacements -----------------------------------------------------
    for tag, vi, mi in REPLACE:
        src = srcs[tag]
        s_over, s_default = src_cts[tag]
        import_slide(dst, src, src_orders[tag][vi - 1], order[mi - 1],
                     cache[tag], s_over, s_default, dst_default,
                     link_map.get(tag, {}))
    print("replaced %d slides from the videos" % len(REPLACE))

    # --- put the production-table links back (Nico's exception 2) --------
    table_part = order[13]          # main display 14, the table slide
    for mi, run_text in TABLE_LINKS:
        ok = attach_table_link(dst, order[mi - 1], run_text, table_part)
        print("   table link on main %-3d (%r): %s"
              % (mi, run_text[:40], "re-attached" if ok else "RUN NOT FOUND"))

    # --- insertions -------------------------------------------------------
    inserted = {}
    for tag, vi, after in INSERT:
        src = srcs[tag]
        s_over, s_default = src_cts[tag]
        part = free_name(dst, "ppt/slides", "slide", "xml")
        add_ct_override(dst, part,
                        "application/vnd.openxmlformats-officedocument."
                        "presentationml.slide+xml")
        import_slide(dst, src, src_orders[tag][vi - 1], part,
                     cache[tag], s_over, s_default, dst_default,
                     link_map.get(tag, {}))
        inserted.setdefault(after, []).append(part)
        new_parts["%s.%d" % (tag, vi)] = part
    print("inserted %d new slides" % len(INSERT))

    # --- appendix ---------------------------------------------------------
    divider = clone_slide(dst, order[DIVIDER_SRC - 1], dup_charts=False)
    set_divider_title(dst, divider)

    appendix_parts = [divider]
    for mi in APPENDIX:
        appendix_parts.append(clone_slide(dst, order[mi - 1]))
    print("appendix: divider + %d copies of cut slides" % len(APPENDIX))

    # --- assemble the new slide order ------------------------------------
    final = []
    for i, part in enumerate(order, 1):
        final.append(part)
        for extra in inserted.get(i, []):
            final.append(extra)
    final.extend(appendix_parts)
    rebuild_sldIdLst(dst, final)
    prune_presentation_rels(dst, final)
    print("slide count %d -> %d" % (len(order), len(final)))
    prune_orphans(dst)

    if dry:
        print("(dry run - nothing written)")
        return final

    # --- roll backups, then write ----------------------------------------
    base = MAIN[:-5]
    if os.path.exists(base + "_t-1.pptx"):
        shutil.copy2(base + "_t-1.pptx", base + "_t-2.pptx")
    shutil.copy2(MAIN, base + "_t-1.pptx")

    tmp = MAIN + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, blob in dst.parts.items():
            z.writestr(name, blob)
    os.replace(tmp, MAIN)
    print("written:", MAIN)
    return final


if __name__ == "__main__":
    main(dry="--dry-run" in sys.argv)
