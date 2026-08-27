# -*- coding: utf-8 -*-
"""Split Module 3 - Revised.pptx into one .pptx per taped video (Nico,
2026-08-26).

Reads the boundaries from the video title cards ("Module 3  ·  Video k")
already in the deck, then writes seven files:

    Module 3 - Video 1 - Introduction.pptx
    Module 3 - Video 2 - Production Function.pptx
    Module 3 - Video 3 - Short Run Hiring.pptx
    Module 3 - Video 4 - Wage Searchers.pptx
    Module 3 - Video 5 - Long Run.pptx
    Module 3 - Video 6 - Cost Concepts.pptx
    Module 3 - Video 7 - Economies of Scale and Scope.pptx

Each output is a full copy of the source deck with the slides that don't
belong to that video (and their per-slide notes / rels) removed and the
sldIdLst / relationships / content-types trimmed to match. Shared parts
(masters, layouts, theme, media pool) stay put -- PowerPoint tolerates
unreferenced media, and leaving the media pool intact avoids risking a
slide losing a picture in a video where several images share one media
part.
"""
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

sys.stdout.reconfigure(encoding="utf-8")

HERE = Path(__file__).parent
SRC = HERE / "Module 3 - Revised.pptx"

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def q(ns, t):
    return "{%s}%s" % (ns, t)


VIDEO_TITLES = [
    "Introduction",
    "Production Function",
    "Short Run Hiring",
    "Wage Searchers",
    "Long Run",
    "Cost Concepts",
    "Economies of Scale and Scope",
]

VIDEO_CARD_RX = re.compile(r"Module 3  ·  Video (\d+)")


def slide_text(blob):
    return " ".join(t.text or "" for t in ET.fromstring(blob).iter(q(A, "t")))


def load_deck(path):
    z = zipfile.ZipFile(path)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    return data


def order_of(data):
    """Return the list of slide part basenames in display order."""
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target") for r in
             ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    return [os.path.basename(rid2t[s.get(q(R, "id"))])
            for s in pres.find(q(P, "sldIdLst"))]


def find_video_boundaries(data):
    """Return list of (video_number, start_display_index_1based) sorted by
    video number. The end is either the next start or len(slides)+1."""
    boundaries = []
    for i, base in enumerate(order_of(data), 1):
        text = slide_text(data["ppt/slides/" + base])
        m = VIDEO_CARD_RX.search(text)
        if m:
            boundaries.append((int(m.group(1)), i))
    return sorted(boundaries)


def parts_referenced_by(data, slide_part):
    """Given 'ppt/slides/slideN.xml', return the set of package-absolute
    part paths its rels reference (notesSlide, tags, etc.). Does NOT walk
    transitively -- notes rels bring in their own downstream via
    downstream() below."""
    rels_key = "ppt/slides/_rels/%s.rels" % os.path.basename(slide_part)
    out = set()
    if rels_key not in data:
        return out
    for rel in ET.fromstring(data[rels_key]):
        target = rel.get("Target")
        if target.startswith("http"):
            continue
        # rels targets are relative to the slide part's folder
        # (ppt/slides/), so join and normalise
        target_abs = os.path.normpath(os.path.join(
            "ppt/slides", target)).replace("\\", "/")
        out.add(target_abs)
    return out


def rels_of(data, part):
    """Return the rels part key for a given part, or None."""
    d, base = os.path.split(part)
    key = "%s/_rels/%s.rels" % (d, base)
    return key if key in data else None


def build_subset(data, keep_display):
    """Return a NEW dict-of-blobs holding only the slides whose 1-based
    display index is in *keep_display*. Shared parts (masters, layouts,
    theme, media, notesMaster) are copied through. Per-slide notes,
    slide-level tags, and per-slide rels for the dropped slides are
    removed, and presentation.xml / rels / [Content_Types].xml are
    trimmed accordingly."""
    out = dict(data)               # shallow copy of blob dict

    # rId -> target for the presentation part
    pres_rels = ET.fromstring(out["ppt/_rels/presentation.xml.rels"])
    rid2target = {r.get("Id"): r.get("Target") for r in pres_rels}
    slide_rids = [(r.get("Id"), os.path.basename(r.get("Target")))
                  for r in pres_rels
                  if r.get("Type").endswith("/relationships/slide")]

    all_slide_bases = order_of(out)
    keep_bases = {all_slide_bases[i - 1] for i in keep_display}
    drop_bases = [b for b in all_slide_bases if b not in keep_bases]

    # Collect parts to physically drop: the dropped slides + their _rels +
    # any notesSlide / slide-level tags reachable ONLY from dropped slides.
    drop_parts = set()
    for base in drop_bases:
        slide_part = "ppt/slides/" + base
        drop_parts.add(slide_part)
        rels_key = rels_of(out, slide_part)
        if rels_key:
            drop_parts.add(rels_key)
        for ref in parts_referenced_by(out, slide_part):
            # Only slide-private targets: notesSlides + tags. Keep media,
            # theme, layouts, embeddings, etc.
            if (ref.startswith("ppt/notesSlides/") or
                    ref.startswith("ppt/tags/")):
                drop_parts.add(ref)
                sub_rels = rels_of(out, ref)
                if sub_rels:
                    drop_parts.add(sub_rels)

    # A notesSlide / tag can be shared -- restore any part still referenced
    # by a KEPT slide.
    keep_refs = set()
    for base in keep_bases:
        keep_refs |= parts_referenced_by(out, "ppt/slides/" + base)
        # notesSlide rels can themselves reference tags -- fold those in
        for ref in list(keep_refs):
            if ref.startswith("ppt/notesSlides/"):
                sub_rels = rels_of(out, ref)
                if sub_rels:
                    for r in ET.fromstring(out[sub_rels]):
                        t = r.get("Target")
                        if t and not t.startswith("http"):
                            abs_t = os.path.normpath(os.path.join(
                                "ppt/notesSlides", t)).replace("\\", "/")
                            keep_refs.add(abs_t)
    drop_parts -= keep_refs
    # Never drop the shared rels of a kept part
    drop_parts = {p for p in drop_parts
                  if not (p.endswith(".rels") and
                          os.path.dirname(os.path.dirname(p)) + "/" +
                          os.path.basename(p)[:-5] in keep_refs)}

    for p in drop_parts:
        out.pop(p, None)

    # Rebuild presentation.xml sldIdLst: keep only sldId whose rId points
    # to a kept slide base.
    pres = ET.fromstring(out["ppt/presentation.xml"])
    lst = pres.find(q(P, "sldIdLst"))
    for sid in list(lst):
        rid = sid.get(q(R, "id"))
        target_base = os.path.basename(rid2target[rid])
        if target_base not in keep_bases:
            lst.remove(sid)
    out["ppt/presentation.xml"] = ET.tostring(
        pres, xml_declaration=True, encoding="UTF-8", standalone=True)

    # Trim presentation.xml.rels: drop rIds for removed slides.
    kept_rids = {sid.get(q(R, "id")) for sid in lst}
    for r in list(pres_rels):
        if r.get("Type").endswith("/relationships/slide"):
            if r.get("Id") not in kept_rids:
                pres_rels.remove(r)
    out["ppt/_rels/presentation.xml.rels"] = ET.tostring(
        pres_rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    # Trim [Content_Types].xml: drop Overrides for removed parts.
    ct = ET.fromstring(out["[Content_Types].xml"])
    for ov in list(ct):
        if ov.tag == q(CT, "Override"):
            pn = ov.get("PartName", "").lstrip("/")
            if pn in drop_parts:
                ct.remove(ov)
    out["[Content_Types].xml"] = ET.tostring(
        ct, xml_declaration=True, encoding="UTF-8", standalone=True)

    return out


def write_deck(data, path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, blob in data.items():
            z.writestr(name, blob)


def main():
    if not SRC.exists():
        raise SystemExit("source deck not found: %s" % SRC)
    data = load_deck(SRC)
    bounds = find_video_boundaries(data)
    n_slides = len(order_of(data))
    print("source: %s  (%d slides)" % (SRC.name, n_slides))
    if len(bounds) != 7:
        raise SystemExit("expected 7 video title cards, found %d: %s"
                         % (len(bounds), bounds))

    # video_number -> (start_display, end_display_inclusive)
    ranges = {}
    for k, (vnum, start) in enumerate(bounds):
        end = bounds[k + 1][1] - 1 if k + 1 < len(bounds) else n_slides
        ranges[vnum] = (start, end)

    for vnum in sorted(ranges):
        start, end = ranges[vnum]
        title = VIDEO_TITLES[vnum - 1]
        out_path = HERE / ("Module 3 - Video %d - %s.pptx"
                           % (vnum, title))
        keep = set(range(start, end + 1))
        subset = build_subset(data, keep)
        write_deck(subset, out_path)
        size_mb = out_path.stat().st_size / 1_048_576
        print("  Video %d  %-32s slides %2d-%2d  (%2d slides)  "
              "-> %s  [%.1f MB]"
              % (vnum, title, start, end, end - start + 1,
                 out_path.name, size_mb))


if __name__ == "__main__":
    main()
