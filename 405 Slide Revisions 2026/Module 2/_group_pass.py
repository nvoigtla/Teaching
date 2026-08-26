# -*- coding: utf-8 -*-
"""Grouping pass for "Module 2 - In Class Revised.pptx" — merges shape
pairs that belong together into <p:grpSp> groups (Teaching CLAUDE.md):

  1. box+text callouts   — a filled roundRect with NO text of its own +
                           the text box layered on top (convention /
                           caution / callout pattern);
  2. figure shades       — a white backing rect carrying an outerShdw +
                           the graphicFrame (chart/table) on top of it;
  3. picture+caption     — a picture + the small all-italic (≤16 pt)
                           text box sitting directly beneath it (multi-
                           picture captions group all pictures they span).

Detection is geometric, never by shape name. Spliced slides (polls +
Excel embed) are never touched. Groups get off/ext = the members'
bounding box with chOff/chExt equal, so children keep absolute
positions. GROUPING INVALIDATES ANIMATIONS — run BEFORE _animate.py:

  _build_Module2InClass.py -> _splice_media.py -> _group_pass.py
                           -> _animate.py all apply
"""
import os
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Module 2 - In Class Revised.pptx"
_args = [a for a in sys.argv[1:]]
for _a in _args:
    if _a.endswith(".pptx"):
        DECK = Path(_a) if Path(_a).is_absolute() else HERE / _a
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0

# Slides spliced in verbatim from the old deck: the group pass leaves
# them alone.  Keyed by deck (2026-08-24) - the In-Class list used to be
# the module-wide default, which silently skipped VIDEO slide 4 and cost
# it its manual groups.
SPLICED_BY_DECK = {
    "Module 2 - In Class Revised": {4, 5, 11, 12, 13, 32, 36, 45, 46,
                                    57, 58, 65, 66},
    # 2026-08-25: 4, 5, 6 and 22 now come over verbatim from the
    # final video decks, groups and all
    "Module 2 - Video Part Revised": {4, 5, 6, 19, 21},
}

# --------------------------------------------------------------------------
# Explicit groups ported from Nico's hand-edits (2026-08-23).  The
# geometric rules above only know callout/shade/caption patterns; these
# are the pairings he made by hand in PowerPoint.  Members are matched by
# their rendered (x, y) in INCHES to 0.01" among the top-level shapes of
# that slide, then grouped in document order.
#
# KEYED BY DECK, because this pass is shared by the In-Class and the
# Video decks and the display numbers collide: In-Class 9/18/19/20/21 are
# entirely different slides from Video 9/18/19/20/21.  Without the deck
# key the video build silently picks up the In-Class groupings.
# --------------------------------------------------------------------------
_M2_INCLASS_GROUPS = {
    9: [   # the whole demand-curve mini figure = one object
        [(9.350, 2.570), (9.350, 6.050), (8.979, 2.285),
         (11.800, 6.150), (9.590, 3.080), (12.155, 5.357)],
    ],
    18: [  # MB = MC anchor star + its label (star moved 2026-08-25)
        [(11.410, 5.280), (11.710, 5.660)],
    ],
    20: [  # aggregation: each "=" with the aggregate dot it produces,
           # all four "+" signs as one beat, aggregate legend swatch+label
        [(3.790, 2.856), (4.291, 2.961)],
        [(5.038, 3.714), (5.955, 3.819)],
        [(6.286, 4.573), (7.619, 4.678)],
        [(2.958, 2.856), (3.790, 3.714), (4.622, 4.573), (5.454, 5.431)],
        [(7.534, 5.431), (9.283, 5.536)],
        [(9.550, 2.870), (9.820, 2.790)],
    ],
    34: [  # 2026-08-25 (Nico): the "absolute value" label and its
           # arrow are one object
        [(1.380, 3.460), (2.360, 3.850)],
    ],
    28: [  # 2026-08-25 (Nico, from CT): the ringed percentage symbol,
           # the arrow and the label are one object
        [(8.100, 1.320), (7.200, 1.430), (6.625, 1.745)],
    ],
    21: [  # demand-shift figure: each shifted curve + its arrow + label
        [(7.791, 1.920), (8.254, 2.645), (9.226, 3.395)],   # rising
        [(6.928, 2.331), (7.708, 2.881), (7.188, 3.653)],   # falling
    ],
}

# Manual groups applied AFTER the geometric rules, so a member may itself
# be a group that rule 1 just built (slide 19: the MPV convention callout
# nests inside a group with the arrow that points at the curve).
_M2_INCLASS_GROUPS_POST = {
    # 2026-08-25 (Nico): each elasticity card is ONE object - rule 1 has
    # already paired the box with its first line, so this pass folds the
    # second line and the verdict into that group
    34: [
        [(0.800, 3.980), (0.800, 4.660), (0.800, 5.280)],
        [(4.820, 3.980), (4.820, 4.660), (4.820, 5.280)],
        [(8.840, 3.980), (8.840, 4.660), (8.840, 5.280)],
    ],
    19: [
        [(5.500, 2.870), (4.092, 3.850)],
    ],
}

# Video deck (2026-08-24, Nico's hand-grouping on slide 4): the two
# axes with their P / Q labels are one object, and the demand line, its
# D label and the $400 / 1600 ticks are another (built as a nested pair,
# exactly the way he made it).
_M2_VIDEO_GROUPS = {
    # 2026-08-25 (Nico): on slide 3 the demand line, its two tick labels
    # and the "D" label are one object, revealed on one click
    3: [
        [(1.10, 2.62), (0.03, 2.48), (5.79, 5.91), (6.43, 5.51)],
    ],
    # slide 37's "best fit?" callout: its text box is TALLER than the
    # cream rect (his own copy is a hand-scaled group), so the geometric
    # rule cannot see the pair - name the two explicitly
    32: [
        [(8.480, 3.410), (8.590, 3.245)],
    ],
}
# slide 4's manual groups are gone: from 2026-08-25 that slide arrives
# spliced from the final Video 1 deck, already grouped.
_M2_VIDEO_GROUPS_POST = {}

# Filled boxes that must NOT be merged with the text sitting on them,
# because the text has to animate paragraph by paragraph.  Keyed by deck
# then slide, listing the box's rendered (x, y) in inches.
# (slide 4's entry is gone: from 2026-08-25 that slide arrives spliced
# from the final Video 1 deck, already grouped the way he wants it.)
NO_GROUP_BOXES_BY_DECK = {}

MANUAL_GROUPS_BY_DECK = {
    "Module 2 - In Class Revised": _M2_INCLASS_GROUPS,
    "Module 2 - Video Part Revised": _M2_VIDEO_GROUPS,
}
MANUAL_GROUPS_POST_BY_DECK = {
    "Module 2 - In Class Revised": _M2_INCLASS_GROUPS_POST,
    "Module 2 - Video Part Revised": _M2_VIDEO_GROUPS_POST,
}
# resolve for the deck actually being processed (a side-path build keeps
# the canonical name plus a "_test" suffix)
_STEM = DECK.stem[:-5] if DECK.stem.endswith("_test") else DECK.stem
MANUAL_GROUPS = MANUAL_GROUPS_BY_DECK.get(_STEM, {})
MANUAL_GROUPS_POST = MANUAL_GROUPS_POST_BY_DECK.get(_STEM, {})
NO_GROUP_BOXES = NO_GROUP_BOXES_BY_DECK.get(_STEM, {})
SPLICED = SPLICED_BY_DECK.get(_STEM, set())
for _a in _args:
    if _a.startswith("--spliced="):
        SPLICED = {int(x) for x in _a.split("=", 1)[1].split(",")}


def q(ns, t):
    return "{%s}%s" % (ns, t)


def bbox(el):
    tag = ET.QName(el).localname
    if tag == "graphicFrame":
        xf = el.find(q(P, "xfrm"))
    else:
        sppr = el.find(q(P, "spPr"))
        if sppr is None:
            sppr = el.find(q(P, "grpSpPr"))
        xf = sppr.find(q(A, "xfrm")) if sppr is not None else None
    if xf is None:
        return None
    off = xf.find(q(A, "off"))
    ext = xf.find(q(A, "ext"))
    if off is None or ext is None:
        return None
    x, y = int(off.get("x")), int(off.get("y"))
    w, h = int(ext.get("cx")), int(ext.get("cy"))
    return (x, y, w, h)


def has_text(el):
    for t in el.iter(q(A, "t")):
        if t.text and t.text.strip():
            return True
    for t in el.iter(q(M, "t")):
        if t.text and t.text.strip():
            return True
    return False


def prst_of(el):
    g = el.find(".//" + q(A, "prstGeom"))
    return g.get("prst") if g is not None else None


def sppr_solidfill(el):
    sppr = el.find(q(P, "spPr"))
    return sppr is not None and sppr.find(q(A, "solidFill")) is not None


def has_outer_shadow(el):
    sppr = el.find(q(P, "spPr"))
    if sppr is None:
        return False
    eff = sppr.find(q(A, "effectLst"))
    return eff is not None and eff.find(q(A, "outerShdw")) is not None


def is_caption(el):
    """Small all-italic text box (picture caption / source line)."""
    if not has_text(el):
        return False
    b = bbox(el)
    if b is None or b[3] > 0.5 * EMU:
        return False
    rprs = [r for r in el.iter(q(A, "rPr"))]
    if not rprs:
        return False
    for r in rprs:
        if r.get("i") != "1":
            return False
        sz = r.get("sz")
        if sz and int(sz) > 1600:
            return False
    return True


def contains(outer, inner, slack=0.08 * EMU):
    ox, oy, ow, oh = outer
    ix, iy, iw, ih = inner
    return (ix >= ox - slack and iy >= oy - slack
            and ix + iw <= ox + ow + slack
            and iy + ih <= oy + oh + slack)


def make_group(spTree, members, gid, anchor="first"):
    """Wrap members (document-order list of elements) in a p:grpSp.

    ``anchor`` picks the tree position of the new group: "first" (the
    historical behaviour of rules 1-3) or "last", which is what
    PowerPoint itself does — the group takes the z-order of the topmost
    member.  Manual groups use "last" so a rebuild reproduces Nico's
    hand-made groups shape-for-shape."""
    boxes = [bbox(m) for m in members]
    x0 = min(b[0] for b in boxes)
    y0 = min(b[1] for b in boxes)
    x1 = max(b[0] + b[2] for b in boxes)
    y1 = max(b[1] + b[3] for b in boxes)
    grp = ET.SubElement(spTree, q(P, "grpSp"))
    nv = ET.SubElement(grp, q(P, "nvGrpSpPr"))
    pr = ET.SubElement(nv, q(P, "cNvPr"))
    pr.set("id", str(gid))
    pr.set("name", "Group pair %d" % gid)
    ET.SubElement(nv, q(P, "cNvGrpSpPr"))
    ET.SubElement(nv, q(P, "nvPr"))
    gpr = ET.SubElement(grp, q(P, "grpSpPr"))
    # schema order matters: off, ext, chOff, chExt — anything else is
    # silently misparsed (children collapse / scale to zero)
    xf = ET.SubElement(gpr, q(A, "xfrm"))
    for tag, a1, v1, a2, v2 in (("off", "x", x0, "y", y0),
                                ("ext", "cx", x1 - x0, "cy", y1 - y0),
                                ("chOff", "x", x0, "y", y0),
                                ("chExt", "cx", x1 - x0, "cy", y1 - y0)):
        e = ET.SubElement(xf, q(A, tag))
        e.set(a1, str(v1))
        e.set(a2, str(v2))
    # move grp to the anchor member's position, then move members inside
    at = members[0] if anchor == "first" else members[-1]
    spTree.remove(grp)
    at.addprevious(grp)
    for m in members:
        spTree.remove(m)
        grp.append(m)
    return grp


def apply_manual_groups(spTree, disp, gid, post=False):
    """Group the explicit member sets in MANUAL_GROUPS[disp].  Members are
    addressed by rendered (x, y) in inches; unlike the geometric rules
    this also reaches connectors (cxnSp), which carry the chart curves
    and axes.  With ``post=True`` it reads MANUAL_GROUPS_POST and may also
    take an existing grpSp as a member (nested groups).
    Returns (n_groups, next_gid, used_ids)."""
    specs = (MANUAL_GROUPS_POST if post else MANUAL_GROUPS).get(disp)
    if not specs:
        return 0, gid, set()
    tags = ("sp", "pic", "graphicFrame", "cxnSp")
    if post:
        tags += ("grpSp",)
    n = 0
    used = set()
    # snapshot the candidates ONCE and keep the references: matching by
    # id() against elements pulled fresh from spTree each pass is unsafe,
    # because lxml frees and RECYCLES proxy ids, so an already-consumed id
    # can spuriously match an untouched shape.
    cands = [c for c in spTree if ET.QName(c).localname in tags]
    for spec in specs:
        members = []
        for want_x, want_y in spec:
            hit = None
            for c in cands:
                if id(c) in used:
                    continue
                b = bbox(c)
                if b is None:
                    continue
                if (abs(b[0] / EMU - want_x) < 0.011
                        and abs(b[1] / EMU - want_y) < 0.011):
                    hit = c
                    break
            assert hit is not None, (
                "s%02d manual group: no shape at (%.3f, %.3f)"
                % (disp, want_x, want_y))
            members.append(hit)
            used.add(id(hit))
        members.sort(key=lambda e: cands.index(e))
        make_group(spTree, members, gid, anchor="last")
        gid += 1
        n += 1
        print("  s%02d manual%s group of %d"
              % (disp, " post" if post else "", len(members)))
    return n, gid, used


def process_slide(tree, disp):
    spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
    kids = [c for c in spTree
            if ET.QName(c).localname in ("sp", "pic", "graphicFrame")]
    info = []
    for c in kids:
        tag = ET.QName(c).localname
        b = bbox(c)
        if b is None:
            continue
        info.append({"el": c, "tag": tag, "b": b,
                     "text": has_text(c) if tag == "sp" else False,
                     "prst": prst_of(c) if tag == "sp" else None})
    max_id = 0
    for pr in tree.iter(q(P, "cNvPr")):
        max_id = max(max_id, int(pr.get("id")))
    gid = max_id + 100
    n_groups, gid, used = apply_manual_groups(spTree, disp, gid)

    # rule 1: bare filled roundRect + contained text box
    for box in info:
        if id(box["el"]) in used or box["tag"] != "sp":
            continue
        if box["prst"] != "roundRect" or box["text"]:
            continue
        if not sppr_solidfill(box["el"]):
            continue
        if box["b"][2] > 10 * EMU:
            continue          # agenda highlight bands, not callouts
        if any(abs(box["b"][0] / EMU - bx) < 0.011
               and abs(box["b"][1] / EMU - by) < 0.011
               for bx, by in NO_GROUP_BOXES.get(disp, ())):
            continue          # its text must stay separately animatable
        for txt in info:
            if id(txt["el"]) in used or txt["tag"] != "sp":
                continue
            if not txt["text"] or txt["el"] is box["el"]:
                continue
            if contains(box["b"], txt["b"]):
                make_group(spTree, [box["el"], txt["el"]], gid)
                used.update((id(box["el"]), id(txt["el"])))
                gid += 1
                n_groups += 1
                print("  s%02d rule1 box+text grouped" % disp)
                break

    # rule 2: shadow backing + graphicFrame
    for back in info:
        if id(back["el"]) in used or back["tag"] != "sp":
            continue
        if back["prst"] != "rect" or back["text"]:
            continue
        if not has_outer_shadow(back["el"]):
            continue
        for gf in info:
            if id(gf["el"]) in used or gf["tag"] != "graphicFrame":
                continue
            if contains(back["b"], gf["b"], slack=0.25 * EMU):
                make_group(spTree, [back["el"], gf["el"]], gid)
                used.update((id(back["el"]), id(gf["el"])))
                gid += 1
                n_groups += 1
                print("  s%02d rule2 shade+frame grouped" % disp)
                break

    # rule 3: picture(s) + caption beneath
    caps = [c for c in info if c["tag"] == "sp"
            and id(c["el"]) not in used and is_caption(c["el"])]
    for cap in caps:
        cx, cy, cw, ch = cap["b"]
        matches = []
        for pic in info:
            if pic["tag"] != "pic" or id(pic["el"]) in used:
                continue
            px, py, pw, phh = pic["b"]
            gap_below = cy - (py + phh)          # caption under picture
            gap_above = py - (cy + ch)           # title caption above it
            x_overlap = min(cx + cw, px + pw) - max(cx, px)
            ok_below = (-0.35 * EMU <= gap_below <= 0.4 * EMU
                        and x_overlap > 0.3 * pw)
            # above-captions additionally must not be much wider than the
            # picture (keeps slide-wide subtitle notes out)
            ok_above = (-0.2 * EMU <= gap_above <= 0.4 * EMU
                        and x_overlap > 0.3 * pw and cw <= 1.5 * pw)
            if ok_below or ok_above:
                matches.append(pic)
        if matches:
            members = sorted(matches, key=lambda m: list(spTree).index(
                m["el"]))
            els = [m["el"] for m in members] + [cap["el"]]
            els.sort(key=lambda e: list(spTree).index(e))
            make_group(spTree, els, gid)
            for e in els:
                used.add(id(e))
            gid += 1
            n_groups += 1
            print("  s%02d rule3 %d pic(s)+caption grouped"
                  % (disp, len(els) - 1))
    n_post, gid, _ = apply_manual_groups(spTree, disp, gid, post=True)
    n_groups += n_post
    return n_groups


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target") for r in
             ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    order = [os.path.basename(rid2t[s.get(q(R, "id"))])
             for s in pres.find(q(P, "sldIdLst"))]
    total = 0
    for disp in range(1, len(order) + 1):
        if disp in SPLICED:
            continue
        part = "ppt/slides/" + order[disp - 1]
        tree = ET.fromstring(data[part])
        n = process_slide(tree, disp)
        if n:
            data[part] = ET.tostring(tree, xml_declaration=True,
                                     encoding="UTF-8", standalone=True)
            total += n
    tmp = DECK.with_suffix(".group_tmp.pptx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in data.items():
            zout.writestr(name, blob)
    shutil.move(str(tmp), str(DECK))
    print("total groups: %d — saved %s" % (total, DECK))


if __name__ == "__main__":
    main()
