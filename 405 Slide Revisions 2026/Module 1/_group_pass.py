# -*- coding: utf-8 -*-
"""Grouping pass for "Module 1 - Revised.pptx" — merges shape
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
DECK = HERE / "Module 1 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0

SPLICED = {22, 25, 45}      # shifted 2026-08-20 (AC + copper inserts)


def q(ns, t):
    return "{%s}%s" % (ns, t)


def bbox(el):
    tag = ET.QName(el).localname
    if tag == "graphicFrame":
        xf = el.find(q(P, "xfrm"))
    else:
        sppr = el.find(q(P, "spPr"))
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


def make_group(spTree, members, gid):
    """Wrap members (document-order list of elements) in a p:grpSp at the
    first member's tree position."""
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
    # move grp to the first member's position, then move members inside
    first = members[0]
    spTree.remove(grp)
    first.addprevious(grp)
    for m in members:
        spTree.remove(m)
        grp.append(m)
    return grp


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
    used = set()
    n_groups = 0

    # rule 1: bare filled roundRect + contained text box
    for box in info:
        if id(box["el"]) in used or box["tag"] != "sp":
            continue
        if box["prst"] != "roundRect" or box["text"]:
            continue
        if not sppr_solidfill(box["el"]):
            continue
        # M1: skip large hosting panels (e.g. the Homo-Economicus cream
        # panel holds several text blocks + images — not a callout pair)
        if box["b"][3] > 2.5 * EMU:
            continue
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
