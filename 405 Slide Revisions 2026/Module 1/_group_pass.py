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
                           picture captions group all pictures they span);
  4. label+link button   — a text label + the action-button link marker
                           sitting just to its right, on the slides in
                           LINK_LABEL_SLIDES. Scoped to an explicit list
                           because the same geometry on slides 2 / 9 would
                           swallow the whole bullet box, which has to keep
                           animating paragraph by paragraph.
  5. chart curve+label   — on the slides in CHART_GROUPS, each named label
                           is grouped with its nearest connector, matching
                           the grouping in Nico's original slides.

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
import sys as _sys
_a = [x for x in _sys.argv[1:] if not x.startswith("-")]
DECK = Path(_a[0]) if _a else HERE / "Module 1 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0

SPLICED = {7, 8, 24, 25, 28, 29, 49, 50}   # shifted 2026-08-22 (poll pairs)

# Rule 4 (2026-08-23): slides where a text label and its action-button
# link marker must move and reveal as ONE object. Display 12's podcast
# label is revealed on its own animation click, so the button has to ride
# with it instead of sitting there from the start.
# 12 = podcast label + Sound button AND the jump pill + its button;
# 17 = the jump pill + its button. Deliberately NOT the backup
# slides: display 94's two link labels share one bullets box, so a
# containment match there would swallow the box with only the first
# of its two buttons.
OUTLINE_SLIDES = {18, 19, 43, 56, 61, 69, 71, 78, 88}

LINK_LABEL_SLIDES = {12, 17}

# Rule 5 (2026-08-23): display 36's avocado chart adopts the grouping of
# Nico's original slide 30 — each initial curve with its label, and each
# equilibrium guide with its axis label(s). Each inner list is one group;
# every label in it is paired with its nearest unused connector.
CHART_GROUPS = {
    36: {
        "D":  ["sdcurve:D", "sdlabel:D"],
        "S":  ["sdcurve:S", "sdlabel:S"],
        "Q0": ["sdguide:v:0", "sdxlab:Q0"],
        "Q1": ["sdguide:h:1", "sdguide:v:1", "sdxlab:Q1", "sdylab:PPeak"],
        "Q2": ["sdguide:v:2", "sdxlab:Q2"],
    },
    # 79: Nico grouped the ice-cream picture with its header (2026-08-23)
    81: {
        "cones": ["sdlabel:cones", "sdpic:cones"],
    },
    # 81: his two hand-made groups — the D' set and the Q3 set
    83: {
        "Dp": ["sdcurve:Dp", "sdlabel:Dp", "sdarrow:ii", "sdlabel:ii"],
        "Q3": ["sdguide:v:Q3", "sdguide:h:Q3", "sdxlab:Q3"],
    },
    # 82: the AI-chips slide. No Q2 set — Nico removed those guides.
    84: {
        # (the picture+caption pair is grouped by rule 3, which carries
        # the sdpic: name through as sdgroup:chips)
        "D":     ["sdcurve:D", "sdlabel:D"],
        "Dp":    ["sdcurve:Dp", "sdlabel:Dp", "sdarrow:shift"],
        "Q1":    ["sdguide:h:P1", "sdguide:v:Q1", "sdylab:P1",
                  "sdxlab:Q1"],
    },
    # 84: Nico's four hand-made groups (2026-08-23). The S' set carries
    # the ii) arrow and label; the Q3 guides ride with it.
    86: {
        "i":  ["sdarrow:i", "sdlabel:i"],
        "P2": ["sdguide:h:P2", "sdguide:v:Q2", "sdylab:P2", "sdxlab:Q2"],
        "P1": ["sdguide:h:P1", "sdguide:v:Q1", "sdylab:P1", "sdxlab:Q1"],
        "Sp": ["sdcurve:Sp", "sdlabel:Sp", "sdarrow:ii", "sdlabel:ii",
               "sdguide:h:Q3", "sdguide:v:Q3", "sdxlab:Q3"],
    },
}
ACTION_BUTTON_PRSTS = {
    "actionButtonSound", "actionButtonDocument", "actionButtonMovie",
    "actionButtonEnd", "actionButtonBeginning", "actionButtonInformation",
}


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


def text_of(el):
    parts = [x.text or "" for x in el.iter(q(A, "t"))]
    parts += [x.text or "" for x in el.iter(q(M, "t"))]
    return " ".join("".join(parts).split())


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
        # an outline agenda band spans the slide — layout, not a callout
        # (Teaching/CLAUDE.md, Module-Outline section). Scoped to the
        # outline slides rather than a width threshold: the Tapestry quote
        # box on display 74 is 12.08" wide and DOES need grouping.
        if disp in OUTLINE_SLIDES and box["b"][2] > 10.0 * EMU:
            continue
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
            grp3 = make_group(spTree, els, gid)
            # carry an sdpic: name through to the group so animation plans
            # can address it as sdgroup:<key> (2026-08-23)
            for e in els:
                c3 = e.find(".//" + q(P, "cNvPr"))
                nm3 = c3.get("name") if c3 is not None else ""
                if nm3 and nm3.startswith("sdpic:"):
                    g3 = grp3.find(".//" + q(P, "cNvPr"))
                    if g3 is not None:
                        g3.set("name", "sdgroup:" + nm3.split(":", 1)[1])
                    break
            for e in els:
                used.add(id(e))
            gid += 1
            n_groups += 1
            print("  s%02d rule3 %d pic(s)+caption grouped"
                  % (disp, len(els) - 1))

    # --- rule 4: link label + its action button -------------------------
    if disp in LINK_LABEL_SLIDES:
        for btn in info:
            if (id(btn["el"]) in used or btn["tag"] != "sp"
                    or btn["prst"] not in ACTION_BUTTON_PRSTS):
                continue
            bx, by, bw, bh = btn["b"]
            bcx, bcy = bx + bw / 2, by + bh / 2
            best = None
            for lab in info:
                if (id(lab["el"]) in used or lab is btn
                        or lab["tag"] != "sp" or not lab["text"]):
                    continue
                lx, ly, lw, lh = lab["b"]
                # the button either abuts the box or, when the label is
                # centred in a box wider than its text, sits inside it
                slack = 0.60 * EMU
                if (lx - slack <= bcx <= lx + lw + slack
                        and ly - slack <= bcy <= ly + lh + slack):
                    area = lw * lh
                    if best is None or area < best[0]:
                        best = (area, lab)
            if best is None:
                continue
            els = sorted([best[1]["el"], btn["el"]],
                         key=lambda e: list(spTree).index(e))
            make_group(spTree, els, gid)
            for e in els:
                used.add(id(e))
            gid += 1
            n_groups += 1
            print("  s%02d rule4 label+link-button grouped" % disp)

    # --- rule 5: chart curve/guide + its label (paired by name) --------
    spec_map = CHART_GROUPS.get(disp)
    if spec_map:
        by_name = {}
        for c in spTree:
            if ET.QName(c).localname not in ("sp", "cxnSp", "pic"):
                continue
            cnv = c.find(".//" + q(P, "cNvPr"))
            if cnv is None or id(c) in used:
                continue
            nm = cnv.get("name") or ""
            if nm.startswith(("sdcurve:", "sdlabel:", "sdguide:",
                              "sdxlab:", "sdylab:", "sdarrow:",
                              "sdpic:", "sdcap:")):
                by_name[nm] = c
        for key, names in spec_map.items():
            els = [by_name[n] for n in names if n in by_name]
            missing = [n for n in names if n not in by_name]
            if missing:
                print("  s%02d rule5 %s: MISSING %s" % (disp, key, missing))
            if len(els) < 2:
                continue
            els.sort(key=lambda e: list(spTree).index(e))
            grp = make_group(spTree, els, gid)
            gcnv = grp.find(".//" + q(P, "cNvPr"))
            if gcnv is not None:
                gcnv.set("name", "sdgroup:%s" % key)
            for e in els:
                used.add(id(e))
            gid += 1
            n_groups += 1
            print("  s%02d rule5 sdgroup:%s <- %s"
                  % (disp, key, ", ".join(names)))
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
