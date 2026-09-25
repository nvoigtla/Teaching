# -*- coding: utf-8 -*-
"""Grouping pass for "Module 6 - Revised.pptx" — merges shape
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

  _build_Module6.py -> _splice_media.py -> _group_pass.py
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
DECK = Path(_a[0]) if _a else HERE / "Module 6 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0

def _derive_spliced():
    """The PollEverywhere slides, read from the splice map itself.

    Hard-coded, this set went stale twice: the deck's poll slides moved to
    72/85/96 when two slides were added on 2026-09-09, and to 73/86/97 when
    another went in at 27 on 2026-09-13, while the constant still said
    {72, 83, 94}.  A stale set means the pass skips an ordinary slide and
    does NOT skip a real poll slide -- and a poll slide is exactly the one
    that must not be touched, because its notes part is load-bearing for
    the PollEv add-in.  Derive it, the way OUTLINE_SLIDES is derived.
    """
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "_splice_media_map", str(HERE / "_splice_media.py"))
    mod = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(mod)
        return set(mod.SPLICE_MAP)
    except Exception:
        # the splice module guards itself against being imported; fall back
        # to reading the literal
        import re as _re
        txt = (HERE / "_splice_media.py").read_text(encoding="utf-8")
        blk = _re.search(r"SPLICE_MAP\s*=\s*\{(.*?)\}", txt, _re.S)
        return set(int(x) for x in _re.findall(r"^\s*(\d+):", blk.group(1),
                                               _re.M))


SPLICED = _derive_spliced()             # the PollEverywhere slides


def _derive_outline_slides(deck):
    """Every slide titled "Outline of Module 6".

    Module 4 hard-coded this set and it went stale as slides were inserted,
    leaving three highlight bands to be swallowed by rule 1.  Derive it
    instead, so it cannot drift.
    """
    import re as _re
    import zipfile as _zip
    import xml.etree.ElementTree as _ET
    ns = {"p": P, "r": R}
    out = set()
    with _zip.ZipFile(str(deck)) as z:
        pres = _ET.fromstring(z.read("ppt/presentation.xml"))
        rels = {r.get("Id"): r.get("Target") for r in
                _ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
        order = ["ppt/" + rels[s.get("{%s}id" % R)].replace("../", "")
                 for s in pres.find("p:sldIdLst", ns)]
        for n, part in enumerate(order, 1):
            if "Outline of Module 6" in z.read(part).decode("utf-8"):
                out.add(n)
    return out


# The agenda slides: their cream highlight band is a LAYOUT band, not a
# callout, so rule 1 must not swallow it together with the item text.
OUTLINE_SLIDES = _derive_outline_slides(DECK)

# Module 6 has no label + action-button pairs and no hand-authored chart
# groupings to reproduce, so rules 4 and 5 are inert here.
LINK_LABEL_SLIDES = set()
CHART_GROUPS = {}

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


def arrow_ends(el):
    """``(has_arrowhead, (origin, head))`` for a connector, in EMU.

    OOXML draws ``headEnd`` at the START of the line and ``tailEnd`` at the
    end, and ``flipH`` / ``flipV`` decide which corner of the bounding box
    the line starts in -- so the pointed end is whichever corner carries a
    marker.  Same reading as `_check_labels.py`, which is what makes the
    two agree on where an arrow's origin is.
    """
    b = bbox(el)
    if b is None:
        return False, None
    x, y, w, h = b
    sppr = el.find(q(P, "spPr"))
    xf = sppr.find(q(A, "xfrm")) if sppr is not None else None
    fh = xf is not None and xf.get("flipH") == "1"
    fv = xf is not None and xf.get("flipV") == "1"
    ln = sppr.find(q(A, "ln")) if sppr is not None else None
    def _end(name):
        if ln is None:
            return False
        e = ln.find(q(A, name))
        return e is not None and e.get("type") not in (None, "none")
    head, tail = _end("headEnd"), _end("tailEnd")
    if not (head or tail):
        return False, None
    p0 = (x + (w if fh else 0), y + (h if fv else 0))
    p1 = (x + (0 if fh else w), y + (0 if fv else h))
    # the marker end is the HEAD of the arrow; the other end is the origin
    return True, ((p1, p0) if head and not tail else (p0, p1))


def is_dashed(el):
    """A dashed connector -- a `_fig_guide` drop line, not an annotation."""
    sppr = el.find(q(P, "spPr"))
    ln = sppr.find(q(A, "ln")) if sppr is not None else None
    return ln is not None and ln.find(q(A, "prstDash")) is not None


def line_ends(el):
    """Both endpoints of a connector, in EMU, honouring flipH / flipV."""
    b = bbox(el)
    if b is None:
        return None
    x, y, w, h = b
    sppr = el.find(q(P, "spPr"))
    xf = sppr.find(q(A, "xfrm")) if sppr is not None else None
    fh = xf is not None and xf.get("flipH") == "1"
    fv = xf is not None and xf.get("flipV") == "1"
    return ((x + (w if fh else 0), y + (h if fv else 0)),
            (x + (0 if fh else w), y + (0 if fv else h)))


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


def _name_group(grp, name):
    """Stamp a group's name.  `_animate.py` looks for the "sdpair:" prefix
    to index a pair's MEMBERS while targeting the pair itself, so every
    per-slide plan keeps working after a new pair is grouped."""
    cnv = grp.find(".//" + q(P, "cNvPr"))
    if cnv is not None:
        cnv.set("name", name)
    return grp


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
        cnv = c.find(".//" + q(P, "cNvPr"))
        info.append({"el": c, "tag": tag, "b": b,
                     "name": (cnv.get("name") or "") if cnv is not None
                             else "",
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
                              "sdpic:", "sdcap:", "sdbadge:")):
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

    # --- rule 6: an annotation arrow + the text box at its ORIGIN -------
    # 2026-09-10 (Nico): "when we have arrows with a text box, group those
    # into one shape."  The pair is already defined by the arrow-origin
    # rule -- the label sits at the tail and the arrow starts at the box's
    # edge -- so this groups what that rule put together, and the label can
    # no longer be dragged off its own arrow.
    #
    # AXES are excluded, the same way `_check_labels.py` excludes them: an
    # axis is a long, perfectly horizontal or vertical arrow, and its title
    # is anchored to the TIP by a different rule, so the two are not a pair.
    for cxn in [c for c in spTree
                if ET.QName(c).localname == "cxnSp" and id(c) not in used]:
        b = bbox(cxn)
        if b is None or is_dashed(cxn):
            continue                    # a dashed drop line is not an arrow
        x, y, w, h = b
        if min(w, h) < 0.02 * EMU and max(w, h) > 2.5 * EMU:
            continue                                    # an axis
        headed, ends = arrow_ends(cxn)
        # Is this one of the leaders _arrow_from_box stamps?  If so the
        # label is KNOWN -- the leader was drawn from that box's edge --
        # and it must not be re-derived from the size and distance
        # heuristics below.  Those exist to stop an UNSTAMPED arrow from
        # grabbing arbitrary text, and they were throwing away real pairs
        # (2026-09-15, Nico: "the kink arrow needs to be grouped with the
        # corresponding box.  Why did you miss this?"):
        #   * slide 40's kink callout is 2.98" wide, over the 2.60" cap
        #     that decides something is "not a label";
        #   * slide 89's two share arrows are only ~0.3" long and start at
        #     their box's edge, so the box CENTRE sits ~0.75" away, right
        #     on the distance cap.
        # Both are exactly the pair the rule is about.
        cnv = cxn.find(".//" + q(P, "cNvPr"))
        stamped = bool(cnv is not None
                       and (cnv.get("name") or "").startswith("sdleader:"))
        if headed:
            origin, head = ends
        else:
            # A callout LEADER is headless: `_arrow_from_box` runs it from
            # the label's edge to a dot on the curve, so its origin is
            # whichever endpoint sits on a label's boundary.  The pair is as
            # real as an arrow's -- Nico's MPV callouts are built this way.
            # Only the STAMPED leaders qualify: a curve is a solid headless
            # connector too, and its endpoint lands on the axis exactly
            # where a tick label sits, which grouped demand lines with
            # "$20" until the leaders were named.
            if not stamped:
                continue
            ends = line_ends(cxn)
            if ends is None:
                continue
            origin, head = ends
            near = [lb for lb in info
                    if lb["tag"] == "sp" and lb["text"]
                    and id(lb["el"]) not in used
                    and contains(lb["b"], (origin[0], origin[1], 0, 0),
                                 slack=0.16 * EMU)]
            if not near:
                origin, head = head, origin
                near = [lb for lb in info
                        if lb["tag"] == "sp" and lb["text"]
                        and id(lb["el"]) not in used
                        and contains(lb["b"], (origin[0], origin[1], 0, 0),
                                     slack=0.16 * EMU)]
            if not near:
                continue
        best, bd = None, 0.75 * EMU
        if stamped:
            # the box the leader LEAVES: its origin sits on that box's
            # boundary by construction, so take the box that contains the
            # origin (either endpoint -- a headed leader may have been
            # drawn box -> target, a headless one either way round)
            for o in (origin, head):
                own = [lb for lb in info
                       if lb["tag"] == "sp" and lb["text"]
                       and id(lb["el"]) not in used
                       and contains(lb["b"], (o[0], o[1], 0, 0),
                                    slack=0.16 * EMU)]
                if own:
                    best = own[0]
                    break
        for lab in ([] if best is not None else info):
            if (id(lab["el"]) in used or lab["tag"] != "sp"
                    or not lab["text"] or lab["prst"] not in (None, "rect")):
                continue
            lx, ly, lw, lh = lab["b"]
            if lw > 2.60 * EMU or lh > 0.72 * EMU:
                continue                                # not a label
            cx, cy = lx + lw / 2.0, ly + lh / 2.0
            d_o = ((cx - origin[0]) ** 2 + (cy - origin[1]) ** 2) ** 0.5
            d_h = ((cx - head[0]) ** 2 + (cy - head[1]) ** 2) ** 0.5
            if d_o < bd and d_o <= d_h:
                best, bd = lab, d_o
        if best is None:
            continue
        els = sorted([best["el"], cxn], key=lambda e: list(spTree).index(e))
        _name_group(make_group(spTree, els, gid), "sdpair:arrow:%d" % gid)
        for e in els:
            used.add(id(e))
        gid += 1
        n_groups += 1
        print("  s%02d rule6 arrow+label grouped (%s)"
              % (disp, text_of(best["el"])[:24]))

    # --- rule 7: a shape + the label written ON it ----------------------
    # 2026-09-10 (Nico): "group the label of a shape with the shape (like
    # the F* and the triangle on slide 49)."  Two shapes of that pattern:
    #   (a) a region label with the region it names, paired by the KEY
    #       both carry -- `_fig_poly` stamps "sdregion:<key>" and
    #       `_region_label` stamps "sdreglab:<key>", derived from the same
    #       polygon.  A bbox test alone was wrong: it grouped the CURVE
    #       labels "D" and "MR" with whatever shaded region they happened
    #       to sit inside, and those name curves, not areas;
    #   (b) an action-button link marker seated in a labelled box (the
    #       practice-video box, the jump pills) -- the mark belongs to the
    #       box it is seated in.
    # Containers are capped in size so this stays figure regions and small
    # chrome: an outline agenda band is 12.15" wide and a full-slide
    # backing rect wider still, and neither is a shape with a label on it.
    # --- rule 8: a CURVE and the label that names it --------------------
    # 2026-09-14 (Nico): "everywhere in Module 6, group a curve with its
    # label".  Paired by the key BOTH carry -- `_fig_line(curve="D")` and
    # `_fig_curve_label(curve="D")` -- never by proximity: a chart carries
    # several lines and several labels, and the nearest label to a line is
    # regularly the wrong one.  That is the mistake the region pairing made
    # in its first version, and it is the same mistake here.
    curves = {}
    for c in spTree:
        if ET.QName(c).localname not in ("cxnSp", "sp"):
            continue
        cnv = c.find(".//" + q(P, "cNvPr"))
        nm = (cnv.get("name") or "") if cnv is not None else ""
        if nm.startswith("sdcurve:"):
            curves[nm.split(":", 1)[1]] = c
    for lab in info:
        if id(lab["el"]) in used or not lab["name"].startswith("sdcurvelab:"):
            continue
        key = lab["name"].split(":", 1)[1]
        crv = curves.get(key)
        if crv is None or id(crv) in used:
            continue
        els = sorted([crv, lab["el"]], key=lambda e: list(spTree).index(e))
        _name_group(make_group(spTree, els, gid), "sdpair:curve:%s" % key)
        for e in els:
            used.add(id(e))
        gid += 1
        n_groups += 1
        print("  s%02d rule8 curve+label grouped (%s)" % (disp, key))

    regions = {r["name"]: r for r in info
               if r["name"].startswith("sdregion:")}
    for lab in info:
        if id(lab["el"]) in used or not lab["name"].startswith("sdreglab:"):
            continue
        reg = regions.get("sdregion:" + lab["name"].split(":", 1)[1])
        if reg is None or id(reg["el"]) in used:
            continue
        els = sorted([reg["el"], lab["el"]],
                     key=lambda e: list(spTree).index(e))
        _name_group(make_group(spTree, els, gid), "sdpair:%s" % lab["name"])
        for e in els:
            used.add(id(e))
        gid += 1
        n_groups += 1
        print("  s%02d rule7 shape+label grouped (%s)"
              % (disp, text_of(lab["el"])[:24]))

    for btn in info:
        if (id(btn["el"]) in used or btn["tag"] != "sp"
                or btn["prst"] not in ACTION_BUTTON_PRSTS):
            continue
        bx, by, bw, bh = btn["b"]
        bcx, bcy = bx + bw / 2.0, by + bh / 2.0
        best = None
        for host in info:
            if (id(host["el"]) in used or host is btn
                    or host["tag"] != "sp" or not host["text"]):
                continue
            hx, hy, hw, hh = host["b"]
            if hw > 6.5 * EMU or hh > 2.0 * EMU:
                continue
            if not (hx <= bcx <= hx + hw and hy <= bcy <= hy + hh):
                continue
            if best is None or hw * hh < best[0]:
                best = (hw * hh, host)
        if best is None:
            continue
        els = sorted([best[1]["el"], btn["el"]],
                     key=lambda e: list(spTree).index(e))
        _name_group(make_group(spTree, els, gid), "sdpair:btn:%d" % gid)
        for e in els:
            used.add(id(e))
        gid += 1
        n_groups += 1
        print("  s%02d rule7 box+link-button grouped (%s)"
              % (disp, text_of(best[1]["el"])[:24]))
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
