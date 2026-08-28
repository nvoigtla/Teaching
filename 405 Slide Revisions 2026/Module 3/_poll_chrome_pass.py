# -*- coding: utf-8 -*-
"""Normalise ALL poll chrome in `Module 3 - Revised.pptx` (2026-08-27).

Applies the Teaching CLAUDE.md poll rules deck-wide:

  * every "Poll Break" parallelogram is rebuilt at the FIXED position
    left 10.238", top 6.769", 2.95 x 0.533" on the 13.33 x 7.5" canvas
    (the badge straddles the footer rule at y 7.15");
  * box + label are ONE `<p:grpSp>`;
  * the badge is the LAST shape in the spTree, so it renders in FRONT of
    the footer rule and the page number;
  * poll chrome is NEVER animated - any `<p:timing>` effect targeting a
    badge is dropped;
  * the round gold POLL pill on the PollEverywhere slides is normalised
    to one position (the majority one already in the deck) and is also
    the last shape.

The canonical badge XML is lifted verbatim from the two slides that
already carry it (the Rivian-designer pair), so nothing about its look
is invented here.

Direct zip + lxml surgery - the deck is never round-tripped through
python-pptx (that would drop the Amazon NULL video rel and the
PollEverywhere tag parts).

Usage:  python _poll_chrome_pass.py [--dry-run]
"""
import copy
import shutil
import sys
import zipfile

from lxml import etree as ET

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0
GOLD = "E09F3E"

DECK = "Module 3 - Revised.pptx"
DRY = "--dry-run" in sys.argv

from _deck_guard import require_committed

# Fixed badge geometry (Teaching CLAUDE.md, 2026-08-27).
BADGE_XY = (9361444, 6190030)                # 10.238", 6.769"
BADGE_WH = (2697480, 487009)                 # 2.95",   0.533"
BADGE_BOX = BADGE_XY + BADGE_WH

# Round POLL pill: the position 4 of the 6 poll slides already use.
PILL_X, PILL_Y = 10366225, 5673852           # 11.334", 6.203"

SHAPES = ("sp", "pic", "graphicFrame", "cxnSp", "grpSp")


def q(ns, t):
    return "{%s}%s" % (ns, t)


def txt(el):
    return "".join(t.text or "" for t in el.iter(q(A, "t"))).strip()


class Pkg(object):
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.parts = {n: z.read(n) for n in z.namelist()}
        self.order = list(z.namelist())
        z.close()

    def xml(self, n):
        return ET.fromstring(self.parts[n])

    def set_xml(self, n, tree):
        self.parts[n] = ET.tostring(tree, xml_declaration=True,
                                    encoding="UTF-8", standalone=True)

    def slides(self):
        pres = self.xml("ppt/presentation.xml")
        rels = self.xml("ppt/_rels/presentation.xml.rels")
        rid2t = {r.get("Id"): r.get("Target") for r in rels}
        return ["ppt/slides/" + rid2t[s.get(q(R, "id"))].split("/")[-1]
                for s in pres.find(q(P, "sldIdLst"))]

    def write(self, path):
        zo = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
        for n in self.order:
            zo.writestr(n, self.parts[n])
        zo.close()


def shape_kids(spTree):
    return [c for c in spTree if ET.QName(c).localname in SHAPES]


def geom(el):
    """(x, y, cx, cy) of a top-level shape, or None."""
    xf = el.find(".//" + q(A, "xfrm"))
    if xf is None:
        return None
    o = xf.find(q(A, "off"))
    e = xf.find(q(A, "ext"))
    if o is None or e is None or o.get("x") is None or e.get("cx") is None:
        return None
    return (int(o.get("x")), int(o.get("y")),
            int(e.get("cx")), int(e.get("cy")))


def ids_of(el):
    return {nv.get("id") for nv in el.iter(q(P, "cNvPr"))}


def max_id(spTree):
    m = 1
    for nv in spTree.iter(q(P, "cNvPr")):
        try:
            m = max(m, int(nv.get("id")))
        except (TypeError, ValueError):
            pass
    return m


def find_badge(spTree):
    """Top-level elements making up the Poll Break badge on this slide:
    a single grpSp, or a loose label plus its gold parallelogram."""
    kids = shape_kids(spTree)
    for c in kids:
        if ET.QName(c).localname == "grpSp" and txt(c) == "Poll Break":
            return [c]
    label = None
    for c in kids:
        if ET.QName(c).localname == "sp" and txt(c) == "Poll Break":
            label = c
            break
    if label is None:
        return []
    out = [label]
    for c in kids:
        if c is label or ET.QName(c).localname != "sp":
            continue
        if c.find(".//" + q(A, "custGeom")) is None:
            continue
        fill = c.find(".//" + q(A, "solidFill") + "/" + q(A, "srgbClr"))
        g = geom(c)
        if (fill is not None and fill.get("val") == GOLD
                and g is not None and g[1] / EMU > 5.8):
            out.append(c)
    return out


def strip_timing(tree, dead_ids):
    """Drop every animation effect targeting one of `dead_ids`.

    Removes a whole click-level `<p:par>` when the badge was its only
    target, otherwise just the effect's own `<p:par>`.  Returns the
    number of clicks removed.
    """
    tm = tree.find(q(P, "timing"))
    if tm is None:
        return 0
    removed = 0
    for seq in tm.iter(q(P, "seq")):
        ctn = seq.find(q(P, "cTn"))
        if ctn is None:
            continue
        ctl = ctn.find(q(P, "childTnLst"))
        if ctl is None:
            continue
        for click in list(ctl):
            tg = [e.get("spid") for e in click.iter(q(P, "spTgt"))]
            if not tg or not (set(tg) & dead_ids):
                continue
            if set(tg) <= dead_ids:
                ctl.remove(click)
                removed += 1
                continue
            for par in list(click.iter(q(P, "par"))):
                sub = [e.get("spid") for e in par.iter(q(P, "spTgt"))]
                if sub and set(sub) <= dead_ids:
                    parent = par.getparent()
                    if parent is not None:
                        parent.remove(par)
    for bl in list(tm.iter(q(P, "bldLst"))):
        for b in list(bl):
            if b.get("spid") in dead_ids:
                bl.remove(b)
        if len(bl) == 0:
            bl.getparent().remove(bl)
    return removed


def canonical_template(pkg, slides):
    for pn in slides:
        t = pkg.xml(pn)
        for c in shape_kids(t.find(q(P, "cSld")).find(q(P, "spTree"))):
            nv = c.find(".//" + q(P, "cNvPr"))
            if (ET.QName(c).localname == "grpSp"
                    and nv is not None
                    and nv.get("name") == "PollBreakBadge"
                    and geom(c) == BADGE_BOX):
                return copy.deepcopy(c)
    return None


def main():
    if not DRY:
        require_committed(DECK)
    pkg = Pkg(DECK)
    slides = pkg.slides()
    template = canonical_template(pkg, slides)
    if template is None:
        sys.exit("canonical PollBreakBadge not found - aborting")

    n_badge = n_moved = n_grouped = n_front = n_anim = n_pill = 0
    log = []

    for i, pn in enumerate(slides, 1):
        tree = pkg.xml(pn)
        spTree = tree.find(q(P, "cSld")).find(q(P, "spTree"))
        dirty = False

        members = find_badge(spTree)
        if members:
            was_group = len(members) == 1
            before = geom(members[0]) if was_group else None
            kids = shape_kids(spTree)
            was_last = kids[-1] in members
            dead = set()
            for el in members:
                dead |= ids_of(el)
            clicks = strip_timing(tree, dead)
            for el in members:
                spTree.remove(el)
            badge = copy.deepcopy(template)
            nid = max_id(spTree)
            nvs = list(badge.iter(q(P, "cNvPr")))
            for k, nv in enumerate(nvs, 1):
                nv.set("id", str(nid + k))
            nvs[0].set("name", "PollBreakBadge")
            spTree.append(badge)
            dirty = True
            n_badge += 1
            notes = []
            if not was_group:
                n_grouped += 1
                notes.append("grouped")
            if before != BADGE_BOX:
                n_moved += 1
                notes.append("repositioned")
            if not was_last:
                n_front += 1
                notes.append("moved to front")
            if clicks:
                n_anim += 1
                notes.append("de-animated (%d click)" % clicks)
            log.append("  slide %-4d badge : %s"
                       % (i, ", ".join(notes) if notes else "already correct"))

        for c in shape_kids(spTree):
            if ET.QName(c).localname != "grpSp" or txt(c) != "POLL":
                continue
            g = geom(c)
            if g and (g[0], g[1]) != (PILL_X, PILL_Y):
                dx, dy = PILL_X - g[0], PILL_Y - g[1]
                xf = c.find(q(P, "grpSpPr")).find(q(A, "xfrm"))
                o = xf.find(q(A, "off"))
                ch = xf.find(q(A, "chOff"))
                o.set("x", str(PILL_X))
                o.set("y", str(PILL_Y))
                ch.set("x", str(int(ch.get("x")) + dx))
                ch.set("y", str(int(ch.get("y")) + dy))
                for sub in c.iter(q(A, "off")):
                    if sub is o or sub is ch:
                        continue
                    sub.set("x", str(int(sub.get("x")) + dx))
                    sub.set("y", str(int(sub.get("y")) + dy))
                dirty = True
                n_pill += 1
                log.append("  slide %-4d pill  : repositioned "
                           "(%.3f, %.3f) -> (%.3f, %.3f)"
                           % (i, g[0] / EMU, g[1] / EMU,
                              PILL_X / EMU, PILL_Y / EMU))
            if shape_kids(spTree)[-1] is not c:
                spTree.remove(c)
                spTree.append(c)
                dirty = True
                log.append("  slide %-4d pill  : moved to front" % i)

        if dirty and not DRY:
            pkg.set_xml(pn, tree)

    print("\n".join(log))
    print("")
    print("Poll Break badges rebuilt : %d" % n_badge)
    print("  grouped (box + label)   : %d" % n_grouped)
    print("  repositioned            : %d" % n_moved)
    print("  moved in front of footer: %d" % n_front)
    print("  de-animated             : %d" % n_anim)
    print("POLL pills normalised     : %d" % n_pill)

    if DRY:
        print("")
        print("(dry run - nothing written)")
        return
    tmp = DECK + ".tmp"
    pkg.write(tmp)
    shutil.move(tmp, DECK)
    print("")
    print("wrote %s" % DECK)


if __name__ == "__main__":
    main()
