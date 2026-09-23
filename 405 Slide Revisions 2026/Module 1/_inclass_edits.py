# -*- coding: utf-8 -*-
"""Two edits to "Module 1 - In Class.pptx" (2026-09-22, Nico):

 1. Slide 4's inserted picture gets the deck's house style - rounded
    corners (roundRect, adj 8000) plus the soft drop shadow. Read off the
    11 pictures in this deck that already carry it (reference: display 17).

 2. After each slide that introduces a poll - the three that carry the
    Poll Break badge, displays 20, 23 and 36 - an EMPTY poll slide is
    inserted, copying display 34 of "Module 4 - Revised.pptx": navy top
    bar + its section tag, the footer chrome, and the round gold POLL
    pill; no title, no content. Nico will drop the PollEverywhere sheet
    onto these later.

    Each placeholder is built from the set-up slide it follows, so it
    inherits that slide's own top-bar tag and footer - which is exactly
    how Module 4's slide 34 relates to its slide 33.

Pure zip + lxml surgery; the deck is never round-tripped through
python-pptx. Rerunnable: it refuses to act twice (it checks for the pill).
"""
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Module 1 - In Class.pptx"
TEMPLATE_DECK = HERE.parent / "Module 4" / "Module 4 - Revised.pptx"
TEMPLATE_DISPLAY = 34

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
EMU = 914400.0
CT_SLIDE = ("application/vnd.openxmlformats-officedocument."
            "presentationml.slide+xml")

# the content band: everything between the title and the footer chrome is
# dropped from the placeholder (title, its rule and gold strip, the body,
# and the Poll Break badge at y 6.77)
BAND_TOP, BAND_BOTTOM = 0.50, 7.10

SETUP_DISPLAYS = [20, 23, 36]


def q(ns, t):
    return "{%s}%s" % (ns, t)


def slide_parts(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = {r.get("Id"): r.get("Target") for r in
            ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    return ["ppt/" + rels[s.get(q(R, "id"))].lstrip("/").replace("../", "")
            for s in pres.find(q(P, "sldIdLst"))]


# ---------------------------------------------------------------------------
# 1. picture style
# ---------------------------------------------------------------------------

def style_picture(tree):
    """Rounded corners + drop shadow on every picture of the slide that
    has neither. Returns the number of pictures changed."""
    n = 0
    for pic in tree.iter(q(P, "pic")):
        spPr = pic.find(q(P, "spPr"))
        if spPr is None:
            continue
        g = spPr.find(q(A, "prstGeom"))
        already = (g is not None and g.get("prst") == "roundRect"
                   and spPr.find(q(A, "effectLst")) is not None)
        if already:
            continue
        for old in spPr.findall(q(A, "prstGeom")):
            spPr.remove(old)
        geom = ET.SubElement(spPr, q(A, "prstGeom"))
        geom.set("prst", "roundRect")
        av = ET.SubElement(geom, q(A, "avLst"))
        gd = ET.SubElement(av, q(A, "gd"))
        gd.set("name", "adj")
        gd.set("fmla", "val 8000")
        xfrm = spPr.find(q(A, "xfrm"))
        if xfrm is not None:
            xfrm.addnext(geom)
        else:
            spPr.insert(0, geom)
        for old in spPr.findall(q(A, "effectLst")):
            spPr.remove(old)
        eff = ET.SubElement(spPr, q(A, "effectLst"))
        sh = ET.SubElement(eff, q(A, "outerShdw"))
        sh.set("blurRad", "50800")
        sh.set("dist", "38100")
        sh.set("dir", "2700000")
        sh.set("algn", "tl")
        sh.set("rotWithShape", "0")
        clr = ET.SubElement(sh, q(A, "srgbClr"))
        clr.set("val", "000000")
        alpha = ET.SubElement(clr, q(A, "alpha"))
        alpha.set("val", "50000")
        n += 1
    return n


# ---------------------------------------------------------------------------
# 2. empty poll placeholder slides
# ---------------------------------------------------------------------------

def top_inches(el, tag):
    if tag == "graphicFrame":
        xf = el.find(q(P, "xfrm"))
    else:
        pr = el.find(q(P, "spPr"))
        if pr is None:
            pr = el.find(q(P, "grpSpPr"))
        xf = pr.find(q(A, "xfrm")) if pr is not None else None
    if xf is None:
        return None
    off = xf.find(q(A, "off"))
    if off is None or off.get("y") is None:
        return None
    return int(off.get("y")) / EMU


def load_pill(zt):
    """The two shapes of Module 4's POLL pill, as detached elements.
    Matched on the pill's own corner (x >= 11", 6.2 <= y <= 6.8) so the
    footer rule and the gold strip, which are also empty, stay out."""
    parts = slide_parts(zt)
    tree = ET.fromstring(zt.read(parts[TEMPLATE_DISPLAY - 1]))
    spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
    out = []
    for sp in spTree:
        if ET.QName(sp).localname != "sp":
            continue
        pr = sp.find(q(P, "spPr"))
        xf = pr.find(q(A, "xfrm")) if pr is not None else None
        if xf is None or xf.find(q(A, "off")) is None:
            continue
        x = int(xf.find(q(A, "off")).get("x")) / EMU
        y = int(xf.find(q(A, "off")).get("y")) / EMU
        if x >= 11.0 and 6.2 <= y <= 6.8:
            out.append(sp)
    if len(out) != 2:
        raise SystemExit("expected 2 pill shapes in the template, got %d"
                         % len(out))
    return out


def build_placeholder(src_tree, pill_shapes, next_id):
    """A copy of the set-up slide stripped to its chrome, with the POLL
    pill grouped and appended."""
    tree = ET.fromstring(ET.tostring(src_tree))
    for t in tree.findall(q(P, "timing")):
        tree.remove(t)
    spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
    for el in list(spTree):
        tag = ET.QName(el).localname
        if tag not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
            continue
        y = top_inches(el, tag)
        if y is None:                       # layout placeholder - keep
            continue
        if BAND_TOP <= y < BAND_BOTTOM:
            spTree.remove(el)

    # the pill: box + label as ONE object, the way this deck's own Poll
    # Break badge is grouped (Teaching CLAUDE.md: a filled box and the
    # text on it are one object)
    grp = ET.SubElement(spTree, q(P, "grpSp"))
    nv = ET.SubElement(grp, q(P, "nvGrpSpPr"))
    cnv = ET.SubElement(nv, q(P, "cNvPr"))
    cnv.set("id", str(next_id))
    cnv.set("name", "PollPill")
    ET.SubElement(nv, q(P, "cNvGrpSpPr"))
    ET.SubElement(nv, q(P, "nvPr"))
    gpr = ET.SubElement(grp, q(P, "grpSpPr"))
    xf = ET.SubElement(gpr, q(A, "xfrm"))
    box = pill_shapes[0].find(q(P, "spPr")).find(q(A, "xfrm"))
    off_x = int(box.find(q(A, "off")).get("x"))
    off_y = int(box.find(q(A, "off")).get("y"))
    ext_cx = int(box.find(q(A, "ext")).get("cx"))
    ext_cy = int(box.find(q(A, "ext")).get("cy"))
    for tag, a1, v1, a2, v2 in (("off", "x", off_x, "y", off_y),
                                ("ext", "cx", ext_cx, "cy", ext_cy),
                                ("chOff", "x", off_x, "y", off_y),
                                ("chExt", "cx", ext_cx, "cy", ext_cy)):
        e = ET.SubElement(xf, q(A, tag))
        e.set(a1, str(v1))
        e.set(a2, str(v2))
    for i, sp in enumerate(pill_shapes):
        copy = ET.fromstring(ET.tostring(sp))
        c = copy.find(".//" + q(P, "cNvPr"))
        c.set("id", str(next_id + 1 + i))
        grp.append(copy)
    return tree


def main():
    if not DECK.exists():
        raise SystemExit("missing %s" % DECK)
    zin = zipfile.ZipFile(DECK)
    items = {n: zin.read(n) for n in zin.namelist()}
    parts = slide_parts(zin)
    zin.close()

    if any(b"PollPill" in items[p] for p in parts):
        raise SystemExit("PollPill already present - nothing to do")

    # ---- 1. slide 4's picture -------------------------------------
    s4 = parts[3]
    tree4 = ET.fromstring(items[s4])
    n = style_picture(tree4)
    items[s4] = ET.tostring(tree4, xml_declaration=True, encoding="UTF-8",
                            standalone=True)
    print("slide 4: %d picture(s) given rounded corners + drop shadow" % n)

    # ---- 2. the three placeholders --------------------------------
    zt = zipfile.ZipFile(TEMPLATE_DECK)
    pill = load_pill(zt)
    zt.close()
    print("template: Module 4 display %d, pill %.3f x %.3f in"
          % (TEMPLATE_DISPLAY,
             int(pill[0].find(q(P, "spPr")).find(q(A, "xfrm"))
                 .find(q(A, "ext")).get("cx")) / EMU,
             int(pill[0].find(q(P, "spPr")).find(q(A, "xfrm"))
                 .find(q(A, "ext")).get("cy")) / EMU))

    pres = ET.fromstring(items["ppt/presentation.xml"])
    prels = ET.fromstring(items["ppt/_rels/presentation.xml.rels"])
    ctypes = ET.fromstring(items["[Content_Types].xml"])
    sldIdLst = pres.find(q(P, "sldIdLst"))

    used_nums = [int(p.rsplit("slide", 1)[1].split(".")[0])
                 for p in items if p.startswith("ppt/slides/slide")
                 and p.endswith(".xml")]
    next_part = max(used_nums) + 1
    next_sldid = max(int(s.get("id")) for s in sldIdLst) + 1
    next_rid = max(int(r.get("Id")[3:]) for r in prels
                   if r.get("Id").startswith("rId")) + 1

    # descending, so the earlier source positions stay valid
    for disp in sorted(SETUP_DISPLAYS, reverse=True):
        src_part = parts[disp - 1]
        src_tree = ET.fromstring(items[src_part])
        max_id = max(int(c.get("id")) for c in src_tree.iter(q(P, "cNvPr")))
        new_tree = build_placeholder(src_tree, pill, max_id + 1)

        new_part = "ppt/slides/slide%d.xml" % next_part
        items[new_part] = ET.tostring(new_tree, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)

        # rels: the layout only (no notes, no media on a blank slide)
        src_relp = src_part.replace("slides/", "slides/_rels/") + ".rels"
        src_rels = ET.fromstring(items[src_relp])
        # lxml wants the default namespace via nsmap, not register_namespace
        newrels = ET.Element(q(REL, "Relationships"), nsmap={None: REL})
        for r in src_rels:
            if r.get("Type").endswith("/slideLayout"):
                e = ET.SubElement(newrels, q(REL, "Relationship"))
                e.set("Id", "rId1")
                e.set("Type", r.get("Type"))
                e.set("Target", r.get("Target"))
        if not len(newrels):
            raise SystemExit("no slideLayout rel on display %d" % disp)
        items[new_part.replace("slides/", "slides/_rels/") + ".rels"] = \
            ET.tostring(newrels, xml_declaration=True, encoding="UTF-8",
                        standalone=True)

        ov = ET.SubElement(ctypes, q(CT, "Override"))
        ov.set("PartName", "/" + new_part)
        ov.set("ContentType", CT_SLIDE)

        rel = ET.SubElement(prels, q(REL, "Relationship"))
        rel.set("Id", "rId%d" % next_rid)
        rel.set("Type", ("http://schemas.openxmlformats.org/officeDocument/"
                         "2006/relationships/slide"))
        rel.set("Target", "slides/slide%d.xml" % next_part)

        sld = ET.Element(q(P, "sldId"))
        sld.set("id", str(next_sldid))
        sld.set(q(R, "id"), "rId%d" % next_rid)
        sldIdLst[disp - 1].addnext(sld)

        print("  placeholder after display %d  ->  %s (sldId %d, rId%d)"
              % (disp, new_part.split("/")[-1], next_sldid, next_rid))
        next_part += 1
        next_sldid += 1
        next_rid += 1

    items["ppt/presentation.xml"] = ET.tostring(
        pres, xml_declaration=True, encoding="UTF-8", standalone=True)
    items["ppt/_rels/presentation.xml.rels"] = ET.tostring(
        prels, xml_declaration=True, encoding="UTF-8", standalone=True)
    items["[Content_Types].xml"] = ET.tostring(
        ctypes, xml_declaration=True, encoding="UTF-8", standalone=True)

    tmp = DECK.with_suffix(".tmp.pptx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in items.items():
            zout.writestr(name, blob)
    shutil.move(str(tmp), str(DECK))
    print("saved %s" % DECK.name)


def renumber_cache():
    """Refresh the cached text of every footer slide-number field so the
    numbering reads correctly before PowerPoint recomputes the live
    fields (Teaching CLAUDE.md: keep the cached <a:t> current). Inserting
    the three placeholders left every later slide's cache one to three
    behind."""
    zin = zipfile.ZipFile(DECK)
    items = {n: zin.read(n) for n in zin.namelist()}
    parts = slide_parts(zin)
    zin.close()

    fixed = 0
    for disp, part in enumerate(parts, 1):
        tree = ET.fromstring(items[part])
        changed = False
        for fld in tree.iter(q(A, "fld")):
            if fld.get("type") != "slidenum":
                continue
            t = fld.find(q(A, "t"))
            if t is not None and t.text != str(disp):
                t.text = str(disp)
                changed = True
        if changed:
            items[part] = ET.tostring(tree, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            fixed += 1

    tmp = DECK.with_suffix(".tmp.pptx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in items.items():
            zout.writestr(name, blob)
    shutil.move(str(tmp), str(DECK))
    print("refreshed the cached page number on %d slide(s)" % fixed)


if __name__ == "__main__":
    if "--dry-run" in sys.argv:
        raise SystemExit("dry-run: no changes written")
    if "--renumber" in sys.argv:
        renumber_cache()
    else:
        main()
