# -*- coding: utf-8 -*-
"""
Group each colored/filled box with the text box(es) layered on top of it into
one PowerPoint group, so the box and its text move/animate as a single object
(no more "box fades in first, then the text"). Same procedure as the
figure+shade grouping, applied to callout boxes: takeaway bars, cards, etc.
Any existing <p:timing> on a touched slide is dropped (rerun _animate.py after).

Detects: a filled "Rounded Rectangle" (with a solid fill and NO text of its
own) + the text box(es) whose centre falls inside it. Names the group so the
animation engine can place it (TakeawayGroup / CardGroup / BoxGroup).
Skips backup slides (105-135). Pass 'apply' to write; else dry-run.
"""
import os
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0
# Only the animated CONTENT slides (skip title, agenda/dividers, polls, backups —
# their filled badges are navigation chrome, not callouts, and they aren't animated).
SKIP = ({1, 2} | {10, 12, 18, 30, 32, 41, 46, 56, 65, 68, 92}
        | {11, 13, 21, 35, 48, 69, 132} | set(range(105, 136)))


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def box(el):
    xf = el.find(q(P, "xfrm")) or el.find(".//" + q(A, "xfrm"))
    if xf is None:
        return None
    o = xf.find(q(A, "off")); e = xf.find(q(A, "ext"))
    if o is None or e is None:
        return None
    return (int(o.get("x")), int(o.get("y")), int(e.get("cx")), int(e.get("cy")))


def has_text(el):
    return any((t.text or "").strip() for t in el.iter(q(A, "t")))


def is_chrome(name, b):
    if name in ("TextBox 2", "TextBox 3"):
        return True
    if name.startswith("Rectangle "):
        return True
    return False


def centre_in(inner, outer, tol=0.15 * EMU):
    ix, iy, iw, ih = inner
    ox, oy, ow, oh = outer
    cx, cy = ix + iw / 2, iy + ih / 2
    return (ox - tol) <= cx <= (ox + ow + tol) and (oy - tol) <= cy <= (oy + oh + tol)


def classify(b):
    x, y, w, h = [v / EMU for v in b]
    if y > 5.8:                       # bottom bar (widths vary 8.0"-12.7")
        return "TakeawayGroup"
    if 4.5 < w < 7.5:                 # body column card
        return "CardGroup"
    return "BoxGroup"


def main(apply):
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target")
             for r in ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    order = [os.path.basename(rid2t[s.get(q(R, "id"))]) for s in pres.find(q(P, "sldIdLst"))]

    report = []
    for disp, part in enumerate(order, 1):
        if disp in SKIP:
            continue
        root = ET.fromstring(data[f"ppt/slides/{part}"])
        tree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
        sps = [el for el in tree if ET.QName(el).localname == "sp"]

        def name_of(el):
            return el.find(".//" + q(P, "cNvPr")).get("name") or ""

        # filled ROUNDED-rect boxes (by geometry, not name) with a solid fill and
        # no text of their own. geom="roundRect" excludes the flat chrome rects.
        boxes = []
        for el in sps:
            spPr = el.find(q(P, "spPr"))
            if spPr is None or spPr.find(q(A, "solidFill")) is None:
                continue
            geom = spPr.find(q(A, "prstGeom"))
            if geom is None or geom.get("prst") != "roundRect":
                continue
            if has_text(el):
                continue
            boxes.append(el)

        pairs = []
        for bx in boxes:
            bb = box(bx)
            members = []
            for el in sps:
                if el is bx:
                    continue
                nm = name_of(el)
                if not has_text(el) or is_chrome(nm, box(el)):
                    continue
                if centre_in(box(el), bb):
                    members.append(el)
            if members:
                pairs.append((bx, members))

        if not pairs:
            continue
        names = []
        max_id = max(int(c.get("id")) for c in root.iter(q(P, "cNvPr")))
        for bx, members in pairs:
            allb = [box(bx)] + [box(m) for m in members]
            x = min(b[0] for b in allb); y = min(b[1] for b in allb)
            w = max(b[0] + b[2] for b in allb) - x
            h = max(b[1] + b[3] for b in allb) - y
            kind = classify(box(bx))
            names.append(kind)
            if apply:
                max_id += 1
                grp = ET.fromstring(
                    f'<p:grpSp xmlns:a="{A}" xmlns:p="{P}"><p:nvGrpSpPr>'
                    f'<p:cNvPr id="{max_id}" name="{kind}"/><p:cNvGrpSpPr/><p:nvPr/>'
                    f'</p:nvGrpSpPr><p:grpSpPr><a:xfrm>'
                    f'<a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/>'
                    f'<a:chOff x="{x}" y="{y}"/><a:chExt cx="{w}" cy="{h}"/>'
                    f'</a:xfrm></p:grpSpPr></p:grpSp>')
                idx = min([list(tree).index(bx)] + [list(tree).index(m) for m in members])
                tree.remove(bx)
                for m in members:
                    tree.remove(m)
                grp.append(bx)
                for m in members:
                    grp.append(m)
                tree.insert(idx, grp)
        if apply:
            for tm in root.findall(q(P, "timing")):
                root.remove(tm)
            data[f"ppt/slides/{part}"] = ser(root)
        report.append((disp, names))

    for disp, names in report:
        print(f"slide {disp}: {', '.join(names)}")
    print(f"== {len(report)} slides, {sum(len(n) for _,n in report)} box groups ==")

    if apply:
        tmp = DECK.with_suffix(".pptx.tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for name, blob in data.items():
                out.writestr(name, blob)
        with zipfile.ZipFile(tmp) as chk:
            assert chk.testzip() is None
        shutil.move(str(tmp), str(DECK))
        print("APPLIED")
    else:
        print("DRY RUN (pass 'apply' to write)")


if __name__ == "__main__":
    main(len(sys.argv) > 1 and sys.argv[1] == "apply")
