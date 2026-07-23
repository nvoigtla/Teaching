# -*- coding: utf-8 -*-
"""
Group each table/chart with its separate shadow-backing rectangle into one
PowerPoint group, so the shade always travels (and animates) with the figure.
Auto-detects: a plain rect with an outer shadow and no text, overlapping a
graphicFrame -> group [backing (behind) + figure (front)].
Any existing <p:timing> on a touched slide is dropped (rerun _animate.py after).
"""
import os
import shutil
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def box(el):
    xf = el.find(q(P, "xfrm"))          # graphicFrame: p:xfrm (direct child)
    if xf is None:
        xf = el.find(".//" + q(A, "xfrm"))  # sp / pic: a:xfrm under spPr
    if xf is None:
        return None
    o = xf.find(q(A, "off")); e = xf.find(q(A, "ext"))
    if o is None or e is None:
        return None
    return (int(o.get("x")), int(o.get("y")), int(e.get("cx")), int(e.get("cy")))


def overlap(a, b):
    ax, ay, aw, ah = a; bx, by, bw, bh = b
    ix = max(0, min(ax + aw, bx + bw) - max(ax, bx))
    iy = max(0, min(ay + ah, by + bh) - max(ay, by))
    return ix * iy


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target")
             for r in ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    order = [os.path.basename(rid2t[s.get(q(R, "id"))]) for s in pres.find(q(P, "sldIdLst"))]

    touched = []
    for disp, part in enumerate(order, 1):
        root = ET.fromstring(data[f"ppt/slides/{part}"])
        tree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
        children = list(tree)
        gfs = [el for el in children if ET.QName(el).localname == "graphicFrame"]
        if not gfs:
            continue
        backings = []
        for el in children:
            if ET.QName(el).localname != "sp":
                continue
            geom = el.find(".//" + q(A, "prstGeom"))
            prst = geom.get("prst") if geom is not None else ""
            has_shadow = el.find(".//" + q(A, "outerShdw")) is not None
            txt = "".join(t.text or "" for t in el.iter(q(A, "t"))).strip()
            if prst == "rect" and has_shadow and not txt:
                backings.append(el)
        pairs = []
        used_gf = set()
        for bk in backings:
            bb = box(bk)
            best, ba = None, 0
            for gf in gfs:
                if id(gf) in used_gf:
                    continue
                ov = overlap(bb, box(gf))
                if ov > ba:
                    best, ba = gf, ov
            if best is not None and ba > 0:
                pairs.append((bk, best))
                used_gf.add(id(best))
        if not pairs:
            continue

        max_id = max(int(c.get("id")) for c in root.iter(q(P, "cNvPr")))
        for bk, gf in pairs:
            bx0, by0, bw0, bh0 = box(bk)
            gx0, gy0, gw0, gh0 = box(gf)
            x = min(bx0, gx0); y = min(by0, gy0)
            w = max(bx0 + bw0, gx0 + gw0) - x
            h = max(by0 + bh0, gy0 + gh0) - y
            max_id += 1
            gname = (gf.find(".//" + q(P, "cNvPr")).get("name") or "Figure")
            grp = ET.fromstring(
                f'<p:grpSp xmlns:a="{A}" xmlns:p="{P}"><p:nvGrpSpPr>'
                f'<p:cNvPr id="{max_id}" name="{gname} + shade"/><p:cNvGrpSpPr/><p:nvPr/>'
                f'</p:nvGrpSpPr><p:grpSpPr><a:xfrm>'
                f'<a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/>'
                f'<a:chOff x="{x}" y="{y}"/><a:chExt cx="{w}" cy="{h}"/>'
                f'</a:xfrm></p:grpSpPr></p:grpSp>')
            idx = min(list(tree).index(bk), list(tree).index(gf))
            tree.remove(bk); tree.remove(gf)
            grp.append(bk)   # backing first (behind)
            grp.append(gf)   # figure second (front)
            tree.insert(idx, grp)

        # drop stale timing (animations rerun afterward)
        for tm in root.findall(q(P, "timing")):
            root.remove(tm)
        data[f"ppt/slides/{part}"] = ser(root)
        touched.append((disp, len(pairs)))

    for disp, n in touched:
        print(f"slide {disp}: grouped {n} figure(s) with shade")

    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print("done")


if __name__ == "__main__":
    main()
