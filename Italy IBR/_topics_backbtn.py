"""
In-place edits (2026-07-19, round 3) to "Class 1 - Revised.pptx", slide 98
(the moved Presentation-Topics table):
  1. Tighten the table's data-row heights so the table ends higher, leaving
     space underneath; shrink the shadow backing to match.
  2. Add a "← Back" button (identical to slide 97's) linking back to slide 8.

OOXML surgery only — no rebuild, no python-pptx round-trip.
"""

import shutil
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
RT_SLIDE = f"{R}/slide"

HEADER_H = 384048
NEW_ROW_H = 411480                       # was 512064 (~0.56") -> ~0.45"
NEW_CY = HEADER_H + 9 * NEW_ROW_H        # = 4087368
# Back-button geometry copied verbatim from slide 97:
BB_X, BB_Y, BB_CX, BB_CY = 10721340, 6035040, 1417320, 420624


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def disp_map(data):
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rels = ET.fromstring(data["ppt/_rels/presentation.xml.rels"])
    rid2t = {r.get("Id"): r.get("Target") for r in rels}
    order = [rid2t[s.get(q(R, "id"))].split("/")[-1]
             for s in pres.find(q(P, "sldIdLst"))]
    return {i + 1: p for i, p in enumerate(order)}


def ensure_slide_rel(data, src_part, tgt_part):
    key = f"ppt/slides/_rels/{src_part}.rels"
    rels = ET.fromstring(data[key])
    tgt = f"../slides/{tgt_part}"
    for r in rels:
        if r.get("Type") == RT_SLIDE and r.get("Target") == tgt:
            return r.get("Id"), key, rels
    used = {r.get("Id") for r in rels}
    n = 1
    while f"rId{n}" in used:
        n += 1
    rid = f"rId{n}"
    rel = ET.SubElement(rels, q(PKG, "Relationship"))
    rel.set("Id", rid); rel.set("Type", RT_SLIDE); rel.set("Target", tgt)
    return rid, key, rels


def back_button_sp(new_id, rid):
    return ET.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">'
        f'<p:nvSpPr><p:cNvPr id="{new_id}" name="BackButton">'
        f'<a:hlinkClick r:id="{rid}" action="ppaction://hlinksldjump"/></p:cNvPr>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{BB_X}" y="{BB_Y}"/><a:ext cx="{BB_CX}" cy="{BB_CY}"/></a:xfrm>'
        f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 30000"/></a:avLst></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="0B2B4E"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>'
        f'<a:p><a:pPr algn="ctr"/><a:r>'
        f'<a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        f'<a:latin typeface="Calibri"/></a:rPr><a:t>← Back</a:t></a:r></a:p></p:txBody></p:sp>')


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    d2p = disp_map(data)
    p98, p8 = d2p[98], d2p[8]

    root = ET.fromstring(data[f"ppt/slides/{p98}"])

    # 1a. tighten table data rows + shrink graphicFrame height
    for gf in root.iter(q(P, "graphicFrame")):
        name = gf.find(f".//{q(P, 'cNvPr')}")
        if name is None or name.get("name") != "TopicsTable":
            continue
        tbl = gf.find(f".//{q(A, 'tbl')}")
        for i, tr in enumerate(tbl.findall(q(A, "tr"))):
            if i > 0:
                tr.set("h", str(NEW_ROW_H))
        ext = gf.find(q(P, "xfrm") + "/" + q(A, "ext"))
        ext.set("cy", str(NEW_CY))

    # 1b. shrink the shadow backing to match
    for sp in root.iter(q(P, "sp")):
        nm = sp.find(f".//{q(P, 'cNvPr')}")
        if nm is not None and nm.get("name") == "TableBacking":
            ext = sp.find(f".//{q(A, 'ext')}")
            ext.set("cy", str(NEW_CY))

    # 2. add the back button -> slide 8
    rid, rk, rels = ensure_slide_rel(data, p98, p8)
    ids = [int(e.get("id")) for e in root.iter(q(P, "cNvPr")) if e.get("id")]
    new_id = (max(ids) + 1) if ids else 100
    spTree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
    spTree.append(back_button_sp(new_id, rid))

    data[f"ppt/slides/{p98}"] = ser(root)
    data[rk] = ser(rels)

    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print(f"Slide 98: rows tightened (cy={NEW_CY}); back button -> slide 8 added")


if __name__ == "__main__":
    main()
