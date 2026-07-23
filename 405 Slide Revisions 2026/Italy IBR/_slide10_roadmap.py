"""
In-place edit (2026-07-19, round 4): reformat slide 10's roadmap "legend"
to match slide 12's numbered-badge divider style — but with ALL nine sections
shown active (cream band + navy badge + navy bold text), none dimmed, and no
"Section N of 9" tag. OOXML surgery only.
"""

import shutil
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ROADMAP = [
    "(Very) Early History of Italy", "The Roman Empire", "The Dark Ages",
    "Independent City States (Communes)", "The Renaissance",
    "Decline and Napoleon", "Italian Unification", "Fascism and World War II",
    "The Italian Economy Today",
]

# Geometry copied from _divider (identical to slide 12), EMU.
MARGIN, RULE_W = 252059, 11687882
ROW_H, BADGE_D, BAND_H = 512064, 402336, 457200
START_Y = 1627632
BX, TEXT_X, TEXT_W = 544667, 1184747, 10499162
NAVY, WHITE, CREAM, GOLD = "0B2B4E", "FFFFFF", "FDF6E6", "E09F3E"


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


def _band(nid, cy):
    y = cy - BAND_H // 2
    return (f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr>'
            f'<p:cNvPr id="{nid}" name="RmBand"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{MARGIN}" y="{y}"/><a:ext cx="{RULE_W}" cy="{BAND_H}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16000"/></a:avLst></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{CREAM}"/></a:solidFill>'
            f'<a:ln w="19050"><a:solidFill><a:srgbClr val="{GOLD}"/></a:solidFill></a:ln>'
            f'<a:effectLst><a:outerShdw blurRad="40000" dist="25400" dir="2700000" rotWithShape="0">'
            f'<a:srgbClr val="000000"><a:alpha val="22000"/></a:srgbClr></a:outerShdw></a:effectLst>'
            f'</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>')


def _oval(nid, cy):
    y = cy - BADGE_D // 2
    return (f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr>'
            f'<p:cNvPr id="{nid}" name="RmOval"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{BX}" y="{y}"/><a:ext cx="{BADGE_D}" cy="{BADGE_D}"/></a:xfrm>'
            f'<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{NAVY}"/></a:solidFill>'
            f'<a:ln w="22225"><a:solidFill><a:srgbClr val="{GOLD}"/></a:solidFill></a:ln>'
            f'</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>')


def _txt(nid, x, y, cx, cy_, text, size, color, align, anchor="ctr", bold=True):
    b = ' b="1"' if bold else ""
    return (f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr>'
            f'<p:cNvPr id="{nid}" name="RmTxt"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy_}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="{align}"/><a:r><a:rPr lang="en-US" sz="{size}"{b}>'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="Calibri"/></a:rPr><a:t>{escape(text)}</a:t></a:r></a:p></p:txBody></p:sp>')


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    p10 = disp_map(data)[10]
    root = ET.fromstring(data[f"ppt/slides/{p10}"])
    spTree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))

    # remove the old bullet-list textbox (contains the whole roadmap)
    for sp in list(spTree.findall(q(P, "sp"))):
        joined = "".join(t.text or "" for t in sp.iter(q(A, "t")))
        if "The Renaissance" in joined and "Fascism and World War II" in joined:
            spTree.remove(sp)

    ids = [int(e.get("id")) for e in root.iter(q(P, "cNvPr")) if e.get("id")]
    nid = (max(ids) + 1) if ids else 100
    frag = []
    for i, sec in enumerate(ROADMAP):
        cy = START_Y + ROW_H * i + ROW_H // 2
        frag.append(_band(nid, cy)); nid += 1
        frag.append(_oval(nid, cy)); nid += 1
        frag.append(_txt(nid, BX, cy - BADGE_D // 2, BADGE_D, BADGE_D,
                         str(i + 1), 1500, WHITE, "ctr")); nid += 1
        frag.append(_txt(nid, TEXT_X, cy - ROW_H // 2, TEXT_W, ROW_H,
                         sec, 2300, NAVY, "l")); nid += 1
    for x in frag:
        spTree.append(ET.fromstring(x))

    data[f"ppt/slides/{p10}"] = ser(root)
    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print(f"Slide 10 ({p10}) roadmap reformatted to slide-12 badge style (all active)")


if __name__ == "__main__":
    main()
