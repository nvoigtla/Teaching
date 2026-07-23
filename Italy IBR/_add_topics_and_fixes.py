"""
In-place edits (2026-07-19, round 2) to "Class 1 - Revised.pptx":
  1. Slide 5: taxi flat rate €110 -> €114 (official Malpensa site).
  2. Slide 96: add a native table of the 9 debate topics from
     "Presentation Topics -- Italy.xlsx" (Team A vs. Team B).

OOXML surgery only — no rebuild, no python-pptx round-trip (preserves the
7 PollEverywhere <p:tags> and all hand-edits).
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
NAVY, WHITE, CREAM = "0B2B4E", "FFFFFF", "FDF6E6"
EMU = 914400


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


HEADER = ["Topic", "Team A", "Team B"]
TOPICS = [
    ("1   Will Southern Italy catch up?", "Yes – it will converge with the North", "No – it will stay behind"),
    ("2   Italy’s future in tech & AI", "Italy is lagging far behind", "Italy is already on a good path"),
    ("3   Corruption and trust", "A serious problem, with real economic cost", "Its reputation is overblown"),
    ("4   Italy and the EU / euro", "The EU and euro are good for Italy", "Italy would be better off outside"),
    ("5   Immigration and growth", "Italy should encourage more immigration", "More immigration is a net negative"),
    ("6   Lavazza in the U.S.", "Teach Americans the Italian way of espresso", "Build a new brand for U.S. tastes"),
    ("7   Family businesses", "A major competitive advantage", "They hold Italian companies back"),
    ("8   Tourism", "Italy should keep expanding tourism", "Italy already depends on it too much"),
    ("Alt.   Pick a company: invest or not?", "It has a bright future – invest", "Be cautious about investing"),
]
COLW = [int(3.2 * EMU), int(4.75 * EMU), int(4.75 * EMU)]
X = int(0.33 * EMU)
Y = int(1.55 * EMU)
H_HEAD = int(0.42 * EMU)
H_ROW = int(0.56 * EMU)


def _cell(text, fill, color, bold, size, align="l"):
    b = ' b="1"' if bold else ""
    return (f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>'
            f'<a:p><a:pPr algn="{align}"/><a:r>'
            f'<a:rPr lang="en-US" sz="{size}"{b}>'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="Calibri"/></a:rPr>'
            f'<a:t>{escape(text)}</a:t></a:r></a:p></a:txBody>'
            f'<a:tcPr marL="82296" marR="82296" marT="27432" marB="27432" anchor="ctr">'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></a:tcPr></a:tc>')


def _table_gf(new_id):
    grid = "".join(f'<a:gridCol w="{w}"/>' for w in COLW)
    rows = [f'<a:tr h="{H_HEAD}">'
            + "".join(_cell(HEADER[i], NAVY, WHITE, True, 1500) for i in range(3))
            + "</a:tr>"]
    for r, (topic, a, b) in enumerate(TOPICS):
        fill = WHITE if r % 2 == 0 else CREAM
        rows.append(
            f'<a:tr h="{H_ROW}">'
            + _cell(topic, fill, NAVY, True, 1400)
            + _cell(a, fill, NAVY, False, 1300)
            + _cell(b, fill, NAVY, False, 1300)
            + "</a:tr>")
    cx = sum(COLW)
    cy = H_HEAD + H_ROW * len(TOPICS)
    xml = (
        f'<p:graphicFrame xmlns:p="{P}" xmlns:a="{A}">'
        f'<p:nvGraphicFramePr><p:cNvPr id="{new_id}" name="TopicsTable"/>'
        f'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
        f'<p:nvPr/></p:nvGraphicFramePr>'
        f'<p:xfrm><a:off x="{X}" y="{Y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm>'
        f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
        f'<a:tbl><a:tblPr firstRow="1" bandRow="0"/><a:tblGrid>{grid}</a:tblGrid>'
        f'{"".join(rows)}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>')
    return ET.fromstring(xml), cx, cy


def _backing(new_id, cx, cy):
    xml = (
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}">'
        f'<p:nvSpPr><p:cNvPr id="{new_id}" name="TableBacking"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{X}" y="{Y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{WHITE}"/></a:solidFill><a:ln><a:noFill/></a:ln>'
        f'<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" rotWithShape="0">'
        f'<a:srgbClr val="000000"><a:alpha val="30000"/></a:srgbClr></a:outerShdw></a:effectLst>'
        f'</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>')
    return ET.fromstring(xml)


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    d2p = disp_map(data)

    # 1. taxi price
    root5 = ET.fromstring(data[f"ppt/slides/{d2p[5]}"])
    for t in root5.iter(q(A, "t")):
        if t.text and "€110" in t.text:
            t.text = t.text.replace("€110", "€114")
    data[f"ppt/slides/{d2p[5]}"] = ser(root5)

    # 2. topics table on slide 96
    p96 = d2p[96]
    root96 = ET.fromstring(data[f"ppt/slides/{p96}"])
    ids = [int(e.get("id")) for e in root96.iter(q(P, "cNvPr")) if e.get("id")]
    base_id = (max(ids) + 1) if ids else 100
    spTree = root96.find(q(P, "cSld") + "/" + q(P, "spTree"))
    gf, cx, cy = _table_gf(base_id + 1)
    spTree.append(_backing(base_id, cx, cy))   # behind
    spTree.append(gf)                          # table on top
    data[f"ppt/slides/{p96}"] = ser(root96)

    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print(f"Done: taxi -> €114; topics table added to slide 96 ({p96})")


if __name__ == "__main__":
    main()
