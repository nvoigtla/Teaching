"""
In-place edits to the canonical "Class 1 - Revised.pptx" (2026-07-19).

The user hand-edited the deck in PowerPoint (TA email on slide 4; transport
rewording + NH Milano Touring + two hotel images on slide 6; pasted schedule
image on slide 99). We therefore DO NOT rebuild — we do OOXML surgery in place
so those hand-edits are preserved, and python-pptx is avoided (it would strip
the 7 PollEverywhere <p:tags> rels).

Changes:
  1. Slide 4: remove the erroneous slide-jump on the TITLE run; make the bullet
     word "Schedule" jump to slide 99 (the pasted schedule).
  2. Slide 99: add a "← Back" button that jumps to slide 4.
  3. Slide 5: refresh transport prices (Malpensa Express €13→€15, ~50 min;
     taxi €104→€110). Coach €10 unchanged. (Verified July 2026.)
  4. Slide 6: give the two hand-added hotel photos rounded corners + soft shadow.
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
NAVY, WHITE = "0B2B4E", "FFFFFF"


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


def add_shadow(spPr):
    for el in spPr.findall(q(A, "effectLst")):
        spPr.remove(el)
    lst = ET.SubElement(spPr, q(A, "effectLst"))
    sh = ET.SubElement(lst, q(A, "outerShdw"))
    sh.set("blurRad", "50800"); sh.set("dist", "38100")
    sh.set("dir", "2700000"); sh.set("rotWithShape", "0")
    c = ET.SubElement(sh, q(A, "srgbClr")); c.set("val", "000000")
    ET.SubElement(c, q(A, "alpha")).set("val", "32000")


def round_and_shadow(spPr):
    geom = spPr.find(q(A, "prstGeom"))
    if geom is None:
        geom = ET.SubElement(spPr, q(A, "prstGeom"))
    geom.set("prst", "roundRect")
    av = geom.find(q(A, "avLst"))
    if av is None:
        av = ET.SubElement(geom, q(A, "avLst"))
    for gd in av.findall(q(A, "gd")):
        av.remove(gd)
    gd = ET.SubElement(av, q(A, "gd"))
    gd.set("name", "adj"); gd.set("fmla", "val 5500")
    add_shadow(spPr)


def back_button_sp(new_id, rid):
    x, y, cx, cy = 10561320, 548640, 1417320, 420624   # top-right, EMU
    return ET.fromstring(f"""<p:sp xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">
  <p:nvSpPr>
    <p:cNvPr id="{new_id}" name="BackButton">
      <a:hlinkClick r:id="{rid}" action="ppaction://hlinksldjump"/>
    </p:cNvPr>
    <p:cNvSpPr/><p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 30000"/></a:avLst></a:prstGeom>
    <a:solidFill><a:srgbClr val="{NAVY}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r>
      <a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="{WHITE}"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr>
      <a:t>← Back</a:t>
    </a:r></a:p>
  </p:txBody>
</p:sp>""")


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    d2p = disp_map(data)
    p4, p5, p6, p99 = d2p[4], d2p[5], d2p[6], d2p[99]

    # ---- 1. Slide 4 link fix ------------------------------------------------
    rid99, rk, rels4 = ensure_slide_rel(data, p4, p99)
    root4 = ET.fromstring(data[f"ppt/slides/{p4}"])
    for r in root4.iter(q(A, "r")):
        t = r.find(q(A, "t"))
        if t is None or not t.text:
            continue
        rPr = r.find(q(A, "rPr"))
        if t.text == "Logistics – Schedule" and rPr is not None:
            for hl in rPr.findall(q(A, "hlinkClick")):
                rPr.remove(hl)                       # drop erroneous title link
        elif t.text == "Schedule" and rPr is not None:
            for hl in rPr.findall(q(A, "hlinkClick")):
                rPr.remove(hl)                       # drop old syllabus-PDF link
            hl = ET.Element(q(A, "hlinkClick"))
            hl.set(q(R, "id"), rid99)
            hl.set("action", "ppaction://hlinksldjump")
            rPr.insert(0, hl)
    data[f"ppt/slides/{p4}"] = ser(root4)
    data[rk] = ser(rels4)

    # ---- 2. Slide 99 back button -------------------------------------------
    rid4, rk9, rels9 = ensure_slide_rel(data, p99, p4)
    root99 = ET.fromstring(data[f"ppt/slides/{p99}"])
    ids = [int(e.get("id")) for e in root99.iter(q(P, "cNvPr")) if e.get("id")]
    new_id = (max(ids) + 1) if ids else 100
    spTree = root99.find(q(P, "cSld") + "/" + q(P, "spTree"))
    spTree.append(back_button_sp(new_id, rid4))
    data[f"ppt/slides/{p99}"] = ser(root99)
    data[rk9] = ser(rels9)

    # ---- 3. Slide 5 prices --------------------------------------------------
    root5 = ET.fromstring(data[f"ppt/slides/{p5}"])
    for t in root5.iter(q(A, "t")):
        if t.text and "54 min / €13" in t.text:
            t.text = t.text.replace("54 min / €13", "~50 min / €15")
        if t.text and "€104" in t.text:
            t.text = t.text.replace("€104", "€110")
    data[f"ppt/slides/{p5}"] = ser(root5)

    # ---- 4. Slide 6 hotel photos -------------------------------------------
    root6 = ET.fromstring(data[f"ppt/slides/{p6}"])
    npic = 0
    for pic in root6.iter(q(P, "pic")):
        spPr = pic.find(q(P, "spPr"))
        if spPr is not None:
            round_and_shadow(spPr)
            npic += 1
    data[f"ppt/slides/{p6}"] = ser(root6)

    # ---- write out ----------------------------------------------------------
    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print(f"Applied edits: slide4 link->99, slide99 back button, "
          f"slide5 prices, slide6 photos styled ({npic})")


if __name__ == "__main__":
    main()
