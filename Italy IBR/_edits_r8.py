"""
Edits (2026-07-20, round 8):
  1. Move back buttons on backup slides 101,102,103,104,105,109 to the
     lower-right corner (matching slides 97/98).
  2. Slide 15: replace the "New Museum" body bullet with a bordered link box
     in the bottom-right ("Museum of Etruscan history in Milan", same URL).
  3. Featured-research slides (32,46,72,82,88): "Featured research" 40pt,
     title 28pt, authors 20pt (+ reposition the banner/title).
OOXML surgery only.
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
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NAVY, WHITE, GOLD = "0B2B4E", "FFFFFF", "E09F3E"
BB_X, BB_Y = 10721340, 6035040
MUSEUM_URL = "https://www.fondazioneluigirovati.org/en/"

BACK_MOVE = [101, 102, 103, 104, 105, 109]
FEATURED = [32, 46, 72, 82, 88]


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def order_of(data):
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target")
             for r in ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    return [rid2t[s.get(q(R, "id"))].split("/")[-1]
            for s in pres.find(q(P, "sldIdLst"))]


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    order = order_of(data)

    # ---- 1. move back buttons ----
    for disp in BACK_MOVE:
        base = order[disp - 1]
        root = ET.fromstring(data[f"ppt/slides/{base}"])
        for sp in root.iter(q(P, "sp")):
            cnv = sp.find(".//" + q(P, "cNvPr"))
            if cnv is not None and cnv.get("name") == "BackButton":
                off = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "off"))
                off.set("x", str(BB_X)); off.set("y", str(BB_Y))
        data[f"ppt/slides/{base}"] = ser(root)
    print(f"back buttons moved on {BACK_MOVE}")

    # ---- 2. slide 15 museum link box ----
    base = order[15 - 1]
    rk = f"ppt/slides/_rels/{base}.rels"
    root = ET.fromstring(data[f"ppt/slides/{base}"])
    spTree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
    # remove the New Museum body bullet
    for sp in root.iter(q(P, "sp")):
        tb = sp.find(q(P, "txBody"))
        if tb is None:
            continue
        for pp in tb.findall(q(A, "p")):
            txt = "".join(t.text or "" for t in pp.iter(q(A, "t")))
            if "Museum on Etruscan history in Milan" in txt:
                tb.remove(pp)
    # external rel for the URL
    rels = ET.fromstring(data[rk])
    used = {r.get("Id") for r in rels}
    n = 1
    while f"rId{n}" in used:
        n += 1
    rid = f"rId{n}"
    rel = ET.SubElement(rels, q(PKG, "Relationship"))
    rel.set("Id", rid); rel.set("Type", f"{R}/hyperlink")
    rel.set("Target", MUSEUM_URL); rel.set("TargetMode", "External")
    data[rk] = ser(rels)
    # bordered link box, bottom-right (shape-level hyperlink)
    ids = [int(e.get("id")) for e in root.iter(q(P, "cNvPr")) if e.get("id")]
    nid = (max(ids) + 1) if ids else 200
    bw, bh = 3657600, 420624                              # 4.0" x 0.46"
    bx, by = 12192000 - 252059 - bw, 6035040             # right-aligned, y6.6"
    box = (f'<p:sp xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:nvSpPr>'
           f'<p:cNvPr id="{nid}" name="MuseumLink">'
           f'<a:hlinkClick r:id="{rid}"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
           f'<p:spPr><a:xfrm><a:off x="{bx}" y="{by}"/><a:ext cx="{bw}" cy="{bh}"/></a:xfrm>'
           f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 18000"/></a:avLst></a:prstGeom>'
           f'<a:solidFill><a:srgbClr val="{WHITE}"/></a:solidFill>'
           f'<a:ln w="15875"><a:solidFill><a:srgbClr val="{GOLD}"/></a:solidFill></a:ln></p:spPr>'
           f'<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>'
           f'<a:r><a:rPr lang="en-US" sz="1500" u="sng"><a:solidFill><a:srgbClr val="{NAVY}"/></a:solidFill>'
           f'<a:latin typeface="Calibri"/><a:hlinkClick r:id="{rid}"/></a:rPr>'
           f'<a:t>{escape("Museum of Etruscan history in Milan")}</a:t></a:r></a:p></p:txBody></p:sp>')
    spTree.append(ET.fromstring(box))
    data[f"ppt/slides/{base}"] = ser(root)
    print("slide 15: museum link box added, old bullet removed")

    # ---- 3. featured-research slides ----
    for disp in FEATURED:
        base = order[disp - 1]
        root = ET.fromstring(data[f"ppt/slides/{base}"])
        for sp in root.iter(q(P, "sp")):
            tb = sp.find(q(P, "txBody"))
            if tb is None:
                continue
            joined = "".join(t.text or "" for t in sp.iter(q(A, "t")))
            runs = sp.findall(q(P, "txBody") + "/" + q(A, "p") + "/" + q(A, "r"))
            off = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "off"))
            ext = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "ext"))
            if joined.strip() == "Featured research":
                for r in runs:
                    r.find(q(A, "rPr")).set("sz", "4000")
                if off is not None:
                    off.set("y", "1645920")           # 1.8"
                if ext is not None:
                    ext.set("cy", "731520")           # 0.8"
            else:
                for r in runs:
                    rPr = r.find(q(A, "rPr"))
                    if rPr is None:
                        continue
                    sz = rPr.get("sz")
                    if sz == "3400":                  # title
                        rPr.set("sz", "2800")
                        if off is not None:
                            off.set("y", "2606040")   # 2.85"
                    elif sz == "2200":                # authors
                        rPr.set("sz", "2000")
        data[f"ppt/slides/{base}"] = ser(root)
    print(f"featured-research resized on {FEATURED}")

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
