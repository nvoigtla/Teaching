"""
SAMPLE (2026-07-19, round 6): restore exact original wording + references on
slide 14, and rework its three backup picture targets (101 Matera, 102/103
Valcamonica) to full-size images with a new-style "← Back" button above them.
If approved, the same approach rolls out to all content + backup slides.

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
RT_SLIDE = f"{R}/slide"
NAVY = "0B2B4E"

LVL_SIZE = {0: 2400, 1: 2200, 2: 2000}
LVL_MARL = {0: 342900, 1: 731520, 2: 1097280}
LVL_CHAR = {0: "▪", 1: "–", 2: "·"}


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
    return {i + 1: p for i, p in enumerate(order)}, order


class Rels:
    """Manage a slide's .rels: hand out rIds for slide-jumps / external URLs."""
    def __init__(self, data, part, order):
        self.key = f"ppt/slides/_rels/{part}.rels"
        self.el = ET.fromstring(data[self.key])
        self.order = order

    def _new_id(self):
        used = {r.get("Id") for r in self.el}
        n = 1
        while f"rId{n}" in used:
            n += 1
        return f"rId{n}"

    def slide(self, disp):
        tgt = f"../slides/{self.order[disp - 1]}"
        for r in self.el:
            if r.get("Type") == RT_SLIDE and r.get("Target") == tgt:
                return r.get("Id")
        rid = self._new_id()
        rel = ET.SubElement(self.el, q(PKG, "Relationship"))
        rel.set("Id", rid); rel.set("Type", RT_SLIDE); rel.set("Target", tgt)
        return rid

    def url(self, u):
        rid = self._new_id()
        rel = ET.SubElement(self.el, q(PKG, "Relationship"))
        rel.set("Id", rid); rel.set("Type", f"{R}/hyperlink")
        rel.set("Target", u); rel.set("TargetMode", "External")
        return rid

    def save(self, data):
        data[self.key] = ser(self.el)


def para_xml(level, runs, first, rels):
    """runs: list of (text, target) where target is None | int(display) | url."""
    size = LVL_SIZE[level]
    spc = "" if first else (f'<a:spcBef><a:spcPts val="{1600 if level == 0 else 500}"/></a:spcBef>')
    lvl = f' lvl="{level}"' if level > 0 else ""
    pPr = (f'<a:pPr marL="{LVL_MARL[level]}" indent="-274320"{lvl}>{spc}'
           f'<a:buClr><a:srgbClr val="{NAVY}"/></a:buClr>'
           f'<a:buFont typeface="Calibri"/><a:buChar char="{LVL_CHAR[level]}"/></a:pPr>')
    rs = ""
    for text, tgt in runs:
        u = ' u="sng"' if tgt is not None else ""
        hl = ""
        if isinstance(tgt, int):
            hl = f'<a:hlinkClick r:id="{rels.slide(tgt)}" action="ppaction://hlinksldjump"/>'
        elif isinstance(tgt, str):
            hl = f'<a:hlinkClick r:id="{rels.url(tgt)}"/>'
        rs += (f'<a:r><a:rPr lang="en-US" sz="{size}"{u}>'
               f'<a:solidFill><a:srgbClr val="{NAVY}"/></a:solidFill>'
               f'<a:latin typeface="Calibri"/>{hl}</a:rPr>'
               f'<a:t>{escape(text)}</a:t></a:r>')
    return f'<a:p xmlns:a="{A}" xmlns:r="{R}">{pPr}{rs}</a:p>'


def replace_body(root, items, rels, needle):
    """Find the bullet text box (joined text contains needle) and rewrite it."""
    for sp in root.iter(q(P, "sp")):
        joined = "".join(t.text or "" for t in sp.iter(q(A, "t")))
        if needle in joined:
            tb = sp.find(q(P, "txBody"))
            for pp in tb.findall(q(A, "p")):
                tb.remove(pp)
            for lvl, content in items:
                runs = [(content, None)] if isinstance(content, str) else content
                first = (lvl, content) == items[0]
                tb.append(ET.fromstring(para_xml(lvl, runs, first, rels)))
            return True
    return False


def back_button(new_id, rid, x=252059, y=468630):
    return ET.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:nvSpPr>'
        f'<p:cNvPr id="{new_id}" name="BackButton">'
        f'<a:hlinkClick r:id="{rid}" action="ppaction://hlinksldjump"/></p:cNvPr>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="1417320" cy="420624"/></a:xfrm>'
        f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 30000"/></a:avLst></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{NAVY}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>'
        f'<a:p><a:pPr algn="ctr"/><a:r>'
        f'<a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        f'<a:latin typeface="Calibri"/></a:rPr><a:t>← Back</a:t></a:r></a:p></p:txBody></p:sp>')


def rework_backup(data, order, disp, back_to):
    """Enlarge the image to near-full-slide; drop title/rule; add back button."""
    part = order[disp - 1]
    root = ET.fromstring(data[f"ppt/slides/{part}"])
    spTree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
    # remove title + rule + gold strip (shapes between the top bar and image)
    for sp in list(spTree.findall(q(P, "sp"))):
        off = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "off"))
        if off is not None and 500000 < int(off.get("y")) < 1300000:
            spTree.remove(sp)
    # enlarge the picture, preserving aspect, into a near-full region
    X0, Y0, X1, Y1 = 137160, 950000, 12054720, 6446520
    for pic in root.iter(q(P, "pic")):
        ext = pic.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "ext"))
        off = pic.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "off"))
        w, h = int(ext.get("cx")), int(ext.get("cy"))
        s = min((X1 - X0) / w, (Y1 - Y0) / h)
        nw, nh = int(w * s), int(h * s)
        off.set("x", str(int(X0 + (X1 - X0 - nw) / 2)))
        off.set("y", str(int(Y0 + (Y1 - Y0 - nh) / 2)))
        ext.set("cx", str(nw)); ext.set("cy", str(nh))
    # back button -> source slide
    rels = Rels(data, part, order)
    rid = rels.slide(back_to)
    ids = [int(e.get("id")) for e in root.iter(q(P, "cNvPr")) if e.get("id")]
    spTree.append(back_button((max(ids) + 1) if ids else 100, rid))
    rels.save(data)
    data[f"ppt/slides/{part}"] = ser(root)


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    d2p, order = disp_map(data)

    # ---- slide 14: exact original wording + references ----
    ITEMS = [
        (0, "Many different cultures across the peninsula"),
        (0, "City of Matera in Southern Italy"),
        (1, "Settlement since 10,000BC"),
        (1, [("Famous “Sassi” (", None), ("ancient town", 101), (")", None)]),
        (0, "Rock Drawings in Valcamonica"),
        (1, "People of Camuni"),
        (1, "From around 8,000BC-1,000BC"),
        (1, "Largest collections of prehistoric petroglyphs in the world"),
        (1, [("Some famous pieces such as “", None), ("The Astronauts", 102),
             ("” with the ", None), ("Camunian Rose", 103)]),
    ]
    p14 = d2p[14]
    root14 = ET.fromstring(data[f"ppt/slides/{p14}"])
    rels14 = Rels(data, p14, order)
    ok = replace_body(root14, ITEMS, rels14, "peninsula")
    rels14.save(data)
    data[f"ppt/slides/{p14}"] = ser(root14)
    print("slide 14 body replaced:", ok)

    # ---- backup targets: big image + back button ----
    for disp in (101, 102, 103):
        rework_backup(data, order, disp, 14)
    print("backup slides 101,102,103 reworked")

    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print("sample applied")


if __name__ == "__main__":
    main()
