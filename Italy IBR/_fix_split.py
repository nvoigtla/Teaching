"""
Fix (2026-07-19, round 7b):
  - Slides 25, 44, 52, 60, 61: their original body was split across TWO text
    boxes; re-extract by concatenating both (ordered top-to-bottom).
  - Slide 15: remove the leftover standalone "New Museum" caption now that the
    restored body bullet already carries it.
OOXML surgery only.
"""

import shutil
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
ORIG = HERE / "Class 1.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
RT_SLIDE = f"{R}/slide"
NAVY = "0B2B4E"
LVL_SIZE = {0: 2400, 1: 2200, 2: 2000}
LVL_MARL = {0: 342900, 1: 731520, 2: 1097280}
LVL_CHAR = {0: "▪", 1: "–", 2: "·"}
LINKMAP = {"slide98.xml": 101, "slide99.xml": 102, "slide100.xml": 103,
           "slide101.xml": 104, "slide102.xml": 105, "slide106.xml": 109,
           "slide95.xml": 97, "slide93.xml": 98}
SPLIT = [25, 44, 52, 60, 61]


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def order_of(pres_xml, rels_xml):
    pres = ET.fromstring(pres_xml)
    rid2t = {r.get("Id"): r.get("Target") for r in ET.fromstring(rels_xml)}
    return [rid2t[s.get(q(R, "id"))].split("/")[-1]
            for s in pres.find(q(P, "sldIdLst"))]


class Rels:
    def __init__(self, data, part, order):
        self.key = f"ppt/slides/_rels/{part}.rels"
        self.el = ET.fromstring(data[self.key]); self.order = order

    def _new(self):
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
        rid = self._new()
        rel = ET.SubElement(self.el, q(PKG, "Relationship"))
        rel.set("Id", rid); rel.set("Type", RT_SLIDE); rel.set("Target", tgt)
        return rid

    def url(self, u):
        for r in self.el:
            if r.get("Target") == u and r.get("TargetMode") == "External":
                return r.get("Id")
        rid = self._new()
        rel = ET.SubElement(self.el, q(PKG, "Relationship"))
        rel.set("Id", rid); rel.set("Type", f"{R}/hyperlink")
        rel.set("Target", u); rel.set("TargetMode", "External")
        return rid

    def save(self, data):
        data[self.key] = ser(self.el)


def para_xml(level, runs, first, rels):
    size = LVL_SIZE.get(level, 2000)
    spc = "" if first else f'<a:spcBef><a:spcPts val="{1600 if level == 0 else 500}"/></a:spcBef>'
    lvl = f' lvl="{level}"' if level > 0 else ""
    pPr = (f'<a:pPr marL="{LVL_MARL.get(level, 1097280)}" indent="-274320"{lvl}>{spc}'
           f'<a:buClr><a:srgbClr val="{NAVY}"/></a:buClr>'
           f'<a:buFont typeface="Calibri"/><a:buChar char="{LVL_CHAR.get(level, "·")}"/></a:pPr>')
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
               f'<a:latin typeface="Calibri"/>{hl}</a:rPr><a:t>{escape(text)}</a:t></a:r>')
    return f'<a:p xmlns:a="{A}" xmlns:r="{R}">{pPr}{rs}</a:p>'


def body_boxes(root, ymin, ymax, min_para):
    out = []
    for sp in root.iter(q(P, "sp")):
        off = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "off"))
        tb = sp.find(q(P, "txBody"))
        if off is None or tb is None:
            continue
        y = int(off.get("y"))
        if not (ymin < y < ymax):
            continue
        n = sum(1 for pp in tb.findall(q(A, "p"))
                if "".join(t.text or "" for t in pp.iter(q(A, "t"))).strip())
        if n >= min_para:
            out.append((y, sp))
    out.sort(key=lambda x: x[0])
    return [sp for _, sp in out]


def extract_concat(orig_root, srels):
    items = []
    for sp in body_boxes(orig_root, 900000, 6900000, 2):
        for pp in sp.find(q(P, "txBody")).findall(q(A, "p")):
            pPr = pp.find(q(A, "pPr"))
            lvl = int(pPr.get("lvl")) if (pPr is not None and pPr.get("lvl")) else 0
            runs = []
            for r in pp.findall(q(A, "r")):
                t = r.find(q(A, "t")); tx = (t.text if t is not None else "") or ""
                hl = r.find(q(A, "rPr") + "/" + q(A, "hlinkClick"))
                tgt = None
                if hl is not None:
                    target, mode = srels.get(hl.get(q(R, "id")), (None, None))
                    if target:
                        tgt = target if (mode == "External" or target.startswith(("http", "mailto"))) \
                            else LINKMAP.get(target.split("/")[-1])
                runs.append((tx, tgt))
            if "".join(x[0] for x in runs).strip():
                items.append((lvl, runs))
    return items


def main():
    zr = zipfile.ZipFile(DECK)
    data = {n: zr.read(n) for n in zr.namelist()}
    zr.close()
    order = order_of(data["ppt/presentation.xml"], data["ppt/_rels/presentation.xml.rels"])
    zo = zipfile.ZipFile(ORIG)
    o_order = order_of(zo.read("ppt/presentation.xml"), zo.read("ppt/_rels/presentation.xml.rels"))

    for disp in SPLIT:
        ob = o_order[disp - 1]
        o_root = ET.fromstring(zo.read(f"ppt/slides/{ob}"))
        srels = {r.get("Id"): (r.get("Target"), r.get("TargetMode"))
                 for r in ET.fromstring(zo.read(f"ppt/slides/_rels/{ob}.rels"))}
        items = extract_concat(o_root, srels)
        part = order[disp - 1]
        root = ET.fromstring(data[f"ppt/slides/{part}"])
        boxes = body_boxes(root, 1300000, 6900000, 1)
        target = max(boxes, key=lambda sp: len(sp.find(q(P, "txBody")).findall(q(A, "p"))))
        tb = target.find(q(P, "txBody"))
        for pp in tb.findall(q(A, "p")):
            tb.remove(pp)
        rels = Rels(data, part, order)
        for i, (lvl, runs) in enumerate(items):
            tb.append(ET.fromstring(para_xml(lvl, runs, i == 0, rels)))
        rels.save(data)
        data[f"ppt/slides/{part}"] = ser(root)
        print(f"slide {disp}: body concatenated ({len(items)} paras)")
    zo.close()

    # slide 15: drop the standalone "New Museum" caption
    p15 = order[15 - 1]
    root = ET.fromstring(data[f"ppt/slides/{p15}"])
    spTree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
    for sp in list(spTree.findall(q(P, "sp"))):
        paras = [pp for pp in sp.iter(q(A, "p"))
                 if "".join(t.text or "" for t in pp.iter(q(A, "t"))).strip()]
        joined = "".join(t.text or "" for t in sp.iter(q(A, "t"))).strip()
        if len(paras) == 1 and joined == "New Museum on Etruscan history in Milan":
            spTree.remove(sp)
            print("slide 15: removed duplicate New Museum caption")
    data[f"ppt/slides/{p15}"] = ser(root)

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
