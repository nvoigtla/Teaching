"""
Wire internal slide-jump hyperlinks in "Class 1 - Revised.pptx".

The original deck linked certain words to backup slides. Now that the revised
deck holds all 123 slides in the same order, those jumps resolve to display
slide N. Run AFTER build + splice. Not re-entrant (run on a fresh build+splice).
"""

import shutil
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
PNS = "http://schemas.openxmlformats.org/presentationml/2006/main"
RT_SLIDE = f"{R}/slide"


def q(ns, tag):
    return f"{{{ns}}}{tag}"


# (source_display_slide, link_text_needle, target_display_slide) — read from
# the build manifest so link positions stay correct after slide insertions.
import json  # noqa: E402
LINKS = [tuple(x) for x in
         json.loads((HERE / "_manifest.json").read_text())["links"]]


def _disp_map(data):
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rels = ET.fromstring(data["ppt/_rels/presentation.xml.rels"])
    rid2t = {r.get("Id"): r.get("Target") for r in rels}
    order = [rid2t[s.get(q(R, "id"))].split("/")[-1]
             for s in pres.find(q(PNS, "sldIdLst"))]
    return {i + 1: p for i, p in enumerate(order)}


def _ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8",
                       standalone=True)


def _ensure_slide_rel(data, src_part, target_part):
    key = f"ppt/slides/_rels/{src_part}.rels"
    rels = ET.fromstring(data[key])
    tgt = f"../slides/{target_part}"
    for r in rels:
        if r.get("Type") == RT_SLIDE and r.get("Target") == tgt:
            return r.get("Id"), key, rels
    used = {r.get("Id") for r in rels}
    n = 1
    while f"rId{n}" in used:
        n += 1
    rid = f"rId{n}"
    rel = ET.SubElement(rels, q(PKG, "Relationship"))
    rel.set("Id", rid)
    rel.set("Type", RT_SLIDE)
    rel.set("Target", tgt)
    return rid, key, rels


def _apply_hlink(root, needle, rid):
    for r in root.iter(q(A, "r")):
        t = r.find(q(A, "t"))
        if t is not None and t.text and needle in t.text:
            rPr = r.find(q(A, "rPr"))
            if rPr is None:
                rPr = ET.Element(q(A, "rPr"))
                r.insert(0, rPr)
            for old in rPr.findall(q(A, "hlinkClick")):
                rPr.remove(old)
            # hlinkClick must be the first child of rPr
            hl = ET.Element(q(A, "hlinkClick"))
            hl.set(q(R, "id"), rid)
            hl.set("action", "ppaction://hlinksldjump")
            rPr.insert(0, hl)
            return True
    return False


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    disp2part = _disp_map(data)

    done = 0
    for src_disp, needle, tgt_disp in LINKS:
        src_part = disp2part[src_disp]
        tgt_part = disp2part[tgt_disp]
        rid, rk, rels = _ensure_slide_rel(data, src_part, tgt_part)
        skey = f"ppt/slides/{src_part}"
        root = ET.fromstring(data[skey])
        if _apply_hlink(root, needle, rid):
            data[skey] = _ser(root)
            data[rk] = _ser(rels)
            done += 1
        else:
            print(f"  WARN: '{needle}' not found on slide {src_disp}")

    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print(f"Wired {done}/{len(LINKS)} internal slide-jump links")


if __name__ == "__main__":
    main()
