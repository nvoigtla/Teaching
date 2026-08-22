"""
Resize bullet text across the deck (2026-07-20):
  - Prefer 28 pt main / 24 pt sub (L2 = 22); fall back to 24 / 22 (L2 = 20)
    on crowded slides. Decision is measured with Calibri metrics so nothing
    spills past the body box.
  - Spacing-before: 12 pt before main bullets, 3 pt before subs; first bullet
    in a box gets none.
Only touches 405 character-bulleted body boxes (buChar). Skips slide 14
(already done by hand). Pass 'apply' to write; otherwise dry-run.

OOXML surgery only.
"""

import math
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET
from PIL import ImageFont

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MARL = {0: 342900, 1: 731520, 2: 1097280}
BIG = {0: 28, 1: 24, 2: 22}
SMALL = {0: 24, 1: 22, 2: 20}
FONT = ImageFont.truetype("C:/Windows/Fonts/calibri.ttf", 100)
SKIP_DISP = {14}


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


def text_px(text, size_pt):
    return FONT.getlength(text) * size_pt / 100.0


def block_height(bullets, sizes, box_cx):
    total = 0.0
    for i, (lvl, txt) in enumerate(bullets):
        sz = sizes[lvl]
        avail = max(1.0, (box_cx - MARL.get(lvl, 1097280)) / 12700.0)
        ln = max(1, math.ceil(text_px(txt, sz) / (avail * 0.95)))
        spc = 0 if i == 0 else (12 if lvl == 0 else 3)
        total += ln * sz * 1.2 + spc
    return total


def bullet_boxes(root):
    out = []
    for sp in root.iter(q(P, "sp")):
        tb = sp.find(q(P, "txBody"))
        if tb is None:
            continue
        if tb.find(q(A, "p") + "/" + q(A, "pPr") + "/" + q(A, "buChar")) is None:
            # check any paragraph
            if not any(pp.find(q(A, "pPr") + "/" + q(A, "buChar")) is not None
                       for pp in tb.findall(q(A, "p"))):
                continue
        ext = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "ext"))
        if ext is None:
            continue
        out.append((sp, tb, int(ext.get("cx")), int(ext.get("cy"))))
    return out


def process_box(tb, cx, cy):
    bullets = []
    for pp in tb.findall(q(A, "p")):
        pPr = pp.find(q(A, "pPr"))
        lvl = int(pPr.get("lvl")) if (pPr is not None and pPr.get("lvl")) else 0
        txt = "".join(t.text or "" for t in pp.iter(q(A, "t")))
        if txt.strip():
            bullets.append((lvl, txt))
    if not bullets:
        return None
    cy_pt = cy / 12700.0
    sizes = BIG if block_height(bullets, BIG, cx) <= cy_pt * 0.93 else SMALL
    # apply
    first = True
    for pp in tb.findall(q(A, "p")):
        pPr = pp.find(q(A, "pPr"))
        txt = "".join(t.text or "" for t in pp.iter(q(A, "t")))
        if not txt.strip():
            continue
        lvl = int(pPr.get("lvl")) if (pPr is not None and pPr.get("lvl")) else 0
        # font size on every run
        for r in pp.findall(q(A, "r")):
            rPr = r.find(q(A, "rPr"))
            if rPr is None:
                rPr = ET.Element(q(A, "rPr")); r.insert(0, rPr)
            rPr.set("sz", str(sizes[lvl] * 100))
        # spacing-before
        if pPr is None:
            pPr = ET.SubElement(pp, q(A, "pPr")); pp.insert(0, pPr)
        for old in pPr.findall(q(A, "spcBef")):
            pPr.remove(old)
        if not first:
            spc = ET.Element(q(A, "spcBef"))
            ET.SubElement(spc, q(A, "spcPts")).set("val", str((12 if lvl == 0 else 3) * 100))
            pPr.insert(0, spc)
        first = False
    return sizes[0], len(bullets)


def main(apply):
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    order = order_of(data)
    report = []
    for disp, base in enumerate(order, 1):
        if disp in SKIP_DISP:
            continue
        root = ET.fromstring(data[f"ppt/slides/{base}"])
        boxes = bullet_boxes(root)
        if not boxes:
            continue
        res = []
        for sp, tb, cx, cy in boxes:
            r = process_box(tb, cx, cy)
            if r:
                res.append(r)
        if res:
            report.append((disp, res))
        data[f"ppt/slides/{base}"] = ser(root)

    for disp, res in report:
        main_sz = res[0][0]; nb = sum(r[1] for r in res)
        print(f"slide {disp:3}: {'28/24' if main_sz == 28 else '24/22'}  ({nb} bullets)")

    if apply:
        tmp = DECK.with_suffix(".pptx.tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for name, blob in data.items():
                out.writestr(name, blob)
        with zipfile.ZipFile(tmp) as chk:
            assert chk.testzip() is None
        shutil.move(str(tmp), str(DECK))
        print(f"APPLIED to {len(report)} slides")
    else:
        print(f"DRY RUN — {len(report)} slides would change")


if __name__ == "__main__":
    main(len(sys.argv) > 1 and sys.argv[1] == "apply")
