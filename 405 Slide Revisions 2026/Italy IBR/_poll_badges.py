"""
In-place edit (2026-07-19, round 5): move the "POLL" badge to the bottom-right
corner on every poll slide (matching slide 11, which the user already fixed),
and delete the small gold accent square in front of the badge. OOXML surgery.
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
GOLD = "E09F3E"
SMALL = 228600   # 0.25" — gold dot is ~0.14"; gold strips are 2.2" wide


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def off_of(sp):
    return sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "off"))


def ext_of(sp):
    return sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "ext"))


def poll_slides(data):
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rels = ET.fromstring(data["ppt/_rels/presentation.xml.rels"])
    rid2t = {r.get("Id"): r.get("Target") for r in rels}
    order = [rid2t[s.get(q(R, "id"))].split("/")[-1]
             for s in pres.find(q(P, "sldIdLst"))]
    out = []
    for disp, base in enumerate(order, 1):
        rk = f"ppt/slides/_rels/{base}.rels"
        types = [r.get("Type").split("/")[-1] for r in ET.fromstring(data[rk])]
        if "tags" in types:
            out.append((disp, base))
    return out


def poll_text_off(root):
    for sp in root.iter(q(P, "sp")):
        if any((t.text or "").strip() == "POLL" for t in sp.iter(q(A, "t"))):
            return off_of(sp)
    return None


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    polls = poll_slides(data)
    print("poll slides:", [d for d, _ in polls])

    # reference target from slide 11
    p11 = dict(polls)[11]
    ref = poll_text_off(ET.fromstring(data[f"ppt/slides/{p11}"]))
    TX, TY = ref.get("x"), ref.get("y")

    for disp, base in polls:
        if disp == 11:
            continue
        root = ET.fromstring(data[f"ppt/slides/{base}"])
        cur = poll_text_off(root)
        if cur is None:
            print(f"  slide {disp}: no POLL text, skipped")
            continue
        cx0, cy0 = cur.get("x"), cur.get("y")
        # move badge (pill + text share this off) to the reference corner
        for sp in root.iter(q(P, "sp")):
            off = off_of(sp)
            if off is not None and off.get("x") == cx0 and off.get("y") == cy0:
                off.set("x", TX); off.set("y", TY)
        # delete the small gold accent square
        spTree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
        for sp in list(spTree.findall(q(P, "sp"))):
            fill = sp.find(q(P, "spPr") + "/" + q(A, "solidFill") + "/" + q(A, "srgbClr"))
            ext = ext_of(sp)
            if (fill is not None and fill.get("val") == GOLD and ext is not None
                    and int(ext.get("cx")) < SMALL and int(ext.get("cy")) < SMALL):
                spTree.remove(sp)
        data[f"ppt/slides/{base}"] = ser(root)

    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print("POLL badges moved to bottom-right; gold squares removed")


if __name__ == "__main__":
    main()
