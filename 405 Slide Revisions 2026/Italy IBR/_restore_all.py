"""
Full pass (2026-07-19, round 7): restore EXACT original bullet wording + all
references on the history/paper content slides, keeping the 405 format; rework
the remaining backup picture targets (104, 105, 109).

Pulls original text programmatically from Class 1.pptx (no transcription).
Excludes logistics (3-9), roadmap (10), topics (98), polls, the deliberately
updated modern slides (brain drain / corruption), and chart-image slides.

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

# original backup-target part  ->  current display slide
LINKMAP = {
    "slide98.xml": 101, "slide99.xml": 102, "slide100.xml": 103,
    "slide101.xml": 104, "slide102.xml": 105, "slide106.xml": 109,
    "slide95.xml": 97, "slide93.xml": 98,
}

# content-bullet slides to restore:  current display -> original display
PAIRS = {n: n for n in (
    15, 16, 18, 19, 21, 23, 24, 25, 27, 29, 30, 33, 36, 38, 42, 44, 45, 47,
    48, 52, 53, 55, 56, 60, 61, 64, 65, 67, 68, 69, 70, 71, 73, 79, 81)}
PAIRS[93] = 89   # Strong Brands
PAIRS[94] = 90   # Quiz

# backup picture slides to enlarge + add back button:  disp -> source slide
BACKUPS = {104: 15, 105: 15, 109: 29}


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def order_of(data):
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rels = ET.fromstring(data["ppt/_rels/presentation.xml.rels"])
    rid2t = {r.get("Id"): r.get("Target") for r in rels}
    return [rid2t[s.get(q(R, "id"))].split("/")[-1]
            for s in pres.find(q(P, "sldIdLst"))]


class Rels:
    def __init__(self, data, part, order):
        self.key = f"ppt/slides/_rels/{part}.rels"
        self.el = ET.fromstring(data[self.key])
        self.order = order

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


def _body_sp(root, ymin, ymax):
    """The text box with the most paragraphs whose top is in (ymin, ymax) EMU."""
    best, best_n = None, 0
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
        if n > best_n:
            best, best_n = sp, n
    return best


def extract_items(orig_root, orig_srels):
    sp = _body_sp(orig_root, 900000, 7000000)
    if sp is None:
        return None
    items = []
    for pp in sp.find(q(P, "txBody")).findall(q(A, "p")):
        pPr = pp.find(q(A, "pPr"))
        lvl = int(pPr.get("lvl")) if (pPr is not None and pPr.get("lvl")) else 0
        runs = []
        for r in pp.findall(q(A, "r")):
            t = r.find(q(A, "t"))
            tx = t.text if t is not None else ""
            if tx is None:
                tx = ""
            hl = r.find(q(A, "rPr") + "/" + q(A, "hlinkClick"))
            tgt = None
            if hl is not None:
                target, mode = orig_srels.get(hl.get(q(R, "id")), (None, None))
                if target:
                    if mode == "External" or target.startswith(("http", "mailto")):
                        tgt = target
                    else:
                        tgt = LINKMAP.get(target.split("/")[-1])
            runs.append((tx, tgt))
        if "".join(x[0] for x in runs).strip():
            items.append((lvl, runs))
    return items


def replace_current_body(root, items, rels):
    sp = _body_sp(root, 1300000, 6900000)
    if sp is None:
        return False
    tb = sp.find(q(P, "txBody"))
    for pp in tb.findall(q(A, "p")):
        tb.remove(pp)
    for i, (lvl, runs) in enumerate(items):
        tb.append(ET.fromstring(para_xml(lvl, runs, i == 0, rels)))
    return True


def back_button(new_id, rid):
    return ET.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:nvSpPr>'
        f'<p:cNvPr id="{new_id}" name="BackButton">'
        f'<a:hlinkClick r:id="{rid}" action="ppaction://hlinksldjump"/></p:cNvPr>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="252059" y="468630"/><a:ext cx="1417320" cy="420624"/></a:xfrm>'
        f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 30000"/></a:avLst></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{NAVY}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>'
        f'<a:p><a:pPr algn="ctr"/><a:r>'
        f'<a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        f'<a:latin typeface="Calibri"/></a:rPr><a:t>← Back</a:t></a:r></a:p></p:txBody></p:sp>')


def rework_backup(data, order, disp, back_to):
    part = order[disp - 1]
    root = ET.fromstring(data[f"ppt/slides/{part}"])
    spTree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
    for sp in list(spTree.findall(q(P, "sp"))):
        off = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "off"))
        if off is not None and 500000 < int(off.get("y")) < 1300000:
            spTree.remove(sp)
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
    rels = Rels(data, part, order)
    rid = rels.slide(back_to)
    ids = [int(e.get("id")) for e in root.iter(q(P, "cNvPr")) if e.get("id")]
    spTree.append(back_button((max(ids) + 1) if ids else 100, rid))
    rels.save(data)
    data[f"ppt/slides/{part}"] = ser(root)


def main():
    zr = zipfile.ZipFile(DECK)
    data = {n: zr.read(n) for n in zr.namelist()}
    zr.close()
    zo = zipfile.ZipFile(ORIG)
    orig_order = order_of({"ppt/presentation.xml": zo.read("ppt/presentation.xml"),
                           "ppt/_rels/presentation.xml.rels": zo.read("ppt/_rels/presentation.xml.rels")})
    order = order_of(data)

    done, skipped = [], []
    for cur, orig in sorted(PAIRS.items()):
        ob = orig_order[orig - 1]
        o_root = ET.fromstring(zo.read(f"ppt/slides/{ob}"))
        o_srels = {r.get("Id"): (r.get("Target"), r.get("TargetMode"))
                   for r in ET.fromstring(zo.read(f"ppt/slides/_rels/{ob}.rels"))}
        items = extract_items(o_root, o_srels)
        if not items:
            skipped.append(cur); continue
        part = order[cur - 1]
        c_root = ET.fromstring(data[f"ppt/slides/{part}"])
        rels = Rels(data, part, order)
        if replace_current_body(c_root, items, rels):
            rels.save(data)
            data[f"ppt/slides/{part}"] = ser(c_root)
            done.append(cur)
        else:
            skipped.append(cur)
    zo.close()

    for disp, src in BACKUPS.items():
        rework_backup(data, order, disp, src)

    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print(f"restored bodies on {len(done)} slides: {done}")
    if skipped:
        print(f"SKIPPED (no body found): {skipped}")
    print(f"backups reworked: {list(BACKUPS)}")


if __name__ == "__main__":
    main()
