"""
Audit-driven fixes (2026-07-20, round 9):
  A. Slide 15 museum box -> move to bottom-LEFT, light-grey fill + soft shadow,
     link corrected to the original (museo.fondazioneluigirovati.org/en).
  B. Restore missing substring hyperlinks dropped during the wording restore:
     s27 "Medieval castles"->106, "many small states"->107;
     s48 "newly created operas"->110; s50 "quality of operas"->111;
     s52 "1,300 years"->112.
  C. Re-clone images that were missed because they sat inside GROUPS:
     slides 39, 40, 54, 56, 66, 75 (group-aware, absolute positions).
OOXML surgery only.
"""

import copy
import os
import shutil
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
ORIG = HERE / "Class 1.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
RT_SLIDE = f"{R}/slide"
NAVY, GREY = "0B2B4E", "E7E9EC"
MUSEUM_URL = "https://museo.fondazioneluigirovati.org/en"

LINKS = [(27, "Medieval castles", 106), (27, "many small states", 107),
         (48, "newly created operas", 110), (50, "quality of operas", 111),
         (52, "1,300 years", 112)]
# slide -> (image region EMU x0,y0,x1,y1); orig slide = same number
IMG_SLIDES = {
    39: (320040, 1417320, 11871960, 6263640),
    40: (320040, 1417320, 11871960, 5943600),
    54: (320040, 1417320, 11871960, 5943600),
    56: (8345000, 1509000, 11887200, 6355000),
    66: (320040, 1417320, 11871960, 6263640),
    75: (320040, 1417320, 11871960, 5943600),
}


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def order_of(pres_xml, rels_xml):
    pres = ET.fromstring(pres_xml)
    rid2t = {r.get("Id"): r.get("Target") for r in ET.fromstring(rels_xml)}
    return [rid2t[s.get(q(R, "id"))].split("/")[-1]
            for s in pres.find(q(P, "sldIdLst"))]


def new_rid(rels):
    used = {r.get("Id") for r in rels}
    n = 1
    while f"rId{n}" in used:
        n += 1
    return f"rId{n}"


def slide_rel(rels, order, disp):
    tgt = f"../slides/{order[disp - 1]}"
    for r in rels:
        if r.get("Type") == RT_SLIDE and r.get("Target") == tgt:
            return r.get("Id")
    rid = new_rid(rels)
    rel = ET.SubElement(rels, q(PKG, "Relationship"))
    rel.set("Id", rid); rel.set("Type", RT_SLIDE); rel.set("Target", tgt)
    return rid


# ---- group-aware picture extraction --------------------------------------

def pic_abs(pic, srels, zo):
    xf = pic.find(q(P, "spPr") + "/" + q(A, "xfrm"))
    off = xf.find(q(A, "off")); ext = xf.find(q(A, "ext"))
    x, y = int(off.get("x")), int(off.get("y"))
    cx, cy = int(ext.get("cx")), int(ext.get("cy"))
    node = pic.getparent()
    while node is not None and node.tag == q(P, "grpSp"):
        g = node.find(q(P, "grpSpPr") + "/" + q(A, "xfrm"))
        go, ge = g.find(q(A, "off")), g.find(q(A, "ext"))
        co, ce = g.find(q(A, "chOff")), g.find(q(A, "chExt"))
        sx = int(ge.get("cx")) / max(1, int(ce.get("cx")))
        sy = int(ge.get("cy")) / max(1, int(ce.get("cy")))
        x = int(go.get("x")) + (x - int(co.get("x"))) * sx
        y = int(go.get("y")) + (y - int(co.get("y"))) * sy
        cx *= sx; cy *= sy
        node = node.getparent()
    blip = pic.find(q(P, "blipFill") + "/" + q(A, "blip"))
    tgt = srels[blip.get(q(R, "embed"))]
    blob = zo.read(os.path.normpath("ppt/slides/" + tgt).replace(os.sep, "/"))
    ext = os.path.splitext(tgt)[1].lower() or ".png"
    return int(x), int(y), int(cx), int(cy), blob, ext


def shadow(spPr):
    lst = ET.SubElement(spPr, q(A, "effectLst"))
    sh = ET.SubElement(lst, q(A, "outerShdw"))
    sh.set("blurRad", "50800"); sh.set("dist", "38100")
    sh.set("dir", "2700000"); sh.set("rotWithShape", "0")
    c = ET.SubElement(sh, q(A, "srgbClr")); c.set("val", "000000")
    ET.SubElement(c, q(A, "alpha")).set("val", "32000")


def place_pics(spTree, pics, region, start_id):
    x0, y0, x1, y1 = region
    L = min(p[0] for p in pics); T = min(p[1] for p in pics)
    Rr = max(p[0] + p[2] for p in pics); B = max(p[1] + p[3] for p in pics)
    bw, bh = max(1, Rr - L), max(1, B - T)
    s = min((x1 - x0) / bw, (y1 - y0) / bh)
    ox = x0 + (x1 - x0 - bw * s) / 2
    oy = y0 + (y1 - y0 - bh * s) / 2
    nid = start_id
    for (l, t, w, h, blob, ext) in pics:
        nl, nt = int(ox + (l - L) * s), int(oy + (t - T) * s)
        nw, nh = int(w * s), int(h * s)
        yield nid, nl, nt, nw, nh, blob, ext
        nid += 1


def main():
    zc = zipfile.ZipFile(DECK)
    data = {n: zc.read(n) for n in zc.namelist()}
    zc.close()
    order = order_of(data["ppt/presentation.xml"], data["ppt/_rels/presentation.xml.rels"])
    zo = zipfile.ZipFile(ORIG)
    o_order = order_of(zo.read("ppt/presentation.xml"), zo.read("ppt/_rels/presentation.xml.rels"))

    # ---- A. slide 15 museum box ----
    p15 = order[14]
    root = ET.fromstring(data[f"ppt/slides/{p15}"])
    rk = f"ppt/slides/_rels/{p15}.rels"
    rels = ET.fromstring(data[rk])
    used_rid = None
    for sp in root.iter(q(P, "sp")):
        cnv = sp.find(".//" + q(P, "cNvPr"))
        if cnv is None or cnv.get("name") != "MuseumLink":
            continue
        off = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "off"))
        off.set("x", "252059"); off.set("y", "6035040")          # bottom-left
        spPr = sp.find(q(P, "spPr"))
        fill = spPr.find(q(A, "solidFill") + "/" + q(A, "srgbClr"))
        fill.set("val", GREY)                                    # light grey
        ln = spPr.find(q(A, "ln"))                               # drop gold border
        if ln is not None:
            for ch in list(ln):
                ln.remove(ch)
            ET.SubElement(ln, q(A, "noFill"))
        for e in spPr.findall(q(A, "effectLst")):
            spPr.remove(e)
        shadow(spPr)
        for hl in sp.iter(q(A, "hlinkClick")):
            if hl.get(q(R, "id")):
                used_rid = hl.get(q(R, "id"))
    for r in rels:                                              # fix URL
        if r.get("Id") == used_rid:
            r.set("Target", MUSEUM_URL)
    data[rk] = ser(rels)
    data[f"ppt/slides/{p15}"] = ser(root)
    print("A: slide 15 box -> bottom-left, grey fill+shadow, link fixed")

    # ---- B. restore substring links ----
    by_slide = {}
    for disp, needle, tgt in LINKS:
        by_slide.setdefault(disp, []).append((needle, tgt))
    for disp, items in by_slide.items():
        part = order[disp - 1]
        root = ET.fromstring(data[f"ppt/slides/{part}"])
        rk = f"ppt/slides/_rels/{part}.rels"
        rels = ET.fromstring(data[rk])
        for needle, tgt in items:
            rid = slide_rel(rels, order, tgt)
            done = False
            for r in list(root.iter(q(A, "r"))):
                t = r.find(q(A, "t"))
                if t is None or not t.text or needle not in t.text:
                    continue
                s = t.text; i = s.index(needle)
                before, after = s[:i], s[i + len(needle):]
                p = r.getparent(); idx = list(p).index(r)
                base_rPr = r.find(q(A, "rPr"))
                new = []
                if before:
                    rb = copy.deepcopy(r); rb.find(q(A, "t")).text = before
                    new.append(rb)
                rl = copy.deepcopy(r); rl.find(q(A, "t")).text = needle
                rpr = rl.find(q(A, "rPr"))
                if rpr is None:
                    rpr = ET.Element(q(A, "rPr")); rl.insert(0, rpr)
                rpr.set("u", "sng")
                for old in rpr.findall(q(A, "hlinkClick")):
                    rpr.remove(old)
                hl = ET.SubElement(rpr, q(A, "hlinkClick"))
                hl.set(q(R, "id"), rid); hl.set("action", "ppaction://hlinksldjump")
                new.append(rl)
                if after:
                    ra = copy.deepcopy(r); ra.find(q(A, "t")).text = after
                    new.append(ra)
                p.remove(r)
                for j, nr in enumerate(new):
                    p.insert(idx + j, nr)
                done = True
                break
            print(f"B: slide {disp} '{needle}' -> {tgt}  {'ok' if done else 'NOT FOUND'}")
        data[rk] = ser(rels)
        data[f"ppt/slides/{part}"] = ser(root)

    # ---- C. re-clone grouped images ----
    for disp, region in IMG_SLIDES.items():
        part = order[disp - 1]
        root = ET.fromstring(data[f"ppt/slides/{part}"])
        spTree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
        # remove current pics
        for pic in list(spTree.iter(q(P, "pic"))):
            pic.getparent().remove(pic)
        # extract from original (group-aware)
        ob = o_order[disp - 1]
        o_root = ET.fromstring(zo.read(f"ppt/slides/{ob}"))
        o_srels = {r.get("Id"): r.get("Target")
                   for r in ET.fromstring(zo.read(f"ppt/slides/_rels/{ob}.rels"))}
        pics = [pic_abs(pic, o_srels, zo) for pic in o_root.iter(q(P, "pic"))]
        # add media parts + rels, place pics
        rk = f"ppt/slides/_rels/{part}.rels"
        rels = ET.fromstring(data[rk])
        ids = [int(e.get("id")) for e in root.iter(q(P, "cNvPr")) if e.get("id")]
        start_id = (max(ids) + 1) if ids else 300
        media_idx = 500 + disp * 10
        for nid, nl, nt, nw, nh, blob, ext in place_pics(spTree, pics, region, start_id):
            mname = f"ppt/media/aud{media_idx}{ext}"
            data[mname] = blob
            rid = new_rid(rels)
            rel = ET.SubElement(rels, q(PKG, "Relationship"))
            rel.set("Id", rid); rel.set("Type", f"{R}/image")
            rel.set("Target", f"../media/{os.path.basename(mname)}")
            pic_xml = (
                f'<p:pic xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:nvPicPr>'
                f'<p:cNvPr id="{nid}" name="Pic{nid}"/><p:cNvPicPr>'
                f'<a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
                f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
                f'<p:spPr><a:xfrm><a:off x="{nl}" y="{nt}"/><a:ext cx="{nw}" cy="{nh}"/></a:xfrm>'
                f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 5500"/></a:avLst></a:prstGeom>'
                f'<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" rotWithShape="0">'
                f'<a:srgbClr val="000000"><a:alpha val="32000"/></a:srgbClr></a:outerShdw></a:effectLst>'
                f'</p:spPr></p:pic>')
            spTree.append(ET.fromstring(pic_xml))
            media_idx += 1
        # register png default in content types (already present) — ensure
        data[rk] = ser(rels)
        data[f"ppt/slides/{part}"] = ser(root)
        print(f"C: slide {disp} re-cloned {len(pics)} images")
    zo.close()

    # ensure image content-type defaults exist
    ct = data["[Content_Types].xml"].decode("utf8")
    for ext, mime in (("png", "image/png"), ("jpeg", "image/jpeg"),
                      ("jpg", "image/jpeg"), ("gif", "image/gif")):
        if f'Extension="{ext}"' not in ct:
            ct = ct.replace("</Types>",
                            f'<Default Extension="{ext}" ContentType="{mime}"/></Types>')
    data["[Content_Types].xml"] = ct.encode("utf8")

    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob if isinstance(blob, bytes) else blob.encode("utf8"))
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print("done")


if __name__ == "__main__":
    main()
