"""
Round/shade edits + new images + new 'Empire at its Height' slide (2026-07-20).
 - Slide 15: round the relief picture; round + shade the takeaway bar.
 - Slide 25: round + shade takeaway; narrow bullets; add denarius (right).
 - Slide 28: round + shade takeaway; narrow bullets; add 592 fragmentation map.
 - Insert 'Roman Empire at its Height' slide (from _empireslide.pptx) after 25.
Run _resize_bullets.py apply afterward.
"""
import os
import shutil
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
EMP = HERE / "_empireslide.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
RT_SLIDE, RT_LAYOUT, RT_IMAGE = f"{R}/slide", f"{R}/slideLayout", f"{R}/image"
GOLD = "E09F3E"


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def round_geom(spPr, adj=10000):
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
    gd.set("name", "adj"); gd.set("fmla", f"val {adj}")


def add_shadow(spPr):
    for e in spPr.findall(q(A, "effectLst")):
        spPr.remove(e)
    lst = ET.SubElement(spPr, q(A, "effectLst"))
    sh = ET.SubElement(lst, q(A, "outerShdw"))
    sh.set("blurRad", "50800"); sh.set("dist", "38100")
    sh.set("dir", "2700000"); sh.set("rotWithShape", "0")
    c = ET.SubElement(sh, q(A, "srgbClr")); c.set("val", "000000")
    ET.SubElement(c, q(A, "alpha")).set("val", "30000")


def style_takeaway(root):
    """Round + shade the full-width gold takeaway rectangle."""
    for sp in root.iter(q(P, "sp")):
        spPr = sp.find(q(P, "spPr"))
        if spPr is None:
            continue
        fill = spPr.find(q(A, "solidFill") + "/" + q(A, "srgbClr"))
        ext = spPr.find(q(A, "xfrm") + "/" + q(A, "ext"))
        if (fill is not None and fill.get("val") == GOLD and ext is not None
                and int(ext.get("cx")) > 5000000):
            round_geom(spPr, 12000)
            add_shadow(spPr)
            return True
    return False


def round_pictures(root):
    for pic in root.iter(q(P, "pic")):
        round_geom(pic.find(q(P, "spPr")), 5500)


def narrow_body(root, new_cx):
    best, best_n = None, 0
    for sp in root.iter(q(P, "sp")):
        tb = sp.find(q(P, "txBody"))
        if tb is None:
            continue
        n = sum(1 for pp in tb.findall(q(A, "p"))
                if pp.find(q(A, "pPr") + "/" + q(A, "buChar")) is not None)
        if n > best_n:
            best, best_n = sp, n
    if best is not None:
        best.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "ext")).set("cx", str(new_cx))


def add_pic(root, data, rels, part_no, blob, ext, x, y, w, h, nid):
    mname = f"ppt/media/pol{part_no}_{nid}{ext}"
    data[mname] = blob
    used = {r.get("Id") for r in rels}
    n = 1
    while f"rId{n}" in used:
        n += 1
    rid = f"rId{n}"
    rel = ET.SubElement(rels, q(PKG, "Relationship"))
    rel.set("Id", rid); rel.set("Type", RT_IMAGE)
    rel.set("Target", f"../media/{os.path.basename(mname)}")
    spTree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
    pic = (f'<p:pic xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:nvPicPr>'
           f'<p:cNvPr id="{nid}" name="Pic{nid}"/><p:cNvPicPr>'
           f'<a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
           f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
           f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
           f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 5500"/></a:avLst></a:prstGeom>'
           f'<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" rotWithShape="0">'
           f'<a:srgbClr val="000000"><a:alpha val="32000"/></a:srgbClr></a:outerShdw></a:effectLst>'
           f'</p:spPr></p:pic>')
    spTree.append(ET.fromstring(pic))


def cap(root, text, x, y, w):
    spTree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
    sp = (f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{9000}" name="cap"/>'
          f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
          f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="274320"/></a:xfrm>'
          f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
          f'<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle/>'
          f'<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1100" i="1">'
          f'<a:solidFill><a:srgbClr val="555B66"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr>'
          f'<a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>')
    from xml.sax.saxutils import escape
    spTree.append(ET.fromstring(sp.replace(text, escape(text))))


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = ET.fromstring(data["ppt/presentation.xml"])
    prels = ET.fromstring(data["ppt/_rels/presentation.xml.rels"])
    rid2t = {r.get("Id"): r.get("Target") for r in prels}
    sldlst = pres.find(q(P, "sldIdLst"))
    sldids = list(sldlst)
    order = [os.path.basename(rid2t[s.get(q(R, "id"))]) for s in sldids]
    p15, p25, p28 = order[14], order[24], order[27]

    den = (HERE / "Images" / "denarius_crop.jpg").read_bytes()
    frag = (HERE / "Images" / "s110_0.png").read_bytes()

    # --- slide 15 ---
    r15 = ET.fromstring(data[f"ppt/slides/{p15}"])
    round_pictures(r15)
    style_takeaway(r15)
    data[f"ppt/slides/{p15}"] = ser(r15)

    # --- slide 25 (market) + denarius ---
    r25 = ET.fromstring(data[f"ppt/slides/{p25}"])
    style_takeaway(r25)
    narrow_body(r25, int(7132320))  # 7.8"
    k25 = f"ppt/slides/_rels/{p25}.rels"
    rels25 = ET.fromstring(data[k25])
    add_pic(r25, data, rels25, 25, den, ".jpg",
            x=7635240, y=2651760, w=4114800, h=1485000, nid=901)  # ~x8.35,y2.9,4.5x1.62
    cap(r25, "Roman silver denarius (PAS / British Museum, CC BY-SA 2.0)",
        7635240, 4200000, 4114800)
    data[k25] = ser(rels25)
    data[f"ppt/slides/{p25}"] = ser(r25)

    # --- slide 28 (fall) + 592 map ---
    r28 = ET.fromstring(data[f"ppt/slides/{p28}"])
    style_takeaway(r28)
    narrow_body(r28, int(6766560))  # 7.4"
    k28 = f"ppt/slides/_rels/{p28}.rels"
    rels28 = ET.fromstring(data[k28])
    add_pic(r28, data, rels28, 28, frag, ".png",
            x=7315200, y=2286000, w=4572000, h=2713000, nid=902)  # ~x8,y2.5,5.0x2.97
    cap(r28, "Italy fragmented, c. 592 AD (Lombard, Byzantine, and duchies)",
        7315200, 5120000, 4572000)
    data[k28] = ser(rels28)
    data[f"ppt/slides/{p28}"] = ser(r28)

    # --- splice empire slide after display 25 ---
    zt = zipfile.ZipFile(EMP)
    tdata = {n: zt.read(n) for n in zt.namelist()}
    zt.close()
    tp = os.path.basename(
        {r.get("Id"): r.get("Target")
         for r in ET.fromstring(tdata["ppt/_rels/presentation.xml.rels"])}
        [ET.fromstring(tdata["ppt/presentation.xml"]).find(q(P, "sldIdLst"))[0].get(q(R, "id"))])
    layout_tgt = next(r.get("Target")
                      for r in ET.fromstring(data[f"ppt/slides/_rels/{order[20]}.rels"])
                      if r.get("Type") == RT_LAYOUT)
    nn = 210
    while f"ppt/slides/slide{nn}.xml" in data:
        nn += 1
    newpart = f"slide{nn}.xml"
    data[f"ppt/slides/{newpart}"] = tdata[f"ppt/slides/{tp}"]
    trels = ET.fromstring(tdata[f"ppt/slides/_rels/{tp}.rels"])
    nrels = ET.Element(q(PKG, "Relationships"))
    for r in trels:
        typ, rid, t = r.get("Type"), r.get("Id"), r.get("Target")
        nr = ET.SubElement(nrels, q(PKG, "Relationship"))
        nr.set("Id", rid); nr.set("Type", typ)
        if typ == RT_LAYOUT:
            nr.set("Target", layout_tgt)
        elif typ == RT_IMAGE:
            blob = tdata[os.path.normpath("ppt/slides/" + t).replace(os.sep, "/")]
            mn = f"ppt/media/empire{nn}.png"
            data[mn] = blob
            nr.set("Target", f"../media/{os.path.basename(mn)}")
        else:
            nr.set("Target", t)
    data[f"ppt/slides/_rels/{newpart}.rels"] = ser(nrels)
    ct = ET.fromstring(data["[Content_Types].xml"])
    ov = ET.SubElement(ct, q(CT, "Override"))
    ov.set("PartName", f"/ppt/slides/{newpart}")
    ov.set("ContentType",
           "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
    data["[Content_Types].xml"] = ser(ct)
    used = {r.get("Id") for r in prels}
    n = 1
    while f"rId{n}" in used:
        n += 1
    prid = f"rId{n}"
    pr = ET.SubElement(prels, q(PKG, "Relationship"))
    pr.set("Id", prid); pr.set("Type", RT_SLIDE); pr.set("Target", f"slides/{newpart}")
    max_id = max(int(s.get("id")) for s in sldids)
    ns = ET.Element(q(P, "sldId"))
    ns.set("id", str(max_id + 1)); ns.set(q(R, "id"), prid)
    sldids[24].addnext(ns)  # after display 25
    data["ppt/presentation.xml"] = ser(pres)
    data["ppt/_rels/presentation.xml.rels"] = ser(prels)

    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print(f"styled 15/25/28, added denarius+map, inserted {newpart} after 25")


if __name__ == "__main__":
    main()
