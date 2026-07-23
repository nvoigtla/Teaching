"""Splice the 5 new slides (_moreslides.pptx) into the canonical deck.
Positions (by current display): Renaissance econ -> after 42; Industry/IRI ->
after 62; Economic Miracle + Third Italy -> after 87 (in that order);
Paradox -> after 97."""
import os
import shutil
import zipfile
from pathlib import Path
from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
SRC = HERE / "_moreslides.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
RT_SLIDE, RT_LAYOUT, RT_IMAGE = f"{R}/slide", f"{R}/slideLayout", f"{R}/image"


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = ET.fromstring(data["ppt/presentation.xml"])
    prels = ET.fromstring(data["ppt/_rels/presentation.xml.rels"])
    rid2t = {r.get("Id"): r.get("Target") for r in prels}
    sldids = list(pres.find(q(P, "sldIdLst")))
    order = [os.path.basename(rid2t[s.get(q(R, "id"))]) for s in sldids]
    layout_tgt = next(r.get("Target")
                      for r in ET.fromstring(data[f"ppt/slides/_rels/{order[20]}.rels"])
                      if r.get("Type") == RT_LAYOUT)
    ct = ET.fromstring(data["[Content_Types].xml"])

    zt = zipfile.ZipFile(SRC)
    tdata = {n: zt.read(n) for n in zt.namelist()}
    zt.close()
    tpres = ET.fromstring(tdata["ppt/presentation.xml"])
    trid = {r.get("Id"): r.get("Target") for r in ET.fromstring(tdata["ppt/_rels/presentation.xml.rels"])}
    tparts = [os.path.basename(trid[s.get(q(R, "id"))]) for s in tpres.find(q(P, "sldIdLst"))]
    # temp order: [renaissance, industry, miracle, thirditaly, paradox]

    used_rids = {r.get("Id") for r in prels}
    max_id = max(int(s.get("id")) for s in sldids)

    def next_name(base=220):
        n = base
        while f"ppt/slides/slide{n}.xml" in data:
            n += 1
        return n

    def next_rid():
        n = 1
        while f"rId{n}" in used_rids:
            n += 1
        used_rids.add(f"rId{n}")
        return f"rId{n}"

    new_sldid = {}
    for i, tp in enumerate(tparts):
        nn = next_name()
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
                ext = os.path.splitext(t)[1]
                mn = f"ppt/media/more{nn}{ext}"
                data[mn] = blob
                nr.set("Target", f"../media/{os.path.basename(mn)}")
            else:
                nr.set("Target", t)
        data[f"ppt/slides/_rels/{newpart}.rels"] = ser(nrels)
        ov = ET.SubElement(ct, q(CT, "Override"))
        ov.set("PartName", f"/ppt/slides/{newpart}")
        ov.set("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
        prid = next_rid()
        pr = ET.SubElement(prels, q(PKG, "Relationship"))
        pr.set("Id", prid); pr.set("Type", RT_SLIDE); pr.set("Target", f"slides/{newpart}")
        max_id += 1
        el = ET.Element(q(P, "sldId"))
        el.set("id", str(max_id)); el.set(q(R, "id"), prid)
        new_sldid[i] = el

    # insert: 0 Ren after 42; 1 Ind after 62; 3 ThirdItaly then 2 Miracle after 87; 4 Paradox after 97
    sldids[41].addnext(new_sldid[0])
    sldids[61].addnext(new_sldid[1])
    sldids[86].addnext(new_sldid[3])   # third italy first
    sldids[86].addnext(new_sldid[2])   # miracle -> ends up before third italy
    sldids[96].addnext(new_sldid[4])

    data["[Content_Types].xml"] = ser(ct)
    data["ppt/presentation.xml"] = ser(pres)
    data["ppt/_rels/presentation.xml.rels"] = ser(prels)

    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print("spliced 5 slides (renaissance/industry/miracle/thirditaly/paradox)")


if __name__ == "__main__":
    main()
