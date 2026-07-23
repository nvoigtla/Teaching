"""
Insert the 3 new slides (from _newslides.pptx) into the canonical deck and
apply the 5 content enrichments. OOXML surgery only — no python-pptx round-trip
of the canonical deck (preserves poll tags and hand-edits).

New slides:  Geography -> after display 14;  First Integrated Market -> after 23;
             The Fall -> after 25.
Enrichments (by current part): 15 Etruscan, 16 Ethnic groups, 19 Republic,
             24 Decline, 25 End of Rome.
Run _resize_bullets.py apply afterward to re-fit sizes.
"""
import os
import shutil
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
TMP = HERE / "_newslides.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
RT_SLIDE = f"{R}/slide"
RT_LAYOUT = f"{R}/slideLayout"
RT_IMAGE = f"{R}/image"
LVL_MARL = {0: 342900, 1: 731520, 2: 1097280}
LVL_CHAR = {0: "▪", 1: "–", 2: "·"}


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def para_xml(level, text):
    lvl = f' lvl="{level}"' if level else ""
    return (f'<a:p xmlns:a="{A}"><a:pPr marL="{LVL_MARL[level]}" indent="-274320"{lvl}>'
            f'<a:buClr><a:srgbClr val="0B2B4E"/></a:buClr>'
            f'<a:buFont typeface="Calibri"/><a:buChar char="{LVL_CHAR[level]}"/></a:pPr>'
            f'<a:r><a:rPr lang="en-US" sz="2400"><a:solidFill><a:srgbClr val="0B2B4E"/></a:solidFill>'
            f'<a:latin typeface="Calibri"/></a:rPr><a:t>{escape(text)}</a:t></a:r></a:p>')


def body_box(root):
    best, best_n = None, 0
    for sp in root.iter(q(P, "sp")):
        tb = sp.find(q(P, "txBody"))
        if tb is None:
            continue
        n = sum(1 for pp in tb.findall(q(A, "p"))
                if pp.find(q(A, "pPr") + "/" + q(A, "buChar")) is not None)
        if n > best_n:
            best, best_n = tb, n
    return best


def enrich(data, part, *, edits=None, inserts=None, appends=None):
    root = ET.fromstring(data[f"ppt/slides/{part}"])
    tb = body_box(root)
    if edits:
        for old, new in edits:
            for t in tb.iter(q(A, "t")):
                if t.text and old in t.text:
                    t.text = t.text.replace(old, new)
                    break
    if inserts:
        for after_needle, text, level in inserts:
            target = None
            for pp in tb.findall(q(A, "p")):
                if after_needle in "".join(t.text or "" for t in pp.iter(q(A, "t"))):
                    target = pp
                    break
            newp = ET.fromstring(para_xml(level, text))
            if target is not None:
                target.addnext(newp)
            else:
                tb.append(newp)
    if appends:
        for text, level in appends:
            tb.append(ET.fromstring(para_xml(level, text)))
    data[f"ppt/slides/{part}"] = ser(root)


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

    # canonical content layout target (reuse an existing content slide's layout)
    ref_rels = ET.fromstring(data[f"ppt/slides/_rels/{order[20]}.rels"])   # display 21
    layout_tgt = next(r.get("Target") for r in ref_rels
                      if r.get("Type") == RT_LAYOUT)

    # ---- enrichments (by current part) ----
    enrich(data, order[14],  # 15 Etruscan
           edits=[("Economy based on mining and metal trade",
                   "Wealth from iron (Elba) and copper — Etruria’s Mediterranean metal trade")],
           inserts=[("Transition from chiefdom",
                     "A league of independent cities — not one unified state", 0)],
           appends=[("Passed on to Rome: the arch, engineering, and the alphabet", 0)])
    enrich(data, order[15],  # 16 Ethnic groups
           appends=[("Magna Graecia (the Greek south) brought coinage, vines, and olives — still core exports", 0),
                    ("The South was then the richer, more urban half — the North–South gap is not primordial", 0)])
    enrich(data, order[18],  # 19 Roman Republic
           appends=[("Roman law — property, contracts, the legal “person” — the ancestor of European commercial law (revived at Bologna, ~1088)", 0)])
    enrich(data, order[23],  # 24 Decline
           appends=[("Fiscal overreach and coin debasement → inflation; Diocletian’s price controls (301 AD) failed", 0)])
    enrich(data, order[24],  # 25 End of Rome
           appends=[("As the empire fragmented, the single market broke into local economies — trade, cities, and coinage shrank", 0)])
    print("enrichments applied to slides 15,16,19,24,25")

    # ---- load temp new slides ----
    zt = zipfile.ZipFile(TMP)
    tdata = {n: zt.read(n) for n in zt.namelist()}
    zt.close()
    tpres = ET.fromstring(tdata["ppt/presentation.xml"])
    trid2t = {r.get("Id"): r.get("Target")
              for r in ET.fromstring(tdata["ppt/_rels/presentation.xml.rels"])}
    tparts = [os.path.basename(trid2t[s.get(q(R, "id"))])
              for s in tpres.find(q(P, "sldIdLst"))]   # [geo, market, fall]
    targets = [14, 23, 25]

    # unique names / ids
    def free_slide_name():
        n = 200
        while f"ppt/slides/slide{n}.xml" in data:
            n += 1
        return n
    used_rids = {r.get("Id") for r in prels}
    def free_pres_rid():
        n = 1
        while f"rId{n}" in used_rids:
            n += 1
        rid = f"rId{n}"
        used_rids.add(rid)
        return rid
    max_id = max(int(s.get("id")) for s in sldids)

    ct = ET.fromstring(data["[Content_Types].xml"])

    for i, (tp, tgt_disp) in enumerate(zip(tparts, targets)):
        n = free_slide_name()
        newpart = f"slide{n}.xml"
        data[f"ppt/slides/{newpart}"] = tdata[f"ppt/slides/{tp}"]
        # rels: remap layout + image, keep rIds
        trels = ET.fromstring(tdata[f"ppt/slides/_rels/{tp}.rels"])
        rels = ET.Element(q(PKG, "Relationships"))
        for r in trels:
            typ, rid, t = r.get("Type"), r.get("Id"), r.get("Target")
            nr = ET.SubElement(rels, q(PKG, "Relationship"))
            nr.set("Id", rid); nr.set("Type", typ)
            if typ == RT_LAYOUT:
                nr.set("Target", layout_tgt)
            elif typ == RT_IMAGE:
                blob = tdata[os.path.normpath("ppt/slides/" + t).replace(os.sep, "/")]
                mname = f"ppt/media/relief{n}.jpg"
                data[mname] = blob
                nr.set("Target", f"../media/{os.path.basename(mname)}")
            else:
                nr.set("Target", t)
        data[f"ppt/slides/_rels/{newpart}.rels"] = ser(rels)
        # content-type override
        ov = ET.SubElement(ct, q(CT, "Override"))
        ov.set("PartName", f"/ppt/slides/{newpart}")
        ov.set("ContentType",
               "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
        # presentation rel + sldId
        prid = free_pres_rid()
        prel = ET.SubElement(prels, q(PKG, "Relationship"))
        prel.set("Id", prid); prel.set("Type", RT_SLIDE)
        prel.set("Target", f"slides/{newpart}")
        max_id += 1
        newsld = ET.Element(q(P, "sldId"))
        newsld.set("id", str(max_id)); newsld.set(q(R, "id"), prid)
        sldids[tgt_disp - 1].addnext(newsld)
        print(f"inserted {newpart} after display {tgt_disp}")

    # ensure jpg default content type
    if not any(d.get("Extension") == "jpg" for d in ct.findall(q(CT, "Default"))) \
       and not any(d.get("Extension") == "jpeg" for d in ct.findall(q(CT, "Default"))):
        d = ET.SubElement(ct, q(CT, "Default"))
        d.set("Extension", "jpg"); d.set("ContentType", "image/jpeg")

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
    print("done; run _resize_bullets.py apply next")


if __name__ == "__main__":
    main()
