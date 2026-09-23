# -*- coding: utf-8 -*-
"""Port one or more slides from the In-Class deck into "Module 1 - Revised".

    python _ic_port.py IC:RV [IC:RV ...]          e.g.  python _ic_port.py 8:42

The slide part is copied WHOLESALE - XML, notes, images, tags, layout - so
hand-built choreography survives exactly. Two rules learned the hard way:

  * KEEP the source's own relationship Ids. Renumbering them sequentially
    and remapping the slide XML in place renamed the same `r:embed` twice
    (In-Class slide 33 stores its rels as rId3, rId2, rId1, rId4) and two
    pictures shipped as "The picture can't be displayed" (2026-09-23).
  * The footer page number is a live field; reset its cached text to the
    Revised display number or it shows the in-class one until PowerPoint
    recomputes.

A slide carrying a slide-to-slide jump needs its target decided, so this
refuses those and points at the round-2 script, which has the link map.

`Module 1 - Revised.pptx` is the source of truth, so this is one of the very
few tools allowed to write it - it edits single slides rather than
regenerating the deck. It rolls the _t-1 / _t-2 backups first.
"""
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
SRC = HERE / "Module 1 - In Class.pptx"
DST = HERE / "Module 1 - Revised.pptx"

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

CT_BY_EXT = {"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg",
             "gif": "image/gif", "emf": "image/x-emf", "wmf": "image/x-wmf",
             "tiff": "image/tiff", "webp": "image/webp", "bmp": "image/bmp"}
CT_NOTES = ("application/vnd.openxmlformats-officedocument."
            "presentationml.notesSlide+xml")
CT_TAGS = ("application/vnd.openxmlformats-officedocument."
           "presentationml.tags+xml")


def q(ns, t):
    return "{%s}%s" % (ns, t)


def slide_parts(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = {r.get("Id"): r.get("Target") for r in
            ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    return ["ppt/" + rels[s.get(q(R, "id"))].lstrip("/").replace("../", "")
            for s in pres.find(q(P, "sldIdLst"))]


def rels_of(z, part):
    try:
        return ET.fromstring(
            z.read(part.replace("slides/", "slides/_rels/") + ".rels"))
    except KeyError:
        return None


def main():
    pairs = []
    for a in sys.argv[1:]:
        if a.startswith("-"):
            continue
        ic, rv = a.split(":")
        pairs.append((int(ic), int(rv)))
    if not pairs:
        raise SystemExit(__doc__)

    zs = zipfile.ZipFile(SRC)
    zd = zipfile.ZipFile(DST)
    src_parts, dst_parts = slide_parts(zs), slide_parts(zd)
    items = {n: zd.read(n) for n in zd.namelist()}
    ct = ET.fromstring(items["[Content_Types].xml"])
    have_ext = {d.get("Extension").lower() for d in ct.findall(q(CT, "Default"))}
    have_over = {o.get("PartName") for o in ct.findall(q(CT, "Override"))}

    report = []
    for ic, rv in pairs:
        sp, dp = src_parts[ic - 1], dst_parts[rv - 1]
        tree = ET.fromstring(zs.read(sp))
        srels = rels_of(zs, sp)
        drels = rels_of(zd, dp)
        layout = None
        for r in (drels if drels is not None else []):
            if r.get("Type").endswith("/slideLayout"):
                layout = r.get("Target")

        new = ET.Element(q(REL, "Relationships"))
        kinds = {}
        for r in (srels if srels is not None else []):
            typ, tgt = r.get("Type"), r.get("Target")
            short = typ.rsplit("/", 1)[-1]
            if short == "slide":
                raise SystemExit(
                    "IC %d carries a slide-to-slide jump; its Revised target "
                    "has to be decided by hand. The wiring in use is written "
                    "up in Session-Notes.md, 2026-09-23 round 4." % ic)
            e = ET.SubElement(new, q(REL, "Relationship"))
            e.set("Id", r.get("Id"))            # keep the source's ids
            e.set("Type", typ)
            if short == "slideLayout":
                e.set("Target", layout or tgt)
                continue
            if r.get("TargetMode") == "External":
                e.set("Target", tgt)
                e.set("TargetMode", "External")
                continue
            src_name = "ppt/" + tgt.replace("../", "")
            base = src_name.rsplit("/", 1)[-1]
            folder = src_name.rsplit("/", 2)[-2]
            dst_name = "ppt/%s/ic%02d_%s" % (folder, rv, base)
            items[dst_name] = zs.read(src_name)
            e.set("Target", "../%s/%s" % (folder, dst_name.split("/")[-1]))
            kinds[folder] = kinds.get(folder, 0) + 1
            ext = base.rsplit(".", 1)[-1].lower()
            pn = "/" + dst_name
            if folder == "media" and ext not in have_ext and ext in CT_BY_EXT:
                d = ET.SubElement(ct, q(CT, "Default"))
                d.set("Extension", ext)
                d.set("ContentType", CT_BY_EXT[ext])
                have_ext.add(ext)
            elif folder in ("notesSlides", "tags") and pn not in have_over:
                o = ET.SubElement(ct, q(CT, "Override"))
                o.set("PartName", pn)
                o.set("ContentType",
                      CT_NOTES if folder == "notesSlides" else CT_TAGS)
                have_over.add(pn)
            if folder == "notesSlides":
                nrels = rels_of(zs, src_name)
                if nrels is not None:
                    for rr in nrels:
                        if rr.get("Type").endswith("/slide"):
                            rr.set("Target", "../slides/%s" % dp.split("/")[-1])
                    items["ppt/notesSlides/_rels/%s.rels"
                          % dst_name.split("/")[-1]] = ET.tostring(
                        nrels, xml_declaration=True, encoding="UTF-8",
                        standalone=True)

        for fld in tree.iter(q(A, "fld")):
            if fld.get("type") == "slidenum":
                t = fld.find(q(A, "t"))
                if t is not None:
                    t.text = str(rv)

        items[dp] = ET.tostring(tree, xml_declaration=True, encoding="UTF-8",
                                standalone=True)
        items[dp.replace("slides/", "slides/_rels/") + ".rels"] = ET.tostring(
            new, xml_declaration=True, encoding="UTF-8", standalone=True)
        report.append((ic, rv, kinds))

    items["[Content_Types].xml"] = ET.tostring(ct, xml_declaration=True,
                                               encoding="UTF-8",
                                               standalone=True)
    zs.close()
    zd.close()

    t1 = DST.with_name(DST.stem + "_t-1.pptx")
    t2 = DST.with_name(DST.stem + "_t-2.pptx")
    if t1.exists():
        shutil.copy2(t1, t2)
    shutil.copy2(DST, t1)
    tmp = DST.with_suffix(".tmp.pptx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for n, data in items.items():
            out.writestr(n, data)
    tmp.replace(DST)

    for ic, rv, kinds in report:
        print("  IC %2d -> RV %2d   %s" % (ic, rv, ", ".join(
            "%s=%d" % kv for kv in sorted(kinds.items())) or "no media"))
    print("backup rolled to %s" % t1.name)


if __name__ == "__main__":
    main()
