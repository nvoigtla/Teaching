"""
Inject the live PollEverywhere embeds into "Class 1 - Revised.pptx".

Run AFTER _build_Italy_Class1.py (which builds the 3 poll slides visually:
chrome + POLL pill + poll snapshot image + verbatim notes).  This step adds
the PollEverywhere plumbing that python-pptx cannot safely round-trip:
a `tags` part carrying __PE_POLL_EMBED_ID (the GUID the PollEv plugin uses to
relink the slide to the online poll), its relationship, the [Content_Types]
override, and the <p:custDataLst><p:tags/></p:custDataLst> reference.

Always run on a freshly-built deck (build -> splice); it is not re-entrant.
"""

import shutil
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"

RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
PNS = "http://schemas.openxmlformats.org/presentationml/2006/main"
ET.register_namespace("", PKG_REL)
CT_TAGS = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml"

# display slide -> (tag filename, __PE_POLL_EMBED_ID GUID), read from the
# build manifest so poll positions stay correct after any slide insertion.
import json  # noqa: E402
_mf = json.loads((HERE / "_manifest.json").read_text())["polls"]  # {guid: disp}
POLLS = {}
for _i, (_guid, _disp) in enumerate(sorted(_mf.items(), key=lambda kv: kv[1]), 1):
    POLLS[_disp] = (f"tag{_i}.xml", _guid)

TAG_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:tag name="__PE_POLL_EMBED_ID" val="{guid}"/></p:tagLst>'
)


def _display_to_part(names_bytes):
    """Map display slide number -> slideN.xml basename via sldIdLst order."""
    pres = ET.fromstring(names_bytes["ppt/presentation.xml"])
    rels = ET.fromstring(names_bytes["ppt/_rels/presentation.xml.rels"])
    rid2t = {r.get("Id"): r.get("Target") for r in rels}
    order = []
    lst = pres.find(f"{{{PNS}}}sldIdLst")
    for sld in lst:
        rid = sld.get(f"{{{RT}}}id")
        order.append(rid2t[rid].split("/")[-1])
    return {i + 1: p for i, p in enumerate(order)}


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()

    disp2part = _display_to_part(data)

    for disp, (tagfile, guid) in POLLS.items():
        part = disp2part[disp]                       # e.g. slide11.xml
        slide_key = f"ppt/slides/{part}"
        rels_key = f"ppt/slides/_rels/{part}.rels"

        # 1. tag part
        data[f"ppt/tags/{tagfile}"] = TAG_XML.format(guid=guid).encode("utf8")

        # 2. relationship (new unique rId)
        rels = ET.fromstring(data[rels_key])
        used = {r.get("Id") for r in rels}
        n = 1
        while f"rId{n}" in used:
            n += 1
        new_rid = f"rId{n}"
        rel = ET.SubElement(rels, f"{{{PKG_REL}}}Relationship")
        rel.set("Id", new_rid)
        rel.set("Type", f"{RT}/tags")
        rel.set("Target", f"../tags/{tagfile}")
        body_rels = ET.tostring(rels, encoding="unicode")
        data[rels_key] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            + body_rels).encode("utf8")

        # 3. custDataLst reference at end of <p:cSld>
        body = data[slide_key].decode("utf8")
        inject = f'<p:custDataLst><p:tags r:id="{new_rid}"/></p:custDataLst>'
        assert "</p:cSld>" in body, part
        body = body.replace("</p:cSld>", inject + "</p:cSld>", 1)
        data[slide_key] = body.encode("utf8")

    # 4. [Content_Types].xml overrides
    ct = data["[Content_Types].xml"].decode("utf8")
    adds = ""
    for tagfile, _ in POLLS.values():
        ov = f'<Override PartName="/ppt/tags/{tagfile}" ContentType="{CT_TAGS}"/>'
        if ov not in ct:
            adds += ov
    ct = ct.replace("</Types>", adds + "</Types>", 1)
    data["[Content_Types].xml"] = ct.encode("utf8")

    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)

    # integrity check
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print(f"Spliced {len(POLLS)} PollEverywhere embeds into {DECK.name}")


if __name__ == "__main__":
    main()
