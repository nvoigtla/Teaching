"""
Splice the PollEverywhere slide from the 2023 deck into "Class 2 - Revised.pptx".

The poll slide cannot be created with python-pptx: the add-in identifies its
activity through a `tags` relationship on the picture shape
(`__PE_POLL_EMBED_ID`), and python-pptx drops it on round-trip.  It also needs
its notes part -- the add-in reads the poll URL from the speaker notes, and a
poll slide whose tag is present but whose notes part is missing crashes the
slideshow renderer deck-wide.  So slide XML + tags + image + notes travel
together, by raw OOXML surgery.

Idempotent: always starts from the fresh build output, so re-running the whole
pipeline never accumulates duplicates.

    python _build_Class2.py && python _splice_poll.py
"""

import os
import re
import shutil
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
DECK = os.path.join(HERE, "Class 2 - Revised.pptx")
SRC = os.path.join(HERE, "Class 2.pptx")           # 2023 original (read-only)

SRC_SLIDE = "ppt/slides/slide3.xml"                # the poll slide
SRC_TAGS = "ppt/tags/tag2.xml"
SRC_IMG = "ppt/media/image1.png"
SRC_NOTES = "ppt/notesSlides/notesSlide3.xml"

POSITION = 4          # 1-based display position the poll slide should land on
SLIDE_W = 12192000

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def next_free(names, pattern):
    used = {int(m.group(1)) for n in names for m in [re.match(pattern, n)] if m}
    i = 1
    while i in used:
        i += 1
    return i


def main():
    zin = zipfile.ZipFile(DECK)
    names = zin.namelist()
    if any(n.startswith("ppt/tags/") for n in names):
        raise SystemExit("deck already carries a tags part - rebuild first")

    n_slide = next_free(names, r"ppt/slides/slide(\d+)\.xml$")
    n_notes = next_free(names, r"ppt/notesSlides/notesSlide(\d+)\.xml$")
    n_img = next_free(names, r"ppt/media/image(\d+)\.png$")
    new_slide = "ppt/slides/slide%d.xml" % n_slide
    new_notes = "ppt/notesSlides/notesSlide%d.xml" % n_notes
    new_img = "ppt/media/image%d.png" % n_img
    new_tags = "ppt/tags/tag1.xml"

    zsrc = zipfile.ZipFile(SRC)
    slide_xml = zsrc.read(SRC_SLIDE).decode("utf8")
    tags_xml = zsrc.read(SRC_TAGS)
    img_bytes = zsrc.read(SRC_IMG)
    notes_xml = zsrc.read(SRC_NOTES).decode("utf8")

    # the poll graphic came off a 10 x 7.5" canvas -- recentre it on 13.33 x 7.5"
    m = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>',
                  slide_xml)
    _, y, cx, cy = m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4))
    slide_xml = slide_xml[:m.start()] + (
        '<a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/>'
        % (int((SLIDE_W - cx) / 2), y, cx, cy)) + slide_xml[m.end():]

    # ---- parts written fresh -------------------------------------------------
    slide_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships">'
        '<Relationship Id="rId1" Type="%s/tags" Target="../tags/tag1.xml"/>'
        '<Relationship Id="rId2" Type="%s/slideLayout" '
        'Target="../slideLayouts/slideLayout7.xml"/>'
        '<Relationship Id="rId3" Type="%s/notesSlide" Target="../notesSlides/%s"/>'
        '<Relationship Id="rId4" Type="%s/image" Target="../media/%s"/>'
        "</Relationships>"
        % (REL, REL, REL, os.path.basename(new_notes), REL,
           os.path.basename(new_img))
    )
    notes_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships">'
        '<Relationship Id="rId1" Type="%s/notesMaster" '
        'Target="../notesMasters/notesMaster1.xml"/>'
        '<Relationship Id="rId2" Type="%s/slide" Target="../slides/%s"/>'
        "</Relationships>" % (REL, REL, os.path.basename(new_slide))
    )

    # ---- [Content_Types].xml -------------------------------------------------
    ctypes = zin.read("[Content_Types].xml").decode("utf8")
    add = (
        '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-'
        'officedocument.presentationml.slide+xml"/>'
        '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-'
        'officedocument.presentationml.notesSlide+xml"/>'
        '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-'
        'officedocument.presentationml.tags+xml"/>' % (new_slide, new_notes, new_tags)
    )
    assert new_slide not in ctypes and new_tags not in ctypes
    ctypes = ctypes.replace("</Types>", add + "</Types>")

    # ---- presentation.xml.rels ----------------------------------------------
    prels = zin.read("ppt/_rels/presentation.xml.rels").decode("utf8")
    rid = "rId%d" % (max(int(x) for x in re.findall(r'Id="rId(\d+)"', prels)) + 1)
    prels = prels.replace("</Relationships>",
                          '<Relationship Id="%s" Type="%s/slide" Target="slides/%s"/>'
                          "</Relationships>" % (rid, REL, os.path.basename(new_slide)))

    # ---- presentation.xml : insert into <p:sldIdLst> at POSITION -------------
    pres = zin.read("ppt/presentation.xml").decode("utf8")
    ids = [int(x) for x in re.findall(r'<p:sldId id="(\d+)"', pres)]
    new_id = max(ids) + 1
    entries = re.findall(r"<p:sldId [^>]*/>", pres)
    entry = '<p:sldId id="%d" r:id="%s"/>' % (new_id, rid)
    entries.insert(POSITION - 1, entry)
    pres = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>",
                  "<p:sldIdLst>" + "".join(entries) + "</p:sldIdLst>", pres,
                  flags=re.S)

    # ---- write the new package ---------------------------------------------
    tmp = DECK + ".tmp"
    replaced = {"[Content_Types].xml": ctypes.encode("utf8"),
                "ppt/_rels/presentation.xml.rels": prels.encode("utf8"),
                "ppt/presentation.xml": pres.encode("utf8")}
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, replaced.get(item.filename, zin.read(item.filename)))
        zout.writestr(new_slide, slide_xml.encode("utf8"))
        zout.writestr("ppt/slides/_rels/%s.rels" % os.path.basename(new_slide),
                      slide_rels.encode("utf8"))
        zout.writestr(new_notes, notes_xml.encode("utf8"))
        zout.writestr("ppt/notesSlides/_rels/%s.rels" % os.path.basename(new_notes),
                      notes_rels.encode("utf8"))
        zout.writestr(new_tags, tags_xml)
        zout.writestr(new_img, img_bytes)
    zin.close()
    zsrc.close()
    shutil.move(tmp, DECK)
    print("spliced poll slide as display #%d  (%s, %s, %s, %s)"
          % (POSITION, new_slide, new_tags, new_img, new_notes))


if __name__ == "__main__":
    main()
