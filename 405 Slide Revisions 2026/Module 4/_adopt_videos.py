"""Adopt Nico's hand edits from the recorded per-video decks back into
`Module 4 - Revised.pptx`, then append the not-in-video slides again behind
two dividers.

WHY A WHOLESALE SLIDE TRANSPLANT.  Replaying individual hand edits would
miss things: paragraph-level bullet properties (buChar / marL / indent) are
invisible to a geometry diff, and the video decks also carry re-choreographed
`<p:timing>` blocks and new shape groupings.  Copying each video slide's part
verbatim carries geometry, run formatting, OMML math, grouping, animation and
notes in one move, so nothing can be left behind.  This is safe because the
video decks were carved out of this very deck: their theme, layout and master
are byte-identical to the deck's apart from a cached date-placeholder string.

Part 1 - adopt
    For each of the 54 mapped pairs in `_video_map.py`, the video deck's
    slide part replaces the deck's, with its images, notes slide and tags
    part copied across (media deduplicated by SHA-1) and the cached
    slide-number field reset to the deck's own number.

Part 2 - append
    A divider "Slides Not Included in the Videos", then a second copy of
    each of the 36 slides no video contains, then a divider "Some
    Applications of the Material Covered in Module 4, Videos 1 - 5".
    Duplicated slides get their own copy of their notes part (a notes slide
    carries a relationship back to its one slide, so it cannot be shared).

Pure zip + ElementTree surgery.  python-pptx would strip the PollEverywhere
`tags` relationships.

Usage:
    python _adopt_videos.py --dry-run
    python _adopt_videos.py --out "Module 4 - Revised_new.pptx"
"""
import hashlib
import re
import shutil
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

import _video_map as VM

HERE = Path(__file__).resolve().parent
DECK = HERE / "Module 4 - Revised.pptx"
M1_DECK = HERE.parent / "Module 1" / "Module 1 - Revised.pptx"

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
R_ID = "{%s}id" % NS_R

CT_SLIDE = ("application/vnd.openxmlformats-officedocument."
            "presentationml.slide+xml")
CT_NOTES = ("application/vnd.openxmlformats-officedocument."
            "presentationml.notesSlide+xml")
CT_TAGS = ("application/vnd.openxmlformats-officedocument."
           "presentationml.tags+xml")

REL_SLIDE = NS_R + "/slide"
REL_NOTES = NS_R + "/notesSlide"
REL_NOTESMASTER = NS_R + "/notesMaster"
REL_LAYOUT = NS_R + "/slideLayout"
REL_IMAGE = NS_R + "/image"
REL_TAGS = NS_R + "/tags"

FOOTER_M4 = ("Management 405  ·  Module 4  ·  "
             "Competitive Markets and Market Interventions")

DIVIDER_A = ("Slides Not Included", "in the Videos")
DIVIDER_B = ("Some Applications of the Material",
             "Covered in Module 4, Videos 1 – 5")


# ------------------------------------------------------------------ helpers

def rels_name(part):
    head, tail = part.rsplit("/", 1)
    return "%s/_rels/%s.rels" % (head, tail)


def read_rels(blobs, part):
    """[(Id, Type, Target, TargetMode)] for a part, or []."""
    rn = rels_name(part)
    if rn not in blobs:
        return []
    root = ET.fromstring(blobs[rn])
    return [(c.get("Id"), c.get("Type"), c.get("Target"), c.get("TargetMode"))
            for c in root]


def write_rels(blobs, part, rels):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<Relationships xmlns="%s">' % NS_REL]
    for rid, typ, tgt, mode in rels:
        m = ' TargetMode="%s"' % mode if mode else ""
        out.append('<Relationship Id="%s" Type="%s" Target="%s"%s/>'
                   % (rid, typ, tgt, m))
    out.append("</Relationships>")
    blobs[rels_name(part)] = "".join(out).encode("utf-8")


def display_to_part(blobs):
    pres = ET.fromstring(blobs["ppt/presentation.xml"])
    rmap = {r[0]: r[2] for r in read_rels(blobs, "ppt/presentation.xml")}
    lst = pres.find("{%s}sldIdLst" % NS_P)
    return ["ppt/" + rmap[s.get(R_ID)].replace("../", "").lstrip("/")
            for s in lst]


def resolve(target, from_part):
    """Resolve a rels Target relative to the part holding the rels."""
    base = from_part.rsplit("/", 1)[0]
    t = target
    while t.startswith("../"):
        t = t[3:]
        base = base.rsplit("/", 1)[0]
    return (base + "/" + t) if base else t


SLIDENUM_RE = re.compile(
    r'(<a:fld[^>]*type="slidenum"[^>]*>.*?<a:t>)(\d*)(</a:t>)', re.S)


def set_pagenum(xml_bytes, n):
    s = xml_bytes.decode("utf-8")
    s2, count = SLIDENUM_RE.subn(lambda m: m.group(1) + str(n) + m.group(3),
                                 s, count=1)
    return s2.encode("utf-8"), count


def load(path):
    z = zipfile.ZipFile(str(path))
    blobs = {n: z.read(n) for n in z.namelist()}
    z.close()
    return blobs


# --------------------------------------------------------------- divider xml

def build_divider(line1, line2, cached_num):
    """Module 1 slide 50's divider, with Module 4's footer."""
    def para(text):
        return ('<a:p><a:pPr algn="ctr"/><a:r>'
                '<a:rPr sz="5400" b="1" i="0" u="none">'
                '<a:solidFill><a:srgbClr val="0B2B4E"/></a:solidFill>'
                '<a:latin typeface="Calibri"/></a:rPr>'
                '<a:t>%s</a:t></a:r></a:p>' % text)

    def strip(sid, name, x, y, cx, cy, colour):
        return (
            '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/>'
            '<p:nvPr/></p:nvSpPr><p:spPr>'
            '<a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:ln><a:noFill/></a:ln><a:effectLst/></p:spPr>'
            '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            '<a:p><a:pPr algn="ctr"/></a:p></p:txBody></p:sp>'
            % (sid, name, x, y, cx, cy, colour))

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/'
        'main" xmlns:p="%s" xmlns:r="%s"><p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
        '</p:nvGrpSpPr><p:grpSpPr/>'
        # centred two-line title
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/>'
        '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'
        '<a:xfrm><a:off x="0" y="2400300"/>'
        '<a:ext cx="12191695" cy="1691640"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" lIns="0" rIns="0" tIns="0" '
        'bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>%s%s</p:txBody>'
        '</p:sp>'
        # gold accent strip under the title
        '%s'
        # footer rule + gold strip
        '%s%s'
        # footer course line
        '<p:sp><p:nvSpPr><p:cNvPr id="6" name="TextBox 5"/>'
        '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'
        '<a:xfrm><a:off x="251999" y="6583680"/>'
        '<a:ext cx="10058400" cy="292608"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" lIns="0" rIns="0" tIns="0" '
        'bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>'
        '<a:p><a:pPr algn="l"/><a:r>'
        '<a:rPr sz="1200" b="0" i="0" u="none">'
        '<a:solidFill><a:srgbClr val="555B66"/></a:solidFill>'
        '<a:latin typeface="Calibri"/></a:rPr><a:t>%s</a:t></a:r></a:p>'
        '</p:txBody></p:sp>'
        # live page number
        '<p:sp><p:nvSpPr><p:cNvPr id="7" name="TextBox 6"/>'
        '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'
        '<a:xfrm><a:off x="11475720" y="6583680"/>'
        '<a:ext cx="502920" cy="292608"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        '<p:txBody><a:bodyPr wrap="none" lIns="0" rIns="0" tIns="0" '
        'bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>'
        '<a:p><a:pPr algn="r"/>'
        '<a:fld id="{B38D1FD4-E120-5B60-A5FE-8016C3F29565}" '
        'type="slidenum"><a:rPr lang="en-US" sz="1200" dirty="0">'
        '<a:solidFill><a:srgbClr val="555B66"/></a:solidFill>'
        '<a:latin typeface="Calibri"/></a:rPr><a:t>%d</a:t></a:fld>'
        '</a:p></p:txBody></p:sp>'
        '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sld>'
        % (NS_P, NS_R,
           para(line1), para(line2),
           strip(3, "Rectangle 2", 975207, 4229100, 10241280, 54864,
                 "E09F3E"),
           strip(4, "Rectangle 3", 0, 6537960, 12191695, 18288, "C8CDD3"),
           strip(5, "Rectangle 4", 251999, 6524244, 2011680, 45720,
                 "E09F3E"),
           FOOTER_M4, cached_num)).encode("utf-8")


# -------------------------------------------------------------------- main

class Builder(object):
    def __init__(self, blobs):
        self.b = blobs
        self.parts = display_to_part(blobs)
        self.media_by_hash = {}
        for n, data in blobs.items():
            if n.startswith("ppt/media/"):
                self.media_by_hash[hashlib.sha1(data).hexdigest()] = n
        self.log = []
        self.new_overrides = []      # (partname, contenttype)

    # -- media / part import ------------------------------------------------

    def import_media(self, src_blobs, src_part, prefix):
        data = src_blobs[src_part]
        h = hashlib.sha1(data).hexdigest()
        if h in self.media_by_hash:
            return self.media_by_hash[h]
        name = "ppt/media/%s%s" % (prefix, src_part.split("/")[-1])
        i = 1
        while name in self.b:
            name = "ppt/media/%s%d_%s" % (prefix, i, src_part.split("/")[-1])
            i += 1
        self.b[name] = data
        self.media_by_hash[h] = name
        return name

    def import_notes(self, src_blobs, src_part, prefix, owner_slide):
        name = "ppt/notesSlides/%s%s" % (prefix, src_part.split("/")[-1])
        i = 1
        while name in self.b:
            name = "ppt/notesSlides/%s%d_%s" % (prefix, i,
                                                src_part.split("/")[-1])
            i += 1
        self.b[name] = src_blobs[src_part]
        rels = [("rId1", REL_NOTESMASTER,
                 "../notesMasters/notesMaster1.xml", None),
                ("rId2", REL_SLIDE,
                 "../slides/" + owner_slide.split("/")[-1], None)]
        # carry any images the notes page itself uses
        extra = 3
        for rid, typ, tgt, mode in read_rels(src_blobs, src_part):
            if typ == REL_IMAGE:
                got = self.import_media(src_blobs,
                                        resolve(tgt, src_part), prefix)
                rels.append(("rId%d" % extra, REL_IMAGE,
                             "../" + got[len("ppt/"):], None))
                extra += 1
        write_rels(self.b, name, rels)
        self.new_overrides.append((name, CT_NOTES))
        return name

    def import_tags(self, src_blobs, src_part, prefix):
        name = "ppt/tags/%s%s" % (prefix, src_part.split("/")[-1])
        i = 1
        while name in self.b:
            name = "ppt/tags/%s%d_%s" % (prefix, i, src_part.split("/")[-1])
            i += 1
        self.b[name] = src_blobs[src_part]
        self.new_overrides.append((name, CT_TAGS))
        return name

    # -- part 1: adopt ------------------------------------------------------

    def adopt(self):
        for deck_name, pairs in VM.VIDEO_DECKS:
            src = load(HERE / VM.VIDEO_DIR / deck_name)
            src_parts = display_to_part(src)
            vk = re.search(r"Video (\d)", deck_name).group(1)
            prefix = "v%s_" % vk
            for v, r in pairs:
                vpart = src_parts[v - 1]
                dpart = self.parts[r - 1]
                xml = src[vpart]
                new_rels = []
                for rid, typ, tgt, mode in read_rels(src, vpart):
                    if mode == "External":
                        new_rels.append((rid, typ, tgt, mode))
                        continue
                    abs_src = resolve(tgt, vpart)
                    if typ == REL_LAYOUT:
                        new_rels.append(
                            (rid, typ, "../slideLayouts/slideLayout1.xml",
                             None))
                    elif typ == REL_IMAGE:
                        got = self.import_media(src, abs_src, prefix)
                        new_rels.append((rid, typ,
                                         "../" + got[len("ppt/"):], None))
                    elif typ == REL_NOTES:
                        got = self.import_notes(src, abs_src, prefix, dpart)
                        new_rels.append((rid, typ,
                                         "../" + got[len("ppt/"):], None))
                    elif typ == REL_TAGS:
                        got = self.import_tags(src, abs_src, prefix)
                        new_rels.append((rid, typ,
                                         "../" + got[len("ppt/"):], None))
                    else:
                        raise SystemExit(
                            "unhandled rel %s on %s slide %d" % (typ, vk, v))
                xml, n = set_pagenum(xml, r)
                self.b[dpart] = xml
                write_rels(self.b, dpart, new_rels)
                self.log.append("adopt  video %s.%-2d -> slide %-2d  "
                                "rels=%d pagenum=%s"
                                % (vk, v, r, len(new_rels),
                                   "set" if n else "none"))

    # -- part 2: append -----------------------------------------------------

    def append_block(self):
        # string surgery on presentation.xml: serialising it through
        # ElementTree would rewrite the namespace prefixes on a part
        # PowerPoint is fussy about.
        pres_txt = self.b["ppt/presentation.xml"].decode("utf-8")
        prels = read_rels(self.b, "ppt/presentation.xml")
        next_rid = max(int(re.sub(r"\D", "", r[0])) for r in prels) + 1
        next_sid = max(int(m) for m in
                       re.findall(r'<p:sldId id="(\d+)"', pres_txt)) + 1
        next_slide = 1 + max(
            int(re.search(r"slide(\d+)\.xml", n).group(1))
            for n in self.b if re.match(r"ppt/slides/slide\d+\.xml$", n))
        new_sldids = []

        def add_slide(part_name, xml, rels, label):
            nonlocal next_rid, next_sid
            self.b[part_name] = xml
            write_rels(self.b, part_name, rels)
            self.new_overrides.append((part_name, CT_SLIDE))
            prels.append(("rId%d" % next_rid, REL_SLIDE,
                          "slides/" + part_name.split("/")[-1], None))
            new_sldids.append('<p:sldId id="%d" r:id="rId%d"/>'
                              % (next_sid, next_rid))
            self.log.append("append %-22s %s" % (label, part_name))
            next_rid += 1
            next_sid += 1

        display = len(self.parts) + 1

        # divider A
        name = "ppt/slides/slide%d.xml" % next_slide
        next_slide += 1
        add_slide(name, build_divider(DIVIDER_A[0], DIVIDER_A[1], display),
                  [("rId1", REL_LAYOUT,
                    "../slideLayouts/slideLayout1.xml", None)],
                  "divider A")
        display += 1

        # the 36 duplicates, in original order
        for r in VM.DUPLICATE_AT_END:
            srcpart = self.parts[r - 1]
            name = "ppt/slides/slide%d.xml" % next_slide
            next_slide += 1
            xml, _ = set_pagenum(self.b[srcpart], display)
            rels = []
            for rid, typ, tgt, mode in read_rels(self.b, srcpart):
                if typ == REL_NOTES and mode != "External":
                    got = self.import_notes(self.b,
                                            resolve(tgt, srcpart),
                                            "dup%d_" % r, name)
                    rels.append((rid, typ, "../" + got[len("ppt/"):], None))
                elif typ == REL_TAGS and mode != "External":
                    got = self.import_tags(self.b, resolve(tgt, srcpart),
                                           "dup%d_" % r)
                    rels.append((rid, typ, "../" + got[len("ppt/"):], None))
                else:
                    rels.append((rid, typ, tgt, mode))
            add_slide(name, xml, rels, "dup of slide %d" % r)
            display += 1

        # divider B
        name = "ppt/slides/slide%d.xml" % next_slide
        add_slide(name, build_divider(DIVIDER_B[0], DIVIDER_B[1], display),
                  [("rId1", REL_LAYOUT,
                    "../slideLayouts/slideLayout1.xml", None)],
                  "divider B")

        write_rels(self.b, "ppt/presentation.xml", prels)
        assert pres_txt.count("</p:sldIdLst>") == 1
        pres_txt = pres_txt.replace("</p:sldIdLst>",
                                    "".join(new_sldids) + "</p:sldIdLst>")
        self.b["ppt/presentation.xml"] = pres_txt.encode("utf-8")

    # -- content types ------------------------------------------------------

    def finish_content_types(self):
        ct = self.b["[Content_Types].xml"].decode("utf-8")
        add = []
        for part, typ in self.new_overrides:
            pn = "/" + part
            if 'PartName="%s"' % pn not in ct:
                add.append('<Override PartName="%s" ContentType="%s"/>'
                           % (pn, typ))
        ct = ct.replace("</Types>", "".join(add) + "</Types>")
        self.b["[Content_Types].xml"] = ct.encode("utf-8")
        return len(add)


def save(blobs, path):
    with zipfile.ZipFile(str(path), "w", zipfile.ZIP_DEFLATED) as z:
        # content types first, as PowerPoint expects
        z.writestr("[Content_Types].xml", blobs["[Content_Types].xml"])
        for n in sorted(k for k in blobs if k != "[Content_Types].xml"):
            z.writestr(n, blobs[n])


def main():
    dry = "--dry-run" in sys.argv
    out = None
    if "--out" in sys.argv:
        out = HERE / sys.argv[sys.argv.index("--out") + 1]

    n_map, n_not = VM.check()
    src = DECK
    if "--from" in sys.argv:
        src = HERE / sys.argv[sys.argv.index("--from") + 1]
    print("source deck: %s" % src.name)
    blobs = load(src)
    bd = Builder(blobs)
    print("deck: %d slides   mapped: %d   not in video: %d"
          % (len(bd.parts), n_map, n_not))

    bd.adopt()
    bd.append_block()
    n_ct = bd.finish_content_types()

    total = len(display_to_part(bd.b))
    print("\n".join(bd.log[:6]))
    print("   ... %d log lines" % len(bd.log))
    print("new content-type overrides: %d" % n_ct)
    print("resulting slide count: %d" % total)

    if dry:
        print("\n--dry-run: nothing written")
        return
    target = out or (HERE / "Module 4 - Revised.pptx")
    save(bd.b, target)
    print("\nwrote %s (%.1f MB)"
          % (target.name, target.stat().st_size / 1e6))


if __name__ == "__main__":
    main()
