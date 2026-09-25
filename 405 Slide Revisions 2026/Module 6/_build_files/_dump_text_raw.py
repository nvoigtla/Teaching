# -*- coding: utf-8 -*-
"""Dump every slide's text straight from the raw OOXML.

WHY THIS EXISTS.  python-pptx does not enumerate a shape wrapped in
``mc:AlternateContent`` — which is what PowerPoint does to any textbox
containing OMML math when the deck is saved.  So `_dump_deck.py` reported
Module 6 slides 52, 56, 59, 71 and 76 as "title only" or "empty
placeholder" when they are nothing of the sort (found 2026-09-09 via the
animation extractor: slide 52 had five animation clicks against a
placeholder the inventory said was not there).

This walks the XML itself, so nothing can hide: <a:t> for drawing text and
<m:t> for math runs, in document order, with a marker on the parts that
live inside an AlternateContent block.

Usage:  python _dump_text_raw.py NV [52 56 ...]
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'p': P_NS, 'r': R_NS}

MODULE = Path(__file__).resolve().parents[1]  # _build_files/ since 2026-09-24
# The NV source decks were deleted on 2026-09-24; pass a .pptx PATH
# instead of a deck key to dump any deck (e.g. "Module 6 - Full.pptx").
SRC = MODULE / "Module 6 - NV Slides"
DECKS = {
    "NV": "Module 6.pptx",
    "APP": "Module 6 -- Slides On-Campus Applications with Solutions.pptx",
}
for _k in range(1, 9):
    DECKS["V%d" % _k] = "Module 6 - Video %d.pptx" % _k


def slide_parts(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = {r.get('Id'): r.get('Target')
            for r in ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    return ["ppt/" + rels[s.get('{%s}id' % R_NS)].replace("../", "")
            for s in pres.find('p:sldIdLst', NS)]


SP_RE = re.compile(r'<p:(sp|pic|graphicFrame|grpSp)\b.*?</p:\1>', re.S)
NAME_RE = re.compile(r'<p:cNvPr[^>]*name="([^"]*)"')
T_RE = re.compile(r'<(?:a|m):t(?:\s[^>]*)?>(.*?)</(?:a|m):t>', re.S)
OFF_RE = re.compile(r'<a:off x="(-?\d+)" y="(-?\d+)"/>')
EXT_RE = re.compile(r'<a:ext cx="(\d+)" cy="(\d+)"/>')
UNESC = [('&amp;', '&'), ('&lt;', '<'), ('&gt;', '>'), ('&quot;', '"'),
         ('&apos;', "'")]


def unescape(s):
    for a, b in UNESC:
        s = s.replace(a, b)
    return s


def dump(tag, wanted=None):
    path = Path(tag) if tag.lower().endswith(".pptx") else SRC / DECKS[tag]
    z = zipfile.ZipFile(str(path))
    parts = slide_parts(z)
    print("# Raw-XML text dump: %s  (%d slides)" % (path.name, len(parts)))
    print("# AC = the shape sits inside mc:AlternateContent, so python-pptx "
          "does NOT see it.\n")
    for n, part in enumerate(parts, 1):
        if wanted and n not in wanted:
            continue
        xml = z.read(part).decode("utf-8")
        body = xml[xml.find("<p:spTree"):xml.find("</p:spTree>")]
        ac_spans = [(m.start(), m.end()) for m in
                    re.finditer(r'<mc:AlternateContent.*?</mc:AlternateContent>',
                                body, re.S)]
        print("## %s slide %d" % (tag, n))
        for m in SP_RE.finditer(body):
            blk = m.group()
            texts = [unescape(t).strip() for t in T_RE.findall(blk)]
            texts = [t for t in texts if t]
            nm = NAME_RE.search(blk)
            off, ext = OFF_RE.search(blk), EXT_RE.search(blk)
            geo = ""
            if off and ext:
                geo = ("[%.2f,%.2f %.2f x %.2f]"
                       % (int(off.group(1)) / 914400, int(off.group(2)) / 914400,
                          int(ext.group(1)) / 914400, int(ext.group(2)) / 914400))
            in_ac = any(a <= m.start() < b for a, b in ac_spans)
            if not texts and not in_ac:
                continue
            print("  - %s%s %s: %s"
                  % ("AC " if in_ac else "", nm.group(1) if nm else "?", geo,
                     " ⏎ ".join(texts) if texts else "(no text)"))
        print()


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        raise SystemExit(__doc__)
    dump(args[0], {int(a) for a in args[1:]} or None)
