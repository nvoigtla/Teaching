# -*- coding: utf-8 -*-
"""Check that every rebuilt slide carries the text of the source slide it
came from — so nothing goes missing the way slides 49 and 52 did.

The source side is read from the RAW OOXML (`_dump_text_raw` logic), not
via python-pptx, because python-pptx cannot see a shape wrapped in
mc:AlternateContent and would report content that is there as absent.

Usage:  python _check_coverage.py
"""
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).parent
MODULE = HERE.parent  # _build_files/ since 2026-09-24
# The NV source decks were deleted on 2026-09-24 (the rebuild is over).
# Restore them from git history to run this check again.
SRC = MODULE / "Module 6 - NV Slides"
P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'p': P_NS, 'r': R_NS}
T_RE = re.compile(r'<(?:a|m):t(?:\s[^>]*)?>(.*?)</(?:a|m):t>', re.S)
# PowerPoint splits a paragraph across many runs at every formatting
# boundary, so comparing run by run reports fragments like "will" / "watch"
# as missing lines.  Compare PARAGRAPHS: join the runs inside each <a:p>.
P_RE = re.compile(r'<a:p>.*?</a:p>', re.S)

# rebuilt slide -> (source deck, source slide).  Cards, agendas and the
# slides adopted from PG have no NV source and are listed as None.
MAP = {
    1: ("NV", 1), 2: ("NV", 2), 3: ("NV", 3),
    5: ("NV", 4), 6: ("NV", 5), 7: ("NV", 6), 8: ("NV", 7), 9: ("NV", 8),
    10: ("NV", 9), 12: ("NV", 10),
    14: ("NV", 11), 15: ("NV", 12), 16: ("NV", 14), 17: ("NV", 15),
    18: ("V2", 6),
    20: ("NV", 17), 21: ("NV", 18), 22: ("NV", 19), 23: ("NV", 20),
    24: ("NV", 21), 25: ("NV", 22), 26: ("NV", 23), 28: ("V3", 9),
    30: ("NV", 30),
    32: ("NV", 31), 33: ("NV", 32), 34: ("NV", 33), 35: ("NV", 34),
    36: ("NV", 59), 37: ("NV", 37), 38: ("NV", 42),
    40: ("NV", 43), 41: ("NV", 44), 42: ("NV", 45), 43: ("NV", 46),
    44: ("NV", 47), 46: ("NV", 49),
    48: ("NV", 50), 49: ("NV", 51), 50: ("NV", 52), 51: ("NV", 53),
    53: ("NV", 60), 54: ("NV", 61), 55: ("NV", 62),
    58: ("NV", 64), 59: ("NV", 74),
}
DECKS = {"NV": "Module 6.pptx",
         "APP": "Module 6 -- Slides On-Campus Applications with Solutions.pptx"}
for _k in range(1, 9):
    DECKS["V%d" % _k] = "Module 6 - Video %d.pptx" % _k

# Chrome and boilerplate the rebuild deliberately does not reproduce.
IGNORE = re.compile(
    r'^(module 6|management 405|\d{1,3}|outline of module 6|poll everywhere '
    r'break|prof\.?|professor.*|complex pricing)$')


def texts(path, want):
    z = zipfile.ZipFile(str(path))
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = {r.get('Id'): r.get('Target')
            for r in ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    order = ["ppt/" + rels[s.get('{%s}id' % R_NS)].replace("../", "")
             for s in pres.find('p:sldIdLst', NS)]
    xml = z.read(order[want - 1]).decode("utf-8")
    body = xml[xml.find("<p:spTree"):xml.find("</p:spTree>")]
    out = []
    for para in P_RE.findall(body):
        s = "".join(T_RE.findall(para))
        for a, b in (('&amp;', '&'), ('&lt;', '<'), ('&gt;', '>'),
                     ('&quot;', '"'), ('&apos;', "'")):
            s = s.replace(a, b)
        if s.strip():
            out.append(s)
    return out


def norm(s):
    # PowerPoint writes math italics as separate codepoints (U+1D44F etc.);
    # fold them back to ASCII so an OMML rebuild is not a phantom mismatch.
    s = "".join(chr(ord('A') + (ord(c) - 0x1D434)) if 0x1D434 <= ord(c) <= 0x1D44D
                else chr(ord('a') + (ord(c) - 0x1D44E))
                if 0x1D44E <= ord(c) <= 0x1D467 else c for c in s)
    s = s.lower().replace('’', "'").replace('‘', "'")
    s = s.replace('“', '"').replace('”', '"')
    s = re.sub(r'[^a-z0-9$%<>=./\' ]+', ' ', s)
    return set(w for w in s.split() if len(w) > 2)


def main():
    new = zipfile.ZipFile(str(MODULE / "Module 6 - Full.pptx"))
    pres = ET.fromstring(new.read("ppt/presentation.xml"))
    rels = {r.get('Id'): r.get('Target') for r in
            ET.fromstring(new.read("ppt/_rels/presentation.xml.rels"))}
    order = ["ppt/" + rels[s.get('{%s}id' % R_NS)].replace("../", "")
             for s in pres.find('p:sldIdLst', NS)]
    bad = 0
    for dst in sorted(MAP):
        deck, src = MAP[dst]
        xml = new.read(order[dst - 1]).decode("utf-8")
        body = xml[xml.find("<p:spTree"):xml.find("</p:spTree>")]
        mine = norm(" ".join(T_RE.findall(body)))
        for line in texts(SRC / DECKS[deck], src):
            line = line.strip()
            if not line or IGNORE.match(line.lower().strip()):
                continue
            toks = norm(line)
            # a 1-3 token fragment carries no information about coverage
            if len(toks) < 4:
                continue
            hit = len(toks & mine) / float(len(toks))
            if hit < 0.60:
                bad += 1
                print("slide %-3d (<- %s %d)  %3d%%  %s"
                      % (dst, deck, src, hit * 100, line[:88]))
    print("\n%d source lines under 60%% coverage" % bad)


if __name__ == "__main__":
    main()
