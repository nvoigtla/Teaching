# -*- coding: utf-8 -*-
"""Per-slide inventory: title, on-slide text, and current speaker notes."""
import sys, zipfile
from lxml import etree as ET
import _diff_slides as D

DECK = sys.argv[1] if len(sys.argv) > 1 else "Module 2 - In Class Revised.pptx"
mode = sys.argv[2] if len(sys.argv) > 2 else "short"
z = zipfile.ZipFile(DECK)
pres = ET.fromstring(z.read("ppt/presentation.xml"))
n = len(pres.find(D.q(D.P, "sldIdLst")))
for d in range(1, n + 1):
    shapes, notes = D.dump(DECK, d)
    title = ""
    for s in shapes:
        if s[0] == "sp" and 0.45 < s[3] < 1.30 and s[6]:
            title = s[6]; break
    body = " | ".join(s[6] for s in shapes
                      if s[6] and s[6] != title
                      and not s[6].startswith("Management 405")
                      and not s[6].startswith("Module 2 ·")
                      and not (s[6].isdigit() and len(s[6]) <= 2))
    nt = D.norm(notes)
    if mode == "short":
        print("%2d | %-52s | notes=%4d | %s" % (d, title[:52], len(nt), nt[:70]))
    else:
        print("=" * 78)
        print("SLIDE %d: %s" % (d, title))
        print("  BODY: %s" % body[:900])
        print("  NOTES(%d): %s" % (len(nt), nt[:1500]))
