# -*- coding: utf-8 -*-
"""Compact on-slide text dump for the given display slides."""
import sys
import _diff_slides as D
DECK = "Module 2 - In Class Revised.pptx"
for d in [int(x) for x in sys.argv[1:]]:
    shapes, notes = D.dump(DECK, d)
    title = ""
    for s in shapes:
        if s[0] == "sp" and 0.45 < s[3] < 1.30 and s[6]:
            title = s[6]; break
    print("### %d  %s" % (d, title))
    seen = set()
    for s in shapes:
        t = s[6]
        if not t or t == title or t in seen:
            continue
        if t.startswith("Management 405") or t.startswith("Module 2 ·"):
            continue
        if t.isdigit() and len(t) <= 2:
            continue
        seen.add(t)
        print("    [%s] %s" % (s[0], t[:520]))
    npic = sum(1 for s in shapes if s[0] == "pic")
    if npic:
        print("    (%d picture(s))" % npic)
