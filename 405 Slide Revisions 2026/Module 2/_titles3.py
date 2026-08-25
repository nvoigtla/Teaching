# -*- coding: utf-8 -*-
"""Title列 for up to 3 decks side by side."""
import sys, zipfile
from lxml import etree as ET
import _diff_slides as D

def nsl(deck):
    z = zipfile.ZipFile(deck); n = len(ET.fromstring(z.read("ppt/presentation.xml")).find(D.q(D.P, "sldIdLst"))); z.close(); return n

def titles(deck):
    out = []
    for d in range(1, nsl(deck) + 1):
        s, _ = D.dump(deck, d)
        t = ""
        for sh in s:
            if sh[0] == "sp" and 0.3 < sh[3] < 1.60 and sh[6]:
                t = sh[6]; break
        if not t:
            for sh in s:
                if sh[6] and len(sh[6]) > 3:
                    t = sh[6]; break
        out.append(t)
    return out

decks = sys.argv[1:]
cols = [titles(d) for d in decks]
n = max(len(c) for c in cols)
print(" # | " + " | ".join("%-46s" % d[:46] for d in decks))
for i in range(n):
    print("%2d | " % (i + 1) + " | ".join("%-46s" % (c[i][:46] if i < len(c) else "") for c in cols))
