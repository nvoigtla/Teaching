"""Full-deck member-level geometry+text diff: canonical (hand-edited)
vs. the fresh side-path build. Prints ONLY mismatching slides.
Reuses the decode/dump machinery in _diff_slides.py."""
import os
import sys
import zipfile

import _diff_slides as D

HERE = os.path.dirname(os.path.abspath(__file__))
CAN = os.path.join(HERE, "Module 2 - In Class Revised.pptx")
TEST = os.path.join(HERE, "Module 2 - In Class Revised_test.pptx")
TOL = 0.011  # inches


def n_slides(deck):
    z = zipfile.ZipFile(deck)
    from lxml import etree as ET
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    n = len(pres.find(D.q(D.P, "sldIdLst")))
    z.close()
    return n


def key(s):
    return (s[0], s[1], s[6], s[7])


def main():
    n1, n2 = n_slides(CAN), n_slides(TEST)
    print("slides: canonical %d | build %d" % (n1, n2))
    if n1 != n2:
        print("!! SLIDE COUNT DIFFERS")
    bad = 0
    for d in range(1, min(n1, n2) + 1):
        s1, note1 = D.dump(CAN, d)
        s2, note2 = D.dump(TEST, d)
        msgs = []
        if len(s1) != len(s2):
            msgs.append("shape count %d vs %d" % (len(s1), len(s2)))
        # pair positionally; flag any member whose signature or geometry moved
        for i in range(min(len(s1), len(s2))):
            a, b = s1[i], s2[i]
            if key(a) != key(b):
                msgs.append("[%d] signature: %r/%r  vs  %r/%r"
                            % (i, a[0], a[6][:45], b[0], b[6][:45]))
                continue
            for j, lbl in ((2, "x"), (3, "y"), (4, "w"), (5, "h")):
                if abs(a[j] - b[j]) > TOL:
                    msgs.append("[%d] %-22s %s: %.3f (canonical) vs %.3f (build)"
                                % (i, a[6][:22] or a[0], lbl, a[j], b[j]))
        if D.norm(note1) != D.norm(note2):
            msgs.append("NOTES differ")
        if msgs:
            bad += 1
            print("=== display %d" % d)
            for m in msgs:
                print("    " + m)
    print("\n%d slide(s) with differences" % bad)


if __name__ == "__main__":
    main()
