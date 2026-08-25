# -*- coding: utf-8 -*-
"""Offset-aware member diff for the Video deck: canonical display d vs
build display MAP(d).  Nico deleted build slides 8-10 (the Wrigley
chewing-gum example + its poll + solution) on 2026-08-24."""
import sys
import _diff_slides as D

CAN = "Module 2 - Video Part Revised.pptx"
TEST = "Module 2 - Video Part Revised_test.pptx"
TOL = 0.011

def MAP(d):
    return d

def key(s):
    return (s[0], s[1], s[6], s[7])

def main():
    rng = [int(x) for x in sys.argv[1:]] or list(range(1, 55))
    bad = 0
    for d in rng:
        s1, n1 = D.dump(CAN, d)
        s2, n2 = D.dump(TEST, MAP(d))
        msgs = []
        if len(s1) != len(s2):
            msgs.append("shape count %d (canon) vs %d (build)" % (len(s1), len(s2)))
        for i in range(min(len(s1), len(s2))):
            a, b = s1[i], s2[i]
            if key(a) != key(b):
                msgs.append("[%d] SIG %r %r | %r  <>  %r %r | %r"
                            % (i, a[0], a[6][:46], a[7], b[0], b[6][:46], b[7]))
                continue
            for j, lbl in ((2, "x"), (3, "y"), (4, "w"), (5, "h")):
                if abs(a[j] - b[j]) > TOL:
                    msgs.append("[%d] %-22s %s: %.3f canon vs %.3f build"
                                % (i, (a[6][:22] or a[0]), lbl, a[j], b[j]))
        if D.norm(n1) != D.norm(n2):
            msgs.append("NOTES DIFFER")
        if msgs:
            bad += 1
            print("=== canon %d (build %d)" % (d, MAP(d)))
            for m in msgs:
                print("    " + m)
    print("\n%d slide(s) with differences" % bad)

main()
