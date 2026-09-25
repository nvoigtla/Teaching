# -*- coding: utf-8 -*-
"""Read-only: CONCLUSIVE check - normalized raw slide XML, pair by pair.

`_diff2.py` sees position / size / text / run formats. `_style_diff.py` adds
geometry, fills, lines, effects and table-cell fills. Neither sees paragraph
spacing, bullet properties, text-box insets, autofit, picture crops or
z-order. This compares the whole slide XML instead, so nothing can hide.

Normalized away, because they differ between the two decks BY DESIGN:
  * shape ids and shape names   (PowerPoint renumbers them per deck)
  * relationship ids            (r:embed / r:id / r:link)
  * the cached text of a <a:fld type="slidenum"> (the live page number)
  * the creationId / editing extLst blobs PowerPoint stamps per save
  * whitespace and attribute order (lxml c14n)

Usage:  python _xml_diff.py [--show IC]
"""
# NOTE (2026-09-24): this reads "Module 1 - Revised.pptx", which the
# step-1b cleanup deleted. To use it again, restore that deck from git
# into the MODULE folder first:
#   git checkout <commit> -- "405 Slide Revisions 2026/Module 1/Module 1 - Revised.pptx"
import re
import sys
import zipfile
from io import BytesIO

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, ".")
from lxml import etree as ET                                # noqa: E402
import _diff2 as D                                          # noqa: E402
from _ic_vs_rev import MAP                                  # noqa: E402

A = D.A
P = D.P


def q(ns, t):
    return "{%s}%s" % (ns, t)


def normalize(xml):
    tree = ET.fromstring(xml)
    # shape ids / names
    for nv in tree.iter(q(P, "cNvPr")):
        nv.set("id", "0")
        nv.set("name", "x")
    # relationship ids
    for el in tree.iter():
        for k in list(el.attrib):
            if k.endswith("}embed") or k.endswith("}id") or k.endswith("}link"):
                el.set(k, "rId")
    # the live page number's cached text
    for fld in tree.iter(q(A, "fld")):
        if fld.get("type") == "slidenum":
            fld.set("id", "{GUID}")
            t = fld.find(q(A, "t"))
            if t is not None:
                t.text = "#"
    # per-save extension blobs (creationId, modId, and friends)
    for ext in list(tree.iter(q(A, "extLst"))) + list(tree.iter(q(P, "extLst"))):
        ext.getparent().remove(ext)
    buf = BytesIO()
    ET.ElementTree(tree).write_c14n(buf)
    out = buf.getvalue().decode("utf-8", "replace")
    return re.sub(r"\s+", " ", out)


def main():
    show = None
    if "--show" in sys.argv:
        show = int(sys.argv[sys.argv.index("--show") + 1])
    zi = zipfile.ZipFile("Module 1 - In Class.pptx")
    zr = zipfile.ZipFile("Module 1 - Revised.pptx")
    pi, pr = D.slide_parts(zi), D.slide_parts(zr)
    bad = []
    for i in sorted(MAP):
        j = MAP[i]
        if j is None:
            continue
        a = normalize(zi.read(pi[i - 1]))
        b = normalize(zr.read(pr[j - 1]))
        if a != b:
            bad.append((i, j, len(a), len(b)))
            if show == i:
                import difflib
                for line in difflib.unified_diff(
                        a.split("><"), b.split("><"),
                        "IC %d" % i, "RV %d" % j, lineterm="", n=1):
                    print(line[:300])
    print("pairs compared: %d" % len([1 for v in MAP.values() if v]))
    if not bad:
        print("IDENTICAL at the XML level - nothing is missing.")
    else:
        print("differ: %d" % len(bad))
        for i, j, la, lb in bad:
            print("   IC %2d -> RV %2d   (%d vs %d chars)" % (i, j, la, lb))
        print("\nrun with --show <IC number> to see one in detail")


if __name__ == "__main__":
    main()
