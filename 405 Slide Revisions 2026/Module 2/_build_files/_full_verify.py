# -*- coding: utf-8 -*-
"""Create the full slide version, step 1: verify "Module 2 - Full.pptx".

Every slide of Full must diff clean against its source - the taped slides
against their video deck, everything else against "Module 1 - Revised.pptx".
The footer slide-number field's cached text is excluded, since it is a live
slidenum field whose cache moves with the slide's position.

Read-only. Run from the module folder (or from _build_files/, where
MODULE = HERE.parent).
"""
# NOTE (2026-09-25): "Module 2 - Full.pptx" and "Module 2 - Revised.pptx"
# were deleted by the step-1b cleanup. Full was a byte copy of Revised, so
# to use this again restore Revised from git and copy it:
#   git checkout b5484ae5 -- "405 Slide Revisions 2026/Module 2/Module 2 - Revised.pptx"
#   copy "Module 2 - Revised.pptx" "Module 2 - Full.pptx"
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).resolve().parent
MODULE = HERE.parent              # _build_files/ since 2026-09-25
NP = "http://schemas.openxmlformats.org/presentationml/2006/main"
NA = "http://schemas.openxmlformats.org/drawingml/2006/main"
NR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0

# Module 2 (2026-09-25, Nico's option 1): the taped decks carry NO edits made
# during taping - they are byte-identical to the "Videos Final" decks of
# 2026-08-25, and Revised 72-116 was built from them plus later approved edits
# (tags, coverage pills, title case, dark-red answers, Fade).  So Full is
# Revised as it stands: every slide's source is the same Revised slide.
# The taped-vs-Revised comparison lives in _full_inventory.py.
MAP = {_i: ("R", _i) for _i in range(1, 117)}


def q(ns, t):
    return "{%s}%s" % (ns, t)


class Deck:
    def __init__(self, path):
        self.name = Path(path).name
        self.z = zipfile.ZipFile(path)
        pres = ET.fromstring(self.z.read("ppt/presentation.xml"))
        rels = {r.get("Id"): r.get("Target") for r in
                ET.fromstring(self.z.read("ppt/_rels/presentation.xml.rels"))}
        self.order = ["ppt/" + rels[s.get(q(NR, "id"))].lstrip("/").replace("../", "")
                      for s in pres.find(q(NP, "sldIdLst"))]

    def root(self, i):
        return ET.fromstring(self.z.read(self.order[i - 1]))

    def hidden(self, i):
        return self.root(i).get("show") == "0"

    def text(self, i):
        out = []
        for p in self.root(i).iter(q(NA, "p")):
            t = " ".join("".join(x.text or "" for x in p.iter(q(NA, "t"))).split())
            if t and not t.isdigit():
                out.append(t)
        return out

    def runs(self, i):
        out = []
        for rn in self.root(i).iter(q(NA, "r")):
            t = (rn.findtext(q(NA, "t")) or "").strip()
            if not t or t.isdigit():
                continue
            pr = rn.find(q(NA, "rPr"))
            sz = b = it = u = col = ""
            if pr is not None:
                sz, b, it, u = (pr.get("sz") or "", pr.get("b") or "",
                                pr.get("i") or "", pr.get("u") or "")
                sf = pr.find(q(NA, "solidFill"))
                if sf is not None:
                    c = sf.find(q(NA, "srgbClr"))
                    col = c.get("val") if c is not None else "scheme"
            out.append((t, sz, b, it, u, col))
        return out

    def shapes(self, i):
        out = []

        def walk(node, off, ext, choff, chext):
            for ch in node:
                tag = ch.tag.split("}")[-1]
                if tag not in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
                    continue
                xf = ch.find(".//" + q(NA, "xfrm"))
                if xf is None:
                    continue
                o, e = xf.find(q(NA, "off")), xf.find(q(NA, "ext"))
                if o is None or e is None:
                    continue
                x, y = int(o.get("x")), int(o.get("y"))
                cx, cy = int(e.get("cx")), int(e.get("cy"))
                if chext:
                    sx = ext[0] / chext[0] if chext[0] else 1
                    sy = ext[1] / chext[1] if chext[1] else 1
                    x, y = off[0] + (x - choff[0]) * sx, off[1] + (y - choff[1]) * sy
                    cx, cy = cx * sx, cy * sy
                if tag == "grpSp":
                    co, ce = xf.find(q(NA, "chOff")), xf.find(q(NA, "chExt"))
                    if co is not None and ce is not None:
                        walk(ch, (x, y), (cx, cy),
                             (int(co.get("x")), int(co.get("y"))),
                             (int(ce.get("cx")), int(ce.get("cy"))))
                    continue
                # a live slidenum field recomputes with the slide's position,
                # so its cached text is not a difference between the decks
                if any(f.get("type") == "slidenum" for f in ch.iter(q(NA, "fld"))):
                    continue
                txt = " ".join(
                    "".join(t.text or "" for t in ch.iter(q(NA, "t"))).split())
                geom = ch.find(".//" + q(NA, "prstGeom"))
                out.append((tag, round(x / EMU, 2), round(y / EMU, 2),
                            round(cx / EMU, 2), round(cy / EMU, 2),
                            geom.get("prst") if geom is not None else "",
                            txt[:60]))

        walk(self.root(i).find(q(NP, "cSld") + "/" + q(NP, "spTree")),
             (0, 0), (0, 0), (0, 0), (0, 0))
        return sorted(out)

    def rels(self, i):
        relp = self.order[i - 1].replace("slides/", "slides/_rels/") + ".rels"
        return sorted(r.get("Type").rsplit("/", 1)[-1]
                      for r in ET.fromstring(self.z.read(relp)))

    def notes(self, i):
        relp = self.order[i - 1].replace("slides/", "slides/_rels/") + ".rels"
        for r in ET.fromstring(self.z.read(relp)):
            if r.get("Type").endswith("/notesSlide"):
                n = ET.fromstring(
                    self.z.read("ppt/" + r.get("Target").replace("../", "")))
                t = " ".join("".join(x.text or "" for x in n.iter(q(NA, "t"))).split())
                return " ".join(w for w in t.split() if not w.isdigit())
        return ""

    def timing(self, i):
        """every animation node under the slide's timing tree, in order."""
        tm = self.root(i).find(q(NP, "timing"))
        if tm is None:
            return []
        return [(e.tag.split("}")[-1], e.get("presetID") or "",
                 e.get("presetClass") or "", e.get("transition") or "")
                for e in tm.iter()
                if e.tag.split("}")[-1] in ("animEffect", "set", "anim", "animMotion")]


def main():
    full = Deck(MODULE / "Module 2 - Full.pptx")
    src = {"R": Deck(MODULE / "Module 2 - Revised.pptx")}
    if len(full.order) != 116:
        print("FAIL slide count: %d, expected 116" % len(full.order))
        return 1
    print("slide count 116 = Revised 116   OK")

    checks = ("text", "runs", "shapes", "notes", "timing", "rels", "hidden")
    bad = 0
    for i in sorted(MAP):
        key, j = MAP[i]
        d = src[key]
        ch = [w for w in checks if getattr(full, w)(i) != getattr(d, w)(j)]
        if ch:
            bad += 1
            print("  DIFF %3d  <- %-40s %-3d  %s" % (i, d.name, j, ", ".join(ch)))
    print("\nslides diffing clean against their source: %d of %d"
          % (len(MAP) - bad, len(MAP)))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
