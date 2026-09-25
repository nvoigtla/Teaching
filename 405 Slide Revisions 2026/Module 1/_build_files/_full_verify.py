# -*- coding: utf-8 -*-
"""Create the full slide version, step 1: verify "Module 1 - Full.pptx".

Every slide of Full must diff clean against its source - the taped slides
against their video deck, everything else against "Module 1 - Revised.pptx".
The footer slide-number field's cached text is excluded, since it is a live
slidenum field whose cache moves with the slide's position.

Read-only. Run from the module folder (or from _build_files/, where
MODULE = HERE.parent).
"""
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).resolve().parent
MODULE = HERE.parent              # _build_files/ since 2026-09-24
NP = "http://schemas.openxmlformats.org/presentationml/2006/main"
NA = "http://schemas.openxmlformats.org/drawingml/2006/main"
NR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0

V1 = "Module 1 - Video 1 - Introduction.pptx"
V2 = "Module 1 - Video 2 - Markets.pptx"
V3 = "Module 1 - Video 3 - Demand and Supply.pptx"
V4 = "Module 1 - Video 4 - Equilibrium.pptx"

# Full slide -> (source deck, source slide).  "R" is Module 1 - Revised.pptx.
MAP = {}
for _i in range(1, 9):
    MAP[_i] = (V1, _i + 1)        # Video 1 slides 2-9; its title card is dropped
for _i in range(9, 16):
    MAP[_i] = (V2, _i - 8)
for _i in range(16, 26):
    MAP[_i] = (V3, _i - 15)
for _i in range(26, 33):
    MAP[_i] = (V4, _i - 25)
for _i in range(33, 86):
    MAP[_i] = ("R", _i + 1)       # R34-R86: the hidden slide and the in-class half
MAP[86] = (V1, 10)                # BACKUP divider, taped with Video 1
MAP[87] = ("R", 88)
MAP[88] = ("R", 89)
MAP[89] = ("R", 90)
MAP[90] = (V1, 11)                # People Respond to Incentives, taped with Video 1
MAP[91] = ("R", 92)


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
    full = Deck(MODULE / "Module 1 - Full.pptx")
    src = {"R": Deck(MODULE / "Module 1 - Revised.pptx")}
    for v in (V1, V2, V3, V4):
        src[v] = Deck(MODULE / "Recorded Video Slides" / v)

    if len(full.order) != 91:
        print("FAIL slide count: %d, expected 91" % len(full.order))
        return 1
    print("slide count 91 = 92 Revised - 1 dropped Video 1 title card   OK")

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
