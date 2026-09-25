# -*- coding: utf-8 -*-
"""2026-09-24 (Nico): displays 12 and 13 of "Module 1 - Final.pptx" must read
`Module 1 · Video 2 · Markets` in the top bar.

Both slides were pulled from "Module 1 - Example Candidates.pptx" into the
taped Video 2 deck and kept that deck's chrome, so they arrived in Full and
then Final reading `Module 1 · Candidates · Market Definition`, over a footer
reading `Management 405 · Module 1 · Example Candidates (for review)`. The
footer is the same leftover and is corrected to the deck's own line, which
every other slide carries.

The In-Class deck does NOT have this defect: its Market Definition block
already reads `Module 1 · In Class · Examples · Markets`, the in-class
convention, and neither of these two slides exists there.

Pure zip + lxml surgery: only the two slide parts are rewritten, every other
part is copied through byte for byte, and the deck is never round-tripped
through python-pptx. Each target run is a single run whose rPr already
matches the correct neighbour (display 15), so only the text changes.

Refuses to act unless it finds exactly the expected wording.
"""
import shutil
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).resolve().parent
MODULE = HERE.parent              # _build_files/ since 2026-09-24
DECK = MODULE / "Module 1 - Final.pptx"
TARGETS = [12, 13]

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

OLD_TAG = "Module 1 · Candidates · Market Definition"
NEW_TAG = "Module 1 · Video 2 · Markets"
OLD_FOOT = ("Management 405  ·  Module 1  ·  "
            "Example Candidates (for review)")
NEW_FOOT = ("Management 405  ·  Module 1  ·  "
            "Basic Concepts and Economic Principles")


def q(ns, t):
    return "{%s}%s" % (ns, t)


def roll_backups(deck):
    t1 = deck.with_name(deck.stem + "_t-1.pptx")
    t2 = deck.with_name(deck.stem + "_t-2.pptx")
    if t1.exists():
        shutil.copy2(t1, t2)
    shutil.copy2(deck, t1)


def main():
    zin = zipfile.ZipFile(DECK)
    items = {n: zin.read(n) for n in zin.namelist()}
    pres = ET.fromstring(items["ppt/presentation.xml"])
    prels = {r.get("Id"): r.get("Target") for r in
             ET.fromstring(items["ppt/_rels/presentation.xml.rels"])}
    order = ["ppt/" + prels[s.get(q(R, "id"))].lstrip("/").replace("../", "")
             for s in pres.find(q(P, "sldIdLst"))]
    zin.close()

    changed = []
    for disp in TARGETS:
        part = order[disp - 1]
        tree = ET.fromstring(items[part])
        hits = []
        for t in tree.iter(q(A, "t")):
            if t.text == OLD_TAG:
                t.text = NEW_TAG
                hits.append("top bar")
            elif t.text == OLD_FOOT:
                t.text = NEW_FOOT
                hits.append("footer")
        if sorted(hits) != ["footer", "top bar"]:
            raise SystemExit(
                "display %d: expected one top bar and one footer to replace, "
                "found %r - stopping rather than guessing" % (disp, hits))
        items[part] = ET.tostring(tree, xml_declaration=True,
                                  encoding="UTF-8", standalone=True)
        changed.append((disp, part.split("/")[-1]))

    roll_backups(DECK)
    tmp = DECK.with_suffix(".tmp.pptx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in items.items():
            zout.writestr(name, blob)
    shutil.move(str(tmp), str(DECK))
    for disp, part in changed:
        print("display %-3d %-14s top bar -> %r" % (disp, part, NEW_TAG))
        print("               %-14s footer  -> %r" % ("", NEW_FOOT))
    print("parts rewritten: %d of %d" % (len(changed), len(items)))


if __name__ == "__main__":
    main()
