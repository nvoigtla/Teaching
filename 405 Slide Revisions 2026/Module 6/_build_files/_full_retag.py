# -*- coding: utf-8 -*-
"""Create the full slide version, step after _assemble_full.ps1: retag the in-class COPIES.

`Module 6 · Video k · <topic>` -> `Module 6 · In Class · Examples · <topic>`
on the slides between the in-class divider and the moved backup slides.
String-level edit of the <a:t> text only, so the tag box, its autofit
setting and the run formatting stay exactly as they were (a COM edit of
the text re-autofits the box).  Slides whose tag is already four-level
are left alone.
"""
import json, re, shutil, zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import _full_inventory as FI

HERE = Path(__file__).resolve().parent
MODULE = HERE.parent  # _build_files/ since 2026-09-24
FULL = MODULE / "Module 6 - Full.pptx"
plan = json.loads((HERE / "_full_plan.json").read_text(encoding="utf-8"))

with zipfile.ZipFile(FULL) as z:
    items = {n: z.read(n) for n in z.namelist()}
    order = [s["part"] for s in FI.deck(FULL)]

n_main = 93                                  # R2..R94
first = n_main + 2                           # after the divider
last = first + len(plan["copies"]) - 1
pat = re.compile(r"(<a:t>Module 6 · )Video \d+ · ")
n = 0
for disp in range(first, last + 1):
    part = order[disp - 1]
    x = items[part].decode("utf-8")
    x2, k = pat.subn(r"\1In Class · Examples · ", x, count=1)
    if k:
        items[part] = x2.encode("utf-8")
        n += 1
        print("retagged slide", disp)

tmp = FULL.with_suffix(".retag_tmp.pptx")
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
    for name, data in items.items():
        zo.writestr(name, data)
shutil.move(str(tmp), str(FULL))
print("retagged:", n)
