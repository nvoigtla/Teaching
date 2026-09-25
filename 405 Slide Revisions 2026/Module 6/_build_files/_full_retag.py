# -*- coding: utf-8 -*-
"""Finalize routine, step after _finalize.ps1: retag the in-class COPIES.

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
import _final_inventory as FI

HERE = Path(__file__).resolve().parent
FINAL = HERE / "Module 6 - Final.pptx"
plan = json.loads((HERE / "_final_plan.json").read_text(encoding="utf-8"))

with zipfile.ZipFile(FINAL) as z:
    items = {n: z.read(n) for n in z.namelist()}
    order = [s["part"] for s in FI.deck(FINAL)]

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

tmp = FINAL.with_suffix(".retag_tmp.pptx")
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
    for name, data in items.items():
        zo.writestr(name, data)
shutil.move(str(tmp), str(FINAL))
print("retagged:", n)
