# -*- coding: utf-8 -*-
"""Finalize routine helper (read-only): print the text of taped slides and
Revised slides side by side, ignoring cached page numbers.

    python _final_detail.py V2.7 R15 R12 ...
    python _final_detail.py --diffs      # every pair whose text differs
"""
import json, re, sys, difflib
from pathlib import Path
import _final_inventory as FI

HERE = Path(__file__).resolve().parent


def clean(texts):
    return [t for t in texts if not re.fullmatch(r"\d+", t)]


rev = FI.deck(FI.REVISED)
vd = {}
for p in FI.VDIR.glob("Module 6 - Video *.pptx"):
    k = int(re.search(r"Video (\d+)", p.name).group(1))
    vd[k] = FI.deck(p)


def get(tok):
    if tok[0] == "R":
        return rev[int(tok[1:]) - 1]
    k, n = tok[1:].split(".")
    return vd[int(k)][int(n) - 1]


if sys.argv[1:] == ["--diffs"]:
    pr = json.loads((HERE / "_final_pairing.json").read_text(encoding="utf-8"))
    for p in pr["pairing"]:
        if not p["rev"]:
            continue
        a = clean(rev[p["rev"] - 1]["texts"])
        b = clean(vd[p["video"]][p["vn"] - 1]["texts"])
        na = rev[p["rev"] - 1]["notes"]
        nb = vd[p["video"]][p["vn"] - 1]["notes"]
        if a != b or na != nb:
            print("=== V%d.%d vs R%d" % (p["video"], p["vn"], p["rev"]))
            for l in difflib.unified_diff(a, b, lineterm="", n=0):
                if not l.startswith(("---", "+++", "@@")):
                    print("   ", l[:160])
            if na != nb:
                print("    [notes differ]")
    sys.exit()

for tok in sys.argv[1:]:
    s = get(tok)
    print("=== %s  clicks=%d pics=%d tags=%s jumps=%d" % (
        tok, s["clicks"], s["pics"], s["tags"], s["jumps"]))
    for t in clean(s["texts"]):
        print("   ", t[:150])
    print("    NOTES:", s["notes"][:300].replace("\n", " / "))
