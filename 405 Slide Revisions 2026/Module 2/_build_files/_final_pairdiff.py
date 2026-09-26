# -*- coding: utf-8 -*-
"""Create the final slide version, step 2, rule 2 (read-only): for every
pairing that `_final_inventory.py` found, say what REALLY differs between
the In-Class slide and its counterpart in Full.

Uses the step-1 reader in `_full_verify.py`, which decodes group transforms
and drops live `slidenum` fields - their cached text moves with a slide's
position and is not a difference between the decks.

Pass --detail to print the changed text lines as well.
"""
# NOTE (2026-09-25): "Module 2 - Full.pptx" and "Module 2 - Revised.pptx"
# were deleted by the step-1b cleanup. Full was a byte copy of Revised, so
# to use this again restore Revised from git and copy it:
#   git checkout b5484ae5 -- "405 Slide Revisions 2026/Module 2/Module 2 - Revised.pptx"
#   copy "Module 2 - Revised.pptx" "Module 2 - Full.pptx"
import difflib
import json
import sys
from pathlib import Path

from _full_verify import Deck, MODULE

HERE = Path(__file__).resolve().parent
CHECKS = ("text", "runs", "shapes", "notes", "timing", "rels", "hidden")


def main():
    detail = "--detail" in sys.argv
    pair = json.loads((HERE / "_final_pairing.json").read_text(encoding="utf8"))
    full = Deck(MODULE / "Module 2 - Full.pptx")
    ic = Deck(MODULE / "Module 2 - In Class.pptx")

    changed, clean, skipped = [], [], []
    for p in pair["pairing"]:
        i, j = p["ic"], p["full"]
        if j is None:
            skipped.append((i, "NEW - no counterpart in Full"))
            continue
        if p["frozen"]:
            ch = [w for w in CHECKS
                  if getattr(ic, w)(i) != getattr(full, w)(j)]
            skipped.append((i, "frozen video block F%d (%s)"
                            % (j, ", ".join(ch) if ch else "identical")))
            continue
        ch = [w for w in CHECKS if getattr(ic, w)(i) != getattr(full, w)(j)]
        (changed if ch else clean).append((i, j, ch))

    print("IN-CLASS slides that REPLACE their Full counterpart (%d):" % len(changed))
    for i, j, ch in changed:
        title = ((ic.text(i) + ["<no text>", ""])[1] if len(ic.text(i)) > 1 else (ic.text(i) or ["<no text>"])[0])[:46]
        print("  IC%-3d -> F%-3d  %-34s %s" % (i, j, ", ".join(ch), title))
        if detail:
            for line in difflib.unified_diff(full.text(j), ic.text(i),
                                             "FULL", "IN CLASS", lineterm="", n=0):
                if line[:3] in ("---", "+++") or line[:2] == "@@":
                    continue
                print("        %s" % line[:100])

    print("\nIN-CLASS slides already identical in Full (%d): %s"
          % (len(clean), ", ".join("IC%d" % i for i, _, _ in clean)))
    print("\nNOT acted on (%d):" % len(skipped))
    for i, why in skipped:
        title = ((ic.text(i) + ["<no text>", ""])[1] if len(ic.text(i)) > 1 else (ic.text(i) or ["<no text>"])[0])[:46]
        print("  IC%-3d  %-44s %s" % (i, why, title))


if __name__ == "__main__":
    main()
