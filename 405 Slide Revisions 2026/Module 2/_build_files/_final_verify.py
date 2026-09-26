# -*- coding: utf-8 -*-
"""Create the final slide version, step 2: verify "Module 2 - Final.pptx".

Two checks, both must pass:
  1. Final 1-64 diff clean against `Module 2 - In Class.pptx` 1-64
     (step 2's own rule 2);
  2. Final 65-109 diff clean against `Module 2 - Full.pptx` 72-116, the
     frozen video slides - the guard that step 2 did not disturb step 1
     (rule 4). In Module 2 these are Revised's video slides (step 1,
     option 1), so Full rather than the taped decks is their reference; the
     taped decks are checked for click counts by _final_verify_anim.ps1.

Reuses the reader in `_full_verify.py`, which decodes group transforms and
drops live `slidenum` fields. Read-only.
"""
# NOTE (2026-09-25): "Module 2 - Full.pptx" and "Module 2 - Revised.pptx"
# were deleted by the step-1b cleanup. Full was a byte copy of Revised, so
# to use this again restore Revised from git and copy it:
#   git checkout b5484ae5 -- "405 Slide Revisions 2026/Module 2/Module 2 - Revised.pptx"
#   copy "Module 2 - Revised.pptx" "Module 2 - Full.pptx"
import sys

from _full_verify import Deck, MODULE

CHECKS = ("text", "runs", "shapes", "notes", "timing", "rels", "hidden")


def main():
    fin = Deck(MODULE / "Module 2 - Final.pptx")
    ic = Deck(MODULE / "Module 2 - In Class.pptx")
    # Full was deleted by the step-1b cleanup once Final superseded it, so
    # check 2 is skipped unless it has been regenerated (see the note above).
    fp = MODULE / "Module 2 - Full.pptx"
    full = Deck(fp) if fp.exists() else None
    if len(fin.order) != 109:
        print("FAIL slide count: %d, expected 109" % len(fin.order))
        return 1
    print("slide count 109 = In Class 64 + Full video slides 45   OK\n")
    bad = 0
    for label, pairs in (
            ("1. Final vs IN CLASS", [(f, ic, f) for f in range(1, 65)]),
            ("2. Final vs FULL (video slides)",
             [(f, full, f + 7) for f in range(65, 110)] if full else [])):
        if not pairs:
            print("%-34s SKIPPED - Module 2 - Full.pptx is gone (regenerate it to run this)" % label)
            continue
        ok = 0
        for f, src, j in pairs:
            ch = [w for w in CHECKS if getattr(fin, w)(f) != getattr(src, w)(j)]
            if ch:
                bad += 1
                print("  DIFF Final %3d <- %-26s %-3d  %s"
                      % (f, src.name, j, ", ".join(ch)))
            else:
                ok += 1
        print("%-34s %3d of %3d clean" % (label, ok, len(pairs)))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
