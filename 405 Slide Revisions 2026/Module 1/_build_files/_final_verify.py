# -*- coding: utf-8 -*-
"""Create the final slide version, step 2: verify "Module 1 - Final.pptx".

Three checks, all of which must pass:
  1. every slide of `Module 1 - In Class.pptx` diffs clean against its
     counterpart in Final (step 2's own rule 2);
  2. every TAPED slide in Final still diffs clean against its video deck -
     the guard that step 2 did not disturb step 1 (rule 4);
  3. the slides In Class dropped but rule 6 kept still diff clean against
     `Module 1 - Full.pptx`.

Reuses the reader in `_full_verify.py`, which decodes group transforms and
drops live `slidenum` fields. Read-only.
"""
# NOTE (2026-09-24): this reads "Module 1 - Full.pptx", which the
# step-1b cleanup deleted once Final superseded it. Full is not in git
# either - it is REGENERATED, not restored:
#   git checkout <commit> -- "405 Slide Revisions 2026/Module 1/Module 1 - Revised.pptx"
#   powershell -File _build_files/_assemble_full.ps1
import sys
from pathlib import Path

from _full_verify import Deck, MODULE, V1, V2, V3, V4

CHECKS = ("text", "runs", "shapes", "notes", "timing", "rels", "hidden")

# Final slide -> In-Class slide.  The in-class block runs 34-85 in IN-CLASS
# order; the backup slides follow at 88-90 and 92.
IC_MAP = {}
for _f, _i in [(34, 1), (35, 2), (36, 3), (37, 4), (38, 5), (39, 6), (40, 7),
               (41, 8), (42, 9), (43, 10), (44, 11), (45, 12), (46, 13),
               (47, 14), (48, 15), (49, 16), (50, 17), (51, 18), (52, 19)]:
    IC_MAP[_f] = _i
for _f in range(53, 86):              # Final 53-85 <- In Class 20-52
    IC_MAP[_f] = _f - 33
IC_MAP[87] = 53                       # BACKUP divider (taped copy, identical)
IC_MAP[88] = 54
IC_MAP[89] = 55
IC_MAP[90] = 56
IC_MAP[92] = 57

# Final slide -> (video deck, slide).  Step 1's map, shifted by the one
# in-class-only slide inserted at 40.
TAPED = {}
for _i in range(1, 9):
    TAPED[_i] = (V1, _i + 1)
for _i in range(9, 16):
    TAPED[_i] = (V2, _i - 8)
for _i in range(16, 26):
    TAPED[_i] = (V3, _i - 15)
for _i in range(26, 33):
    TAPED[_i] = (V4, _i - 25)
TAPED[87] = (V1, 10)                  # BACKUP divider
TAPED[91] = (V1, 11)                  # People Respond to Incentives

# Final slide -> Full slide, for what In Class dropped and rule 6 kept.
KEPT = {33: 33, 86: 85, 91: 90}

# Deliberate deviations from a taped deck, each with the reason. These are
# the ONLY slides allowed to differ from their source; anything else is a
# failure. 2026-09-24 (Nico): both slides were pulled from "Module 1 -
# Example Candidates.pptx" into the taped Video 2 deck and kept that deck's
# chrome, so they arrived reading `Module 1 · Candidates · Market Definition`
# over an "Example Candidates (for review)" footer. `_fix_candidates_chrome.py`
# corrected both to the deck's own chrome; the taped deck still carries the
# original, so this diff is expected and must stay visible.
ACCEPTED = {
    12: "top bar + footer corrected from Example-Candidates chrome",
    13: "top bar + footer corrected from Example-Candidates chrome",
}


def main():
    fin = Deck(MODULE / "Module 1 - Final.pptx")
    ic = Deck(MODULE / "Module 1 - In Class.pptx")
    vd = {v: Deck(MODULE / "Recorded Video Slides" / v) for v in (V1, V2, V3, V4)}
    # Full was deleted by the step-1b cleanup once Final superseded it, so
    # check 3 is skipped unless it has been regenerated (see the note above).
    fp = MODULE / "Module 1 - Full.pptx"
    full = Deck(fp) if fp.exists() else None

    if len(fin.order) != 92:
        print("FAIL slide count: %d, expected 92" % len(fin.order))
        return 1
    print("slide count 92 = 91 Full + 1 in-class-only slide   OK\n")

    bad = 0
    for label, pairs in (
            ("1. Final vs IN CLASS", [(f, ic, i) for f, i in sorted(IC_MAP.items())]),
            ("2. Final vs the TAPED decks",
             [(f, vd[d], j) for f, (d, j) in sorted(TAPED.items())]),
            ("3. Final vs FULL (rule 6 keeps)",
             [(f, full, j) for f, j in sorted(KEPT.items())] if full else [])):
        if not pairs:
            print("%-34s SKIPPED - %s is gone (regenerate it to run this)"
                  % (label, "Module 1 - Full.pptx"))
            continue
        n, ok = 0, 0
        for f, src, j in pairs:
            ch = [w for w in CHECKS if getattr(fin, w)(f) != getattr(src, w)(j)]
            if not ch:
                ok += 1
            elif f in ACCEPTED and src is not ic:
                n += 1
                print("  note Final %3d <- %-34s %-3d  %s  [%s]"
                      % (f, src.name, j, ", ".join(ch), ACCEPTED[f]))
            else:
                bad += 1
                print("  DIFF Final %3d <- %-34s %-3d  %s"
                      % (f, src.name, j, ", ".join(ch)))
        tail = " + %d accepted deviation(s)" % n if n else ""
        print("%-34s %3d of %3d clean%s" % (label, ok, len(pairs), tail))

    print("\nslides verified: %d" % (len(IC_MAP) + len(TAPED) + len(KEPT)))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
