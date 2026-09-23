# -*- coding: utf-8 -*-
"""Read-only: In-Class deck (dominant) vs the in-class part of Revised.

Uses an EXPLICIT map, because a text-similarity matcher pairs in-class
slides with their near-identical twins in the VIDEO block of the Revised
deck (Netflix, the hedgehogs slide, the outline slides all appear twice).
Every pairing is re-scored and anything weak is flagged.

Usage:  python _ic_vs_rev.py [--full]
"""
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, ".")
import _diff2 as D                                          # noqa: E402

# In-Class display -> Revised display  (None = no twin in Revised)
MAP = {
    1: 37, 2: 36, 3: 38, 4: 39, 5: 40, 6: 41, 7: None, 8: 42,
    9: 43, 10: 44, 11: 45, 12: 46, 13: 47, 14: 35, 15: 48, 16: 49,
    17: 50, 18: 51, 19: 52, 20: 53, 21: 54, 22: 55, 23: 56, 24: 57,
    25: 58, 26: 59, 27: 60, 28: 61, 29: 62, 30: 63, 31: 64, 32: 65,
    33: 66, 34: 67, 35: 68, 36: 69, 37: 70, 38: 71, 39: 72, 40: 73,
    41: 74, 42: 75, 43: 76, 44: 77, 45: 78, 46: 79, 47: 80, 48: 81,
    49: 82, 50: 83, 51: 84, 52: 85, 53: 86, 54: 87, 55: 88, 56: 89,
    57: 90, 58: 92,
}

# Revised slides in the in-class part with no In-Class twin
REV_ONLY = {91: "Backup: People Respond to Incentives (kept, RV 3 links to it)"}


def ident(d, i):
    ts = [s for s in d.shapes[i] if s[6] and not D.is_pagenum(s)]
    ts.sort(key=lambda s: (s[3], s[2]))
    return " || ".join(t[6][:40] for t in ts[:2])[:84]


def main():
    full = "--full" in sys.argv
    ic = D.Deck("Module 1 - In Class.pptx")
    rv = D.Deck("Module 1 - Revised.pptx")
    print("In Class %d slides / Revised %d slides"
          % (len(ic.shapes), len(rv.shapes)))

    print("\n=== PAIRING CHECK (sim < 0.50 is suspicious) ===")
    for i in sorted(MAP):
        j = MAP[i]
        if j is None:
            print("  IC %2d -> --   NEW, no twin   %s" % (i, ident(ic, i - 1)))
            continue
        s = D.sim(ic.texts(i - 1), rv.texts(j - 1))
        flag = "  <-- CHECK" if s < 0.50 else ""
        print("  IC %2d -> RV %2d  sim=%.2f  %s%s"
              % (i, j, s, ident(ic, i - 1), flag))

    print("\n=== IN REVISED ONLY (in-class part) ===")
    for j, what in sorted(REV_ONLY.items()):
        print("  RV %2d  %s" % (j, what))

    print("\n=== DIFFERENCES  (IC = dominant, RV = to be updated) ===")
    clean = []
    for i in sorted(MAP):
        j = MAP[i]
        if j is None:
            continue
        d = D.diff_slide(ic.shapes[i - 1], rv.shapes[j - 1])
        nd = ic.notes[i - 1].strip() != rv.notes[j - 1].strip()
        cn, co = ic.clicks[i - 1], rv.clicks[j - 1]
        cd = len(cn) != len(co)
        if not (d or nd or cd):
            clean.append((i, j))
            continue
        print("--- IC %2d -> RV %2d   %s" % (i, j, ident(ic, i - 1)))
        for line in d:
            print("   " + line.strip())
        if cd:
            print("   CLICKS  IC=%d  RV=%d" % (len(cn), len(co)))
        if nd:
            print("   NOTES DIFFER")
            if full:
                import difflib
                for line in difflib.unified_diff(
                        rv.notes[j - 1].splitlines(),
                        ic.notes[i - 1].splitlines(),
                        "revised", "in-class", lineterm="", n=0):
                    print("      " + line[:180])
    print("\n=== IDENTICAL (%d pairs) ===" % len(clean))
    print("  " + ", ".join("IC%d=RV%d" % p for p in clean))


if __name__ == "__main__":
    main()
