"""One-off: adopt the dark-red demand rule in Module 1's build script.

Teaching CLAUDE.md, 2026-08-30 (Nico): "A DEMAND curve is dark red
`C00000` - curve and label alike. Supply stays navy, so the two curves
read apart at a glance."  Module 1 was built on 2026-08-20 and predates
the rule, so it still draws demand in GOLD and supply in STEEL (the
source deck's light blue 95B3D7).  Module 4 is the reference
implementation (`_sd_curves` in `_build_Module4.py`).

WHAT THIS CHANGES, and what it deliberately does not:

  demand curve  GOLD  -> RED    (RED is C00000 in this script's palette;
                                 do NOT use DARKRED, which is A2162A,
                                 the source MC-bar red)
  demand label  GOLD/NAVY -> RED
  supply curve  STEEL -> NAVY
  supply label  already NAVY, untouched

  SHIFTED curves (D', S') keep GREEN_DK / BLUE_PED.  Green marks "this is
  the shifted curve", which is a different job from naming the curve
  type, and Module 4 keeps it for exactly that (`_build_Module4.py`
  line 3643 draws a shifted supply in GREEN_DK on a slide whose base
  demand is RED).

  The GOLD "Excess demand" band on the market-mechanism slide is NOT a
  demand curve and is left alone, as is every other GOLD element on the
  deck (chrome, badges, table headers, roadmap dots).

Run:  python _sweep_demand_color.py --dry-run
      python _sweep_demand_color.py
Then rebuild the deck with the usual pipeline.
"""

import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "_build_Module1.py")

# (line number in the CURRENT file, old fragment, new fragment, what it is)
# Line numbers are checked before the edit and the run aborts if any has
# drifted, so this can never fire against a file it was not written for.
EDITS = [
    # -- slide 72: the market demand curve --------------------------------
    (5424, "(0.8, 8.6), (7.6, 1.4), color=GOLD",
           "(0.8, 8.6), (7.6, 1.4), color=RED", "D curve"),
    (5426, '7.8, 1.9, "D", color=NAVY',
           '7.8, 1.9, "D", color=RED', "D label"),

    # -- slide 73: movement along D vs. shift of D ------------------------
    (5439, "(1, 8), (8, 1), color=GOLD", "(1, 8), (8, 1), color=RED",
     "D curve"),
    (5440, '8.05, 1.35, "D", color=GOLD', '8.05, 1.35, "D", color=RED',
     "D label"),

    # -- slide 74: AI and the demand for computer chips -------------------
    # the label's colour sits on its own continuation line here
    (5528, "(1, 8), (8, 1), color=GOLD", "(1, 8), (8, 1), color=RED",
     "D curve"),
    (5532, 'color=GOLD).name = "sdlabel:D"',
           'color=RED).name = "sdlabel:D"', "D label"),

    # -- slide 75: the market supply curve --------------------------------
    (5608, "(0.8, 1.6), (7.6, 8.6), color=STEEL",
           "(0.8, 1.6), (7.6, 8.6), color=NAVY", "S curve"),

    # -- slide 76: movement along S vs. shift of S ------------------------
    (5631, "(1, 2), (7.5, 8.5), color=STEEL",
           "(1, 2), (7.5, 8.5), color=NAVY", "S curve"),

    # -- slide 79: the market mechanism -----------------------------------
    (5722, "(1, 8), (8, 1), color=GOLD", "(1, 8), (8, 1), color=RED",
     "D curve"),
    (5723, '8.05, 1.35, "D", color=NAVY', '8.05, 1.35, "D", color=RED',
     "D label"),
    (5724, "(1.5, 1.5), (8.5, 8.5), color=STEEL",
           "(1.5, 1.5), (8.5, 8.5), color=NAVY", "S curve"),
    # NOTE: line 5734's GOLD line is the "Excess demand" band, not a
    # demand curve.  Left alone on purpose.

    # -- slide 80 area: changes in market equilibrium ---------------------
    (5862, "(1, 8), (8, 1), color=GOLD", "(1, 8), (8, 1), color=RED",
     "D curve"),
    (5863, '8.1, 1.4, "D", color=GOLD', '8.1, 1.4, "D", color=RED',
     "D label"),
    (5864, "(1.5, 1.5), (9.0, 9.0), color=STEEL",
           "(1.5, 1.5), (9.0, 9.0), color=NAVY", "S curve"),

    # -- AC solution slide ------------------------------------------------
    (6150, "(1, 8), (8, 1), color=GOLD", "(1, 8), (8, 1), color=RED",
     "D curve"),
    (6151, '8.05, 1.35, "D", color=GOLD', '8.05, 1.35, "D", color=RED',
     "D label"),

    # -- tuple-driven charts: (p0, p1, colour, dash, label, label_pos) -----
    # tea, avocados, LA real estate, copper.  The BASE curve takes the new
    # colour; the dashed shifted one keeps its own.
    (4024, '((1, 8), (8, 1), GOLD, None, "D", (8.05, 1.35))',
           '((1, 8), (8, 1), RED, None, "D", (8.05, 1.35))', "D curve"),
    (4026, '((1.5, 1.5), (8.5, 8.5), STEEL, None, "S", (8.55, 8.75))',
           '((1.5, 1.5), (8.5, 8.5), NAVY, None, "S", (8.55, 8.75))',
     "S curve"),
    (4125, '((1, 8), (8, 1), GOLD, None, "D", (8.05, 1.35))',
           '((1, 8), (8, 1), RED, None, "D", (8.05, 1.35))', "D curve"),
    (4127, '((1.5, 1.5), (8.5, 8.5), STEEL, None, "S", (8.55, 8.75))',
           '((1.5, 1.5), (8.5, 8.5), NAVY, None, "S", (8.55, 8.75))',
     "S curve"),
    (4280, '((1, 8), (8, 1), GOLD, None, "D0", (8.05, 1.35))',
           '((1, 8), (8, 1), RED, None, "D0", (8.05, 1.35))', "D curve"),
    (4282, '((1.5, 1.5), (8.5, 8.5), STEEL, None, "S0", (8.55, 8.75))',
           '((1.5, 1.5), (8.5, 8.5), NAVY, None, "S0", (8.55, 8.75))',
     "S curve"),
    (6223, '((1, 8), (8, 1), GOLD, None, "D0", (8.05, 1.35))',
           '((1, 8), (8, 1), RED, None, "D0", (8.05, 1.35))', "D curve"),
    (6225, '((1.5, 1.5), (9, 9), STEEL, None, "S0", (9.05, 9.35))',
           '((1.5, 1.5), (9, 9), NAVY, None, "S0", (9.05, 9.35))',
     "S curve"),
]

# The dashed shifted curves in the tuple charts are GOLD / STEEL too --
# they are the SAME curve type, one period later, so they follow their base
# curve's colour while keeping the dash that marks them as shifted.
DASHED = [
    (4025, '((2, 8.5), (8.6, 1.9), GOLD, \'dash\', "D1", (8.65, 2.2))',
           '((2, 8.5), (8.6, 1.9), RED, \'dash\', "D1", (8.65, 2.2))'),
    (4027, '((0.5, 4), (6, 9.5), STEEL, \'dash\', "S1", (5.6, 10.0))',
           '((0.5, 4), (6, 9.5), NAVY, \'dash\', "S1", (5.6, 10.0))'),
    (4126, '((4, 9), (9.6, 3.4), GOLD, \'dash\', "D1", (9.3, 3.9))',
           '((4, 9), (9.6, 3.4), RED, \'dash\', "D1", (9.3, 3.9))'),
    (4128, '((0.5, 4), (5.8, 9.3), STEEL, \'dash\', "S1", (5.35, 9.8))',
           '((0.5, 4), (5.8, 9.3), NAVY, \'dash\', "S1", (5.35, 9.8))'),
    (4129, '((5, 1), (9.5, 5.5), STEEL, \'dash\', "S2", (9.55, 5.85))',
           '((5, 1), (9.5, 5.5), NAVY, \'dash\', "S2", (9.55, 5.85))'),
    (4281, '((0.9, 6), (6.4, 0.5), GOLD, \'dash\', "D1", (6.15, 1.05))',
           '((0.9, 6), (6.4, 0.5), RED, \'dash\', "D1", (6.15, 1.05))'),
    (4283, '((0.5, 2), (7.5, 9), STEEL, \'dash\', "S1", (7.1, 9.4))',
           '((0.5, 2), (7.5, 9), NAVY, \'dash\', "S1", (7.1, 9.4))'),
    (6224, '((4, 9), (11.4, 1.6), GOLD, \'dash\', "D1", (11.05, 2.1))',
           '((4, 9), (11.4, 1.6), RED, \'dash\', "D1", (11.05, 2.1))'),
    (6226, '((5, 1), (11.5, 7.5), STEEL, \'dash\', "S1", (11.2, 7.95))',
           '((5, 1), (11.5, 7.5), NAVY, \'dash\', "S1", (11.2, 7.95))'),
]


def main():
    dry = "--dry-run" in sys.argv
    # newline="" keeps the file's own CRLF endings inside the split parts,
    # so re-joining does not quietly rewrite every line in the file.
    lines = io.open(SCRIPT, encoding="utf-8", newline="").read().split("\n")

    todo = [(n, o, nw, what) for n, o, nw, what in EDITS] + \
           [(n, o, nw, "shifted curve") for n, o, nw in DASHED]

    # verify every anchor BEFORE touching anything
    bad = []
    for n, old, _new, _what in todo:
        if old not in lines[n - 1]:
            bad.append((n, old, lines[n - 1].strip()[:80]))
    if bad:
        print("ABORT -- %d anchor(s) have drifted; the build script has "
              "changed since this sweep was written:" % len(bad))
        for n, old, got in bad:
            print("  line %d\n    want: %s\n    got:  %s" % (n, old, got))
        return 1

    for n, old, new, what in sorted(todo):
        print("  %5d  %-14s %s" % (n, what, lines[n - 1].strip()[:70]))
        lines[n - 1] = lines[n - 1].replace(old, new)

    if dry:
        print("\n--dry-run: %d line(s) would change, nothing written."
              % len(todo))
        return 0

    io.open(SCRIPT, "w", encoding="utf-8",
            newline="").write("\n".join(lines))
    print("\nrewrote %s (%d lines)" % (SCRIPT, len(todo)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
