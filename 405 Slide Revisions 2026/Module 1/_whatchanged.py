# -*- coding: utf-8 -*-
"""Read-only: what changed between two versions of the In-Class deck.

Usage:  python _whatchanged.py OLD.pptx NEW.pptx [--full]

Matches slides across the two files by text signature (so deletions and
reorders are handled), then reports member-level geometry / text / format
differences, notes differences and click-structure differences.
"""
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, ".")
import _diff2 as D                                          # noqa: E402


def ident(deck, i):
    ts = [s for s in deck.shapes[i] if s[6] and not D.is_pagenum(s)]
    ts.sort(key=lambda s: (s[3], s[2]))
    return " || ".join(t[6][:40] for t in ts[:2])[:92]


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    full = "--full" in sys.argv
    old, new = D.Deck(args[0]), D.Deck(args[1])
    print("old %d slides / new %d slides" % (len(old.shapes), len(new.shapes)))

    pairs = D.match(new, old)          # new slide -> old slide
    matched_old = set(j for _, j, _ in pairs if j >= 0)

    print("\n=== SLIDE MAP (new -> old) ===")
    for i, j, s in pairs:
        print("  new %2d -> old %s  sim=%.2f  %s"
              % (i + 1, ("%2d" % (j + 1)) if j >= 0 else "--", s,
                 ident(new, i)))

    print("\n=== IN OLD, GONE FROM NEW (deleted) ===")
    for j in range(len(old.shapes)):
        if j not in matched_old:
            print("  old %2d  %s" % (j + 1, ident(old, j)))

    print("\n=== PER-SLIDE DIFFERENCES ===")
    any_diff = False
    for i, j, s in pairs:
        if j < 0:
            print("--- new %d is BRAND NEW" % (i + 1))
            any_diff = True
            continue
        d = D.diff_slide(new.shapes[i], old.shapes[j])
        nd = new.notes[i].strip() != old.notes[j].strip()
        cn, co = new.clicks[i], old.clicks[j]
        cd = cn != co
        if d or nd or cd:
            any_diff = True
            print("--- new %d (old %d) sim=%.2f   %s"
                  % (i + 1, j + 1, s, ident(new, i)))
            for line in d:
                print(line)
            if cd:
                print("  CLICKS old=%d new=%d" % (len(co), len(cn)))
                if full or len(co) != len(cn):
                    for k, g in enumerate(co):
                        print("     old click %d: %s" % (k + 1, g))
                    for k, g in enumerate(cn):
                        print("     new click %d: %s" % (k + 1, g))
            if nd:
                print("  NOTES DIFFER")
                if full:
                    import difflib
                    for line in difflib.unified_diff(
                            old.notes[j].splitlines(),
                            new.notes[i].splitlines(),
                            "old", "new", lineterm="", n=0):
                        print("     " + line[:200])
    if not any_diff:
        print("(none)")


if __name__ == "__main__":
    main()
