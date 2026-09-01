"""Shape matching used by `_diff_videos.py`.

A naive match on (kind, text) pairs every blank-text shape with whatever
blank shape happens to come first, which turns a slide full of unlabelled
connectors and strips into dozens of phantom "moved" reports. This does it
in three passes instead:

  1. exact  - same kind, text AND geometry (chrome, untouched shapes)
  2. text   - same kind and non-empty text (a shape that moved or resized)
  3. blanks - same kind, nearest in position+size, within a cutoff

Whatever is left over really is only in one of the two decks.

The footer page number is dropped before matching: its cached value is the
slide's own number, so it always differs between a video deck and the main
deck and never represents a hand edit.
"""

TOL = 0.011
BLANK_CUTOFF = 0.75   # inches of combined pos+size drift

# footer page-number field, on the 13.33 x 7.5" canvas
PAGENUM_BOX = (12.55, 7.20, 0.55, 0.32)


def is_pagenum(sh):
    """sh = (kind, text, x, y, w, h, path)"""
    if sh[0] != "sp":
        return False
    if not sh[1].isdigit():
        return False
    return all(abs(sh[2 + i] - PAGENUM_BOX[i]) < 0.05 for i in range(4))


def strip_pagenum(shapes):
    return [s for s in shapes if not is_pagenum(s)]


def _dist(a, b):
    return sum(abs(a[2 + i] - b[2 + i]) for i in range(4))


def match(A, B):
    """Return (pairs, only_a, only_b).

    pairs is a list of (a, b); only_* are the unmatched shapes.
    """
    a_left, b_left = list(A), list(B)
    pairs = []

    # pass 1 - identical kind, text and geometry
    for a in list(a_left):
        hit = next((b for b in b_left
                    if b[0] == a[0] and b[1] == a[1]
                    and max(abs(a[2 + i] - b[2 + i])
                            for i in range(4)) <= TOL), None)
        if hit is not None:
            a_left.remove(a)
            b_left.remove(hit)
            pairs.append((a, hit))

    # pass 2 - same kind and non-empty text; nearest wins
    for a in list(a_left):
        if not a[1]:
            continue
        cands = [b for b in b_left if b[0] == a[0] and b[1] == a[1]]
        if not cands:
            continue
        hit = min(cands, key=lambda b: _dist(a, b))
        a_left.remove(a)
        b_left.remove(hit)
        pairs.append((a, hit))

    # pass 3 - blanks, by nearest geometry within a cutoff
    blanks = [a for a in a_left if not a[1]]
    cand_pairs = []
    for a in blanks:
        for b in b_left:
            if b[0] != a[0] or b[1]:
                continue
            d = _dist(a, b)
            if d <= BLANK_CUTOFF:
                cand_pairs.append((d, a, b))
    cand_pairs.sort(key=lambda t: t[0])
    used_a, used_b = set(), set()
    for d, a, b in cand_pairs:
        if id(a) in used_a or id(b) in used_b:
            continue
        used_a.add(id(a))
        used_b.add(id(b))
        a_left.remove(a)
        b_left.remove(b)
        pairs.append((a, b))

    return pairs, a_left, b_left


def moved(pairs):
    return [(a, b) for a, b in pairs
            if max(abs(a[2 + i] - b[2 + i]) for i in range(4)) > TOL]
