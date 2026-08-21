# One-off: shift page-number literals in _build_Module1.py for the
# 2026-08-20 inserts (new #23 AC solution: +1 from old 23; new #37-38
# copper: +3 from old 36). Descending order avoids double-shifts;
# replace(…, 1) keeps the copper functions' own literals untouched.
import io, os

HERE = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(HERE, '_build_Module1.py')
src = io.open(path, encoding='utf-8').read()

def shift(old):
    if old >= 36:
        return old + 3
    if old >= 23:
        return old + 1
    return old

pairs = []
# footer literals in direct-built slides
for old in (84, 75, 74, 72, 70, 66, 56, 55, 52, 47, 46, 38, 33, 32, 29,
            26, 23):
    pairs.append(("_draw_footer(slide, FOOTER_TEXT, %d)" % old,
                  "_draw_footer(slide, FOOTER_TEXT, %d)" % shift(old)))
# outline / roadmap makers
for old in (78, 68, 63, 61, 53, 48, 36):
    pairs.append(("make_m1_outline(prs, %d" % old,
                  "make_m1_outline(prs, %d" % shift(old)))
pairs.append(("make_roadmap(prs, 60", "make_roadmap(prs, 63"))
# content_slide / make_stub first args (tag-qualified, descending)
fam = [
    (83, 'TAG_V4'), (82, 'TAG_V4'), (81, 'TAG_V4'), (80, 'TAG_V4'),
    (79, 'TAG_V4'), (76, 'TAG_V3'), (73, 'TAG_V3'), (71, 'TAG_V3'),
    (69, 'TAG_V3'), (65, 'TAG_V2'), (64, 'TAG_V2'), (58, 'TAG_WRAP'),
    (57, 'TAG_WRAP'), (54, 'TAG_CBA'), (51, 'TAG_SUNK'), (50, 'TAG_SUNK'),
    (49, 'TAG_SUNK'), (45, 'TAG_OPP'), (44, 'TAG_OPP'), (43, 'TAG_OPP'),
    (42, 'TAG_OPP'), (41, 'TAG_OPP'), (40, 'TAG_OPP'), (39, 'TAG_OPP'),
    (37, 'TAG_OPP'), (35, 'TAG_SD'), (34, 'TAG_SD'), (31, 'TAG_SD'),
    (30, 'TAG_SD'), (28, 'TAG_SD'), (27, 'TAG_SD'), (25, 'TAG_SD'),
    (24, 'TAG_SD'),
]
for old, tag in fam:
    pairs.append(("prs, %d, %s" % (old, tag),
                  "prs, %d, %s" % (shift(old), tag)))

for old_s, new_s in pairs:
    n = src.count(old_s)
    assert n >= 1, "NOT FOUND: %r" % old_s
    if old_s.startswith("_draw_footer") and ", 38)" in old_s:
        assert n == 2, "expected fruit+copper footers for 38, got %d" % n
        src = src.replace(old_s, new_s, 1)   # fruit table only
    else:
        assert n == 1, "AMBIGUOUS (%d): %r" % (n, old_s)
        src = src.replace(old_s, new_s)

io.open(path, 'w', encoding='utf-8').write(src)
print("renumbered %d literals" % len(pairs))
