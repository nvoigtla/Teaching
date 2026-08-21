# -*- coding: utf-8 -*-
"""Move the 76/77 post-work link boxes to the bottom-right link-box
position (deck convention) and update the animation plans. One-time."""
from pathlib import Path

p = Path("_build_Module2InClass.py")
lines = p.read_text(encoding='utf-8').splitlines(keepends=True)

i75 = next(i for i, l in enumerate(lines)
           if l.startswith("def slide_75_postwork_videos"))
i76 = next(i for i, l in enumerate(lines)
           if l.startswith("def slide_76_postwork_ps2"))
iend = next(i for i, l in enumerate(lines[i76:], start=i76)
            if l.startswith("    return slide")) + 1

new_block = '''def slide_75_postwork_videos(prs):
    slide = make_m2_outline(prs, 76, section_tag=TAG_WRAP,
                            highlight_set={2, 3, 4})
    # bottom-right link box overlaying the footer (deck convention),
    # drawn last so it sits in front
    _add_outlined_box(slide, Inches(8.15), Inches(6.68), Inches(4.9),
                      Inches(0.72),
                      "\\u25b6  Module 2 Videos 1+2   \\u00b7   "
                      "Practice Videos 1+2\\nOn BL under "
                      "\\u201cModule 2 Post-Work\\u201d",
                      line=GOLD, text_color=NAVY, size=15, bold=True,
                      rounded=True, shadow=True, corner_pct=0.20)
    return slide


def slide_76_postwork_ps2(prs):
    slide = make_m2_outline(prs, 77, section_tag=TAG_WRAP,
                            highlight_set={5})
    _add_convention_box(
        slide, Inches(9.05), Inches(1.75), Inches(3.9), Inches(1.5),
        prefix="Note: ",
        body="You do not need to perform the actual estimation "
             "(regression). But you need to understand how to interpret "
             "regression coefficients", size=14)
    _add_outlined_box(slide, Inches(8.15), Inches(6.68), Inches(4.9),
                      Inches(0.72),
                      "\\u25b6  Module 2 Video 3   \\u00b7   "
                      "\\u270e  Problem Set 2\\nOn BL under "
                      "\\u201cModule 2 Post-Work\\u201d",
                      line=GOLD, text_color=NAVY, size=15, bold=True,
                      rounded=True, shadow=True, corner_pct=0.20)
    return slide
'''
out = ''.join(lines[:i75]) + new_block + ''.join(lines[iend:])
p.write_text(out, encoding='utf-8')
print("slide 76/77 builders replaced")

a = Path("_animate.py")
t = a.read_text(encoding='utf-8')
i0 = t.index("    75: [")
i1 = t.index("}", t.index("    76: ["))
i1 = t.index("],", t.index('"osp:4", "t:Note:"')) + 3
new_plans = '''    75: [  # post-work link box (bottom-right, in front)
        ["t:\\u25b6"],
    ],
    76: [  # post-work link box, then the estimation note
        ["t:\\u25b6"],
        ["grp:0"],
    ],
'''
t = t[:i0] + new_plans + t[i1:]
a.write_text(t, encoding='utf-8')
print("animation plans 76/77 replaced")
