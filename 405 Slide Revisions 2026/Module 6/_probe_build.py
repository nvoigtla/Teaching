# -*- coding: utf-8 -*-
"""Build one probe deck per slide function, so PowerPoint can say which
slide it refuses to open.  python-pptx and zipfile.testzip() both accept
files PowerPoint rejects, so the deck must be opened to be verified.
"""
import sys
from pathlib import Path
from pptx import Presentation
import _build_Module6 as B
from _m6_helpers import SLIDE_W, SLIDE_H

CASES = [
    ("title",         lambda p: B.slide_title(p)),
    ("logistics",     lambda p: B.slide_logistics(p, 2)),
    ("roadmap",       lambda p: B.slide_roadmap(p, 3)),
    ("card",          lambda p: B.video_card(p, 0)),
    ("outline_all",   lambda p: B.make_m6_outline(p, 5, descriptions=True)),
    ("outline_one",   lambda p: B.make_m6_outline(p, 6, highlight_idx=0)),
    ("dilemma",       lambda p: B.s_dilemma(p, 7)),
    ("nfx_simple",    lambda p: B.s_netflix_simple(p, 8)),
    ("nfx_lose",      lambda p: B.s_netflix_lose(p, 9)),
    ("complex",       lambda p: B.s_complex(p, 10)),
    ("dumdums",       lambda p: B.s_dumdums(p, 11)),
    ("three_deg",     lambda p: B.s_three_degrees(p, 12)),
]
only = sys.argv[1:] or None
out = Path("_probe"); out.mkdir(exist_ok=True)
for name, fn in CASES:
    if only and name not in only:
        continue
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    try:
        fn(prs)
    except Exception as e:
        print("%-12s BUILD FAILED: %s" % (name, e))
        continue
    prs.save(str(out / ("_p_%s.pptx" % name)))
    print("%-12s -> _probe/_p_%s.pptx" % (name, name))
