# -*- coding: utf-8 -*-
"""The Module 3 outline — ONE source of truth for the agenda slides
(_retrofit_agenda.py), the coverage pills, and the video title cards
(_video_prep.py), so an item's wording and its video number can never
drift apart.

2026-08-26 (Nico): "Introduction to Module 3" joins as item 1 and is
taped as Video 1, so the five teaching sections move to Videos 2-7.
"""

M3_OUTLINE = [
    ("Introduction to Module 3",
     "A brief overview of what we cover in Module 3"),
    ("The Production Function",
     "How inputs (capital and labor) are turned into output"),
    ("Short Run: Hiring Decisions",
     "How many workers to hire when capital is fixed"),
    ("Wage Searchers",
     "When hiring one more worker bids up the wage for everyone"),
    ("Long Run: The Optimal Input Mix",
     "Choosing between capital and labor when both can be adjusted"),
    ("Cost Concepts",
     # shortened 2026-08-26: the full line ("...matter for decisions")
     # ran 9.60" and reached under the Video 6 pill; the clearance rule
     # caps a description at 8.63"
     "Fixed, variable, marginal, and sunk costs, and which ones matter"),
    ("Economies of Scale and Scope",
     "When bigger or broader production lowers cost per unit"),
]

# Module 3 is taped end to end, so every item is a video. A half-video
# module lists its in-class items here instead (Module 1 is the
# reference) and they get the navy "In class" pill.
IN_CLASS_ITEMS = set()
COVERAGE_LABEL = {i: ("In class" if i in IN_CLASS_ITEMS
                      else "Video %d" % (i + 1))
                  for i in range(len(M3_OUTLINE))}

# Where each video's title card goes: immediately BEFORE its anchor.
#   ("top",   None) — the very first slide of the deck
#   ("agenda", k)   — the k-th agenda slide in deck order (0 = the module
#                     overview, k = the agenda that highlights item k)
#   ("text",   s)   — the slide carrying this text
# 2026-08-26 (Nico): the INTRODUCTION card opens the deck, ahead of the
# deck title slide — a video-mode deck starts by naming the video the
# viewer is about to watch. (It also has no agenda slide of its own: the
# module overview is its agenda.)
VIDEO_ANCHOR = {
    0: ("top", None),
    1: ("agenda", 1),
    2: ("agenda", 2),
    3: ("agenda", 3),
    4: ("agenda", 4),
    5: ("agenda", 5),
    6: ("agenda", 6),
}
