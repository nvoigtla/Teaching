# -*- coding: utf-8 -*-
"""Canonical slide order for Module 1.

Nico teaches the module by taping four videos FIRST, then meeting in class.
The deck reflects that sequence: the four `Videos Final` decks open the
module, then the In-Class part, which starts from the module title slide and
carries the front matter (logistics, teaching philosophy) plus the
applications of the video material.

ONE map, imported by every pass in the pipeline (`_build_Module1.py`,
`_splice_media.py`, `_group_pass.py`, `_animate.py`) so a display-keyed
config never has to be hand-shifted again.  `OLD` numbers are the 101-slide
deck of 2026-08-24; `NEW` numbers are the 95-slide deck.

Video -> old-display mapping verified with `_map_vids.py` (id- and
text-keyed, every pairing >= 0.90).  Video 1 contributes 9 of its 11
slides: its trailing BACKUP divider and "People Respond to Incentives"
are duplicates of the deck's own backup section and were dropped from the
front (Nico, 2026-08-26).

2026-08-27 (Nico's hand pass over the 104-slide deck, adopted here):
  * the "Slides Not Used in the Videos" divider is gone — those slides now
    simply belong to the In-Class part;
  * the shift-combination table (old 94) moved up to CLOSE Video 4;
  * a second copy of the "Introduction / about the instructor" slide opens
    the In-Class part, ahead of the module title slide;
  * "Homo Economicus" (old 12) moved down, after the course roadmap;
  * the Kroger-Albertsons pair from `Module 1 - Example Candidates.pptx`
    was adopted after the Netflix slide;
  * DELETED: the ADM mini-case (old 21), the two COVID/tea slides (31-32),
    the three avocado slides (34-36), the two copper slides (41-42), the
    "Related Work by Anderson Faculty" backup (99), and the QUESTION slide
    of both the AC and the diamonds poll pairs (24, 28) — each poll now
    runs set-up slide -> PollEv results -> solution.
"""

# --------------------------------------------------------------------------
# The new sequence.  An int is an OLD display number; a string is a slide
# that does not exist in the old deck.
# --------------------------------------------------------------------------

V1 = [67, 2, 9, 10, 11, 13, 17, 1, 69]      # Video 1 · Introduction
V2 = [70, 71, 72, 73, 74, 75, 76]           # Video 2 · Markets
V3 = [77, 78, 79, 80, 81, 82, 83, 84, 85, 86]   # Video 3 · Demand and Supply
V4 = [87, 88, 89, 90, 91, 92, 93, 94]       # Video 4 · Market Equilibrium
#     ^ 94 = "Effect of Shifts … in Isolation", moved up to close Video 4.

# Sentinels for the slides that are new or duplicated.
DIV_IN_CLASS = "div_in_class"      # divider: In-Class Part
INTRO_AGAIN = "intro_again"        # second copy of old 2, about the instructor
TITLE_AGAIN = "title_again"        # second copy of old 1, the module title
# divider: "Some Applications of the Material Covered in Videos 2 to 4"
# (2026-08-26, Nico). It REPLACES old 19, the Markets section agenda — the
# applications are a category of their own, but deliberately not one of the
# agenda's items, so the slide just sits there instead.
DIV_APPLICATIONS = "div_applications"
KROGER_CASE = "kroger_case"        # adopted from the Example Candidates deck
KROGER_COSTCO = "kroger_costco"    # ditto, the resolution slide

IN_CLASS_FRONT = [
    DIV_IN_CLASS,        # 35
    INTRO_AGAIN,         # 36
    TITLE_AGAIN,         # 37
    3, 4, 5, 6,          # 38-41  Logistics I-III, Questions and Office Hours
    7, 8,                # 42-43  Econ & Coffee poll pair
    14, 15,              # 44-45  Making the Most of the Course 1-2
    16,                  # 46     Teaching Philosophy
    68,                  # 47     Video-1-tagged "Agenda for the Class"
    12,                  # 48     Homo Economicus
    18,                  # 49     outline overview
    DIV_APPLICATIONS,    # 50
    20,                  # 51     Recall from Video 2: Market Definition
    22,                  # 52     Netflix
    KROGER_CASE,         # 53
    KROGER_COSTCO,       # 54
    23,                  # 55     heatwaves set-up
    25,                  # 56     AC poll (results view)
    26,                  # 57     AC solution
    27,                  # 58     Swiftonomics set-up
    29,                  # 59     diamonds poll (results view)
    30,                  # 60     Swiftonomics solution
    33,                  # 61     shortages when disasters loom
    37,                  # 62     steps to analyze shocks
    38,                  # 63     wheat / Ukraine
    39, 40,              # 64-65  LA residential real estate + its market
]

IN_CLASS = list(range(43, 67))               # old 43-66 -> 66-89
BACKUP = [95, 96, 97, 98, 100, 101]          # 90-95 (old 99 dropped)

# Old slides deliberately dropped from the deck, with what replaced them.
DROPPED = {
    19: DIV_APPLICATIONS,   # Markets section agenda
    21: None,               # ADM mini-case
    24: None,               # AC poll, question view
    28: None,               # diamonds poll, question view
    31: None, 32: None,     # COVID / tea pair
    34: None, 35: None, 36: None,   # avocado trio
    41: None, 42: None,     # copper pair
    99: None,               # "Related Work by Anderson Faculty" backup
}

ORDER = V1 + V2 + V3 + V4 + IN_CLASS_FRONT + IN_CLASS + BACKUP

N_SLIDES = len(ORDER)

# --------------------------------------------------------------------------
# old display -> new display.  Old 1 and old 2 each appear twice in the new
# deck (inside the Video 1 block and again in the in-class part); the map
# points at the Video 1 occurrence, which is where the old slide's identity
# lives.  The second copies are reached through NEW_TITLE_AGAIN /
# NEW_INTRO_AGAIN.
# --------------------------------------------------------------------------

OLD_TO_NEW = {}
for _i, _o in enumerate(ORDER, start=1):
    if isinstance(_o, int) and _o not in OLD_TO_NEW:
        OLD_TO_NEW[_o] = _i

NEW_TO_OLD = {i: o for i, o in enumerate(ORDER, start=1)}

NEW_DIV_IN_CLASS = ORDER.index(DIV_IN_CLASS) + 1
NEW_INTRO_AGAIN = ORDER.index(INTRO_AGAIN) + 1
NEW_TITLE_AGAIN = ORDER.index(TITLE_AGAIN) + 1
NEW_DIV_APPLICATIONS = ORDER.index(DIV_APPLICATIONS) + 1
NEW_KROGER_CASE = ORDER.index(KROGER_CASE) + 1
NEW_KROGER_COSTCO = ORDER.index(KROGER_COSTCO) + 1

# Slides that carry no old-deck identity AND no build: the two dividers and
# the second copy of the title slide.  The second copy of the introduction
# slide and the two Kroger slides DO animate, so they stay out of this set.
NEW_ONLY = {NEW_DIV_IN_CLASS, NEW_TITLE_AGAIN, NEW_DIV_APPLICATIONS}


def new(old):
    """New display number for an old (101-deck) display number."""
    return OLD_TO_NEW[old]


def remap_keys(d):
    """Re-key a display-keyed dict from old to new display numbers."""
    return {OLD_TO_NEW[k]: v for k, v in d.items() if k in OLD_TO_NEW}


def remap_set(s):
    """Re-key a display-keyed set from old to new display numbers."""
    return {OLD_TO_NEW[k] for k in s if k in OLD_TO_NEW}


# --------------------------------------------------------------------------
# Self-check: every surviving old slide appears exactly once, nothing is
# invented.
# --------------------------------------------------------------------------
_olds = [o for o in ORDER if isinstance(o, int)]
_expected = [n for n in range(1, 102) if n not in DROPPED]
assert sorted(_olds) == _expected, "old deck not fully covered"
assert len(_olds) == len(set(_olds)) == 101 - len(DROPPED)
assert N_SLIDES == 95, N_SLIDES


if __name__ == "__main__":
    for i, o in enumerate(ORDER, start=1):
        print("%3d  <-  %s" % (i, o))
