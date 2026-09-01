"""Mapping from the recorded per-video decks to `Module 4 - Revised.pptx`.

Derived 2026-08-31 from the five decks in `Recorded Video Slides/`, by
walking each video deck in order against its block in the revised deck
(the blocks are delimited by the video title cards at revised slides
1, 9, 14, 49, 54; video 5's block ends at 63).

VIDEO_DECKS maps a deck file to a list of (video slide, revised slide).
The revised slides that appear in NO video are listed in NOT_IN_VIDEO.
"""

VIDEO_DIR = "Recorded Video Slides"

VIDEO_DECKS = [
    ("Module 4 - Video 1 - Introduction to Market Structures.pptx", [
        (1, 1), (2, 2), (3, 4), (4, 5), (5, 6), (6, 7), (7, 8),
    ]),
    ("Module 4 - Video 2 - Perfect Competition.pptx", [
        (1, 9), (2, 10), (3, 11), (4, 12), (5, 13),
    ]),
    ("Module 4 - Video 3 - Profit Maximization of a Price Taker.pptx", [
        (1, 14), (2, 15), (3, 16), (4, 17), (5, 18), (6, 19), (7, 20),
        (8, 21), (9, 22), (10, 23), (11, 24), (12, 25), (13, 26), (14, 27),
        (15, 28), (16, 30), (17, 31), (18, 32), (19, 33), (20, 37), (21, 38),
        (22, 39), (23, 40), (24, 41), (25, 42), (26, 43), (27, 44), (28, 45),
        (29, 46), (30, 47), (31, 48),
    ]),
    ("Module 4 - Video 4 - Firm-Level and Market Supply.pptx", [
        (1, 49), (2, 50), (3, 51), (4, 52), (5, 53),
    ]),
    ("Module 4 - Video 5 - Long-Run Competitive Equilibrium.pptx", [
        (1, 54), (2, 55), (3, 56), (4, 57), (5, 58), (6, 63),
    ]),
]

# Revised-deck slides that no video contains.
#   3            "Some Logistics"            (dropped from video 1)
#   29           Ross Stores annual report   (dropped from video 3)
#   34, 35, 36   the three-slide PollEv run  (dropped from video 3)
#   59           "Long-Run Equilibrium: The Drug Market"
#   60           its PollEv slide
#   61, 62       "Arrest Drug Dealers" / "Arrest / Punish Drug Users"
#   64 - 90      the in-class part: agenda items 3 and 4 plus the summary
NOT_IN_VIDEO = [3, 29, 34, 35, 36, 59, 60, 61, 62] + list(range(64, 91))

# Of those, only the ones worth showing a SECOND time at the end of the deck.
# Slides 64 - 90 are agenda items 3 and 4 plus the summary: they were
# designated in-class from the start, so they could not possibly have been in
# a video and re-listing them says nothing (Nico, 2026-08-31 - the first pass
# duplicated all 36 and he cut these 27).  What is left is the informative
# set: the 9 slides that sit INSIDE a video block yet were dropped from the
# recording.
DUPLICATE_AT_END = [3, 29, 34, 35, 36, 59, 60, 61, 62]

TOTAL_REVISED = 90


def mapped_pairs():
    """(deck filename, video slide, revised slide) for every mapped slide."""
    for deck, pairs in VIDEO_DECKS:
        for v, r in pairs:
            yield deck, v, r


def check():
    seen = sorted(r for _, _, r in mapped_pairs())
    assert len(seen) == len(set(seen)), "a revised slide is mapped twice"
    both = set(seen) & set(NOT_IN_VIDEO)
    assert not both, "mapped and not-in-video overlap: %s" % sorted(both)
    missing = set(range(1, TOTAL_REVISED + 1)) - set(seen) - set(NOT_IN_VIDEO)
    assert not missing, "unaccounted revised slides: %s" % sorted(missing)
    return len(seen), len(NOT_IN_VIDEO)


if __name__ == "__main__":
    n_map, n_not = check()
    assert set(DUPLICATE_AT_END) <= set(NOT_IN_VIDEO)
    print("mapped video slides: %d" % n_map)
    print("not in any video:    %d" % n_not)
    print("total:               %d" % (n_map + n_not))
    print("not-in-video slides:", NOT_IN_VIDEO)
    print("duplicated at end:  ", DUPLICATE_AT_END)
