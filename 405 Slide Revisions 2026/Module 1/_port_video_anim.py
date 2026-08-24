# -*- coding: utf-8 -*-
"""Third part of the Videos-Final port: Nico's re-choreography, read out
of the polished decks with _extract_timing.py. 2026-08-24.

NOTE: _animate.py holds the LIVE config; _anim_config_m1.txt is a stale
snapshot from the one-off _splice_anim_config.py step. Edit _animate.py.
"""
import io
import os

HERE = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(HERE, "_animate.py")
src = io.open(PATH, encoding="utf-8").read()
edits = []


def rep(old, new, label):
    global src
    n = src.count(old)
    assert n == 1, "expected 1 match, found %d for %s\n%r" % (n, label,
                                                              old[:220])
    src = src.replace(old, new)
    edits.append(label)


# ---------------------------------------------------------------------------
# display 72 (key 64) — a new custom build; it ran on the default rollout
# ---------------------------------------------------------------------------
rep("    65: [  # netflix (video): pic static; options; covid Q; badge\n"
    "        [\"pr:Define Netflix:1:1\"],\n"
    "        [\"pr:Define Netflix:2:2\"],\n"
    "        [\"pr:Define Netflix:3:3\"],\n"
    "        [\"pr:Define Netflix:4:4\"],\n"
    "        [\"osp:3\", \"t:Prepare for Class Discussion\"],\n"
    "    ],\n",
    "    64: [  # market definition — Nico's 2026-08-24 build: \"Extent of\n"
    "           # market\" arrives together with the question under it, then\n"
    "           # the two sub-points one at a time, then geography with its\n"
    "           # example\n"
    "        [\"pr:A company must:1:1\"],\n"
    "        [\"pr:A company must:2:2\"],\n"
    "        [\"pr:A company must:3:4\"],\n"
    "        [\"pr:A company must:5:5\"],\n"
    "        [\"pr:A company must:6:6\"],\n"
    "        [\"pr:A company must:7:8\"],\n"
    "    ],\n"
    "    # 65 (display 75, the video Netflix slide) is STATIC as of\n"
    "    # 2026-08-24 — see the SKIP_STATIC override below. Nico removed\n"
    "    # the build along with the Covid bullet.\n",
    "display 72 plan; display 75 plan removed")

# ---------------------------------------------------------------------------
# display 84 (key 74) — the bullet and the photo are static now
# ---------------------------------------------------------------------------
rep("        [\"pr:Firms are racing:1:1\", \"n:sdpic:chips\"],\n"
    "        [\"n:sdgroup:D\", \"n:sdgroup:Q1\"],\n"
    "        [\"n:sdgroup:Dp\"],\n",
    "        # 2026-08-24 (Nico): the opening bullet and the photo now show\n"
    "        # with the slide; the build is D, then the shift to D’\n"
    "        [\"n:sdgroup:D\", \"n:sdgroup:Q1\"],\n"
    "        [\"n:sdgroup:Dp\"],\n",
    "display 84: drop the first click")

# ---------------------------------------------------------------------------
# displays 91 / 92 / 93 (keys 81 / 82 / 83) — Nico's group-based builds
# ---------------------------------------------------------------------------
rep("    81: [  # shift in demand: setup + old eq; text + D' + arrow; new eq\n"
    "        [\"cxn:6\", \"t:D#1\", \"cxn:7\", \"t:S#2\", \"cxn:2\", \"cxn:3\",\n"
    "         \"t:P0\", \"t:Q0\"],\n"
    "        [\"t:When the demand\", \"cxn:8\", \"t:D’\", \"cxn:9\", \"t:Shift in\"],\n"
    "        [\"cxn:4\", \"cxn:5\", \"t:P1\", \"t:Q1\"],\n"
    "    ],\n"
    "    82: [  # shift in supply\n"
    "        [\"cxn:6\", \"t:D\", \"cxn:7\", \"t:S#2\", \"cxn:2\", \"cxn:3\",\n"
    "         \"t:P0\", \"t:Q0\"],\n"
    "        [\"t:When the supply\", \"cxn:8\", \"t:S’\", \"cxn:9\", \"t:Shift in\"],\n"
    "        [\"cxn:4\", \"cxn:5\", \"t:P1\", \"t:Q1\"],\n"
    "    ],\n"
    "    83: [  # both shifts; note last; PS pointer static\n"
    "        [\"cxn:6\", \"t:D#1\", \"cxn:7\", \"t:S#2\", \"cxn:2\", \"cxn:3\",\n"
    "         \"t:P0\", \"t:Q0\"],\n"
    "        [\"pr:In this case:0:0\", \"cxn:8\", \"t:D’\", \"cxn:9\", \"t:S’\"],\n"
    "        [\"cxn:4\", \"cxn:5\", \"t:P1\", \"t:Q1\"],\n"
    "        [\"pr:In this case:1:1\"],\n"
    "    ],\n",
    "    # 2026-08-24: Nico's choreography on the three equilibrium-change\n"
    "    # slides, read out of the polished Video 4 deck. On all three the\n"
    "    # opening D, S and old equilibrium now show WITH the slide; the\n"
    "    # build is the shift, then the new equilibrium landing together\n"
    "    # with the sentence that describes it.\n"
    "    81: [  # shift in demand\n"
    "        [\"n:sdgroup:Dp\"],\n"
    "        [\"n:sdgroup:Q1\", \"t:When the demand\"],\n"
    "    ],\n"
    "    82: [  # shift in supply\n"
    "        [\"n:sdgroup:Sp\"],\n"
    "        [\"n:sdgroup:Q1\", \"t:When the supply\"],\n"
    "    ],\n"
    "    83: [  # both shifts; the note, then the problem-set pointer last\n"
    "        [\"n:sdgroup:shifts\"],\n"
    "        [\"n:sdgroup:Q1\", \"pr:In this case:0:0\"],\n"
    "        [\"pr:In this case:1:1\"],\n"
    "        [\"t:\u279c Problem Set 1\"],\n"
    "    ],\n",
    "displays 91/92/93: group-based builds")

# ---------------------------------------------------------------------------
# Post-shift overrides, by DISPLAY number (same pattern as PLANS[73]/[74]).
# ---------------------------------------------------------------------------
rep("PLANS[74] = [\n"
    "    [\"t:\\u201cAccessible luxury\\u201d\"],\n"
    "    [\"t:True luxury\"],\n"
    "    [\"t:Combined Tapestry\", \"t:59%\"],\n"
    "    [\"t:77%\", \"t:(figures from documents\"],\n"
    "    [\"t:83%\"],\n"
    "    [\"t:\\u201cBottom line\"],\n"
    "    [\"t:Oct 2024:\"],\n"
    "]\n",
    "PLANS[74] = [\n"
    "    [\"t:\\u201cAccessible luxury\\u201d\"],\n"
    "    [\"t:True luxury\"],\n"
    "    [\"t:Combined Tapestry\", \"t:59%\"],\n"
    "    [\"t:77%\", \"t:(figures from documents\"],\n"
    "    [\"t:83%\"],\n"
    "    [\"t:\\u201cBottom line\"],\n"
    "    [\"t:Oct 2024:\"],\n"
    "]\n"
    "\n"
    "# ---------------------------------------------------------------------------\n"
    "# 2026-08-24: adopted from Nico's polished 'Videos Final' decks.\n"
    "#  * display 75 (video Netflix) loses its build entirely.\n"
    "#  * display 100 (window-tax backup) gains one: the property-tax block,\n"
    "#    then the Back pill. It is the one backup slide that animates.\n"
    "# ---------------------------------------------------------------------------\n"
    "PLANS.pop(75, None)\n"
    "SKIP_STATIC = (SKIP_STATIC | {75}) - {100}\n"
    "SKIP = SKIP_TITLE | SKIP_AGENDA | SKIP_MEDIA | SKIP_STATIC\n"
    "PLANS[100] = [\n"
    "    [\"t:Property tax:\"],\n"
    "    [\"t:\\u2190 Back\"],\n"
    "]\n",
    "post-shift overrides: display 75 static, display 100 animated")

io.open(PATH, "w", encoding="utf-8").write(src)
print("applied %d edit(s):" % len(edits))
for e in edits:
    print("   - " + e)
