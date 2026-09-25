# -*- coding: utf-8 -*-
"""Inject Fade / on-click build animations via OOXML <p:timing> —
Module 4 (2026-08-28). Engine from Module 1 / Module 7 / Italy IBR.

Teaching CLAUDE.md rules implemented:
  - Fade entrance, 0.5 s, on click; one style deck-wide. Chrome (top
    bar, tag, title, rules, footer, page number, chart axes/backings,
    video-link boxes, back buttons) never animates.
  - Text slides: first top-level bullet (with subs) shows with the
    slide; build from the second, one top-level bullet per click.
    Figures + captions ride on the click of the bullet they support.
  - Diagram/chart slides: custom story plans (PLANS below) — curve +
    its label on one click, guides before regions, callouts in story
    order, the gold equilibrium/takeaway element last. Unlisted shapes
    stay visible from the start (axes, ticks that anchor the setup).
  - Skips: title, roadmap, outline/agenda slides, section cards, poll
    and video slides, the practice-video index.

Pipeline:  _build_Module4.py -> _splice_media.py -> _group_pass.py
           -> _animate.py all apply
Verify via PowerPoint COM (MainSequence counts) + eyeball.

Selector language for PLANS beats:
  "t:PREFIX"    first UNUSED text shape whose text starts with PREFIX
  "t:PREFIX#n"  the n-th (1-based, doc order) text match, used or not
  "pic:N" / "cxn:N" / "grp:N" / "osp:N"   n-th picture / connector /
                group / textless shape in document order
  "pr:PREFIX:st:end"  paragraph range st..end of the text box whose
                text starts with PREFIX (adds a bldP for the box)
"""
import os
import unicodedata
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
import sys as _sys
_deckarg = [a for a in _sys.argv[1:]
            if a.endswith(".pptx")]
DECK = Path(_deckarg[0]) if _deckarg else HERE / "Module 6 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
EMU = 914400.0

# --------------------------------------------------------------------------
#  Module 6 configuration — the skip sets are DERIVED, not hard-coded
# --------------------------------------------------------------------------

def _derive_skips(deck):
    """(title/card/divider, agenda, poll) slide numbers, read off the deck.

    Module 4 kept these as literals and its own comments record them going
    stale as slides moved.  With 117 slides that is not a risk worth
    running, so each set is recognised by a structural fact:

      * a slide with NO navy top bar is the deck title, a video title card
        or a section divider — none of which is ever animated;
      * a slide titled "Outline of Module 6" is an agenda slide — static,
        so nothing jumps between consecutive agendas;
      * a slide with a PollEverywhere `tags` relationship is live content
        spliced in verbatim, and must not be touched.
    """
    import zipfile as _zip
    import xml.etree.ElementTree as _ET
    titles, agendas, polls = set(), set(), set()
    with _zip.ZipFile(str(deck)) as z:
        pres = _ET.fromstring(z.read("ppt/presentation.xml"))
        rels = {r.get("Id"): r.get("Target") for r in
                _ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
        order = ["ppt/" + rels[s.get("{%s}id" % R)].replace("../", "")
                 for s in pres.find("{%s}sldIdLst" % P)]
        names = set(z.namelist())
        for n, part in enumerate(order, 1):
            xml = z.read(part).decode("utf-8")
            # the navy top bar is a full-width 0.42" rect at the origin
            if '<a:off x="0" y="0"/><a:ext cx="12192000" cy="384048"/>' \
                    not in xml.replace(" ", " "):
                if "0B2B4E" not in xml[:4000] or "Module 6 ·" not in xml:
                    titles.add(n)
            if "Outline of Module 6" in xml:
                agendas.add(n)
            relp = (part.rsplit("/", 1)[0] + "/_rels/"
                    + part.rsplit("/", 1)[1] + ".rels")
            if relp in names and "tags" in z.read(relp).decode("utf-8"):
                polls.add(n)
    return titles, agendas, polls


SKIP_TITLE, SKIP_AGENDA, SKIP_MEDIA = _derive_skips(DECK)
# single-thought slides that are one object and build nothing.
# Slide 4 (the course roadmap) is static in Nico's own deck too -- it is
# not among the 48 slides that carry a timing tree.
SKIP_STATIC = {4}   # was {3} before the intro card went in, 2026-09-23
SKIP = SKIP_TITLE | SKIP_AGENDA | SKIP_MEDIA | SKIP_STATIC

# figures ride on this animated-bullet index by default (0 = first
# built bullet); per-slide overrides:
FIG_GROUP_DEFAULT = 0
FIG_GROUP = {}

# The zoo example's figures come from the build script, never retyped: a
# plan that names "$20" after the constants moved to $40 does not fail,
# it silently animates the wrong label or raises a KeyError mid-run.
# Importing is safe -- _build_Module6's work sits behind __main__.
import _build_Module6 as _M6
# the per-market profit lines, which now show the AMOUNT and not the
# arithmetic, so the plan has to name the amount
def _bmw_profit(pr, qty):
    return "Profit = $%s" % format(int(round((pr - _M6.BMW_MC) * qty)), ",")

_PROF_US = _bmw_profit(_M6.BMW_PUS, _M6.BMW_QUS)
_PROF_DE = _bmw_profit(_M6.BMW_PDE, _M6.BMW_QDE)
_PROF_J = _bmw_profit(_M6.BMW_PJ, _M6.BMW_QJ)
_Z = {"a": "$%g" % _M6.ZOO_A, "mc": "$%g" % _M6.ZOO_MC,
      "qmc": "%g" % _M6.ZOO_QMC, "qmax": "%g" % _M6.ZOO_QMAX,
      "sp": "$%g" % _M6.ZOO_SIMPLE_P, "sq": "%g" % _M6.ZOO_SIMPLE_Q,
      "sprof": "$%d" % int(_M6.ZOO_SIMPLE_PROFIT),
      "flat": "$%d" % int(_M6.ZOO_FLAT),
      "ffprof": "$%d" % int(_M6.ZOO_FF_FLAT - _M6.ZOO_FF_COST)}

PLANS_PRE = {}

STATIC = {}

# --------------------------------------------------------------------------
#  Custom story plans (selector language in the header docstring).
#  Authored per slide from the inventory printed by:
#      python _animate.py inventory <display>...
#  Anything not selected stays visible from the start, which is how chart
#  axes and tick labels stay static.
# --------------------------------------------------------------------------
# NOTE: the numeric t: selectors below (60,000 / 400 / 358 / 210 / 187.5 /
# 662.5) are chart labels printed from the cost constants in
# _build_Module4.py.  Change those constants and these have to change with
# them - the engine raises KeyError on a selector it cannot resolve, so a
# stale one fails the run rather than animating the wrong shape.
PLANS = {
    # 2026-09-15 (Nico).  The reference slide for the curve-order
    # convention now in the Teaching CLAUDE.md: formula WITH its own line
    # and that line's intercepts, then the other primary line, then the
    # derived line, then the guides that only exist because of it, then
    # the shaded areas.  This slide had broken it three ways -- MR came
    # before MC, the MC LINE was buried in the dashed-guide beat, and its
    # label "MC = 0" was stranded on a click of its own at the very end,
    # a beat after the curve it names.
    10: [["grp:0", "cxn:2", "t:D", "t:$4", "t:12"],
        ["cxn:4", "t:MC = 0"],
        ["cxn:3", "t:MR"],
        # the two dashed guides mark MR = 0, so they cannot precede MR
        ["cxn:6", "cxn:5", "t:6", "t:$2"],
        ["grp:1"],
        ["osp:3", "t:Max. possible TR"]],

    # 9 <- NV 8 (his 7 clicks, order kept exactly):
    #   figure -> bullet -> REVENUES -> bullet -> CONSUMER SURPLUS ->
    #   UNEXPLOITED -> bullet.  He builds revenues before surplus.
    # 34 <- NV 25.  Two panels of the same chart, and it was running on
    # the auto-rollout: 17 clicks, one shape each, in geometric order --
    # both "P" axis titles, then both "P*", then the two MC lines, then the
    # two demands, then the two MRs, then both "Q" and both "Q*".  Nothing
    # about that told the story, and the two panels were interleaved so the
    # class never saw one market finished before the other started.
    # Now: panel by panel, each in the deck's curve order (demand, then the
    # other primary line, then the derived one, then the guides its optimum
    # implies), with the comparison as the closing bar.
    35: [["t:Flight LAX", "cxn:2", "t:D"],
         ["cxn:4", "t:MC"],
         ["cxn:3", "t:MR"],
         ["cxn:5", "cxn:6", "t:P*", "t:Q*"],
         ["t:Direct flight LAX", "cxn:9", "t:D"],
         ["cxn:11", "t:MC"],
         ["cxn:10", "t:MR"],
         ["cxn:12", "cxn:13", "t:P*", "t:Q*"],
         ["t:Higher price for the shorter"]],

    # 2026-09-15 (Nico), two conventions at once on this slide:
    #  * the "6" is NOT one of demand's intercepts.  It is the quantity
    #    at the revenue rectangle's right edge, so it comes with that
    #    rectangle, not with the demand line.  (On slide 9 the same tick
    #    marks MR = 0 and rides with the MR guides; MR is not drawn here
    #    at all, so a "6" on click 1 marked nothing.)
    #  * a legend row naming a COMBINATION of areas comes after the LAST
    #    of them: "wasted profits = consumer surplus + unexploited
    #    markets" used to be revealed two clicks before either area, so
    #    it named two things that were not on screen yet.  A row naming
    #    ONE area still rides with it, as "Revenues" does.
    11: [["grp:0", "cxn:2", "t:D", "t:$4", "t:12"],
        ["t:Simplifying assumption", "n:sdbullmark:0:dot"],
        # the revenue rectangle with the bullet that names it, and the
        # two ticks at its corner -- the new row pushed every mark index
        # up by one
        ["osp:4", "t:Revenues under Simple", "t:Revenues",
         "t:$2", "t:6", "n:sdbullmark:1:rev"],
        ["osp:3", "t:Consumer surplus"],
        ["osp:5", "t:Unexploited"],
        # each bullet rides with the region mark that replaced its dot --
        # the mark IS the label, so one click
        ["t:“Wasted profits”", "n:sdbullmark:2:cs",
         "n:sdbullmark:2:unx"],
        ["t:Can do better", "n:sdbullmark:3:cs",
         "n:sdbullmark:3:rev", "n:sdbullmark:3:unx"]],

    # 11, 12 <- the two "In the News" slides.  2026-09-15 (Nico): "when we
    # have a newspaper or similar item placed on top of the slide, it is
    # always there (no animation)."  The clipping is what the class is
    # reading while the question is put to them, so it is chrome; only the
    # question animates.  The auto-rollout had been fading the clipping in
    # on click 1 and the question on click 2.
    12: [["t:What does recent research"]],
    13: [["t:Describe this pricing strategy"]],

    # ----------------------------------------------------------------
    #  Video 2 — first degree
    # ----------------------------------------------------------------
    # 16 <- NV 14, all nine of his clicks: he alternates bullet and
    #   figure, and the three MPV callouts run high -> median -> low.
    20: [["pr:Assume MC:0:0"],
         ["cxn:2", "t:Demand = MPV"],
         ["pr:Assume MC:1:1"],
         ["osp:4", "cxn:3", "t:Price charged to high"],
         ["osp:5", "cxn:4", "t:Price charged to median"],
         ["osp:6", "cxn:5", "t:Price charged to low"],
         ["osp:3", "t:Revenues and profits", "t:(assume MC = 0)"],
         ["pr:Assume MC:2:2"],
         ["pr:Assume MC:3:3"]],

    # 17 <- NV 15, his seven clicks.  Note the PROFIT region lands LAST,
    #   after "can it work in the real world?" — his order, not the
    #   obvious one.
    21: [["pr:Now: MC:0:0", "cxn:2", "t:Demand = MPV"],
         ["cxn:3", "t:MC"],
         ["osp:4", "cxn:6", "t:Price charged to high"],
         ["osp:5", "cxn:7", "t:Price charged to median"],
         ["osp:6", "cxn:8", "t:Do not sell"],
         ["pr:Now: MC:1:1"],
         ["osp:3", "t:Profits", "cxn:4", "cxn:5"]],

    # ----------------------------------------------------------------
    #  Video 3 — segment pricing
    # ----------------------------------------------------------------
    # 22 <- NV 19: two clicks, one per price band.  Demand and the
    #   right-hand bullets are static, as in his.
    27: [["osp:3", "t:Regular movie-goers", "t:P\u2081"],
         ["osp:4", "t:Seniors", "t:P\u2082"]],

    # 29 — the adopted three-panel BMW figure has no original to follow,
    #   so it builds the way the slide argues: each market, then the
    #   joint market, then the two profit totals that settle it.
    # 2026-09-15 (Nico): "fix the animation timing."  The whole figure was
    # STATIC -- all three panels, every curve, every shaded rectangle up
    # from the first second -- and only five labels faded in over the top,
    # which is the opposite of a build.  Now each market is built in the
    # deck's curve order (demand, then MC, then MR, then the guides its
    # optimum implies, then the rectangle with the profit it is worth),
    # market by market, with each total as its own beat.  The headline,
    # which states the answer, is the LAST click rather than the first
    # thing on screen.
    41: [["t:US market", "cxn:2", "t:D"],
         ["cxn:4", "t:MC"],
         ["cxn:3", "t:MR"],
         ["cxn:5", "cxn:6"],
         ["osp:3", "t:" + _PROF_US],
         ["t:German market", "cxn:9", "t:D"],
         ["cxn:11", "t:MC"],
         ["cxn:10", "t:MR"],
         ["cxn:12", "cxn:13"],
         ["osp:4", "t:" + _PROF_DE],
         # the joint demand kinks, so it is two connectors, and the note
         # about the kink belongs with the curve that has one
         ["t:Joint market", "cxn:16", "cxn:17", "t:D", "t:Kink results"],
         ["cxn:19", "t:MC"],
         ["cxn:18", "t:MR"],
         ["cxn:20", "cxn:21"],
         ["osp:6", "t:" + _PROF_J],
         # the navy bar carries the comparison AND both totals, so it is
         # one closing beat rather than three (2026-09-15)
         ["t:Segment pricing:"]],

    # ----------------------------------------------------------------
    #  Video 4 — versioning
    # ----------------------------------------------------------------
    # 34 <- NV 33: one click per version, then the second bullet.  His
    #   first bullet is visible from the start.
    46: [["osp:3", "t:Version 1", "t:P\u2081"],
         ["osp:4", "t:Version 2", "t:P\u2082"],
         ["osp:5", "t:Version 3", "t:P\u2083"],
         ["osp:6", "t:Version 4", "t:P\u2084"],
         ["osp:7", "t:Version 5", "t:P\u2085"],
         ["pr:Create different:1:1"]],

    # ----------------------------------------------------------------
    #  Video 5 — flat fee
    # ----------------------------------------------------------------
    # 43 <- NV 46, his six clicks: the whole figure first, then the
    #   bullets, with the shaded triangle landing on the same click as
    #   the "area = 16" line (his click 3).
    49: [["out:n:MASK-1"], ["out:n:MASK-2"], ["out:n:MASK-3"]],

    # ----------------------------------------------------------------
    #  55 <- NV 62, his richest build: 15 clicks that lay one block down
    #  at a time, each one introduced by the bullet that explains it.
    #  Order kept exactly; the axes stay static and each block's price
    #  tick, quantity tick and letter ride with the block.
    # ----------------------------------------------------------------
    61: [["t:P = 4", "cxn:2", "t:D", "t:$4", "t:12"],
         # the legend triangle IS the label of the line beside it
         ["osp:4", "pr:Triangle =:0:0"],
         ["osp:3", "t:Revenues = $24", "pr:Triangle =:1:1"],
         ["pr:Triangle =:2:2"],
         ["pr:Triangle =:3:3", "pr:Triangle =:4:4"],
         ["pr:Triangle =:5:5"]],

    # 67-74 all build the same way (2026-09-15, Nico: "follow exactly the
    # animation convention for when the various elements incl. legend show
    # up"): formula and demand, then the other primary line, then the
    # guides its optimum implies, then each area WITH the legend row that
    # names it.  The usage-fee row names the MC LINE, so it rides with MC.
    68: [["cxn:2", "t:D"],
         ["cxn:3", "t:MC", "osp:6", "t:Usage fee P*"],
         ["cxn:4", "cxn:5", "t:P*"],
         ["osp:3", "t:F*", "osp:5", "t:Flat fee F*"],
         ["osp:4", "t:Revenues from usage fee", "osp:7",
          "t:Revenue from usage fee"]],

    70: [["t:P = 1.5", "cxn:2", "t:D", "t:$1.50"],
         ["cxn:3", "t:MC", "t:$0.50"],
         ["cxn:4", "cxn:5", "t:100"],
         ["osp:3", "t:Flat fee = $50", "t:Two-part tariff:",
          "t:Flat fee F", "n:sdbullmark:0:zfee"],
         ["t:Usage fee P", "n:sdbullmark:1:zline"],
         ["osp:4", "t:Revenues from Usage Fee", "t:Revenue from usage fee",
          "n:sdbullmark:2:zuse"],
         ["t:Would a flat fee", "n:sdbullmark:3:dot"],
         ["t:No \u2014 it would lead", "n:sdbullmark:4:rarr"]],

    73: [["t:P = " + ("%g" % _M6.ZOO_A), "cxn:2", "t:D",
          "t:" + _Z["a"], "t:" + _Z["qmax"]],
         ["cxn:3", "t:MC", "t:" + _Z["mc"]],
         ["cxn:4", "cxn:5", "t:" + _Z["qmc"]],
         ["osp:3", "t:Flat fee =", "t:Two-part tariff:", "t:Flat fee F",
          "n:sdbullmark:0:zfee"],
         ["t:Usage fee", "n:sdbullmark:1:zline"],
         ["osp:4", "t:Revenues from usage fee =",
          "t:Revenue from usage fee", "n:sdbullmark:2:zuse"],
         ["t:Revenues from usage fee recover", "n:sdbullmark:3:dot"],
         ["t:Profit:", "n:sdbullmark:4:dot"]],

    74: [["t:P = " + ("%g" % _M6.ZOO_A), "cxn:3", "t:D",
          "t:" + _Z["a"], "t:" + _Z["qmax"]],
         ["cxn:2", "t:MC", "t:" + _Z["mc"]],
         ["cxn:4", "t:MR"],
         ["t:MR = MC", "cxn:5", "cxn:6",
          "t:" + _Z["sp"], "t:" + _Z["sq"]],
         ["osp:3", "t:" + _Z["sprof"], "t:Simple Pricing:",
          "t:Simple pricing profits:", "n:sdbullmark:0:zrev"],
         ["t:Profit:", "n:sdbullmark:1:dot"]],

    # (slide 74's plan lives further down; a duplicate block here was
    # DEAD -- a dict literal keeps the last key, so this one never ran and
    # still named the pre-rescale $20 / $4 zoo.  Removed 2026-09-15.)

    # 100 <- NV 63, the ice-cream blocks: one scoop at a time, each with
    #  its price and quantity, then the marginal-price note.
    80: [["pr:Marginal cost of printing:0:0"],
         ["pr:Marginal cost of printing:1:1"],
         ["cxn:2", "t:D"],
         ["grp:0", "cxn:3", "t:MC", "t:$0.05"],
         ["pr:Marginal cost of printing:2:2"],
         ["osp:3", "t:$0.25", "t:100", "t:A"],
         ["pr:Marginal cost of printing:3:3"],
         ["pr:Marginal cost of printing:4:4"],
         ["osp:4", "t:$0.20", "t:125", "t:B"],
         ["pr:Marginal cost of printing:5:5"],
         ["pr:Marginal cost of printing:6:6"],
         ["osp:5", "t:$0.10", "t:175", "t:C"],
         ["pr:Marginal cost of printing:7:7"]],

    # 58 <- NV 64: the tree walks down the questions, each with the
    #  outcome it rules out, then the direct-PD branch as the payoff.
    81: [["t:Purchase options"],
          ["cxn:2", "cxn:3", "cxn:4", "cxn:5", "t:D"],
          ["cxn:6", "t:MC", "t:$1.00"],
          ["osp:3", "t:$4.00", "t:1", "t:A"],
          ["osp:4", "t:$3.00", "t:2", "t:B"],
          # the three explanatory lines went when his own wording was
          # adopted (2026-09-14); the photo closes the build instead
          ["osp:5", "t:$1.50", "t:3", "t:C"],
          ["pic:0"]],

    # 38 <- NVapp 19.  The takeaway bar was landing on click 1, with the
    # Odyssey facts building after it -- the punchline first.  His own
    # slide 19 animates one highlight and nothing else, so the structure
    # here is the deck's: the point of the slide stays on screen, his two
    # clippings come first, then the card with its source, then the facts
    # in pairs, and the gold line lands last.
    39: [["pic:0", "pic:1"],
         ["grp:0", "pr:The Odyssey:0:0"],
         ["pr:The Odyssey:1:2"],
         ["pr:The Odyssey:3:4"],
         ["t:Premium for:"]],

    # 46 <- NV 34, adopted beat for beat: three clicks, one picture each,
    # in his order (the Windows editions, then the two cans, then the
    # FastPass sign).  He leaves the lead line on screen throughout, and
    # the auto-rollout had been revealing it LAST, after all three
    # pictures -- the slide announced its own point only once the
    # examples were already up.
    47: [["pic:0"], ["pic:1"], ["pic:2"]],

    # 59 <- NV 45.  Nothing to fix -- listed so the next reader can see
    # it was checked: his slide 45 carries no timing at all.

    # 61 <- NV 47, adopted exactly: his ONE click is the closing note,
    # with the plan table and the Netflix shot on screen from the start.
    # The auto-rollout had the note first and the table second.
    62: [["t:This pricing strategy is a mix"]],

    75: [["t:P = " + ("%g" % _M6.ZOO_A), "cxn:2", "t:D",
          "t:" + _Z["a"], "t:" + _Z["qmax"]],
         ["cxn:3", "t:MC", "t:" + _Z["mc"]],
         ["osp:3", "t:Flat fee =", "t:Flat-Fee Pricing:", "t:Flat fee F",
          "n:sdbullmark:0:zfee"],
         ["osp:4", "t:Variable cost", "t:Costs due to visits",
          "n:sdbullmark:1:zcost"],
         # the wedge the flat fee never recovers, with its callout
         ["osp:5", "t:Excess variable cost"],
         ["t:Profit: " + _Z["ffprof"], "n:sdbullmark:2:dot"],
         ["t:Overall " + _Z["qmax"] + " visits", "n:sdbullmark:3:dot"],
         ["t:Two-part tariff: " + _Z["flat"]]],

    # 53.  2026-09-15 (Nico): "adopt EXACTLY the slide as is now.
    # Spacing, location of graphs, animations etc."  Pinned here so the
    # auto-rollout cannot drift it: the clipping and the first three
    # lines together, then the versions button, then the closing two
    # lines.
    54: [["pic:0", "pr:Response by major:0:2"],
         ["t:More seating versions"],
         ["pr:Response by major:3:4"]],

    # 89 <- NVapp 48: the three options are the SET-UP and stay on screen,
    # the way a worked-computation table does.  Then the question, then
    # what the class actually chose, then the punchline.  Without a plan
    # the auto-rollout gave the two shares a click each, which reads as
    # two findings rather than one split.
    90: [["t:What do you think about this"],
         ["t:16%", "t:84%"],
         ["t:Customer attention focused"]],

    # 91 <- NV 73: two PANELS, not three figures.  His own slide gives the
    # left pair one shared caption, so the pair plus that caption is one
    # beat; the baggage table and the caption above it are the other.
    # The auto-rollout had opened with the sale tags alone and only
    # brought their caption up two clicks later, alongside a different
    # picture.
    92: [["pic:0", "pic:1", "t:Prices ending in 9"],
         ["pic:2", "t:Adding a highly priced option"]],
    # 101 — the in-class copy of the tree
    84: [["t:Does the firm have market power", "t:Perfect competition produces"],
         # his 2026-09-15 rewording of the second outcome
         ["t:Can the firm prevent resale", "t:Simple pricing under"],
         ["t:Do the firm", "t:Advanced pricing strategies"],
         ["t:Can the firm directly identify",
          "t:Indirect (second-degree)"],
         ["t:Direct price discrimination"],
         ["t:Complete information on every",
          "t:Information on groups of customers"]],

    # ----------------------------------------------------------------
    #  In-class copies of video charts — the same functions, so the same
    #  plans.  Keeping them in step matters: a student sees the video
    #  build first and the class build second.
    # ----------------------------------------------------------------
}


def q(ns, t):
    return "{%s}%s" % (ns, t)


class Counter:
    def __init__(self, start=3):
        self.n = start

    def __call__(self):
        v = self.n
        self.n += 1
        return v


def sptgt(spid, prg):
    if prg is None:
        return '<p:spTgt spid="%d"/>' % spid
    return ('<p:spTgt spid="%d"><p:txEl><p:pRg st="%d" end="%d"/>'
            '</p:txEl></p:spTgt>' % (spid, prg[0], prg[1]))


def effect_par(spid, prg, node_type, ids, direction="in"):
    """One fade effect.  ``direction`` is "in" (entrance, the default) or
    "out" (exit — used to reveal something by removing a mask laid over it,
    as on Nico's "Versioning: dolls" slide).

    The exit form mirrors what PowerPoint wrote on his original slide 36:
    presetClass="exit", the animEffect FIRST, then the visibility set to
    "hidden" at delay 499.  In the entrance form the set comes first and
    sets "visible" at delay 0 — the order really is reversed.
    """
    a, b, c = ids(), ids(), ids()
    tgt = sptgt(spid, prg)
    if direction == "out":
        return (
            '<p:par><p:cTn id="%d" presetID="10" presetClass="exit" '
            'presetSubtype="0" fill="hold" grpId="0" nodeType="%s">'
            '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            '<p:animEffect transition="out" filter="fade"><p:cBhvr>'
            '<p:cTn id="%d" dur="500"/><p:tgtEl>%s</p:tgtEl></p:cBhvr>'
            '</p:animEffect>'
            '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold">'
            '<p:stCondLst><p:cond delay="499"/></p:stCondLst></p:cTn>'
            '<p:tgtEl>%s</p:tgtEl><p:attrNameLst><p:attrName>'
            'style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:to><p:strVal val="hidden"/></p:to></p:set>'
            '</p:childTnLst></p:cTn></p:par>'
            % (a, node_type, b, tgt, c, tgt))
    return (
        '<p:par><p:cTn id="%d" presetID="10" presetClass="entr" '
        'presetSubtype="0" fill="hold" grpId="0" nodeType="%s">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '<p:tgtEl>%s</p:tgtEl><p:attrNameLst><p:attrName>'
        'style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
        '<p:to><p:strVal val="visible"/></p:to></p:set>'
        '<p:animEffect transition="in" filter="fade"><p:cBhvr>'
        '<p:cTn id="%d" dur="500"/><p:tgtEl>%s</p:tgtEl></p:cBhvr>'
        '</p:animEffect></p:childTnLst></p:cTn></p:par>'
        % (a, node_type, b, tgt, c, tgt))


def click_group(beat, ids):
    """A beat is a list of (spid, prg) or (spid, prg, direction) tuples."""
    outer, inner = ids(), ids()
    effs = "".join(
        effect_par(item[0], item[1],
                   "clickEffect" if i == 0 else "withEffect", ids,
                   item[2] if len(item) > 2 else "in")
        for i, item in enumerate(beat))
    return (
        '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst>'
        '<p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
        '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst>'
        '<p:cond delay="0"/></p:stCondLst><p:childTnLst>%s'
        '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'
        % (outer, inner, effs))


def timing_xml(beats, para_boxes):
    ids = Counter(3)
    groups = "".join(click_group(b, ids) for b in beats)
    bld = "".join('<p:bldP spid="%d" grpId="0" build="p"/>' % spid
                  for spid in sorted(set(para_boxes)))
    bldlst = '<p:bldLst>%s</p:bldLst>' % bld if bld else ""
    return (
        '<p:timing xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:tnLst>'
        '<p:par><p:cTn id="1" dur="indefinite" restart="never" '
        'nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" '
        'nextAc="seek"><p:cTn id="2" dur="indefinite" '
        'nodeType="mainSeq"><p:childTnLst>%s</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl>'
        '<p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst>'
        '<p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl>'
        '</p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn>'
        '</p:par></p:tnLst>%s</p:timing>' % (A, R, P, groups, bldlst))


def shape_info(el, kinds):
    cnv = el.find(".//" + q(P, "cNvPr"))
    xf = el.find(".//" + q(A, "xfrm"))
    x = y = w = h = 0.0
    if xf is not None:
        o = xf.find(q(A, "off"))
        e = xf.find(q(A, "ext"))
        if o is not None and e is not None:
            x, y = int(o.get("x")) / EMU, int(o.get("y")) / EMU
            w, h = int(e.get("cx")) / EMU, int(e.get("cy")) / EMU
    paras = []
    for idx, pp in enumerate(el.iter(q(A, "p"))):
        lvl = 0
        pPr = pp.find(q(A, "pPr"))
        if pPr is not None and pPr.get("lvl"):
            lvl = int(pPr.get("lvl"))
        txt = "".join(t.text or "" for t in pp.iter(q(A, "t"))).strip()
        paras.append((idx, lvl, bool(txt)))
    tag = ET.QName(el).localname
    txt = " ".join(t.text or "" for t in el.iter(q(A, "t")))
    txt = (txt + " " + " ".join(t.text or ""
                                for t in el.iter(q(M, "t")))).strip()
    txt = " ".join(txt.split())      # normalize run-boundary whitespace
    # runs are joined with a space above, which splits symbols the
    # subscript pass broke into "P" + "0". Keep a concatenated variant so
    # t:/pr: selectors can still name them (2026-08-23).
    tight = "".join(t.text or "" for t in el.iter(q(A, "t")))
    tight = (tight + "".join(t.text or "" for t in el.iter(q(M, "t"))))
    tight = " ".join(tight.split())
    # PowerPoint sets OMML variables in the Mathematical Alphanumeric
    # block, so a formula box reads as "𝑃 = 4 - 𝑄 3"
    # and a plain "t:P = 4" selector never matches it.  NFKD folds those
    # codepoints back to ASCII, which is the same normalisation
    # _diff_slides.py already does for the geometry diff (2026-09-15).
    tight = unicodedata.normalize("NFKD", tight)
    # a curve is a custGeom outline with no fill; a chart backing card is
    # a preset rect with a solid fill.  is_chrome needs to tell them apart.
    spPr = el.find(q(P, "spPr"))
    custgeom = spPr is not None and spPr.find(q(A, "custGeom")) is not None
    solidfill = spPr is not None and spPr.find(q(A, "solidFill")) is not None
    info = {
        "custgeom": custgeom, "solidfill": solidfill,
        "id": int(cnv.get("id")), "name": cnv.get("name") or "",
        "tag": tag, "x": x, "y": y, "w": w, "h": h, "text": txt,
        "tight": tight,
        "paras": paras,
        "bul0": sum(1 for (_, l, ne) in paras if l == 0 and ne),
    }
    key = {"pic": "pic", "cxnSp": "cxn", "grpSp": "grp"}.get(tag)
    if key is None:
        key = "osp" if not txt else None
    if key:
        info["idx"] = "%s:%d" % (key, kinds[key])
        kinds[key] += 1
    else:
        info["idx"] = None
    return info


def collect_shapes(spTree):
    kinds = {"pic": 0, "cxn": 0, "grp": 0, "osp": 0}
    shapes = []
    for el in spTree:
        tag = ET.QName(el).localname
        if tag == "AlternateContent":
            ch = el.find(q(MC, "Choice"))
            if ch is not None and len(ch):
                el = ch[0]
                tag = ET.QName(el).localname
        if tag == "graphicFrame":
            cnv = el.find(".//" + q(P, "cNvPr"))
            xf = el.find(q(P, "xfrm"))
            x = y = 0.0
            if xf is not None and xf.find(q(A, "off")) is not None:
                o = xf.find(q(A, "off"))
                x, y = int(o.get("x")) / EMU, int(o.get("y")) / EMU
            shapes.append({"id": int(cnv.get("id")), "name": "",
                           "tag": "graphicFrame", "x": x, "y": y,
                           "w": 0, "h": 0, "text": "", "tight": "",
                           "paras": [], "bul0": 0,
                           "custgeom": False, "solidfill": False,
                           "idx": "gf:%d" % kinds.setdefault("gf", 0)})
            kinds["gf"] += 1
            continue
        if tag not in ("sp", "cxnSp", "pic", "grpSp"):
            continue
        if tag == "grpSp" and (el.find(".//" + q(P, "cNvPr")).get("name")
                               or "").startswith("sdpair:"):
            # A pair grouped by `_group_pass.py` rules 6 and 7 -- an arrow
            # with its label, a region with its label, a box with its link
            # button.  The pair now animates as ONE object, which is what
            # the house rule wants, but the members must stay ADDRESSABLE:
            # every plan in this file names them individually, and if the
            # group were opaque each `osp:` / `cxn:` ordinal after it would
            # shift and the plan would land on the footer rule (2026-09-10).
            #
            # So index the CHILDREN as before and point each of them at the
            # group. The plans keep working untouched, and a beat naming two
            # members of one pair emits one effect on the pair rather than
            # two (deduped in `build_timing`).
            # read the pair's id WITHOUT letting it consume a "grp:"
            # ordinal -- it is transparent, so the next real group on the
            # slide must still be grp:0 (this is what put slide 55's
            # box+text group at grp:3 and broke its plan)
            gid_el = el.find(".//" + q(P, "cNvPr"))
            grp = {"id": int(gid_el.get("id"))}
            for child in el:
                ctag = ET.QName(child).localname
                if ctag not in ("sp", "cxnSp", "pic"):
                    continue
                info = shape_info(child, kinds)
                info["id"] = grp["id"]          # the pair is the target
                info["group"] = True
                shapes.append(info)
            continue
        shapes.append(shape_info(el, kinds))
    return shapes


def is_chrome(s):
    t = s["text"].strip()
    # An axis and its title are chrome by the animation rules ("a chart's
    # axes and axis labels are visible from the start and are never
    # animated").  They are recognised by the name _fig_axes stamps, not
    # by geometry: an axis title sits a few tenths from the arrow tip and
    # so does the outermost TICK label, which does animate, with its
    # curve.  Slide 34 was giving "P" and "Q" a click each.
    if s.get("name", "").startswith(("sdaxis:", "sdaxistitle:")):
        return True
    if s["y"] + s["h"] <= 1.32 and not (1.3 < s["y"]):
        return True                      # top bar / tag / title / rule
    if s["y"] >= 7.05:
        return True                      # footer rules / text / page no
    if t.startswith("Management 405"):
        return True
    if s["x"] > 12.4 and s["tag"] == "sp" and len(t) <= 3:
        return True                      # page number
    if t.startswith("\u25b6"):
        return True                      # practice-video link boxes
    if t.startswith("\u2190 Back") or t.startswith("Tickets"):
        return True                      # navigation pills
    # EVERY bottom-right corner mark is chrome and is never animated
    # (Teaching CLAUDE.md, "Poll chrome is never animated"): the Poll Break
    # parallelogram, the round POLL pill, a Discussion badge, the
    # practice-video box, a problem-set pointer, a Back pill.  The
    # y >= 7.05 test above misses all of them, because a corner mark
    # STRADDLES the footer rule at 7.15" -- it starts ABOVE it.  This is
    # the same geometric test _raise_corner_marks uses to bring these to
    # the foreground, so the two passes agree on what a corner mark is.
    # Pictures are excluded so a full-bleed backup figure still animates,
    # and the width cap keeps it to marks (the widest is the 5.85"
    # practice-video box).  13 slides were fading their Discussion badge
    # in mid-build until this was added (2026-09-14).
    if (s["h"] >= 0.25 and s["y"] < 7.15 < s["y"] + s["h"]
            and s["tag"] != "pic" and s["w"] <= 6.5):
        return True                      # corner marks
    if (s["tag"] == "sp" and not t and s["w"] > 4 and s["h"] > 2.5
            and s.get("solidfill") and not s.get("custgeom")):
        return True                      # big white chart backings
    return False


def resolve(sel, shapes, used):
    """Resolve one selector to (spid, prg) or None.

    A selector prefixed "out:" marks an EXIT fade rather than an entrance
    (used to reveal something by removing a mask laid over it).  The prefix
    is stripped here and re-attached by the caller.

    Text selectors search the CONTENT shapes only: a slide title or a
    top-bar tag must never be animated, and it would otherwise win a
    prefix match ahead of the label the plan meant (Sales Tax / "S").
    """
    if sel.startswith(("t:", "pr:")):
        shapes = [s for s in shapes if not is_chrome(s)]
    if sel.startswith("pr:"):
        rest = sel.split(":", 1)[1]
        prefix, st, end = rest.rsplit(":", 2)
        for s in shapes:
            if (s["text"].startswith(prefix)
                    or s.get("tight", "").startswith(prefix)):
                # one effect per paragraph, so PowerPoint keeps them in
                # this click group (a single multi-para range effect
                # gets re-expanded into separate clicks on open)
                targets = [(s["id"], (i, i))
                           for i in range(int(st), int(end) + 1)]
                return targets, s["id"]
        raise KeyError(sel)
    if sel.startswith("t:"):
        body = sel[2:]
        nth = None
        if "#" in body:
            body, n = body.rsplit("#", 1)
            nth = int(n)
        matches = [s for s in shapes
                   if (s["text"].strip().startswith(body)
                       or s.get("tight", "").strip().startswith(body))]
        if nth is not None:
            if len(matches) < nth:
                raise KeyError(sel)
            return (matches[nth - 1]["id"], None), None
        for s in matches:
            if s["id"] not in used:
                used.add(s["id"])
                return (s["id"], None), None
        raise KeyError(sel)
    if sel.startswith("n:"):
        # exact shape-name selector (names emitted by _sd_chart and the
        # grouping pass) — immune to index shifts from grouping
        want = sel[2:]
        for s in shapes:
            if s["name"] == want and s["id"] not in used:
                used.add(s["id"])
                return (s["id"], None), None
        raise KeyError(sel)
    # indexed selectors
    for s in shapes:
        if s["idx"] == sel:
            # 2026-08-30: a plan that has drifted out of step with the
            # slide will happily land on the footer rule or the top bar and
            # animate it, with nothing in the output to say so (this is how
            # slide 52's osp:7 came to hide the footer rule when LMC turned
            # from a freeform into a connector).  Chrome is never a target.
            if is_chrome(s):
                raise KeyError(
                    "%s resolves to CHROME on this slide "
                    "(%s at [%.2f, %.2f]) - the plan is out of step with "
                    "the shape inventory" % (sel, s["tag"], s["x"], s["y"]))
            return (s["id"], None), None
    raise KeyError(sel)


def para_beats(box):
    groups = []
    cur = None
    for (idx, lvl, ne) in box["paras"]:
        if lvl == 0:
            cur = [(idx, ne)]
            groups.append(cur)
        else:
            if cur is None:
                cur = []
                groups.append(cur)
            cur.append((idx, ne))
    clean = [[i for (i, ne) in g if ne] for g in groups]
    clean = [g for g in clean if g]
    if not clean:
        return [], []
    return clean[0], clean[1:]


def default_plan(shapes, disp):
    """Bullet-driven default: first top-level bullet static, one click
    per further top-level bullet; figures + captions ride on the
    FIG_GROUP bullet; a gold takeaway bar (if any) last."""
    content = [s for s in shapes if not is_chrome(s)]
    for sel in STATIC.get(disp, ()):
        body = sel[2:]
        content = [s for s in content
                   if not s["text"].strip().startswith(body)]
    if not content:
        return [], []
    boxes = [s for s in content if s["tag"] == "sp" and s["bul0"] >= 2]
    bullets = max(boxes, key=lambda s: s["bul0"]) if boxes else None
    beats, para_ids = [], []
    takeaways = [s for s in content
                 if s["name"].startswith("Rounded Rectangle")
                 and s["y"] > 5.8 and s["text"].strip().startswith("\u2192")]
    t_ids = {s["id"] for s in takeaways}
    if bullets is not None:
        _static, groups = para_beats(bullets)
        bullet_beats = [[(bullets["id"], (i, i)) for i in g]
                        for g in groups]
        para_ids.append(bullets["id"])
        figs = [s for s in content
                if s["id"] != bullets["id"] and s["id"] not in t_ids]
        figs.sort(key=lambda s: (0 if s["tag"] in ("pic", "grpSp")
                                 else 1, s["y"], s["x"]))
        fig_targets = [(s["id"], None) for s in figs]
        if fig_targets:
            if not bullet_beats:
                bullet_beats.append(fig_targets)
            else:
                gi = FIG_GROUP.get(disp, FIG_GROUP_DEFAULT)
                gi = max(0, min(gi if gi >= 0
                                else len(bullet_beats) + gi,
                                len(bullet_beats) - 1))
                bullet_beats[gi] += fig_targets
        beats.extend(bullet_beats)
    else:
        # no bullet box: reveal pictures (with nearest caption), then
        # remaining text boxes top-to-bottom
        pics = [s for s in content if s["tag"] in ("pic", "grpSp")]
        caps = [s for s in content if s["tag"] == "sp"
                and s["text"].strip() and s["id"] not in t_ids]
        cap_of = {}
        for c in caps:
            if pics:
                best = min(pics, key=lambda p: (p["x"] + p["w"] / 2
                                                - c["x"] - c["w"] / 2) ** 2
                           + (p["y"] + p["h"] / 2 - c["y"] - c["h"] / 2) ** 2)
                if abs(best["y"] + best["h"] - c["y"]) < 0.4 or \
                   abs(c["y"] + c["h"] - best["y"]) < 0.4:
                    cap_of.setdefault(best["id"], []).append(c["id"])
        capped = {cid for v in cap_of.values() for cid in v}
        for p_ in sorted(pics, key=lambda s: (round(s["y"] / 1.4), s["x"])):
            beats.append([(p_["id"], None)]
                         + [(cid, None) for cid in cap_of.get(p_["id"], [])])
        for s in sorted([c for c in caps if c["id"] not in capped],
                        key=lambda s: s["y"]):
            beats.append([(s["id"], None)])
    for t in takeaways:
        beats.append([(t["id"], None)])
    return beats, para_ids


def custom_plan(shapes, disp):
    used = set()
    beats, para_ids = [], []
    for beat_sel in PLANS[disp]:
        beat = []
        for sel in beat_sel:
            # "out:" marks an EXIT fade (reveal by removing a mask)
            direction = "in"
            if sel.startswith("out:"):
                direction, sel = "out", sel[4:]
            (target, pbox) = resolve(sel, shapes, used)
            if isinstance(target, list):
                beat.extend((spid, prg, direction) for spid, prg in target)
            else:
                beat.append((target[0], target[1], direction))
            if pbox is not None:
                para_ids.append(pbox)
        # A beat that names two members of one PAIR group (the region and
        # its label, say) resolves both to the pair's own id, so drop the
        # duplicate rather than fading the same object twice (2026-09-10).
        seen, uniq = set(), []
        for t in beat:
            if t[1] is None:
                if t[0] in seen:
                    continue
                seen.add(t[0])
            uniq.append(t)
        beats.append(uniq)
    return beats, para_ids


def inventory(order, data, disps):
    """Print the animatable shapes of each slide, in document order, with
    the selector that reaches them.  This is what the PLANS below are
    authored from; chrome is marked so it can be ignored."""
    for disp in disps:
        part = "ppt/slides/" + order[disp - 1]
        tree = ET.fromstring(data[part])
        spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
        print("===== slide %d =====" % disp)
        for s_ in collect_shapes(spTree):
            txt = (s_["text"] or "").strip().replace("\n", " / ")[:52]
            print("  %-9s %-16s [%5.2f,%5.2f %4.2fx%4.2f]%s %s"
                  % (s_["idx"], s_["tag"] + ":" + s_["name"][:14],
                     s_["x"], s_["y"], s_["w"], s_["h"],
                     "  CHROME" if is_chrome(s_) else "        ", txt))


def main():
    args = sys.argv[1:]
    apply = "apply" in args
    sel = [a for a in args
           if a != "apply" and not a.endswith(".pptx")]
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target") for r in
             ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    order = [os.path.basename(rid2t[s.get(q(R, "id"))])
             for s in pres.find(q(P, "sldIdLst"))]

    if sel and sel[0] == "inventory":
        rest = [int(x) for x in sel[1:]]
        inventory(order, data, rest or range(1, len(order) + 1))
        return

    if sel == ["all"] or not sel:
        todo = [d for d in range(1, len(order) + 1) if d not in SKIP]
    else:
        todo = [int(x) for x in sel]

    PLANS.update(PLANS_PRE)
    for disp in todo:
        part = "ppt/slides/" + order[disp - 1]
        tree = ET.fromstring(data[part])
        spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
        shapes = collect_shapes(spTree)
        if disp in PLANS:
            beats, para_ids = custom_plan(shapes, disp)
        else:
            beats, para_ids = default_plan(shapes, disp)
        if not beats:
            print("s%02d: no beats — skipped" % disp)
            continue
        # strip existing timing, then append the new block
        for t_el in tree.findall(q(P, "timing")):
            tree.remove(t_el)
        timing = ET.fromstring(timing_xml(beats, para_ids))
        tree.append(timing)
        data[part] = ET.tostring(tree, xml_declaration=True,
                                 encoding="UTF-8", standalone=True)
        print("s%02d: %d clicks (%d effects)"
              % (disp, len(beats), sum(len(b) for b in beats)))

    if apply:
        tmp = DECK.with_suffix(".anim_tmp.pptx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, blob in data.items():
                zout.writestr(name, blob)
        os.replace(tmp, DECK)
        print("written:", DECK)
    else:
        print("(dry run — pass 'apply' to write)")


if __name__ == "__main__":
    main()
