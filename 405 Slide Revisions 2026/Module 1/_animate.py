# -*- coding: utf-8 -*-
"""Inject Fade / on-click build animations via OOXML <p:timing> —
Module 2 In-Class (2026-08-15). Engine from Module 7 / Italy IBR.

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

Pipeline:  _build_Module2InClass.py -> _splice_media.py -> _animate.py all apply
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
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
import sys as _sys
_deckarg = [a for a in _sys.argv[1:]
            if a.endswith(".pptx")]
DECK = Path(_deckarg[0]) if _deckarg else     HERE / "Module 1 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
EMU = 914400.0

# display numbers AFTER the 2026-08-20 inserts (#23 AC solution,
# #37-38 copper); the PLANS dict below still uses the pre-insert keys
# and is shifted by _m1_shift_key at the bottom of the config.
SKIP_TITLE = {1, 62, 65, 70, 80}         # deck title + 4 video titles
SKIP_AGENDA = {15, 16, 17, 39, 51, 56,   # roadmap + outline slides
               63, 64, 66, 71, 81}
SKIP_MEDIA = {22, 25, 45}                # spliced PollEv slides
SKIP_STATIC = {30, 49, 73}               # single-figure / single-thought
SKIP = SKIP_TITLE | SKIP_AGENDA | SKIP_MEDIA | SKIP_STATIC

# figures ride on this animated-bullet index by default (0 = first
# built bullet); per-slide overrides:
FIG_GROUP_DEFAULT = 0
FIG_GROUP = {}

PLANS_PRE = {}

STATIC = {}

# ---------------------------------------------------------------------------
# Custom story plans (selector language in the header docstring).
# Module 1 - Revised (84 slides), 2026-08-20.
# ---------------------------------------------------------------------------
PLANS = {
    9: [   # find a model: each question+map pair; maps paint over each other
        ["pr:Do you need to drive:2:3", "pic:1"],
        ["pr:Do you need to drive:4:5", "pic:2"],
    ],
    10: [  # homo economicus: panel + the whole Homo-Economicus row are
           # static; row 2, then the podcast label.
           # 2026-08-23 (Nico): he dropped pic:0 from the first beat, and
           # the only effect left on it targeted the slide TITLE (the
           # t: selector matched the action title before the panel
           # heading) - so the slide opened untitled. Beat removed
           # entirely: chrome stays put, and row 1 is fully static.
        ["pic:1", "t:Real Human"],
        ["t:Podcast: Who is Homo"],
    ],
    11: [  # hedgehogs: quote static; foxes + art, hedgehogs + art
        ["pr:“The fox knows:1:1", "pic:0"],
        ["pr:“The fox knows:2:2", "pic:1"],
    ],
    12: [  # making the most 1: bullets static; the 4 illustrations together
        ["pic:0", "pic:1", "pic:2", "pic:3"],
    ],
    20: [  # netflix: pic static; one option per click; badge last
        ["pr:Define Netflix:1:1"],
        ["pr:Define Netflix:2:2"],
        ["pr:Define Netflix:3:3"],
        ["osp:3", "t:Class Discussion"],
    ],
    21: [  # heatwaves: headline + badge static; AC map on click
        ["pic:1"],
    ],
    23: [  # swiftonomics: photos + badge static; question on click
        ["t:How did Taylor Swift"],
    ],
    25: [  # swift solution: photos static; verdict static; quotes build
        ["pr:Taylor Swift’s engagement:1:1"],
        ["pr:Taylor Swift’s engagement:2:2"],
    ],
    26: [  # tea: article + photo static; price chart on click
        ["pic:2"],
    ],
    27: [  # market for tea: D, S, old eq, shifts, new eq
        ["pr:Demand curve::0:0", "cxn:6", "t:D#2"],
        ["pr:Demand curve::1:1", "cxn:8", "t:S#1"],
        ["cxn:2", "cxn:3", "t:P0", "t:Q0"],
        ["cxn:7", "t:D#3", "cxn:10"],
        ["cxn:9", "t:S#2", "cxn:11"],
        ["cxn:4", "cxn:5", "t:P1", "t:Q1"],
    ],
    28: [  # disasters: photo + Q1 static; Q2; badge last
        ["pr:Why are there:1:1"],
        ["osp:3", "t:Class Discussion"],
    ],
    30: [  # avocados: photos static; price facts build
        ["pr:From December 2016:1:1"],
        ["pr:From December 2016:2:2"],
    ],
    31: [  # market for avocados — Nico's original choreography
           # (In Class slide 30), 10 clicks, addressed by shape name.
        ["n:sdgroup:D", "n:sdgroup:S"],              # 1 initial D and S
        ["n:sdguide:h:0", "n:sdgroup:Q0"],           # 2 P0 guide + Q0
        ["pr:Demand curve::0:0"],                    # 3 demand-craze bullet
        ["n:sdcurve:D1", "n:sdlabel:D1"],            # 4 D shifts out
        ["pr:Demand curve::1:1"],                    # 5 dry-weather bullet
        ["n:sdcurve:S1", "n:sdlabel:S1", "n:sdarrow:0"],   # 6 S left
        ["n:sdgroup:Q1"],                            # 7 PPeak / Q1 set
        ["pr:Demand curve::2:2"],                    # 8 pest-control bullet
        ["n:sdcurve:S2", "n:sdlabel:S2", "n:sdarrow:1"],   # 9 S back right
        ["n:sdguide:h:2", "n:sdgroup:Q2", "n:sdylab:P2 = P0"],  # 10 P2 + Q2
    ],
    33: [  # wheat: chart static; war line, red callout, grain line, green
        ["cxn:0", "t:War begins"],
        ["t:Supply shifts left", "cxn:2"],
        ["cxn:1", "t:Black sea grain"],
        ["t:Supply shifts right", "cxn:3"],
    ],
    34: [  # LA case: photo + lead static; the two facts; prompt; badge
        ["pr:From Sep-2021:1:1"],
        ["pr:From Sep-2021:2:2"],
        ["pr:From Sep-2021:3:3"],
        ["osp:3", "t:Class Discussion"],
    ],
    35: [  # LA market: setup + old eq, demand shift, supply shift, new eq
        ["cxn:6", "t:D0", "cxn:8", "t:S0", "cxn:2", "cxn:3",
         "t:P0", "t:Q0"],
        ["pr:Demand curve::0:0", "cxn:7", "t:D1", "cxn:10"],
        ["pr:Demand curve::1:1", "cxn:9", "t:S1", "cxn:11"],
        ["cxn:4", "cxn:5", "t:P1", "t:Q1"],
    ],
    37: [  # opportunity costs: podcast static; implicit block; cost box
        ["pr:The opportunity:2:4"],
        ["grp:0"],
    ],
    38: [  # fruit table: lead static; table; definition; preferences
        ["grp:0"],
        ["pr:The opportunity cost:0:0"],
        ["pr:The opportunity cost:1:1"],
    ],
    40: [  # MBA cost: prompt static; remote question; badge
        ["pr:Distinguish explicit:6:6"],
        ["osp:3", "t:Class Discussion"],
    ],
    41: [  # flip house: flip block static; consulting; question
        ["pr:Flip house::5:7"],
        ["pr:Flip house::8:8"],
    ],
    44: [  # another opp cost: lead static; 500K point; podcast + banner
        ["pr:Additionally::1:1"],
        ["pr:Additionally::2:2", "pic:0"],
    ],
    45: [  # child cost: lead static; figure; which-costs question
        ["pic:0"],
        ["pr:Data from the U.S.:9:9"],
    ],
    51: [  # concorde: photo + innovation static; losses; rational decision
        ["pr:Very coveted:3:5"],
        ["pr:Very coveted:6:6"],
    ],
    52: [  # sunk take-away 2x2: reading order across the columns
        ["pr:Often sunk costs:0:0"],
        ["pr:Make an optimal:1:1"],
        ["pr:Often sunk costs:1:1"],
    ],
    54: [  # CBA: objective static; marginal block; MB>MC; rule; payoff
        ["pr:Objective::1:4"],
        ["pr:Objective::5:5"],
        ["pr:Objective::6:6"],
        ["pr:Objective::7:7"],
    ],
    55: [  # exercise bars: hour labels stay as static setup; bars build
        # hour by hour with their verdicts; optimum + totals last
        ["osp:4", "osp:5", "t:MC#1", "t:Net Benefit of Hour 1", "osp:6"],
        ["osp:7", "osp:8", "t:MC#2", "t:Net Benefit of Hour 2", "osp:9"],
        ["osp:10", "osp:11", "t:MC#3", "osp:12"],
        ["osp:13", "t:MC#4", "t:indifferent"],
        ["osp:14", "osp:15", "t:MC#5", "osp:16", "pic:0"],
        ["t:Optimum where"],
        ["t:Total", "osp:3", "cxn:0"],
    ],
    56: [  # continuous MB=MC: axes + lead static; MB, MC, Q*
        ["cxn:2", "t:MB (Marginal"],
        ["cxn:3", "t:MC (Marginal"],
        ["cxn:4", "t:Q* (optimum)"],
    ],
    64: [  # market definition — Nico's 2026-08-24 build: "Extent of
           # market" arrives together with the question under it, then
           # the two sub-points one at a time, then geography with its
           # example
        ["pr:A company must:1:1"],
        ["pr:A company must:2:2"],
        ["pr:A company must:3:4"],
        ["pr:A company must:5:5"],
        ["pr:A company must:6:6"],
        ["pr:A company must:7:8"],
    ],
    # 65 (display 75, the video Netflix slide) is STATIC as of
    # 2026-08-24 — see the SKIP_STATIC override below. Nico removed
    # the build along with the Covid bullet.
    66: [  # actors: consumers card static; workers; firms
        ["t:WORKERS", "t:Choose a job"],
        ["t:FIRMS", "t:Employ workers"],
    ],
    71: [  # ceteris paribus — Nico's 2026-08-23 choreography: the grouped
           # header+cones panel first, then "assume", then the list
        ["n:sdgroup:cones"],
        ["pr:Want to know::1:1"],
        ["pr:Want to know::2:7"],
    ],
    72: [  # demand curve: def box static; statement + D curve together
        ["t:The demand curve is", "cxn:2", "t:D#2"],
    ],
    73: [  # movement vs shift (D) — Nico's 2026-08-23 choreography:
           # i) movement, then the ii) text, then the D' group, then the
           # Q3 group (which carries his new horizontal dashed segment)
        ["pr:Distinguish between::1:1", "cxn:7", "t:i)", "cxn:4", "cxn:5",
         "t:P2", "t:Q2"],
        ["pr:Distinguish between::2:3"],
        ["n:sdgroup:Dp"],
        ["n:sdgroup:Q3"],
    ],
    74: [  # AI and the demand for chips (2026-08-23): photo, then the
           # starting demand curve at P1, then the outward shift to D’
        # 2026-08-24 (Nico): the opening bullet and the photo now show
        # with the slide; the build is D, then the shift to D’
        ["n:sdgroup:D", "n:sdgroup:Q1"],
        ["n:sdgroup:Dp"],
    ],
    75: [  # supply curve: def box static; statement + S curve together
        ["t:Upward sloping", "cxn:2", "t:S"],
    ],
    76: [  # movement vs shift (S) — Nico's 2026-08-23 choreography:
           # S with its starting point (P1, Q1) static; then i) movement
           # with the new (P2, Q2) guides; then the whole S' set.
        ["pr:Distinguish between::1:2", "n:sdgroup:i", "n:sdgroup:P2"],
        ["pr:Distinguish between::3:4", "n:sdgroup:Sp"],
    ],
    79: [  # market mechanism: eq bullet + curves + eq point; P1 story; P2
        ["pr:Market equilibrium::0:0", "cxn:6", "t:D", "cxn:7", "t:S",
         "cxn:2", "cxn:3", "t:P0", "t:Q0"],
        ["pr:Market equilibrium::1:1", "cxn:4", "cxn:8",
         "t:Excess supply", "t:P1"],
        ["pr:Market equilibrium::2:2", "cxn:5", "cxn:9",
         "t:Excess demand", "t:P2"],
    ],
    # 2026-08-24: Nico's choreography on the three equilibrium-change
    # slides, read out of the polished Video 4 deck. On all three the
    # opening D, S and old equilibrium now show WITH the slide; the
    # build is the shift, then the new equilibrium landing together
    # with the sentence that describes it.
    81: [  # shift in demand
        ["n:sdgroup:Dp"],
        ["n:sdgroup:Q1", "t:When the demand"],
    ],
    82: [  # shift in supply
        ["n:sdgroup:Sp"],
        ["n:sdgroup:Q1", "t:When the supply"],
    ],
    83: [  # both shifts; the note, then the problem-set pointer last
        ["n:sdgroup:shifts"],
        ["n:sdgroup:Q1", "pr:In this case:0:0"],
        ["pr:In this case:1:1"],
        ["t:✎ Problem Set 1"],   # glyph changed 2026-08-27 (➜ -> ✎)
    ],
    84: [  # shift table static; the Important rule box is the payoff
        ["grp:1"],
    ],
}


# ---------------------------------------------------------------------------
# 2026-08-20 inserts: shift the pre-insert PLANS keys (+1 from 23,
# +3 from 36), then add plans for the three NEW slides.
# ---------------------------------------------------------------------------
def _m1_shift_key(k):
    if k >= 36:
        return k + 3
    if k >= 23:
        return k + 1
    return k


PLANS = {_m1_shift_key(k): v for k, v in PLANS.items()}

PLANS[23] = [   # AC solution (MW #51): D + axes static; the answer beat
    ["t:The heatwaves", "cxn:3", "t:D’", "cxn:4"],
]
PLANS[37] = [   # copper case (MW #65): quantity chart static; price
    # overlay rides the price bullet; framework prompt last
    ["pr:Annual consumption:1:1", "pic:1"],
    ["pr:Annual consumption:2:2"],
]
PLANS[38] = [   # copper market (MW #66): setup, D shift, S shift, flat P
    ["cxn:6", "t:D0", "cxn:8", "t:S0", "cxn:2", "cxn:3", "t:Q0"],
    ["cxn:7", "t:D1", "cxn:10"],
    ["cxn:9", "t:S1", "cxn:11"],
    ["cxn:4", "cxn:5", "t:P1 = P0", "t:Q1"],
]


# ---------------------------------------------------------------------------
# 2026-08-22 inserts (deck now 99 slides): Econ&Coffee poll pair (7-8),
# poll results-view slides (25, 29, 50), BACKUP section (93-99). Shift
# the 87-deck keys, then extend the skip sets. Backup slides are static
# per the Animations rule; all 8 poll slides are spliced media.
# ---------------------------------------------------------------------------
def _m1_shift_key2(k):
    if k >= 46:
        return k + 5
    if k >= 26:
        return k + 4
    if k >= 23:
        return k + 3
    if k >= 7:
        return k + 2
    return k


PLANS = {_m1_shift_key2(k): v for k, v in PLANS.items()}
FIG_GROUP = {_m1_shift_key2(k): v for k, v in FIG_GROUP.items()}
STATIC = {_m1_shift_key2(k): v for k, v in STATIC.items()}
SKIP_TITLE = {_m1_shift_key2(k) for k in SKIP_TITLE} | {93}
SKIP_AGENDA = {_m1_shift_key2(k) for k in SKIP_AGENDA}
SKIP_MEDIA = ({_m1_shift_key2(k) for k in SKIP_MEDIA}
              | {7, 8, 25, 29, 50})
SKIP_STATIC = ({_m1_shift_key2(k) for k in SKIP_STATIC}
               | set(range(94, 100)))
SKIP = SKIP_TITLE | SKIP_AGENDA | SKIP_MEDIA | SKIP_STATIC


# ---------------------------------------------------------------------------
# 2026-08-23 insert: Nico copied the two Tapestry-Capri slides in from the
# Example Candidates deck at displays 73-74, so everything from 73 on moves
# down by two. Shift the 99-deck keys, then register the two new plans.
# ---------------------------------------------------------------------------
def _m1_shift_key3(k):
    return k + 2 if k >= 73 else k


PLANS = {_m1_shift_key3(k): v for k, v in PLANS.items()}
FIG_GROUP = {_m1_shift_key3(k): v for k, v in FIG_GROUP.items()}
STATIC = {_m1_shift_key3(k): v for k, v in STATIC.items()}
SKIP_TITLE = {_m1_shift_key3(k) for k in SKIP_TITLE}
SKIP_AGENDA = {_m1_shift_key3(k) for k in SKIP_AGENDA}
SKIP_MEDIA = {_m1_shift_key3(k) for k in SKIP_MEDIA}
SKIP_STATIC = {_m1_shift_key3(k) for k in SKIP_STATIC}
SKIP = SKIP_TITLE | SKIP_AGENDA | SKIP_MEDIA | SKIP_STATIC

# Nico's choreography for the two new slides, read out of his deck with
# _extract_timing.py. Display 73: the setup bullets one at a time, photos
# static. Display 74: the price ladder builds left to right (the mass-market
# card is static), then the three share cards with the provenance line
# riding the middle one, then the internal quote, then the court decision.
PLANS[73] = [
    ["pr:Aug 2023:1:1"],
    ["pr:Aug 2023:2:2"],
    ["pr:Aug 2023:3:3"],
]
PLANS[74] = [
    ["t:\u201cAccessible luxury\u201d"],
    ["t:True luxury"],
    ["t:Combined Tapestry", "t:59%"],
    ["t:77%", "t:(figures from documents"],
    ["t:83%"],
    ["t:\u201cBottom line"],
    ["t:Oct 2024:"],
]

# ---------------------------------------------------------------------------
# 2026-08-24: adopted from Nico's polished 'Videos Final' decks.
#  * display 75 (video Netflix) loses its build entirely.
#  * display 100 (window-tax backup) gains one: the property-tax block,
#    then the Back pill. It is the one backup slide that animates.
# ---------------------------------------------------------------------------
PLANS.pop(75, None)
SKIP_STATIC = (SKIP_STATIC | {75}) - {100}
SKIP = SKIP_TITLE | SKIP_AGENDA | SKIP_MEDIA | SKIP_STATIC
PLANS[100] = [
    ["t:Property tax:"],
    ["t:\u2190 Back"],
]

# ---------------------------------------------------------------------------
# 2026-08-26: the videos-first restructure. Everything above is keyed by the
# OLD (101-deck) display number, which is what the dated comments refer to;
# one map re-keys the whole config. The three slides the reorder ADDS carry
# no build: the two section dividers are single-word slides and the second
# copy of the module title slide is a title slide.
# ---------------------------------------------------------------------------
import sys as _sys                                            # noqa: E402
_sys.path.insert(0, str(HERE))
import _m1_order as _ORDER                                    # noqa: E402

PLANS = _ORDER.remap_keys(PLANS)
FIG_GROUP = _ORDER.remap_keys(FIG_GROUP)
STATIC = _ORDER.remap_keys(STATIC)
SKIP_TITLE = _ORDER.remap_set(SKIP_TITLE) | _ORDER.NEW_ONLY
SKIP_AGENDA = _ORDER.remap_set(SKIP_AGENDA)
SKIP_MEDIA = _ORDER.remap_set(SKIP_MEDIA)
SKIP_STATIC = _ORDER.remap_set(SKIP_STATIC)
SKIP = SKIP_TITLE | SKIP_AGENDA | SKIP_MEDIA | SKIP_STATIC

# ---------------------------------------------------------------------------
# 2026-08-27: the two Kroger–Albertsons slides adopted from the Example
# Candidates deck. Keyed by DISPLAY number (they have no old-deck twin).
# Chronology first: the case slide ends on "market definition would turn out
# to be crucial", and the resolution slide's gold decision bar is its final
# click.
# ---------------------------------------------------------------------------
PLANS[_ORDER.NEW_KROGER_CASE] = [
    # bullet 1 (the deal) and both photos are visible from the start
    ["pr:Oct 2022: Kroger agrees to buy:1:1"],   # "everything would hinge…"
    ["pr:Oct 2022: Kroger agrees to buy:2:2"],   # the firms' market
    ["pr:Oct 2022: Kroger agrees to buy:3:3"],   # the FTC's market
]
PLANS[_ORDER.NEW_KROGER_COSTCO] = [
    ["t:THE MARKET"],
    ["t:Outside the market", "t:Club stores", "t:Limited assortment",
     "t:Dollar & convenience", "t:Online-only sellers"],
    ["t:“A monthly trip to Costco"],
    ["t:Dec 2024: federal and state courts"],
]


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


def effect_par(spid, prg, node_type, ids):
    a, b, c = ids(), ids(), ids()
    tgt = sptgt(spid, prg)
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
    outer, inner = ids(), ids()
    effs = "".join(
        effect_par(spid, prg, "clickEffect" if i == 0 else "withEffect",
                   ids)
        for i, (spid, prg) in enumerate(beat))
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
    info = {
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
                           "idx": "gf:%d" % kinds.setdefault("gf", 0)})
            kinds["gf"] += 1
            continue
        if tag not in ("sp", "cxnSp", "pic", "grpSp"):
            continue
        shapes.append(shape_info(el, kinds))
    return shapes


def is_chrome(s):
    t = s["text"].strip()
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
    if (s["tag"] == "sp" and not t and s["w"] > 4 and s["h"] > 2.5):
        return True                      # big white chart backings
    return False


def resolve(sel, shapes, used):
    """Resolve one selector to (spid, prg) or None."""
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
            (target, pbox) = resolve(sel, shapes, used)
            if isinstance(target, list):
                beat.extend(target)
            else:
                beat.append(target)
            if pbox is not None:
                para_ids.append(pbox)
        beats.append(beat)
    return beats, para_ids


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
