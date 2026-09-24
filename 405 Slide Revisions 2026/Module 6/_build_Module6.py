# -*- coding: utf-8 -*-
# ==========================================================================
#  *** STALE since 2026-09-24 ***  Module 6 is taped and finalized.  The
#  source of truth is now "Module 6 - Final.pptx" (assembled from the taped
#  decks in "Recorded Video Slides" by _finalize.ps1 + _final_retag.py),
#  and it is edited in place.  Re-running this script regenerates only
#  "Module 6 - Revised.pptx", which no longer carries the edits made while
#  taping.
# ==========================================================================
#  _build_Module6.py — build script for "Module 6 - Revised.pptx"
#
#  Complex Pricing and Advanced Pricing Strategies, per the approved outline
#  in "Module 6 - Revised - outline.md".  This script is the SOURCE OF TRUTH
#  for the deck; hand edits made in PowerPoint are ported back into it.
#
#  Module 6 is taught entirely on video (8 videos) with the week-9 class
#  doing applications only, so the deck is built as a TAPED module: 8 video
#  title cards, "Module 6 · Video k · <topic>" tags, gold coverage pills,
#  then an In-Class Examples block carrying the applications material.
#
#  The video block boundaries are NOT guesswork: _map_videos.py matched
#  every slide of the eight taped decks to a main-deck slide (all 44 content
#  slides at >= 0.71 text overlap, most at 1.00).  See section 2 of the
#  outline.  Each slide function names its source in a comment.
#
#  Pipeline (rerunnable, Module 7 pattern):
#      python _build_Module6.py        -> Module 6 - Revised.pptx
#      python _splice_media.py         -> polls, verbatim + notes + tags
#      python _group_pass.py           -> box+text / figure+shade groups
#      python _animate.py all apply    -> fade builds
# ==========================================================================

import sys
from pathlib import Path

from pptx import Presentation

# The final-exam date and slot are READ FROM THE COURSE CALENDAR, never
# typed here.  Slide 2 went stale once already (it still carried the
# March 2022 final), and then a second time when the exam moved from a
# three-day window to one fixed Saturday slot.  _calendar_content.py is
# the file both the calendar and the website build from, so importing it
# means the deck cannot disagree with either.  No fallback on purpose: if
# the calendar ever moves, this should fail loudly rather than quietly
# ship last year's date.
_CAL_DIR = (Path(__file__).resolve().parents[2]
            / "405 Calendar and Website" / "Course Calendar")
sys.path.insert(0, str(_CAL_DIR))
import _calendar_content as _CAL
sys.path.pop(0)

_FINAL_WEEK = next(w for w in _CAL.WEEKS if w.get("kind") == "final")
_FINAL_DATE = _CAL.dt(_FINAL_WEEK["num"], "Sat")
_FINAL_SLOT = _CAL.slot_label(_FINAL_WEEK["exam"])
_FINAL_HOURS = 3

import _m6_helpers as _H
from _m6_helpers import *                                     # noqa: F401,F403
from _m6_helpers import (
    CREAM, DARKRED, DIM, GOLD, GOLD_W, GRAY, MARGIN, NAVY, RED, RULE,
    RULE_W, SLIDE_H, SLIDE_W, WHITE, FADED, BLUE_PED, GREEN_DK,
    Inches, Pt, PP_ALIGN, MSO_ANCHOR, MSO_SHAPE, RGBColor, qn,
    _add_arrow, _add_arrow_shape, _add_convention_box, _add_drop_shadow,
    _add_hierarchical_bullets, _add_math_equation, _add_media_image,
    _add_mixed_textbox, _add_outlined_box, _add_pollbreak_badge,
    _add_ps_pointer, _add_rect, _add_rounded_filled_box, _add_styled_table,
    _add_ext_link_button, _add_jump_pill, _add_takeaway_bar, _add_text,
    _poly_key,
    _welfare_rows,
    _apply_picture_style,
    _blank_slide,
    _draw_action_title, _draw_footer, _draw_top_bar_tc,
    _fig_axes, _fig_curve_label, _fig_guide, _fig_line, _fig_poly,
    _fig_vbrace, _fig_xlab, _fig_ylab, _text_w_in,
    _omml_frac, _omml_run, _omml_sub, _omml_sup, _omml_text,
    _set_fill_alpha, _set_notes, _title_case, apply_symbol_subscripts,
    content_slide, make_diagram_slide, make_stub, SimpleFig,
)
from _m6_notes import NOTES_NV, NOTES_APP, NOTES_V1, NOTES_V2, NOTES_V3
from _m6_written_notes import WRITTEN_NOTES, NOTE_FIXES

OUT = Path(__file__).parent / "Module 6 - Revised.pptx"

# --------------------------------------------------------------------------
#  The OMML helpers concatenate raw XML and do not escape their argument, so
#  a literal "<" or ">" ("MC > 0", "MRPL < w") silently produces an invalid
#  slide part.  Module 6 is full of such comparisons, so wrap once here
#  rather than at every call site (the Module 4 fix).
# --------------------------------------------------------------------------

_omml_text_raw = _omml_text
_omml_run_raw = _omml_run


def _xml_escape(text):
    return (text.replace('&', '&amp;').replace('<', '&lt;')
                .replace('>', '&gt;'))


def _omml_text(text, **kw):                                   # noqa: F811
    return _omml_text_raw(_xml_escape(text), **kw)


def _omml_run(text, **kw):                                    # noqa: F811
    return _omml_run_raw(_xml_escape(text), **kw)


# ==========================================================================
#  Deck constants
# ==========================================================================

_H.FOOTER_TEXT = ("Management 405  ·  Module 6  ·  "
                  "Complex Pricing and Advanced Pricing Strategies")
FOOTER_TEXT = _H.FOOTER_TEXT

DECK_TITLE = "Complex Pricing"
DECK_SUB = "Advanced Pricing Strategies"

# 2026-09-23 (Nico): Module 6 now HAS an introduction video.  It is Video 1,
# its title card opens the deck, and the front matter (logistics, roadmap,
# the descriptive overview) sits inside its block, so per the taping rule
# those slides carry "Video 1".  The eight topic videos move up to 2-9.
# The introduction is NOT an outline row -- the agenda has no room for it
# -- so the outline starts at Video 2.
# (Before: no introduction video, and a two-level front-matter tag.)
INTRO_VIDEO_NAME = "Introduction to Module 6"
TAG_LOG = "Module 6 · Video 1 · Logistics"
TAG_ROADMAP = "Module 6 · Video 1 · Course Roadmap"
TAG_SUMMARY = "Module 6 · Summary"
TAG_BACKUP = "Module 6 · Backup"

# 2026-09-09 (Nico): the problem-set pointers keep the number each source
# slide carries -- the zoo comparison is mine ("Problem Set 5"), the BMW
# take-away is PG's ("Problem Set 4").  Numbers only, never exercise
# numbers, per the pointer convention; one constant each so a Fall 2026
# renumbering is a one-line change.
# The deck-wide surplus washes, fixed by Teaching CLAUDE.md and named
# the same way as in Module 4: consumer surplus is C0201B at ~26%,
# producer surplus 4E79B5 at ~34%.
CS_RED = RGBColor(0xC0, 0x20, 0x1B)
PS_BLUE = RGBColor(0x4E, 0x79, 0xB5)
PS_BLUE_LINE = RGBColor(0x2E, 0x5A, 0xA8)

PS_ZOO = "Problem Set 5"
PS_BMW = "Problem Set 4"

# ==========================================================================
#  The module outline — ONE source of truth for the agenda slides, the
#  coverage pills, the video title cards and the top-bar tags, so an item's
#  wording and its video number can never drift apart.
#
#  Item titles and video names are verbatim from the Fall 2026 course
#  calendar's Session-7 prep block, so the deck and the calendar cannot
#  disagree about what a student is asked to watch.
#
#  2026-09-14 (Nico): "Throughout in the agenda: say 'Perfect Price
#  Discrimination (First Degree)' and similar for the other degrees -- just
#  like in my original slides.  Also, adopt the layout from my original
#  slides for the agenda, with the sub-headers."  His master deck's slide 4
#  groups the topics under three umbrellas, so the agenda now carries HEADER
#  rows and indents the topics beneath them.
#
#  A header costs almost nothing vertically because it carries no number, no
#  description and no pill -- one 17 pt line.  That is what makes the extra
#  rows affordable: ten rows solve to a 0.626" pitch with 21 pt titles over
#  18 pt descriptions.
#
#  2026-09-14 (Nico), second pass: the two umbrellas are NUMBERED items on
#  Module 2's convention -- "2. Direct Price Discrimination" with 2a and 2b
#  under it, exactly as Module 2 carries "3. Demand and revenue" with 3a and
#  3b (M2_OUTLINE in `Module 2/_build_Module2InClass.py`).  The geometry,
#  the banding and the dimming below are that deck's, scaled to this one's
#  smaller pitch.
#
#  The ONE deviation: an umbrella row reserves no DESCRIPTION row.  Module 2
#  gives its parent one, but Module 2 has six rows and this deck has ten --
#  measured, ten full two-line rows floor the type at 19 pt titles over
#  17 pt descriptions, while dropping the two umbrella descriptions holds it
#  at 22 over 19.  An umbrella here is a pure grouping label anyway (it owns
#  no video and no slides), so its children carry the explanation.
#
#  Row = (topic, number, title, description, kind).  ``topic`` is the index
#  every other map in this file is keyed by: the eight taped topics keep the
#  indices 0-7 they had while the outline was flat, so ``highlight_idx=3``
#  and the TAG_* constants below go on meaning what they meant.  ``kind`` is
#  "main", "sub" (indented under the parent above it) or "parent".
# ==========================================================================

M6_ROWS = [
    (0, "1", "Simple vs. Complex Pricing",
     "The pricing dilemma and why complex pricing pays", "main"),
    (None, "2", "Direct Price Discrimination", "", "parent"),
    (1, "2a", "Perfect Price Discrimination (First Degree)",
     "Charging every customer her own willingness to pay", "sub"),
    (2, "2b", "Segment Pricing (Third Degree)",
     "Different prices for groups with different price sensitivity", "sub"),
    (3, "3", "Indirect Price Discrimination (Second Degree)",
     "Versioning and coupons let customers sort themselves", "main"),
    (None, "4", "Advanced Pricing Strategies", "", "parent"),
    (4, "4a", "Flat Fee Pricing",
     "Unlimited access for one all-or-nothing fee", "sub"),
    (5, "4b", "Two-Part Tariff",
     "A flat fee plus a usage fee set at marginal cost", "sub"),
    (6, "4c", "Block Pricing",
     "A lower price as the same customer buys more", "sub"),
    (7, "5", "Summary of Pricing Strategies",
     "Which strategy fits which market conditions", "main"),
]

# An umbrella's own coverage pill names the videos its children span, per
# the coverage-pill rule ("Videos 1+2" rather than a sequence number).
# 2026-09-23: shifted by one for the new introduction video (was
# "Videos 2+3" / "Videos 5-7").
PARENT_PILL = {1: "Videos 3+4", 5: "Videos 6\u20138"}

# the eight taped topics in topic-index order -- what the rest of the deck
# means by "an outline item"
M6_OUTLINE = [(r[1], r[2], r[3]) for r in
              sorted((r for r in M6_ROWS if r[0] is not None),
                     key=lambda r: r[0])]

# Module 6 is taped end to end, so every item is a video and every pill is
# gold.  The names are the calendar's own video names.
VIDEO_NAME = {
    0: "Simple vs. Complex Pricing",
    1: "First-Degree Price Discrimination",
    2: "Segment Pricing",
    3: "Versioning and Coupons",
    4: "Flat Fee Pricing",
    5: "Two-Part Tariffs",
    6: "Block Pricing",
    7: "Summary of Pricing Strategies",
}
IN_CLASS_ITEMS = set()
# topic i is Video i + 2: Video 1 is the introduction (2026-09-23)
VIDEO_OFFSET = 2
COVERAGE_LABEL = {i: ("In class" if i in IN_CLASS_ITEMS else "Video %d" % (i + VIDEO_OFFSET))
                  for i in range(len(M6_OUTLINE))}
ITEM_VIDEO = {i: (None if i in IN_CLASS_ITEMS else i + VIDEO_OFFSET)
              for i in range(len(M6_OUTLINE))}

TAG_BASE = {i: _title_case(row[1][0].upper() + row[1][1:])
            for i, row in enumerate(M6_OUTLINE)}


def _tag_for(i):
    """"Module 6 · Video k · <topic>" inside a video block, two levels
    outside one (Teaching CLAUDE.md, top-bar tag rule)."""
    v = ITEM_VIDEO[i]
    return ("Module 6 · Video %d · %s" % (v, TAG_BASE[i])) if v \
        else ("Module 6 · " + TAG_BASE[i])


TAG = {i: _tag_for(i) for i in range(len(M6_OUTLINE))}

TAG_SIMPLE = TAG[0]     # Simple vs. Complex Pricing
TAG_FIRST = TAG[1]      # First Degree: Perfect Price Discrimination
TAG_SEGMENT = TAG[2]    # Third Degree: Segment Pricing
TAG_VERSION = TAG[3]    # Second Degree: Versioning and Coupons
TAG_FLAT = TAG[4]       # Advanced Pricing: Flat Fee
TAG_TWOPART = TAG[5]    # Advanced Pricing: Two-Part Tariff
TAG_BLOCK = TAG[6]      # Advanced Pricing: Block Pricing
TAG_SUMTAG = TAG[7]     # Summary of Pricing Strategies

# The In-Class Examples block: a four-level tag naming the video topic being
# applied, spliced in rather than retyped, per the in-class tag rule.
def _tag_inclass(i):
    return "Module 6 · In Class · Examples · " + TAG_BASE[i]


TAG_IC = {i: _tag_inclass(i) for i in range(len(M6_OUTLINE))}
TAG_IC_AGENDA = "Module 6 · In Class · Agenda"

# how far a shaded (one-line) row is nudged down inside its reserved
# two-row box, so single-line rows sit centred — see Teaching CLAUDE.md
DIM_DROP = Inches(0.19)

# the coverage pill sits vertically centred in the reserved two-row box, so
# the pill column is identical on every agenda slide
PILL_DROP = Inches(0.16)


# ==========================================================================
#  Outline / agenda slides
# ==========================================================================

def make_m6_outline(prs, page_num, *, tag=None, title="Outline of Module 6",
                    descriptions=False, highlight_idx=None,
                    highlight_set=None, in_class=False):
    """The module outline in the format converged on for Modules 1 – 4: a
    gold circle carrying the item number, the item title beside it in bold
    navy, a one-line grey description underneath, and the coverage pill at
    the right.

    Every item RESERVES the description row, so item positions are
    pixel-identical on every agenda slide.  The description shows only for
    the current topic, or for all of them when ``descriptions=True``.
    Section agendas band the current item in cream and shade the rest.
    """
    slide = _blank_slide(prs)
    if tag is None:
        if in_class:
            tag = TAG_IC_AGENDA
        elif highlight_idx is not None and ITEM_VIDEO.get(highlight_idx):
            tag = "Module 6 · Video %d · Agenda" % ITEM_VIDEO[highlight_idx]
        else:
            tag = "Module 6 · Agenda"
    _draw_top_bar_tc(slide, tag)
    _draw_action_title(slide, title)

    hi = set()
    if highlight_idx is not None:
        hi.add(highlight_idx)
    if highlight_set:
        hi.update(highlight_set)
    if descriptions:
        hi = set(range(len(M6_OUTLINE)))

    n_rows = len(M6_ROWS)
    top = Inches(1.38)
    bottom = Inches(7.06)
    # 0.05" between rows rather than 0.07": each row already reserves its own
    # height, so the inter-row gap is decoration, and the 0.18" it gives back
    # over ten rows is what pays for the headers' lead-in below at 21 pt.
    gap = Inches(0.05)
    PITCH_MAX = Inches(0.91)          # the Module 2 pitch
    # 2026-09-09: relaxed from Module 4's 6.22" -- Module 6's agenda slides
    # carry no bottom-right pointer box (the pills own the right edge), so
    # the last row may sit lower.
    LAST_ROW_MAX = Inches(6.55)
    TITLE_PT_MAX = 24                 # the largest title type this pitch holds
    # An umbrella row is one title line, no description row.
    # A parent belongs to the rows BELOW it.  With only the ordinary gap
    # above it, it sat as close to the previous item's description as to its
    # own first sub-item and read as part of the wrong group, so it gets a
    # lead-in (caught on the PNG probe of slide 5, 2026-09-14).
    HDR_LEAD = Inches(0.10)

    def _parent_h(t_h):
        return int(t_h * 1.22)

    def _layout(pitch):
        """Row tops for a candidate pitch.  A header row is short, so the
        rows are laid out by accumulating their OWN heights rather than by
        multiplying one pitch -- which is what buys the headers their room.
        """
        content = pitch - gap
        # 0.58 rather than 0.53 (2026-09-14, Nico: "make the sub-titles have
        # smaller font").  Moving the split takes the description from 19 pt
        # to 17 and, because the title was being held below its 24 pt cap by
        # the old share, takes the ITEM TITLE up to 24 -- so the two rows now
        # read clearly as heading and gloss rather than as two similar lines.
        t_h = int(content * 0.58)
        d_h = content - t_h
        # A parent reserves ONE title line -- but the line RENDERS about
        # 22% taller than its point size, and reserving the bare t_h let the
        # sub-row's cream band and the sub-row's pill both ride up over the
        # parent (PNG probe of slide 14, 2026-09-14).
        hs = [_parent_h(t_h) if r[4] == "parent" else t_h + d_h
              for r in M6_ROWS]
        leads = [HDR_LEAD if (j and r[4] == "parent") else 0
                 for j, r in enumerate(M6_ROWS)]
        block = sum(hs) + sum(leads) + gap * (n_rows - 1)
        y0 = top + max(0, (bottom - top - block) // 2)
        tops, yy = [], y0
        for h, ld in zip(hs, leads):
            yy += ld
            tops.append(int(yy))
            yy += h + gap
        return tops, block, t_h, d_h

    pitch = PITCH_MAX
    while pitch > Inches(0.55):
        tops, block, title_h, desc_h = _layout(pitch)
        if block <= bottom - top and tops[-1] <= LAST_ROW_MAX:
            break
        pitch -= 4572                  # 0.005" steps
    tops, block, title_h, desc_h = _layout(pitch)

    t_size = min(TITLE_PT_MAX, int(title_h / 914400.0 * 72))
    d_size = min(18, int(desc_h / 914400.0 * 72))
    # Fail loudly rather than shipping type that overflows its row.
    if t_size < 20 or d_size < 16:
        raise ValueError(
            "outline rows too tight: pitch %.3f\" gives %d pt titles and "
            "%d pt descriptions - drop a row" % (pitch / 914400.0, t_size,
                                                 d_size))
    # a one-line (shaded) row is centred in its own reserved two-row box,
    # and the pill is centred on the same box -- both derived rather than
    # carried as fixed constants, so they follow the pitch
    dim_drop = int(desc_h // 2)
    pill_drop = int(max(0, (title_h + desc_h - Inches(0.36)) // 2))
    parent_h = _parent_h(title_h)
    parent_pill_drop = int(max(0, (parent_h - Inches(0.36)) // 2))

    def _children(r_i):
        """The topic indices of the sub-rows under the parent at r_i."""
        nxt = next((j for j in range(r_i + 1, n_rows)
                    if M6_ROWS[j][4] != "sub"), n_rows)
        return set(M6_ROWS[j][0] for j in range(r_i + 1, nxt))

    for r_i, (i, label, item, desc, kind) in enumerate(M6_ROWS):
        y = tops[r_i]
        sub = kind == "sub"
        parent = kind == "parent"

        # An umbrella owns no video, so it is never the "current topic" and
        # never carries a band.  It stays LIT while one of its own children
        # is current, though -- dimmed, the grouping it exists to show goes
        # invisible at exactly the moment the student needs it.
        lit = descriptions or (bool(hi & _children(r_i)) if parent
                               else i in hi)
        # Module 2's row geometry, scaled from its 25/22 pt type to this
        # deck's 22/19 (Module 2/_build_Module2InClass.py, make_m2_outline).
        circ_d = Inches(0.42) if sub else Inches(0.51)
        circ_x = Inches(1.62) if sub else Inches(1.15)
        circ_dy = Inches(0.06) if sub else Inches(0.02)
        text_x = Inches(2.28) if sub else Inches(2.05)
        num_pt = 15 if sub else t_size
        item_pt = t_size - 3 if sub else t_size

        if not descriptions and i in hi and not parent:
            # 2026-09-09: the band is sized from the RENDERED type, not from
            # the row's reserved height.  Eight rows put the pitch at 0.709",
            # whose 0.639" content box is shorter than a 24 pt title over a
            # 21 pt description actually needs -- sized off the reserved
            # height, the band's bottom border cut straight through the
            # description (caught on the PNG probe of slide 6).
            # 2026-09-14: back to the row's own reserved height plus a
            # small pad.  Sized off the RENDERED type (the 2026-09-09 fix),
            # the band was 0.82" tall against a 0.56" row and its bottom
            # border cut straight through the next item's title and circle
            # -- the tighter ten-row pitch turned that headroom into an
            # overlap.  The two lines of ink measure 0.52" inside a 0.56"
            # row, so the reserved height already contains them.
            # +0.115" because the two lines of type RENDER about 0.09"
            # taller than the box that reserves them (21 pt over 19 pt is
            # 0.667" of leading against a 0.564" row).  That is the whole
            # tolerance there is: any taller and the band's bottom border
            # runs into the next item's gold circle.
            band_h = int(title_h + desc_h + Inches(0.115))
            # Module 2's band geometry: an indented row gets an indented,
            # narrower band, so the nesting the umbrellas set up is not
            # undone by a full-width highlight.
            band_x = Inches(1.37) if sub else Inches(0.90)
            band_w = Inches(11.68) if sub else Inches(12.15)
            band = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, int(band_x),
                int(y - Inches(0.05)), int(band_w), band_h)
            try:
                band.adjustments[0] = 0.35
            except Exception:
                pass
            band.fill.solid()
            band.fill.fore_color.rgb = CREAM
            band.line.color.rgb = GOLD
            band.line.width = Pt(1.0)
            band.shadow.inherit = False
            _add_drop_shadow(band)

        circ = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, int(circ_x), int(y + circ_dy),
            int(circ_d), int(circ_d))
        circ.fill.solid()
        circ.fill.fore_color.rgb = GOLD
        circ.line.fill.background()
        circ.shadow.inherit = False
        tf = circ.text_frame
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.add_run()
        run.text = label
        run.font.size = Pt(num_pt)
        run.font.bold = True
        run.font.color.rgb = NAVY if lit else DIM
        run.font.name = "Calibri"

        rows = [([(_title_case(item[0].upper() + item[1:]),
                   {'bold': True, 'size': item_pt,
                    'color': NAVY if lit else DIM})], 0,
                 {'bullet_style': 'none', 'space_before_pts': 0})]
        if i in hi and not parent:
            rows.append(([(desc, {'size': d_size, 'color': GRAY})], 0,
                         {'bullet_style': 'none', 'space_before_pts': 0}))

        # the coverage pill: gold = on video, navy = in class, dimming with
        # its own row on a section agenda (Teaching CLAUDE.md)
        # a sub-item's pill is narrower and right-aligned to the SAME edge,
        # so it reads as part of its parent's rather than as a peer
        pill_w = Inches(1.14) if sub else Inches(1.55)
        pill_label = PARENT_PILL[r_i] if parent else COVERAGE_LABEL[i]
        pill = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, int(Inches(12.85) - pill_w),
            int(y + (parent_pill_drop if parent else pill_drop)),
            int(pill_w), int(Inches(0.36)))
        try:
            pill.adjustments[0] = 0.30
        except Exception:
            pass
        label_txt = pill_label
        on_video = label_txt.startswith("Video")
        pill.fill.solid()
        pill.line.fill.background()
        pill.shadow.inherit = False
        if lit:
            pill.fill.fore_color.rgb = GOLD if on_video else NAVY
            pill_fg = NAVY if on_video else WHITE
            _add_drop_shadow(pill)
        else:
            pill.fill.fore_color.rgb = DIM
            pill_fg = WHITE
        ptf = pill.text_frame
        ptf.margin_left = ptf.margin_right = 0
        ptf.margin_top = ptf.margin_bottom = 0
        ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
        ppara = ptf.paragraphs[0]
        ppara.alignment = PP_ALIGN.CENTER
        prun = ppara.add_run()
        prun.text = label_txt
        prun.font.size = Pt(13)
        prun.font.bold = True
        prun.font.name = "Calibri"
        prun.font.color.rgb = pill_fg

        # a description must clear the pill by at least a five-letter word,
        # or it reads as running into it (course rule)
        if i in hi and not parent:
            avail = (Inches(12.85) - pill_w - Inches(0.10)) - text_x
            if _text_w_in(desc, d_size) > avail / 914400.0:
                raise ValueError(
                    "outline description %r is %.2f\" wide but only %.2f\" "
                    "clears the pill - shorten the wording"
                    % (desc, _text_w_in(desc, d_size), avail / 914400.0))
        # and a TITLE must clear it too -- eight rows makes the titles long
        if _text_w_in(_title_case(item[0].upper() + item[1:]), item_pt,
                      bold=True) > (
                (Inches(12.85) - pill_w - Inches(0.10)) - text_x) / 914400.0:
            raise ValueError(
                "outline title %r does not clear the coverage pill" % item)

        # a parent row reserves only the title line, so it is neither
        # dropped nor given the description row's height
        _add_hierarchical_bullets(
            slide, text_x,
            y if (lit or parent) else int(y + dim_drop),
            int(Inches(12.85) - pill_w - Inches(0.18)) - text_x,
            parent_h if parent else title_h + desc_h,
            rows, size=item_pt, line_spacing_pts=0)

    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


# ==========================================================================
#  Title slide, video title cards, dividers
# ==========================================================================

def slide_title(prs):
    """Deck-standard title slide: centred, no top bar, no page number.

    2026-09-09: the UCLA Anderson wordmark my original slide 1 carries in
    the lower-left corner (NV_s01_3, 2.94" wide at 0.59 / 6.40) is DROPPED
    -- no institutional branding mark goes on any slide of this deck
    (Teaching CLAUDE.md).  The affiliation is in the byline below.
    """
    slide = _blank_slide(prs)
    _add_text(slide, 0, Inches(2.35), SLIDE_W, Inches(1.2), DECK_TITLE,
              size=60, bold=True, color=NAVY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_text(slide, 0, Inches(3.62), SLIDE_W, Inches(0.8), DECK_SUB,
              size=40, bold=True, color=GOLD, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_rect(slide, int((SLIDE_W - Inches(4.0)) / 2), Inches(4.62),
              Inches(4.0), 54864, GOLD)
    _add_text(slide, 0, Inches(4.96), SLIDE_W, Inches(0.55),
              "Management 405", size=26, bold=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _add_text(slide, 0, Inches(5.66), SLIDE_W, Inches(0.5),
              "Prof. Nico Voigtländer  ·  UCLA Anderson",
              size=22, color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
    _add_rect(slide, 0, Inches(7.15), SLIDE_W, Inches(0.02), RULE)
    _add_rect(slide, MARGIN, Inches(7.135), GOLD_W, Inches(0.05), GOLD)
    return slide


def video_card(prs, item_idx):
    """A video title card, placed immediately BEFORE that section's agenda
    slide.  Title-slide layout, no top bar, no footer text, no page number.
    Geometry from _video_title_slide in Module 2/_build_Module2Video.py.
    Carries NO speaker notes, per the taping rule.
    """
    return _video_card(prs, VIDEO_NAME[item_idx], ITEM_VIDEO[item_idx])


def intro_video_card(prs):
    """The introduction video's card: Video 1, the VERY FIRST slide of the
    deck, ahead of the title slide (2026-09-23)."""
    return _video_card(prs, INTRO_VIDEO_NAME, 1)


def _video_card(prs, name, k):
    slide = _blank_slide(prs)
    _add_text(slide, 0, Inches(2.10), SLIDE_W, Inches(1.1),
              name, size=60, bold=True, color=NAVY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _add_text(slide, 0, Inches(3.25), SLIDE_W, Inches(0.75),
              "Module 6  ·  Video %d" % k,
              size=40, bold=True, color=GOLD, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_rect(slide, int((SLIDE_W - Inches(4.0)) / 2), Inches(4.28),
              Inches(4.0), 54864, GOLD)
    _add_text(slide, 0, Inches(4.62), SLIDE_W, Inches(0.55),
              "Management 405", size=26, bold=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _add_text(slide, 0, Inches(5.32), SLIDE_W, Inches(0.5),
              "Prof. Nico Voigtländer  ·  UCLA Anderson",
              size=22, color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
    _add_rect(slide, 0, Inches(7.15), SLIDE_W, Inches(0.02), RULE)
    _add_rect(slide, MARGIN, Inches(7.135), GOLD_W, Inches(0.05), GOLD)
    return slide


def divider(prs, main, sub=None):
    """A section divider: no top bar, no tag, no page number."""
    slide = _blank_slide(prs)
    _add_text(slide, 0, Inches(2.85), SLIDE_W, Inches(1.2), main,
              size=54, bold=True, color=NAVY, font="Calibri",
              align=PP_ALIGN.CENTER)
    if sub:
        # 2026-09-14: he dropped the subtitle below the gold strip (4.10 ->
        # 5.05), so the divider reads title / rule / qualifier rather than
        # title / qualifier / rule.
        _add_text(slide, 0, Inches(5.05), SLIDE_W, Inches(0.7), sub,
                  size=26, color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
    _add_rect(slide, int((SLIDE_W - Inches(4.0)) / 2), Inches(4.95),
              Inches(4.0), 54864, GOLD)
    _add_rect(slide, 0, Inches(7.15), SLIDE_W, Inches(0.02), RULE)
    _add_rect(slide, MARGIN, Inches(7.135), GOLD_W, Inches(0.05), GOLD)
    return slide


# ==========================================================================
#  Front matter
# ==========================================================================

def slide_logistics(prs, page_num):
    """Logistics.  [NV 2]

    2026-09-09 (Nico: "fix the exam dates on slide 2 from the calendar").
    My original slide was the March 2022 final: "5 hours", "any time
    between March 19 and March 20", two windows with guaranteed TA
    availability on Zoom, and "many practice problems on Achieve".  Every
    one of those was out of date.

    2026-09-15 (Nico: "update the final exam date/time to the one listed
    on the calendar.  Also, the practice final will be on the class
    website").  The exam is no longer a three-day window a student picks a
    start inside -- the calendar now carries one fixed Saturday slot -- and
    the practice final has moved off BruinLearn.  Both the date and the
    slot are read from _calendar_content.py at build time (see the import
    at the top of this file), so this slide cannot drift from the calendar
    or the course website again.

    The website address is deliberately NOT on the slide: it differs by
    section (MGMT-405-EMBA and MGMT-405-FEMBA), and this deck is shown to
    both.
    """
    bullets = [
        ("Practice Final Exam on the class website — similar to the "
         "actual final", 0),
        ("Final exam: %s, %s"
         % (_FINAL_DATE.strftime("%A, %B %d, %Y").replace(" 0", " "),
            _FINAL_SLOT), 0),
        ("Online — %d hours to solve it and upload your scanned "
         "solutions" % _FINAL_HOURS, 1),
        ("Open book, open notes; covers Modules 1 – 8", 0),
        ("About 20 multiple-choice questions and 3 – 4 problems", 1),
    ]
    return content_slide(prs, page_num, TAG_LOG, "Some Logistics", bullets,
                         size=26, sub_size=24, line_spacing_pts=14)


def slide_roadmap(prs, page_num):
    """Course roadmap.  [NV 3]

    The deck-standard diamond, identical to Module 1's and Module 4's: box
    1 across the top, boxes 2 and 3 side by side, box 4 across the bottom,
    joined by four faded connector arrows.  Module 6 sits in box 4, so that
    box is navy while the earlier three are faded, and the gold block arrow
    seats just to its LEFT.

    2026-09-09: rebuilt to match the other decks exactly -- the first pass
    had no connector arrows and used a gold "we are here" text label to the
    RIGHT of the box, neither of which is the house format.
    """
    def draw(slide):
        box_h = Inches(0.85)
        narrow_w = Inches(4.6)
        wide_w = Inches(8.6)
        gap = Inches(0.3)
        mid = SLIDE_W // 2

        top_x = mid - wide_w // 2
        top_y = Inches(2.0)
        _add_rounded_filled_box(
            slide, top_x, top_y, wide_w, box_h,
            "1. Basic Principles and Economic Way of Thinking",
            fill=FADED, text_color=WHITE, size=24, bold=True)

        row2_y = Inches(3.65)
        left_x = mid - gap // 2 - narrow_w
        right_x = mid + gap // 2
        _add_rounded_filled_box(slide, left_x, row2_y, narrow_w, box_h,
                                "2. Value and Demand", fill=FADED,
                                text_color=WHITE, size=26, bold=True)
        _add_rounded_filled_box(slide, right_x, row2_y, narrow_w, box_h,
                                "3. Supply and Cost", fill=FADED,
                                text_color=WHITE, size=26, bold=True)

        bot_x = mid - wide_w // 2
        bot_y = Inches(5.5)
        _add_rounded_filled_box(slide, bot_x, bot_y, wide_w, box_h,
                                "4. Markets, Pricing, and Strategy",
                                fill=NAVY, text_color=WHITE, size=24,
                                bold=True)

        # the diamond: 1 -> 2, 1 -> 3, 2 -> 4, 3 -> 4
        top_bottom_y = top_y + box_h
        _add_arrow(slide, (top_x + wide_w // 2, top_bottom_y),
                   (left_x + narrow_w // 2, row2_y),
                   color=FADED, weight_pt=3.0, head=True)
        _add_arrow(slide, (top_x + wide_w // 2, top_bottom_y),
                   (right_x + narrow_w // 2, row2_y),
                   color=FADED, weight_pt=3.0, head=True)
        row2_bottom_y = row2_y + box_h
        _add_arrow(slide, (left_x + narrow_w // 2, row2_bottom_y),
                   (bot_x + wide_w // 2, bot_y),
                   color=FADED, weight_pt=3.0, head=True)
        _add_arrow(slide, (right_x + narrow_w // 2, row2_bottom_y),
                   (bot_x + wide_w // 2, bot_y),
                   color=FADED, weight_pt=3.0, head=True)

        # the "we are here" marker: a gold block arrow, no text
        arrow_w, arrow_h = Inches(0.6), Inches(0.5)
        arrow_left = bot_x - arrow_w - Inches(0.12)
        arrow_top = bot_y + (box_h - arrow_h) // 2
        _add_arrow_shape(slide, arrow_left, arrow_top, arrow_w, arrow_h,
                         direction="right", fill=GOLD)
        # the gold italic label sits to the LEFT of the arrow, right-aligned
        # against it -- geometry copied from Module 4's make_m4_roadmap
        _add_text(slide, arrow_left - Inches(1.55),
                  bot_y + (box_h - Inches(0.32)) // 2,
                  Inches(1.45), Inches(0.32), "we are here",
                  size=16, italic=True, bold=True, color=GOLD,
                  font="Calibri", align=PP_ALIGN.RIGHT)

    return make_diagram_slide(prs, page_num, TAG_ROADMAP,
                              "Agenda for the Class", draw)


# ==========================================================================
#  1 · Simple vs. Complex Pricing — Video 1   [V1 = NV 4, 5, 6, 7, 8, 9, 10]
# ==========================================================================

# How small a region label may be squeezed before it is moved OUT of its
# region instead (2026-09-10, Nico: "If that text does not fit inside the
# shape (at reasonable text sizes), move it outside the shape with an
# arrow pointing into the shape").  11 pt is below the deck's 16 pt
# chart-label floor on purpose: a staircase block or a thin profit
# rectangle is a "narrow object" in the Teaching CLAUDE.md sense, its
# width fixed by the data, and the exception there says fit the text to
# the object rather than vaguen the label or misdraw the object.
REGION_FLOOR_PT = 11
# Clearance kept between the label box and the region's own edges.  Those
# edges ARE curves -- a demand line is drawn at 2.75 pt, which is 0.038"
# wide -- so the box has to stop short of them, not merely touch.
REGION_PAD_IN = 0.03


def _poly_clear(poly, x, y, pad):
    """Is (x, y) at least `pad` inches inside the convex polygon `poly`?

    `poly` is in slide inches, in either winding order: the shoelace sign
    settles which side is the interior, so no caller has to state it --
    and on a slide, where y grows DOWNWARD, guessing it is a coin toss.
    """
    n = len(poly)
    s = sum(poly[i][0] * poly[(i + 1) % n][1]
            - poly[(i + 1) % n][0] * poly[i][1] for i in range(n))
    sgn = 1.0 if s > 0 else -1.0
    for i in range(n):
        x0, y0 = poly[i]
        x1, y1 = poly[(i + 1) % n]
        ex, ey = x1 - x0, y1 - y0
        seg = (ex * ex + ey * ey) ** 0.5
        if seg < 1e-12:
            continue
        if sgn * (ex * (y - y0) - ey * (x - x0)) / seg < pad:
            return False
    return True


def _box_in_poly(poly, cx, cy, w, h, pad=REGION_PAD_IN):
    """Does the whole w x h box centred at (cx, cy) sit inside `poly`?

    Four corners only: every shaded region in this deck is convex, and a
    convex polygon that contains the corners contains the box.
    """
    return all(_poly_clear(poly, cx + sx * w / 2.0, cy + sy * h / 2.0, pad)
               for sx in (-1.0, 1.0) for sy in (-1.0, 1.0))


def _fit_in_poly(poly, ax, ay, w, h, steps=32):
    """The centre nearest (ax, ay) at which a w x h box fits in `poly`.

    The anchor is where the label WANTS to sit (its region's centroid);
    the scan is what keeps a wide label off a sloped edge.  Clamping to
    the region's x-range, which is what this replaced, could not do that
    -- a triangle is only as wide as its x-range down at its base, and
    the label sits further up, where the hypotenuse has closed in.
    """
    if _box_in_poly(poly, ax, ay, w, h):
        return ax, ay
    xs = [p[0] for p in poly]
    ys = [p[1] for p in poly]
    best, bd = None, 1e18
    for i in range(steps + 1):
        x = min(xs) + (max(xs) - min(xs)) * i / float(steps)
        for j in range(steps + 1):
            y = min(ys) + (max(ys) - min(ys)) * j / float(steps)
            d = (x - ax) ** 2 + (y - ay) ** 2
            if d < bd and _box_in_poly(poly, x, y, w, h):
                best, bd = (x, y), d
    return best


def _wrap_two(label):
    """Split a one-line label into two balanced lines, or None."""
    words = label.split()
    if len(words) < 2:
        return None
    best, bd = None, 1e18
    for k in range(1, len(words)):
        a, b = " ".join(words[:k]), " ".join(words[k:])
        if abs(len(a) - len(b)) < bd:
            best, bd = a + "\n" + b, abs(len(a) - len(b))
    return best


def _region_label(slide, fig, pts, label, *, size=17, color=NAVY,
                  bold=True, dy=0.0, dx=0.0, out=None,
                  floor=REGION_FLOOR_PT):
    """Label a shaded region -- INSIDE it, or outside it with an arrow in.

    Teaching CLAUDE.md requires a surplus label to sit inside its triangle
    at the centroid, computed rather than eyeballed; the same applies to
    every lettered/named region.  ``dy`` / ``dx`` nudge the anchor in
    LOGICAL units.

    2026-09-10 (Nico, on slide 9's "Consumer surplus"): text inside a
    shape has to sit FULLY inside it and be crossed by nothing, and where
    it cannot, it belongs outside with an arrow pointing in.  So the box
    is now tested for real CONTAINMENT -- all four corners inside the
    polygon, clear of its edges -- and three things are tried in order:

      1. the label on one line, shrinking toward ``floor``;
      2. the label on TWO lines, shrinking again (this is what fits
         "Version 3" inside a 0.70"-wide staircase block, where one line
         is wider than the block at any readable size);
      3. failing both, the label goes OUTSIDE at full size with an arrow
         from its edge back into the region.

    The earlier version clamped the box into the region's x-range, which
    is not containment: a triangle is only that wide at its base, so
    "Consumer surplus" still lay across the demand line that forms the
    hypotenuse.  Pass ``out=(x, y)`` in logical units to choose the
    outside position yourself.
    """
    xs = [x for x, _ in pts]
    ys = [y for _, y in pts]
    cx = sum(xs) / float(len(pts)) + dx
    cy = sum(ys) / float(len(ys)) + dy

    # a triangle's centroid sits a third of the way toward the
    # hypotenuse; bias back toward the right-angle corner so the curve
    # has room
    if len(pts) == 3:
        corner = None
        for i, (px, py) in enumerate(pts):
            a = pts[(i + 1) % 3]
            b = pts[(i + 2) % 3]
            if (abs(px - a[0]) < 1e-9 and abs(py - b[1]) < 1e-9) or \
               (abs(px - b[0]) < 1e-9 and abs(py - a[1]) < 1e-9):
                corner = (px, py)
                break
        if corner is not None:
            cx += (corner[0] - cx) * 0.22
            cy += (corner[1] - cy) * 0.22

    poly = [(fig.x(px) / 914400.0, fig.y(py) / 914400.0) for px, py in pts]
    ax, ay = fig.x(cx) / 914400.0, fig.y(cy) / 914400.0

    def _try(text):
        """Largest size at or below `size` at which `text` fits, and where."""
        lines = text.split("\n")
        longest = max(lines, key=len)
        for sz in range(int(size), int(floor) - 1, -1):
            w = _text_w_in(longest, sz, bold=bold) + 0.10
            h = sz * 1.35 * len(lines) / 72.0
            spot = _fit_in_poly(poly, ax, ay, w, h)
            if spot:
                return sz, w, h, spot, text
        return None

    got = _try(label)
    if got is None and "\n" not in label:
        two = _wrap_two(label)
        if two:
            got = _try(two)

    if got is not None:
        sz, w, h, (bx, by), text = got
        lab = _add_text(
            slide, int(Inches(bx - w / 2.0)), int(Inches(by - h / 2.0)),
            int(Inches(w)), int(Inches(h)), text, size=sz, bold=bold,
            color=color, font="Calibri", align=PP_ALIGN.CENTER)
        # stamped with its region's own key so the grouping pass pairs the
        # two exactly -- a CURVE label that merely sits inside a region is
        # not stamped, and so is never grouped with it (2026-09-10)
        lab.name = "sdreglab:%s" % _poly_key(pts)
        return lab

    # Nothing readable fits: the label goes outside, with an arrow back in.
    lines = label.split("\n")
    w = _text_w_in(max(lines, key=len), size, bold=bold) + 0.10
    h = size * 1.35 * len(lines) / 72.0
    if out is not None:
        lx = fig.x(out[0]) / 914400.0 - w / 2.0
        ly = fig.y(out[1]) / 914400.0 - h / 2.0
    else:
        # just beyond the region: above it first, then right, below, left
        px0, px1 = min(p[0] for p in poly), max(p[0] for p in poly)
        py0, py1 = min(p[1] for p in poly), max(p[1] for p in poly)
        g = 0.16
        cands = [(ax - w / 2.0, py0 - g - h), (px1 + g, ay - h / 2.0),
                 (ax - w / 2.0, py1 + g), (px0 - g - w, ay - h / 2.0)]
        lx, ly = cands[0]
        for c in cands:
            if (c[0] >= 0.15 and c[0] + w <= 13.18
                    and c[1] >= 1.45 and c[1] + h <= 7.05):
                lx, ly = c
                break
    print("     region label OUTSIDE: %r -> (%.2f, %.2f)"
          % (label.replace("\n", " "), lx, ly))
    box = _add_text(slide, int(Inches(lx)), int(Inches(ly)),
                    int(Inches(w)), int(Inches(h)), label, size=size,
                    bold=bold, color=color, font="Calibri",
                    align=PP_ALIGN.CENTER)
    # a label placed OUTSIDE its region is not grouped with it -- it is
    # grouped with its own leader instead, by the arrow rule
    _arrow_from_box(slide, (Inches(lx), Inches(ly)),
                    (Inches(w), Inches(h)), (fig.x(cx), fig.y(cy)),
                    color=color, weight_pt=1.25, head=True)
    return box


# The three shaded regions of the simple-pricing figure, as marks small
# enough to stand in front of a bullet.  Colours and washes are the
# regions' own; the SHAPES are too -- both triangles have their right
# angle at the bottom left, exactly as the regions on the slide do, which
# is why they are flipped.  Laid out in the arrangement the graph has, so
# all three together read as the whole triangle under demand.
#
#      +--------+           cs  (red)     top-left
#      | cs  /  |
#      +----+---+           rev (gold)    below cs
#      | rev| u |           unx (grey)    right of rev
#      +----+---+
_REGION_MARKS = {
    "cs":  (MSO_SHAPE.RIGHT_TRIANGLE, CS_RED, 26, 0.0, 0.0),
    "rev": (MSO_SHAPE.RECTANGLE,      GOLD,   24, 0.0, 0.5),
    "unx": (MSO_SHAPE.RIGHT_TRIANGLE, GRAY,   26, 0.5, 0.5),
    # the zoo panels (2026-09-14, Nico: "for the agenda, use the
    # shape-symbols").  Each copies its own region's shape, colour and wash,
    # so the line in the list and the area in the chart read as one thing.
    "zfee": (MSO_SHAPE.RIGHT_TRIANGLE, GOLD, 26, 0.0, 0.0),   # flat fee
    "zuse": (MSO_SHAPE.RECTANGLE,      NAVY, 16, 0.0, 0.5),   # usage fee
    "zrev": (MSO_SHAPE.RECTANGLE,      GOLD, 26, 0.0, 0.5),   # simple-price
    "zcost": (MSO_SHAPE.RECTANGLE,  DARKRED, 16, 0.0, 0.5),   # visit costs
    # 2026-09-15 (Nico): "the usage fee itself is the MC line.  The area
    # under it is the revenue from the usage fee, so always show both."
    # The fee therefore gets a LINE mark and the revenue a grey box.
    "zline": (MSO_SHAPE.RECTANGLE,     NAVY, 100, 0.0, 0.5),  # the MC line
}

# a row whose mark is the "⇒" glyph rather than a shape: it does not
# name a region, it says the line FOLLOWS from the one above it
# (2026-09-15, Nico, on slide 69's closing line)
_GLYPH_MARKS = {"rarr": "⇒"}

# marks drawn as a thin BAR rather than a filled area: they stand for a
# line on the chart, not a region, and a square would read as an area
_LINE_MARKS = {"zline"}


def _region_mark(slide, left, top, w, h, parts, row=0):
    """Draw the composite region mark for `parts` in a w x h block.

    A row naming ONE region is a special case (2026-09-14, Nico, on the
    revenue mark of slide 10: "I moved up the square symbol ... make also
    sure it's actually a square").  There is no second mark for it to tile
    with, so its `fx` / `fy` placement inside the block is meaningless -- it
    just left the mark sitting low against its own line of text.  A lone
    mark is therefore drawn SQUARE and vertically centred in the block,
    keeping the left edge so the mark column still lines up down the list.
    """
    lone = len(parts) == 1
    for key in parts:
        if key in _GLYPH_MARKS:
            box = _add_text(slide, int(left), int(top), int(w), int(h),
                            _GLYPH_MARKS[key], size=20, bold=True,
                            color=NAVY, font="Calibri",
                            align=PP_ALIGN.LEFT,
                            anchor=MSO_ANCHOR.MIDDLE)
            box.name = "sdbullmark:%d:%s" % (row, key)
            continue
        geom, colour, alpha, fx, fy = _REGION_MARKS[key]
        if key in _LINE_MARKS:
            bar_h = int(h * 0.14)
            mx, my = int(left), int(top + (h - bar_h) / 2.0)
            mw, mh = int(h * 0.62), bar_h
        elif lone:
            side = int(h * 0.5)
            mx, my = int(left), int(top + (h - side) / 2.0)
            mw = mh = side
        else:
            mx, my = int(left + w * fx), int(top + h * fy)
            mw, mh = int(w * 0.5), int(h * 0.5)
        shp = slide.shapes.add_shape(geom, mx, my, mw, mh)
        # NOT flipped.  DrawingML's `rtTriangle` preset is already
        # moveTo(0,h) -> lnTo(w,h) -> lnTo(0,0), i.e. the right angle at the
        # BOTTOM LEFT and the hypotenuse falling from top-left to
        # bottom-right -- exactly how both regions sit in the figure.  A
        # flipH mirrored them (2026-09-10, judged off a 0.24" mark in a
        # 1400 px render, which is far too small to read an orientation
        # from); it also broke row 3, because two mirrored hypotenuses
        # zigzag instead of lining up into one straight edge.
        shp.fill.solid()
        shp.fill.fore_color.rgb = colour
        if key in _LINE_MARKS:
            shp.line.fill.background()
        else:
            _set_fill_alpha(shp, alpha)
            shp.line.color.rgb = NAVY
            shp.line.width = Pt(0.75)
        shp.shadow.inherit = False
        shp.name = "sdbullmark:%d:%s" % (row, key)


def _region_bullets(slide, left, top, width, height, rows, *, size=28,
                    sub_size=None, mark_w=0.70, mark_h=0.50, gap=0.16,
                    indent=0.30):
    """Bullets whose MARK is a picture of the regions the line names.

    `rows` are ``(parts, text)`` or ``(parts, text, level)``; ``parts=None``
    keeps an ordinary square bullet and level 1 indents the row by `indent`
    and sets it in `sub_size`.  One text box per row, the way `_welfare_rows` does it -- the
    rows do not share a paragraph list because each carries its own drawn
    mark, and a mark has to be positioned against the line it belongs to.

    The block is centred vertically in `height`, with the rows' own
    wrapped heights measured in the real font so a two-line row pushes its
    neighbours apart instead of overlapping them.
    """
    sub_size = size if sub_size is None else sub_size
    rows = [(r if len(r) == 3 else (r[0], r[1], 0)) for r in rows]
    text_dx = mark_w + gap

    def sz(lvl):
        return sub_size if lvl else size

    def line_h_of(lvl):
        return sz(lvl) * 1.22 / 72.0

    def wrap(t, lvl):
        avail = width / 914400.0 - text_dx - (indent if lvl else 0.0)
        out, cur = [], ""
        for word in t.split():
            trial = (cur + " " + word).strip()
            if _text_w_in(trial, sz(lvl)) <= avail or not cur:
                cur = trial
            else:
                out.append(cur)
                cur = word
        out.append(cur)
        return len(out)

    heights = [wrap(t, lvl) * line_h_of(lvl) for _, t, lvl in rows]
    slack = height - sum(heights)
    pad = slack / (2.0 * len(rows))          # even air above, below, between
    y = top + pad
    for i, ((parts, text, lvl), h) in enumerate(zip(rows, heights)):
        line_h = line_h_of(lvl)
        dx = Inches(indent) if lvl else 0
        left_i = int(left + dx)
        if parts is None:
            # the deck's own square bullet, drawn rather than typed so it
            # lines up with the composite marks below it
            d = 0.15
            box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, int(left_i + Inches(0.02)),
                int(Inches(y + (line_h - d) / 2.0)), int(Inches(d)),
                int(Inches(d)))
            box.fill.solid()
            box.fill.fore_color.rgb = NAVY
            box.line.fill.background()
            box.shadow.inherit = False
            box.name = "sdbullmark:%d:dot" % i
        else:
            _region_mark(slide, left_i,
                         int(Inches(y + (line_h - mark_h) / 2.0)),
                         Inches(mark_w), Inches(mark_h), parts, row=i)
        _add_text(slide, int(left_i + Inches(text_dx)), int(Inches(y)),
                  int(width - dx - Inches(text_dx)), Inches(h + 0.06), text,
                  size=sz(lvl), color=NAVY, font="Calibri")
        y += h + 2 * pad


def s_dilemma(prs, page_num, *, tag=None):
    """[NV 6] · [V1 s4] — the dilemma of simple pricing.

    2026-09-13, Nico's hand-edits adopted, three of them:
      * the bullet column narrowed 12.78" -> 5.74" to clear the figure;
      * "high price" and "low price" set BOLD.  Run-level emphasis is
        INVISIBLE to `_diff_slides.py`, which matches shapes on their
        normalised text -- this one only showed up in a run-by-run read of
        the slide XML, where the paragraph had split into three runs;
      * his willingness-to-pay figure added on the right.
    """
    bullets = [
        ("Simple pricing: same price for all customers", 0),
        ("Limitations of simple pricing:", 0),
        ([("If you charge a ", {}),
          ("high price", {"bold": True}),
          (", you prevent customers with relatively low willingness to pay "
           "(marginal personal value) from purchasing your product", {})],
         1, {}),
        ([("If you charge a ", {}),
          ("low price", {"bold": True}),
          (", you are too “generous” to those with high willingness to pay",
           {})],
         1, {}),
    ]

    def extras(slide):
        # His figure at the position and size he set by hand.  The source
        # is 1376 x 768 px, so passing the WIDTH alone reproduces his
        # 3.821" height exactly from the native 1.792 ratio -- width is the
        # constraining dimension here (the picture-sizing rule).  Default
        # rounded corners + soft shadow: he pasted it flat.
        _add_media_image(slide, "NV_s07_wtp_segments.png",
                         left=Inches(6.098), top=Inches(2.620),
                         width=Inches(6.846))

    return content_slide(prs, page_num, tag or TAG_SIMPLE,
                         "Dilemma of Simple Pricing", bullets,
                         size=26, sub_size=24, line_spacing_pts=14,
                         bullets_width=Inches(5.74), extras=extras,
                         notes=NOTES_NV.get(6))


# ONE subscriber's demand for movies in a MONTH, used across Videos 1, 5
# and 6:  P = 4 - Q/3, so the first movie is worth $4 and the 12th is worth
# nothing.  MR = 4 - 2Q/3 (twice the slope), so MR = 0 at Q = 6.  With
# MC = 0 the simple-pricing optimum is 6 movies at $2, TR = $12 a month;
# the whole area under demand is (12 x 4)/2 = $24, which is what a flat fee
# for unlimited access can take -- and about what Netflix's top tier costs.
#
# 2026-09-13 (Nico), rescaled from P = 2 - Q/8 for realism: $1 a movie and
# $8 a month were not Netflix numbers.  4 x 12 is the smallest pair that
# keeps EVERY derived figure a whole dollar -- which needs both maxima even
# and their product divisible by 8 -- while landing in a believable range.
# None of the economics moves: the flat fee is PMAX*QMAX/2 and the
# simple-pricing revenue PMAX*QMAX/4, so "exactly twice as much" holds for
# any pair, and the asserts below keep it honest.
# His demand red, set by hand on 2026-09-14.  It is NOT the palette's
# DARKRED (C00000) -- slide 8's demand line, its "D" and the red formula
# box are all A2162A, a deeper crimson.  Flagged to him: the rest of the
# deck's demand curves are still C00000, and the colour rule says to sweep
# the deck rather than leave one panel odd.
NFX_RED = RGBColor(0xA2, 0x16, 0x2A)

# The MC = 0 line lies ON the x axis, so it is drawn a hair above it -- and
# in NAVY it was indistinguishable from the axis it sits on, which is the
# one thing it must not be.  He asked for blue (2026-09-14); this is the
# deck's concept blue, and slide 8 is a concept-introduction slide, which is
# where the palette allows it.  The "MC = 0" label takes the same blue, per
# the curve-and-label rule.
NFX_BLUE = RGBColor(0x00, 0x70, 0xC0)

# MR runs PAST its zero on slide 8 (he extended it by hand): marginal
# revenue really does go negative beyond Q*, and the extension is what
# makes that visible rather than asserted.
NFX_MR_EXTEND = 0.54          # extra Q units beyond MR = 0 at Q_SIMPLE

NFX_PMAX, NFX_QMAX = 4.0, 12.0
# Everything else is DERIVED, so a figure and its arithmetic cannot drift
# apart the way they would if the optimum were typed in by hand.
NFX_SLOPE_DEN = NFX_QMAX / NFX_PMAX                # P = PMAX - Q/3
NFX_Q_SIMPLE = NFX_QMAX / 2.0                      # MR = 0 at half the choke
NFX_P_SIMPLE = NFX_PMAX / 2.0
NFX_TR_SIMPLE = NFX_Q_SIMPLE * NFX_P_SIMPLE        # $12 a month
NFX_FLAT = NFX_PMAX * NFX_QMAX / 2.0               # $24 a month
NFX_CS_SIMPLE = NFX_Q_SIMPLE * NFX_P_SIMPLE / 2.0  # $6, and the same again
                                                   # in unexploited market
# Plot headroom.  The y factor is the one the figures already had (2.4/2).
# The x factor is LARGER than the old 18/16, because the x-axis title grew
# from "Movies" to "Movies per month": the title is centred on the arrow
# tip, so a long one reaches back over the tick at QMAX and the two
# overprinted ("12" under "Movies per month").  1.25 puts 0.68" of clear
# air between them on the widest of these figures (2026-09-13).
NFX_XMAX, NFX_YMAX = NFX_QMAX * 1.25, NFX_PMAX * 1.20         # 15.0 x 4.8
assert abs(NFX_FLAT - 2.0 * NFX_TR_SIMPLE) < 1e-9, "flat fee must double TR"
assert abs(NFX_TR_SIMPLE + 2 * NFX_CS_SIMPLE - NFX_FLAT) < 1e-9, "areas"
for _v in (NFX_SLOPE_DEN, NFX_Q_SIMPLE, NFX_P_SIMPLE, NFX_TR_SIMPLE,
           NFX_FLAT, NFX_CS_SIMPLE):
    assert float(_v).is_integer(), "every Netflix figure stays a whole dollar"

NFX_AXIS_Q = "Movies per month"      # not bare "Movies": the dollar figures
                                     # are monthly, and the flat fee is a
                                     # subscription price (2026-09-13)


def s_netflix_simple(prs, page_num, *, tag=None):
    """[NV 7] · [V1 s5] — simple pricing per movie, MC = 0."""
    def draw(slide):
        # His figure, moved up and shortened by hand on 2026-09-14: the
        # plot now sits at 0.97 / 6.19 with h 4.29, and the y headroom is
        # down from 20% to 5.2% of PMAX, which is what pulled the whole
        # picture up and made room for the Netflix screen on the right.
        fig = SimpleFig(0.97, 6.19, 6.1, 4.29,
                        xmax=NFX_XMAX, ymax=NFX_PMAX * 1.0515)
        _fig_axes(slide, fig, x_title=NFX_AXIS_Q, y_title="P",
                  label_size=18)
        # Demand solid, MR DASHED and in the SAME red -- they are the same
        # curve family (MR is derived from D), so they share a colour and
        # the dash is what tells them apart (2026-09-14, Nico).
        _fig_line(slide, fig, (0, NFX_PMAX), (NFX_QMAX, 0), color=NFX_RED,
                  weight_pt=2.75, curve="D")
        mr_end = NFX_Q_SIMPLE + NFX_MR_EXTEND
        _fig_line(slide, fig, (0, NFX_PMAX),
                  (mr_end, NFX_PMAX - 2.0 * NFX_PMAX / NFX_QMAX * mr_end),
                  color=NFX_RED, weight_pt=2.25, dash="dash", curve="MR")
        _fig_line(slide, fig, (0, 0.06), (NFX_QMAX + 1.2, 0.06),
                  color=NFX_BLUE, weight_pt=2.25, curve="MC")
        _fig_curve_label(slide, fig, 10.94, 0.598, "D",
                         color=NFX_RED, size=20, curve="D")
        # at the tail of the extended MR, BELOW the axis, where he put it
        _fig_curve_label(slide, fig, 6.71, -0.363, "MR",
                         color=NFX_RED, size=20, curve="MR")
        # the MC label takes the MC line's own colour (2026-09-14, Nico)
        _fig_curve_label(slide, fig, 12.61, 0.176, "MC = 0",
                         color=NFX_BLUE, size=17, bold=False, curve="MC")
        # the max-TR rectangle: P = $2 over Q = 6
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, int(fig.x(0)), int(fig.y(NFX_P_SIMPLE)),
            int(fig.x(NFX_Q_SIMPLE) - fig.x(0)),
            int(fig.y(0) - fig.y(NFX_P_SIMPLE)))
        rect.fill.solid()
        rect.fill.fore_color.rgb = GOLD
        _set_fill_alpha(rect, 24)
        rect.line.color.rgb = GOLD
        rect.line.width = Pt(1.25)
        rect.shadow.inherit = False
        # the LEFT half of the rectangle: MR crosses only its right half
        # (MR meets P = 1 at Q = 4), so nothing is covered here
        # 2026-09-15 (Nico): the box says what the amount IS, not just
        # that it is a maximum -- it is the most a single price can take
        # from this demand, which is the whole point of the slide.  (He
        # typed "simply pricing"; the deck says "simple pricing"
        # everywhere else.)  The label runs to four lines now, so it uses
        # the rectangle's full width rather than its left half.
        _add_text(slide, Inches(1.07), Inches(4.30), Inches(2.24),
                  Inches(1.40),
                  "Max. possible TR under simple pricing = $%g"
                  % NFX_TR_SIMPLE, size=16,
                  bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        _fig_guide(slide, fig, (NFX_Q_SIMPLE, NFX_P_SIMPLE))
        _fig_ylab(slide, fig, NFX_PMAX, "$%g" % NFX_PMAX)
        _fig_ylab(slide, fig, NFX_P_SIMPLE, "$%g" % NFX_P_SIMPLE)
        _fig_xlab(slide, fig, NFX_Q_SIMPLE, "%g" % NFX_Q_SIMPLE)
        _fig_xlab(slide, fig, NFX_QMAX, "%g" % NFX_QMAX)
        # The demand equation, native OMML, in a red box at the spot Nico
        # moved it to by hand (2026-09-13).  Absolute inches, not figure
        # coordinates: he placed it against the slide, not against the plot.
        # The box is emitted FIRST so the equation draws on top of it, and
        # it is sized to the RENDERED glyphs -- a stacked fraction stands
        # about 0.62" tall out of a 0.50" text box, so a box built to the
        # text box's own height would clip the denominator.
        # The BOX is the thing he positions by hand -- he dragged it to
        # (2.13, 2.26) on 2026-09-13 and the formula stayed behind, because
        # the two are separate shapes.  So the box is the anchor now and
        # the equation is placed FROM it; they cannot drift apart again.
        # (They are not grouped: an OMML-math shape wrapped in
        # mc:AlternateContent loses its a14 namespace when extracted for a
        # group, per the grouping exception in the Teaching CLAUDE.md.)
        # his Netflix screen, 1088 x 752 px placed at its native ratio
        _add_media_image(slide, "NV_s08_netflix_tv.png",
                         left=Inches(7.150), top=Inches(1.760),
                         width=Inches(4.127))
        # 2026-09-14: he moved the box down-right and shrank it, because at
        # (2.13, 2.26) x 2.70 the box sat ON the demand line.  The equation
        # keeps its 22 pt and is CENTRED in whatever box he gives it, so the
        # two can never drift apart again.
        box_x, box_y = Inches(2.38), Inches(2.39)
        box_w, box_h = Inches(2.50), Inches(0.85)
        eq_w, eq_h = Inches(2.300), Inches(0.50)
        eq_x = int(box_x + (box_w - eq_w) / 2)
        # the glyphs sit high in their text box (a stacked fraction renders
        # taller than the box), so the box's spare height is split 15:27 the
        # way it was at his previous size rather than centred
        eq_y = int(box_y + (box_h - eq_h) * 0.357)
        _add_outlined_box(
            slide, int(box_x), int(box_y), int(box_w), int(box_h), "",
            line=DARKRED, fill=WHITE, rounded=True, shadow=True,
            corner_pct=0.14, line_w=1.75)
        _add_math_equation(
            slide, int(eq_x), int(eq_y), eq_w,
            eq_h,
            _omml_run("P") + _omml_text(" = %g − " % NFX_PMAX)
            + _omml_frac(_omml_run("Q"), _omml_text("%g" % NFX_SLOPE_DEN)),
            size_pt=22)
        # moved out from under the Netflix screen to the lower right
        # (2026-09-14, his position and size)
        _add_convention_box(
            slide, Inches(8.98), Inches(5.46), Inches(4.17), Inches(1.14),
            prefix="Since MC = 0:  ",
            body="profit maximization implies MR = MC = 0, so Netflix "
                 "charges $%g per movie." % NFX_P_SIMPLE,
            size=18)
    return make_diagram_slide(
        prs, page_num, tag or TAG_SIMPLE,
        "What if Netflix Used Simple Pricing per Movie? (Assume MC = 0)",
        draw)


def s_netflix_lose(prs, page_num, *, tag=None):
    """[NV 8] · [V1 s6] — what simple pricing leaves on the table.

    A1 (adopted from PG1 9): the three regions are named INSIDE the plot
    with leader lines, instead of my three labels floating above it.
    """
    def draw(slide):
        fig = SimpleFig(1.05, 6.55, 5.6, 4.9,
                        xmax=NFX_XMAX, ymax=NFX_YMAX)
        _fig_axes(slide, fig, x_title=NFX_AXIS_Q, y_title="P", label_size=18)
        # Consumer surplus (above P = $2, left of Q = 6), the revenue
        # rectangle, and the unexploited market (right of Q = 6, under D).
        # Each region's sloped edge IS the demand line: the vertices are the
        # curve's own intercept (0, 4), its zero (12, 0) and the simple-
        # pricing point (6, 2), which lies on P = 4 - Q/3 by construction.
        # Consumer surplus keeps the deck's fixed CS wash (C0201B at 26%).
        _fig_poly(slide, fig,
                  [(0, NFX_PMAX), (0, NFX_P_SIMPLE),
                   (NFX_Q_SIMPLE, NFX_P_SIMPLE)],
                  fill=CS_RED, line=CS_RED, alpha=26000, line_w=1.0)
        _fig_poly(slide, fig,
                  [(0, NFX_P_SIMPLE), (0, 0), (NFX_Q_SIMPLE, 0),
                   (NFX_Q_SIMPLE, NFX_P_SIMPLE)],
                  fill=GOLD, line=GOLD, alpha=24000, line_w=1.0)
        _fig_poly(slide, fig,
                  [(NFX_Q_SIMPLE, NFX_P_SIMPLE), (NFX_Q_SIMPLE, 0),
                   (NFX_QMAX, 0)],
                  fill=GRAY, line=GRAY, alpha=26000, line_w=1.0)

        # Each label sits at its own region's computed centroid.  The
        # "Unexploited markets" triangle is thin, so its label is lifted
        # clear and an arrow carries the eye down into the region.
        cs_pts = [(0, NFX_PMAX), (0, NFX_P_SIMPLE),
                  (NFX_Q_SIMPLE, NFX_P_SIMPLE)]
        rev_pts = [(0, NFX_P_SIMPLE), (0, 0), (NFX_Q_SIMPLE, 0),
                   (NFX_Q_SIMPLE, NFX_P_SIMPLE)]
        unx_pts = [(NFX_Q_SIMPLE, NFX_P_SIMPLE), (NFX_Q_SIMPLE, 0),
                   (NFX_QMAX, 0)]
        # the demand line and its label come AFTER the shaded regions,
        # or the regions are painted over the label and it cannot be
        # clicked (2026-09-09)
        _fig_line(slide, fig, (0, NFX_PMAX), (NFX_QMAX, 0), color=NFX_RED,
                  weight_pt=2.75, curve="D")
        _fig_curve_label(slide, fig, NFX_QMAX - 1.05, 0.56, "D",
                         color=NFX_RED, size=20, curve="D",
                         clear_of=((0, NFX_PMAX), (NFX_QMAX, 0)))
        # He copied slide 8's demand equation and its red box onto this
        # slide too (2026-09-14), at the same spot, so the two slides read
        # as one picture developing.
        _add_outlined_box(
            slide, int(Inches(2.13)), int(Inches(2.26)),
            int(Inches(2.70)), Inches(0.92), "",
            line=NFX_RED, fill=WHITE, rounded=True, shadow=True,
            corner_pct=0.14, line_w=1.75)
        _add_math_equation(
            slide, int(Inches(2.33)), int(Inches(2.41)), Inches(2.300),
            Inches(0.5),
            _omml_run("P") + _omml_text(" = %g − " % NFX_PMAX)
            + _omml_frac(_omml_run("Q"), _omml_text("%g" % NFX_SLOPE_DEN)),
            size_pt=22)
        _region_label(slide, fig, cs_pts, "Consumer surplus")
        _region_label(slide, fig, rev_pts, "Revenues")
        # the demand line is this region's own hypotenuse, so the label
        # is biased down toward the right-angle corner and away from it
        # two lines, so the box is narrow enough to sit under the
        # hypotenuse (which is the demand curve) without being crossed
        _region_label(slide, fig, unx_pts, "Unexploited\nmarkets", dy=0.24)
        _fig_ylab(slide, fig, NFX_PMAX, "$%g" % NFX_PMAX)
        _fig_ylab(slide, fig, NFX_P_SIMPLE, "$%g" % NFX_P_SIMPLE)
        _fig_xlab(slide, fig, NFX_Q_SIMPLE, "%g" % NFX_Q_SIMPLE)
        _fig_xlab(slide, fig, NFX_QMAX, "%g" % NFX_QMAX)

        # Nico's box and his 28 pt (2026-09-13), and the bullet DOTS on
        # rows 2 and 3 replaced by marks made of the figure's own regions:
        # row 2 names the two wasted areas, row 3 names all three, which
        # together are the whole triangle under demand.  This is the
        # legend-mark rule -- a mark copies the area's colour AND shape,
        # and several areas keep the arrangement they have in the graph --
        # applied to a bullet instead of a legend row.
        _region_bullets(
            slide, Inches(6.31), 2.15, Inches(6.64), 3.40,
            [(None, "Simplifying assumption: MC = 0"),
             # the revenue rectangle on its own, in the place it
             # occupies in the figure (2026-09-14, Nico)
             (("rev",), "Revenues under Simple Pricing"),
             (("cs", "unx"),
              "“Wasted profits”: consumer surplus + unexploited markets"),
             (("cs", "rev", "unx"),
              "Can do better using complex pricing")],
            size=28)
    # 2026-09-14, he asked for a better heading.  The old one hedged
    # "lose" in scare quotes and never said WHAT is lost; the takeaway
    # rule wants the point itself.  The figure's argument is that one
    # price collects the gold rectangle and leaves the other two areas
    # on the table, so the title now says so.
    return make_diagram_slide(
        prs, page_num, tag or TAG_SIMPLE,
        "One Price Leaves Two Kinds of Money on the Table", draw)


def s_complex(prs, page_num, *, tag=None):
    """[NV 9] · [V1 s7] — requirements for complex pricing.

    A2 (adopted from PG1 10): two sub-bullets make the market-power
    condition explicit -- perfect competition means a price taker with no
    pricing strategy to set, market power means a price searcher that has
    pricing power.
    """
    bullets = [
        ("Requirements for complex pricing:", 0),
        ("Market power", 1),
        ("Downward-sloping demand (i.e., price searcher)", 2),
        # the "perfect competition -> price taker" sub-bullet was cut by
        # hand on 2026-09-13: the line above already says price searcher
        ("Market power → price searchers have pricing power", 2),
        ("Capacity to prevent resale and arbitrage", 1),
        ("Enables sellers to charge different prices for different "
         "volumes/packages, or to different customers", 1),
        ("Complex pricing is usually more profitable than simple pricing", 0),
    ]
    def extras(slide):
        # His figure, at the position and size he set by hand.  1024 x 1024
        # px placed square, so nothing is stretched; default rounded corners
        # and soft shadow, which he had not applied.
        _add_media_image(slide, "NV_s10_segments_3d.png",
                         left=Inches(7.705), top=Inches(1.720),
                         width=Inches(5.110))

    return content_slide(prs, page_num, tag or TAG_SIMPLE,
                         "Complex (Non-Uniform) Pricing", bullets,
                         size=24, sub_size=22, line_spacing_pts=11,
                         bullets_width=Inches(7.17), extras=extras,
                         notes=NOTES_NV.get(9))


def s_dumdums(prs, page_num, *, tag=None):
    """A3 [NEW – PG1 20] — the concrete case for the no-resale condition.

    Bloomberg: rogue sellers order Dum-Dums in bulk from Sam's Club and
    resell them on Amazon at a markup, undercutting Spangler Candy on price.
    Placed straight after the requirements slide, which is where the
    condition is stated.
    """
    def draw(slide):
        _add_media_image(slide, "PG1_s20_rId3.png", left=Inches(1.55),
                         top=Inches(1.60), width=Inches(6.6), rounded=False)
        _add_hierarchical_bullets(
            slide, Inches(8.55), Inches(2.15), Inches(4.4), Inches(3.6),
            [([("Complex pricing needs resale to be hard", {})], 0, {}),
             ([("Sellers bought in bulk at Sam’s Club and drop-shipped "
                "on Amazon", {})], 0, {}),
             ([("Listing costs nothing, so the arbitrage scales", {})], 0,
              {})],
            size=21, line_spacing_pts=17)
        # Dark blue rather than the gold takeaway, and white text on it
        # (2026-09-13, Nico): this is the CONDITION the module leans on,
        # not the punchline of a mini-case.
        _add_takeaway_bar(
            slide,
            "If buyers can resell, complex pricing is ineffective",
            top=Inches(6.30), fill=NAVY, text_color=WHITE, size=19,
            rounded=True, shadow=True)
    sl = make_diagram_slide(prs, page_num, tag or TAG_SIMPLE,
                            "Preventing Resale: The Dum-Dums Case", draw)
    _set_notes(sl, (
        "Spangler Candy is a family business, and it makes Dum-Dums. "
        "Bloomberg's story is the clipping on the slide: rogue sellers "
        "buy Dum-Dums at Sam's Club, list them on Amazon at a markup, "
        "and undercut Spangler's own price while doing it. The headline "
        "puts what that costs the company in the millions.\n"
        "Notice what those sellers are adding: nothing. They are not "
        "manufacturing, branding or improving anything. They are moving "
        "the same candy out of a cheap channel into an expensive one "
        "and keeping the spread. That is arbitrage, and it exists only "
        "because the identical product is available at two prices.\n"
        "Two things make it hard to stop. Listing on Amazon costs a "
        "seller essentially nothing, so one person who spots the gap "
        "can scale to a thousand listings with no fixed cost to "
        "recover, and the arbitrage does not stay small. And Bloomberg "
        "notes that the sellers are violating Amazon's policy, which "
        "is the part worth dwelling on: an explicit contractual ban did "
        "not stop it, because policing thousands of third-party "
        "listings is much harder than writing the rule.\n"
        "The general lesson is the one in the blue bar. A firm can have "
        "all the market power in the world, but if the customer who is "
        "offered the low price can turn round and sell to the customer "
        "who was offered the high one, the two prices collapse into "
        "one. Ask the class what Spangler could actually do about it: "
        "different pack sizes or different products for different "
        "channels, contracts with distributors, tighter allocation to "
        "the clubs, or simply a smaller gap between the two prices. "
        "Every one of those is a way of making resale hard, which is "
        "the condition this slide is here to establish.\n"
        "Source: Bloomberg."))
    return sl


def s_three_degrees(prs, page_num, *, tag=None):
    """[NV 10] · [V1 s8] — the three degrees, as cream concept callouts."""
    def draw(slide):
        # Heights hand-tightened on 2026-09-13: the boxes were 1.62" for
        # every row regardless of how much text was in it, which left the
        # shorter two looking hollow.  The TOPS are unchanged (1.70, 3.52,
        # 5.34, a fixed 1.82" pitch), so only the gaps between the cards
        # grew.
        rows = [
            ("First degree (direct)",
             "“Dream world”: the firm knows each customer’s willingness to "
             "pay (marginal personal value – MPV)"),
            ("Third degree (direct)",
             "The firm can identify specific groups of customers with "
             "different MPV (students, seniors, tourists)"),
            ("Second degree (“indirect”)",
             "The firm cannot identify customers’ MPV, so it uses different "
             "purchase options to infer willingness to pay: versioning, "
             "quantity discounts, coupons"),
        ]
        tops = (1.70, 3.52, 5.34)                  # fixed 1.82" pitch
        heights = (1.12, 1.01, 0.98)               # his, one per row
        # One colour per DEGREE (2026-09-14, Nico): dark red for first,
        # dark blue for third, the house cream kept for second.  The three
        # cards are no longer interchangeable -- the colour is what the
        # rest of the module refers back to.
        fills = (NFX_RED, NAVY, CREAM)
        texts = (WHITE, WHITE, NAVY)
        for (head, body), y, h, fill, tc in zip(rows, tops, heights,
                                                fills, texts):
            _add_convention_box(
                slide, Inches(0.85), Inches(y), Inches(11.65), Inches(h),
                prefix=head + ":  ", body=body, size=20,
                anchor=MSO_ANCHOR.MIDDLE, fill_rgb=fill, text_color=tc,
                border=NAVY if fill is CREAM else fill)
    return make_diagram_slide(prs, page_num, tag or TAG_SIMPLE,
                              "Three Degrees of Price Discrimination", draw)


# ==========================================================================
#  2 · First Degree: Perfect Price Discrimination — Video 2
#     [V2 = NV 11 (agenda), 12, 14, 15, 13 — the dystopian slide comes LAST]
# ==========================================================================

# The first-degree demand used on both chart slides: P = 10 − Q, so the
# choke price is $10 and demand hits zero at Q = 10.  MC = 3 on the second
# slide, so demand crosses MC at Q = 7 — computed from the two lines, never
# eyeballed, and every marked point below is derived from them.
FD_PMAX, FD_QMAX = 10.0, 10.0
FD_MC = 3.0
FD_QMC = FD_QMAX - FD_MC          # where P = 10 − Q meets P = 3  ->  Q = 7


def s_first_degree(prs, page_num, *, tag=None):
    """[NV 12] · [V2 s3] — conditions for first-degree price discrimination."""
    bullets = [
        ("Charge each customer a different price", 0),
        ("Also known as perfect price discrimination", 0),
        ("Conditions:", 0),
        ("Many consumers with different MPV for the firm’s product", 1),
        ("Firm knows each consumer’s MPV (i.e., willingness to pay)", 1),
        ("Firm can prevent resale", 1),
        # the "still a hypothetical scenario" bullet was cut by hand on
        # 2026-09-13 and re-asked as the question in the blue box below
    ]

    def extras(slide):
        _add_media_image(slide, "NV_s15_first_degree.png",
                         left=Inches(7.660), top=Inches(1.605),
                         width=Inches(5.165))
        # His question, bottom-left under the bullets, in the same navy
        # bar the resale condition uses on slide 11 -- a question the class
        # answers next, not a takeaway, so it is navy rather than gold.
        _add_takeaway_bar(
            slide,
            "Is first-degree price discrimination hypothetical or real?",
            left=Inches(0.51), width=Inches(6.47), top=Inches(6.25),
            height=Inches(0.70), fill=NAVY, text_color=WHITE, size=19,
            rounded=True, shadow=True)

    return content_slide(prs, page_num, tag or TAG_FIRST,
                         "First Degree Price Discrimination", bullets,
                         size=26, sub_size=24, line_spacing_pts=12,
                         bullets_width=Inches(6.47),
                         bullets_height=Inches(4.21), extras=extras)


def _arrow_from_box(slide, box_xy, box_wh, target, *, color=NAVY,
                    weight_pt=1.25, head=False, gap=0.06):
    """Draw a leader from the EDGE of a text box to ``target``.

    Teaching CLAUDE.md: an annotation arrow begins at its text box, and the
    text must not be crossed by any line.  Starting the arrow at the box
    CENTRE satisfies the first and breaks the second -- the line is then
    drawn straight through the words, which is what slides 16 and 17 did
    to every "...-MPV customer" callout.

    So: intersect the segment (centre -> target) with the box boundary and
    start there, plus ``gap`` inches of clear air.  The arrow simply comes
    out shorter, which is the right trade.
    """
    bx, by = box_xy
    bw, bh = box_wh
    cx, cy = bx + bw / 2.0, by + bh / 2.0
    tx, ty = target
    dx, dy = tx - cx, ty - cy
    span = (dx * dx + dy * dy) ** 0.5
    if span < 1.0:
        return None
    # scale to the boundary of the box, then step `gap` further out
    ts = []
    if abs(dx) > 1e-9:
        ts.append((bw / 2.0) / abs(dx))
    if abs(dy) > 1e-9:
        ts.append((bh / 2.0) / abs(dy))
    k = min(ts) if ts else 0.0
    k += Inches(gap) / span
    k = min(k, 0.98)
    sx, sy = cx + dx * k, cy + dy * k
    ln = _add_arrow(slide, (int(sx), int(sy)), (int(tx), int(ty)),
                    color=color, weight_pt=weight_pt, head=head)
    # Named so the grouping pass can tell a callout LEADER from a
    # curve: both are solid connectors, and a curve's endpoint lands
    # on the axis right where a tick label sits, so a geometric test
    # grouped demand lines with "$20" (2026-09-10).
    if ln is not None:
        ln.name = "sdleader:%d" % len(slide.shapes._spTree)
    return ln


def _fd_callout(slide, fig, xv, yv, text, *, to_xy, width=1.95, color=None):
    """One of the MPV callouts: a dot ON the demand curve, the label, and a
    leader running from the LABEL'S EDGE to the dot.

    The dot's coordinates are a point of the demand line, so the marker
    cannot drift off the curve; the leader is emitted after the label so
    its start can be computed from the label's real box.

    ``color`` paints the dot, the leader and the text in one tone, so a
    callout reads as one object (2026-09-14, Nico: dark red for the
    high-MPV customer, dark green for the low-MPV one, navy for the median).
    """
    color = NAVY if color is None else color
    dot_d = Inches(0.13)
    sx, sy = int(fig.x(to_xy[0])), int(fig.y(to_xy[1]))
    dot = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, int(sx - dot_d / 2), int(sy - dot_d / 2),
        int(dot_d), int(dot_d))
    dot.fill.solid()
    dot.fill.fore_color.rgb = color
    dot.line.fill.background()
    dot.shadow.inherit = False
    w, h = Inches(width), Inches(0.48)
    left = int(fig.x(xv) - w / 2)
    top = int(fig.y(yv) - h / 2)
    _add_text(slide, left, top, int(w), int(h), text, size=15, bold=False,
              color=color, font="Calibri", align=PP_ALIGN.CENTER)
    _arrow_from_box(slide, (left, top), (w, h), (sx, sy), color=color,
                    weight_pt=1.25, head=False)


def s_fd_mc0(prs, page_num, *, tag=None):
    """[NV 14] · [V2 s4] — perfect price discrimination with MC = 0."""
    def draw(slide):
        fig = SimpleFig(1.45, 6.55, 5.7, 4.85, xmax=11.8, ymax=11.8)
        _fig_axes(slide, fig, x_title="Q", y_title="P", label_size=18)
        # the whole area under demand is revenue; its sloped edge IS the
        # demand line, built from that line's own intercept and zero
        _fig_poly(slide, fig, [(0, FD_PMAX), (0, 0), (FD_QMAX, 0)],
                  fill=GOLD, line=None, alpha=24000)
        _fig_line(slide, fig, (0, FD_PMAX), (FD_QMAX, 0), color=DARKRED,
                  weight_pt=2.75, curve="D")
        # 2026-09-09 (Nico): a curve label belongs at the curve's RIGHT
        # END, or just above/below it -- for a demand curve, the lower
        # right of the plot.  2026-09-14: he moved it right, to (6.23, 6.12)
        # in slide inches, and asked why it was not grouped with its curve.
        # It was not, because the curve-and-label rule was written after
        # this slide was built and only slides 8 and 9 carried the key.
        # `curve="D"` on both is what the grouping pass pairs on.
        # His (6.23, 6.12) in slide inches, inverted through this figure's
        # own transform so the label stays tied to the PLOT rather than to
        # the canvas: xv = (6.23 - 1.45) * 11.8 / 5.7, and
        # yv = (6.55 - 0.16 - 6.12) * 11.8 / 4.85.
        _fig_curve_label(slide, fig, 9.896, 0.657, "Demand = MPV",
                         color=DARKRED, size=18, curve="D",
                         clear_of=((0, FD_PMAX), (FD_QMAX, 0)))
        _region_label(slide, fig, [(0, FD_PMAX), (0, 0), (FD_QMAX, 0)],
                      "Revenues and profits", size=17, dy=-0.5)
        _region_label(slide, fig, [(0, FD_PMAX), (0, 0), (FD_QMAX, 0)],
                      "(assume MC = 0)", size=15, bold=False, dy=-1.35)
        _fd_callout(slide, fig, 3.3, 10.9,
                    "Price charged to\nhigh-MPV customer", to_xy=(0.6, 9.4),
                    color=DARKRED)
        _fd_callout(slide, fig, 8.0, 7.1,
                    "Price charged to\nmedian-MPV customer", to_xy=(5.0, 5.0))
        _fd_callout(slide, fig, 9.9, 3.3,
                    "Price charged to\nlow-MPV customer", to_xy=(9.4, 0.6),
                    color=GREEN_DK)
        # his own side block, verbatim — missing from the first pass
        _add_hierarchical_bullets(
            slide, Inches(7.90), Inches(2.35), Inches(4.95), Inches(3.0),
            [([("Assume MC = 0", {})], 0, {}),
             # "each customer" in dark red (2026-09-14, Nico) -- it is the
             # phrase the whole first degree turns on
             ([("Charge ", {}),
               ("each customer", {'color': DARKRED, 'bold': True}),
               (" a different price", {})], 0, {}),
             ([("The Holy Grail of complex pricing", {})], 0, {}),
             ([("Zero consumer surplus: everything is extracted!", {})], 0,
              {})],
            size=22, line_spacing_pts=18)
    sl = make_diagram_slide(prs, page_num, tag or TAG_FIRST,
                            "First Degree Price Discrimination "
                            "(Assume MC = 0)", draw)
    _set_notes(sl, NOTES_NV.get(14, ""))
    return sl


def s_fd_mcpos(prs, page_num, *, tag=None):
    """[NV 15] · [V2 s5] — the same figure once MC > 0.

    The low-MPV customer is now NOT served: demand meets MC at Q = 7, which
    is computed from the two lines (FD_QMC) rather than typed in.
    """
    def draw(slide):
        fig = SimpleFig(1.45, 6.55, 5.7, 4.85, xmax=11.8, ymax=11.8)
        _fig_axes(slide, fig, x_title="Q", y_title="P", label_size=18)
        # profit is the area between demand and MC, out to their crossing
        _fig_poly(slide, fig,
                  [(0, FD_PMAX), (0, FD_MC), (FD_QMC, FD_MC)],
                  fill=DARKRED, line=None, alpha=22000)
        _fig_line(slide, fig, (0, FD_PMAX), (FD_QMAX, 0), color=DARKRED,
                  weight_pt=2.75, curve="D")
        _fig_line(slide, fig, (0, FD_MC), (FD_QMAX + 0.9, FD_MC), color=NAVY,
                  weight_pt=2.25, curve="MC")
        _fig_curve_label(slide, fig, FD_QMAX + 1.15, FD_MC + 0.15, "MC",
                         color=NAVY, size=18, curve="MC")
        # 2026-09-14 (Nico): "move the Demand label to the same spot as on
        # slide 20."  Both panels use the same SimpleFig, so the anchor is
        # literally the same pair of figure coordinates.
        _fig_curve_label(slide, fig, 9.896, 0.657, "Demand = MPV",
                         color=DARKRED, size=18, curve="D",
                         clear_of=((0, FD_PMAX), (FD_QMAX, 0)))
        _region_label(slide, fig,
                      [(0, FD_PMAX), (0, FD_MC), (FD_QMC, FD_MC)],
                      "Profits (if TFC = 0)", size=17, dy=-0.3)
        _fig_guide(slide, fig, (FD_QMC, FD_MC))
        # slide 20's colour scheme, carried over (2026-09-14, Nico): dark
        # red for the high-MPV customer, navy for the median, dark green for
        # the low-MPV one -- who on this panel is the customer NOT served
        _fd_callout(slide, fig, 3.3, 10.9,
                    "Price charged to\nhigh-MPV customer", to_xy=(0.6, 9.4),
                    color=DARKRED)
        _fd_callout(slide, fig, 8.3, 7.9,
                    "Price charged to\nmedian-MPV customer", to_xy=(4.4, 5.6))
        # above MC (3.0) and right of where demand has fallen below it,
        # which is the only clear air in that corner; the leader is longer
        # as a result, which is the trade the rule allows
        _fd_callout(slide, fig, 9.80, 4.35,
                    "Do not sell to the\nlow-MPV customer", to_xy=(9.2, 0.8),
                    color=GREEN_DK)
        # 2026-09-14: he brought the BMW configurator shot and the marque
        # across from his own slides, and moved the bullets up to clear it.
        _add_hierarchical_bullets(
            slide, Inches(7.87), Inches(2.04), Inches(4.95), Inches(2.0),
            [([("Now: MC > 0", {})], 0, {}),
             ([("Can it work in the real world?", {})], 0, {})],
            size=23, line_spacing_pts=18)
        _add_media_image(slide, "NV_s21_bmw_price.png",
                         left=Inches(8.42), top=Inches(3.41),
                         width=Inches(4.29))
        _add_media_image(slide, "NV_s21_bmw_logo.png",
                         left=Inches(10.95), top=Inches(6.31),
                         width=Inches(1.69), rounded=False, shadow=False)
    sl = make_diagram_slide(prs, page_num, tag or TAG_FIRST,
                            "First Degree Price Discrimination "
                            "(Now MC > 0)", draw)
    _set_notes(sl, NOTES_NV.get(15, ""))
    return sl


def s_dystopian(prs, page_num, *, tag=None):
    """[V2 s6] carrying [NV 13]'s figures, plus A4.

    The tape STATES the point ("By tracking consumer behavior, (tech)
    companies are increasingly able to use first degree price
    discrimination"); my 2024 main-deck slide ASKS it as a discussion
    question.  A video wants the statement, so the wording here is the
    tape's and the question version is kept for the in-class block.

    The figures are my own slide 13's: the Telegraph headline "Japanese
    vending machine tells you what you should drink" and the photo of the
    machine reading the customer's face.

    2026-09-09: PG1 15 ("Personalized Pricing in the Wild") was proposed as
    an adoption and then WITHDRAWN — her headline and photo are BYTE-
    IDENTICAL to my own slide 13's two images (sha1 6471d638 and 4f2ec7c5),
    so there was nothing to adopt.  A4, the WSJ electronic-shelf-labels
    headline, is the genuine addition.
    """
    def draw(slide):
        _add_media_image(slide, "NV_s13_g3.2_4f2ec7c5.png", left=Inches(0.70),
                         top=Inches(1.58), width=Inches(7.0), rounded=False,
                         shadow=False)
        _add_media_image(slide, "NV_s13_g3.1_6471d638.png", left=Inches(1.05),
                         top=Inches(2.45), width=Inches(6.3))
        _add_hierarchical_bullets(
            slide, Inches(8.15), Inches(1.75), Inches(4.65), Inches(2.2),
            [([("By tracking consumer behavior, (tech) companies are "
                "increasingly able to use first degree price discrimination "
                "(personally tailored prices)", {})], 0, {}),
             ([("Example: Uber", {})], 0, {})],
            size=19, line_spacing_pts=13)
        _add_media_image(slide, "PG1_s05_rId2.png", left=Inches(8.15),
                         top=Inches(4.35), width=Inches(4.65), rounded=False,
                         shadow=False)
    sl = make_diagram_slide(prs, page_num, tag or TAG_FIRST,
                            "The Dystopian Future of Price Discrimination",
                            draw)
    _set_notes(sl, NOTES_V2.get(6, "") or NOTES_NV.get(13, ""))
    return sl


# ==========================================================================
#  3 · Third Degree: Segment Pricing — Video 3
#     [V3 = NV 17 (agenda), 18, 19, 20, 21, 22, 23, 29, 30]
# ==========================================================================

def s_segment_concept(prs, page_num, *, tag=None):
    """[NV 18] · [V3 s3] — conditions for segment pricing."""
    def draw(slide):
        _add_hierarchical_bullets(
            slide, Inches(0.70), Inches(1.95), Inches(6.3), Inches(4.4),
            [([("Divide customers into different groups with different "
                "price sensitivities, and charge each group different "
                "prices", {})], 0, {}),
             ([("Conditions:", {})], 0, {}),
             ([("Consumer groups with (on average) different MPV", {})], 1,
              {}),
             ([("Firm can identify these groups", {})], 1, {}),
             ([("Firm has market power and can prevent resale", {})], 1, {})],
            size=23, sub_size=21, line_spacing_pts=13)
        # he swapped the picture on 2026-09-13 (not in the list he sent,
        # but a rebuild would have thrown it away)
        _add_media_image(slide, "NV_s21_groups.png", left=Inches(7.590),
                         top=Inches(1.703), width=Inches(5.137))
    return make_diagram_slide(
        prs, page_num, tag or TAG_SEGMENT,
        "Third Degree Price Discrimination: Segment Pricing "
        "(Group Pricing)", draw)


# The segment-pricing picture: D is P = 10 − Q.  Regular movie-goers pay
# P1 = 7 (so they take the first 3 units); seniors pay P2 = 4 (the next 3).
# Both blocks sit UNDER the demand line by construction.
SEG_PMAX, SEG_QMAX = 10.0, 10.0
SEG_P1, SEG_P2 = 7.0, 4.0
SEG_Q1 = SEG_QMAX - SEG_P1        # 3
SEG_Q2 = SEG_QMAX - SEG_P2        # 6


def s_segment_chart(prs, page_num, *, tag=None):
    """[NV 19] · [V3 s4] — the two price bands under one demand curve."""
    def draw(slide):
        fig = SimpleFig(1.45, 6.55, 5.5, 4.85, xmax=11.8, ymax=11.8)
        _fig_axes(slide, fig, x_title="Q", y_title="Price", label_size=18)
        _fig_poly(slide, fig,
                  [(0, SEG_P1), (0, 0), (SEG_Q1, 0), (SEG_Q1, SEG_P1)],
                  fill=GOLD, line=GOLD, alpha=26000, line_w=1.0)
        _fig_poly(slide, fig,
                  [(SEG_Q1, SEG_P2), (SEG_Q1, 0), (SEG_Q2, 0),
                   (SEG_Q2, SEG_P2)],
                  fill=NAVY, line=NAVY, alpha=18000, line_w=1.0)
        _fig_line(slide, fig, (0, SEG_PMAX), (SEG_QMAX, 0), color=DARKRED,
                  weight_pt=2.75)
        # a demand curve is labelled "D", never spelled out
        # (2026-09-10, Nico)
        _fig_curve_label(slide, fig, SEG_QMAX - 0.85, 1.25, "D",
                         color=DARKRED, size=18,
                         clear_of=((0, SEG_PMAX), (SEG_QMAX, 0)))
        _fig_ylab(slide, fig, SEG_P1, "P₁")
        _fig_ylab(slide, fig, SEG_P2, "P₂")
        _region_label(slide, fig,
                      [(0, SEG_P1), (0, 0), (SEG_Q1, 0), (SEG_Q1, SEG_P1)],
                      "Regular\nmovie-goers", size=15)
        _region_label(slide, fig,
                      [(SEG_Q1, SEG_P2), (SEG_Q1, 0), (SEG_Q2, 0),
                       (SEG_Q2, SEG_P2)], "Seniors", size=15)
        # 2026-09-14: he moved the definition up and left, over the top of
        # the figure, and put his own segments picture in the space it left
        _add_hierarchical_bullets(
            slide, Inches(3.78), Inches(1.76), Inches(5.1), Inches(2.6),
            [([("Divide customers into different groups with different "
                "price sensitivities", {})], 0, {}),
             # "each group" in dark red (2026-09-14, Nico) -- the phrase
             # the whole third degree turns on, as on slide 20
             ([("Charge ", {}),
               ("each group", {'color': DARKRED, 'bold': True}),
               (" different prices (lower prices to more elastic groups)",
                {})], 0, {})],
            size=21, line_spacing_pts=18)
        _add_media_image(slide, "NV_s27_groups.png", left=Inches(8.07),
                         top=Inches(3.71), width=Inches(4.76))
    sl = make_diagram_slide(prs, page_num, tag or TAG_SEGMENT, "Segment Pricing",
                            draw)
    _set_notes(sl, NOTES_NV.get(19, ""))
    return sl


def s_student_discounts(prs, page_num, *, tag=None):
    """[NV 20] · [V3 s5]."""
    def draw(slide):
        _add_media_image(slide, "NV_s20_2_e1dce5dd.png", left=Inches(4.65),
                         top=Inches(1.55), height=Inches(5.15))
    sl = make_diagram_slide(prs, page_num, tag or TAG_SEGMENT,
                            "Segment Pricing: Student Discounts", draw)
    _set_notes(sl, NOTES_NV.get(20, ""))
    return sl


def s_local_discounts(prs, page_num, *, tag=None):
    """[NV 21] · [V3 s6] — Disneyland and the Taj Mahal."""
    def draw(slide):
        # 2026-09-14 (Nico): "arrange pics on the left on top of each
        # other, as in my original slide 21."  His slide shows ONE wide
        # photo as two stacked halves, split by cropping -- the left half
        # (crop right 0.518, bottom 0.094) above the right half (crop left
        # 0.483) -- rather than the single wide strip this had.  The crops
        # are his; the positions are his 4:3 coordinates scaled x1.333
        # across and nudged down to clear the top bar.
        _add_media_image(slide, "NV_s21_4_1f3a388b.png", left=Inches(0.48),
                         top=Inches(1.52), width=Inches(5.79),
                         crop=(0.0, 0.518, 0.0, 0.094))
        _add_media_image(slide, "NV_s21_4_1f3a388b.png", left=Inches(0.41),
                         top=Inches(4.18), width=Inches(5.84),
                         crop=(0.483, 0.0, 0.0, 0.0))
        _add_media_image(slide, "NV_s21_3_c920a8ba.png", left=Inches(6.95),
                         top=Inches(1.60), height=Inches(5.40))
    sl = make_diagram_slide(prs, page_num, tag or TAG_SEGMENT,
                            "Segment Pricing: Discounts for Locals", draw)
    _set_notes(sl, NOTES_NV.get(21, ""))
    return sl


def s_timing(prs, page_num, *, tag=None):
    """[NV 22] · [V3 s7] — intertemporal pricing, EA Sports FC 26.

    2026-09-13, his replacement: the game's own cover beside its Amazon
    price history, instead of the old banner-plus-chart stack.  It is a
    better intertemporal example than a generic gadget -- the same SKU
    falls from $69.99 at launch to about $29.99, and the class can see the
    launch premium being harvested from the impatient buyers first.
    """
    def draw(slide):
        # Both at his positions and sizes.  562 x 603 and 931 x 517 px, so
        # passing the WIDTH reproduces his heights exactly from the native
        # ratios (0.932 and 1.801) -- neither picture is stretched.
        _add_media_image(slide, "NV_s25_eafc26_cover.png",
                         left=Inches(0.812), top=Inches(1.950),
                         width=Inches(3.903))
        _add_media_image(slide, "NV_s25_eafc26_pricehistory.png",
                         left=Inches(5.369), top=Inches(2.161),
                         width=Inches(7.308))
        # the source belongs to the chart, so it sits under the chart
        _add_text(slide, Inches(6.763), Inches(6.255), Inches(5.100),
                  Inches(0.202), "Source:  https://camelcamelcamel.com/",
                  size=12, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)
    sl = make_diagram_slide(prs, page_num, tag or TAG_SEGMENT,
                            "Segment Pricing: Timing", draw)
    _set_notes(sl, NOTES_NV.get(22, ""))
    return sl


def s_ways_to_segment(prs, page_num, *, tag=None):
    """[NV 23] · [V3 s8] — ways to segment consumers."""
    bullets = [
        ("By customer characteristics", 0),
        ("Age, gender, …", 1),
        ("But even by ‘customer’ species: doctors and veterinarians "
         "sometimes use the same medicines", 1),
        ("Discounts for new customers (gym memberships)", 1),
        ("By location", 0),
        ("Pharmaceuticals in the US v. emerging economies", 1),
        ("Intertemporal", 0),
        ("Electronic gadgets: early adopters vs. mass market", 1),
        ("Pricing based on search history", 1),
        ("Black Friday", 1),
    ]
    def extras(slide):
        # "New Member Discount.png", his file, replacing the gym promo he
        # had shared with slide 27 -- the two slides now carry their own
        # picture each (2026-09-13).  Sized by HEIGHT to match slide 27's
        # 5.439": this one is a 2:3 portrait, so matching its WIDTH instead
        # would have run it 0.68" past the footer rule.  Right edge aligned
        # with slide 27's, so the pair reads as one column.
        _add_media_image(slide, "New Member Discount.png",
                         left=Inches(8.924), top=Inches(1.600),
                         height=Inches(5.439))

    return content_slide(prs, page_num, tag or TAG_SEGMENT,
                         "Segment Pricing: Ways to Segment Consumers",
                         bullets, size=23, sub_size=21,
                         line_spacing_pts=8, bullets_width=Inches(7.27),
                         extras=extras, notes=NOTES_NV.get(23))


def s_recall_segment(prs, page_num, *, tag=None):
    """His slide, added by hand on 2026-09-13 and adopted here.

    A short recall of the ways-to-segment list, for use with the
    applications later in the course.  It sits INSIDE the Video 3 block, so
    it keeps that video's tag rather than the four-level In Class · Examples
    one -- the tag rule is about where a slide sits, not what it is for.
    If it moves to the in-class section the tag has to move with it.
    """
    bullets = [
        ("By customer characteristics", 0),
        ("Discounts for new customers (gym memberships)", 1),
        ("By location", 0),
        ("Intertemporal", 0),
    ]

    def draw(slide):
        _add_hierarchical_bullets(
            slide, Inches(0.61), Inches(2.22), Inches(7.65), Inches(3.06),
            [([(t, {})], lvl, {}) for t, lvl in bullets],
            size=23, sub_size=21, line_spacing_pts=8)
        _add_media_image(slide, "NV_s26_gym_newyear.png",
                         left=Inches(8.340), top=Inches(1.511),
                         width=Inches(4.210))
    return make_diagram_slide(
        prs, page_num, tag or TAG_SEGMENT,
        "Recall from Videos: Ways to Segment Consumers", draw)


def s_orbitz(prs, page_num, *, tag=None):
    """A6 [NEW – PG1 41] — Orbitz quoted Mac users pricier hotels.

    Placed straight after "ways to segment consumers", which is where my own
    deck already names "pricing based on search history" — this slide is
    that bullet's evidence.
    """
    def draw(slide):
        _add_media_image(slide, "PG1_s41_rId2.png", left=Inches(1.25),
                         top=Inches(1.60), width=Inches(7.3), rounded=False,
                         shadow=True)
        _add_hierarchical_bullets(
            slide, Inches(9.00), Inches(2.30), Inches(3.85), Inches(3.2),
            [([("Same city, same dates, same moment", {})], 0, {}),
             ([("Mac users were shown costlier hotels first", {})], 0, {}),
             ([("The segment is inferred from the device, never asked for",
                {})], 0, {})],
            size=20, line_spacing_pts=17)
    sl = make_diagram_slide(prs, page_num, tag or TAG_SEGMENT,
                            "Impediments to Segment Pricing: Orbitz", draw)
    _set_notes(sl, (
        "The two columns are the same hotel search, run at the same moment "
        "for the same dates, from a Mac and from a PC. Orbitz found that "
        "Mac users spent more per night on hotels, and started sorting "
        "costlier options to the top for them. Notice what the firm is "
        "doing here: it never asks which segment you belong to, it infers "
        "it from your device. That is third-degree price discrimination "
        "running on a signal you did not know you were sending, and it is "
        "also why this kind of segmentation draws a backlash once "
        "customers find out.\nSource: reported by the Wall Street "
        "Journal."))
    return sl


PV_BOX_XY = (Inches(6.92), Inches(6.83))
PV_BOX_WH = (Inches(5.85), Inches(0.58))
PV_BTN_WH = (Inches(0.46), Inches(0.28))
PV_INSET = Inches(0.15)


def _add_practice_video_box(slide, label=None):
    """The deck-standard practice-video link box.

    Rounded rect, grey face, gold border, soft shadow, navy label -- and
    the FILM symbol in a reserved left inset (2026-09-10, Nico: "follow
    the format that we used for other such boxes (i believe we had a film
    symbol?)").  It did: on 2026-08-23 the gold "▶" text glyph was
    replaced deck-family-wide by the action-button markers, which say what
    a link OPENS rather than which way it points -- `ACTION_BUTTON_MOVIE`
    for a video, `_SOUND` for a podcast, `_DOCUMENT` for an article.  This
    box was still on the old glyph.

    The button is seated the way `_add_jump_pill` seats its jump button:
    in a left inset, vertically centred, with the label's text frame
    indented past it, so the mark reads as part of the box.

    Default corner position, drawn AFTER the footer so it sits in front of
    the rule and the page number.
    """
    left, top = PV_BOX_XY
    w, h = PV_BOX_WH
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, int(left), int(top), int(w), int(h))
    try:
        box.adjustments[0] = 0.28
    except Exception:
        pass
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xF1, 0xF2, 0xF4)
    box.line.color.rgb = GOLD
    box.line.width = Pt(1.75)
    box.shadow.inherit = False
    _add_drop_shadow(box)
    btn_w, btn_h = PV_BTN_WH
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = int(PV_INSET + btn_w + Inches(0.11))
    tf.margin_right = Inches(0.07)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.LEFT
    r = para.add_run()
    r.text = label or "Practice Video:  Optimal Pricing in Two Markets"
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = NAVY
    r.font.name = "Calibri"
    _add_ext_link_button(slide, "movie", left=int(left + PV_INSET),
                         top=int(top + (h - btn_h) / 2),
                         width=int(btn_w), height=int(btn_h))
    return box


def s_two_markets(prs, page_num, *, tag=None):
    """[V3 s9] — segment pricing in two markets, pointing at the practice
    video.  The tape uses this LIGHT version; my heavier main-deck slide 29
    is kept for the in-class block."""
    def draw(slide):
        # narrowed from 8.4" so the text clears his BMW shot at 8.95"
        _add_hierarchical_bullets(
            slide, Inches(0.75), Inches(1.85), Inches(7.9), Inches(2.3),
            [([("A firm produces the same product for two markets", {})], 0,
              {}),
             ([("E.g. cars", {})], 1, {}),
             ([("Should the price be the same in both markets?", {})], 0, {}),
             ([("Not necessarily!", {})], 1, {})],
            size=25, sub_size=23, line_spacing_pts=13)
        # 2026-09-14 (Nico): "make the formula box have also a dark red
        # background.  And adopt my size of that box."  One shape now: the
        # equation helper carries the fill, so the box cannot drift away
        # from the formula it holds, and the glyphs go white on the red.
        _add_math_equation(
            slide, Inches(0.75), Inches(4.19), Inches(5.15), Inches(0.61),
            _omml_sub(_omml_run("MR"), _omml_run("i")) + _omml_text(" = ")
            + _omml_sub(_omml_run("MC"), _omml_run("i"))
            + _omml_text("   in each market ") + _omml_run("i"),
            size_pt=21, color=WHITE, fill=DARKRED, line=DARKRED,
            rounded=True, shadow=True, corner_pct=14000)
        # 2026-09-14 (Nico): "create a bullet point for the second point of
        # the rule, as that's the actual rule."  The first clause only says
        # THAT prices differ; the second says which way round, so it is set
        # on its own bulleted line rather than trailing after a dash.
        _add_convention_box(
            slide, Inches(0.75), Inches(5.25), Inches(7.9), Inches(1.36),
            runs=[("Rule:  ", {'bold': True}),
                  ("prices differ if demand (MR) and/or costs (MC) differ",
                   {}),
                  ("Charge the lower price in the more elastic market",
                   {'newline': True, 'bullet': True,
                    'bullet_char': "⇒"})],
            size=19, anchor=MSO_ANCHOR.MIDDLE)
        _add_media_image(slide, "NV_s34_bmw.png", left=Inches(8.95),
                         top=Inches(2.45), width=Inches(3.77))
    sl = make_diagram_slide(prs, page_num, tag or TAG_SEGMENT,
                            # his header, 2026-09-14
                            "Segment Pricing: Analysis for Two Markets",
                            draw)
    # AFTER the slide is built: the box straddles the footer rule at
    # y 7.15" by design, so it has to be the last shape appended or the
    # rule is drawn straight across its label (2026-09-10) -- the same
    # reason the Poll Break badge is called after _draw_footer
    _add_practice_video_box(sl)
    return sl


# --------------------------------------------------------------------------
#  The BMW two-market example, shared by main-deck slide 35 and the whole
#  practice-video deck.  Prices in $1,000, quantities in 1,000 cars, so
#  profit comes out in $ million.
#
#  2026-09-14 (Nico): "are those the numbers from the practice video?  We
#  need to make those more realistic."  Yes, they were his practice video's
#  own numbers (P = 160 − 4Q in the US, 140 − 5Q in Germany, MC = 40), but
#  they implied a $100,000 average BMW at a Lerner index of 0.60 — a 60 %
#  mark-up over variable cost, against the roughly 20 % gross margin a
#  premium carmaker actually earns, and demand elasticities of only −1.7 and
#  −1.8 where a single model faces something nearer −3 to −4.
#
#  These numbers keep the shape of the example (one market is clearly less
#  elastic and takes the higher price) at realistic magnitudes.  Nico then
#  asked, the same day, for the two prices to be reversed, so the demands
#  swapped and Germany is now the less elastic market:
#
#     US       P = 75 − 0.5Q   MR = 75 − Q   MR = MC = 45  ->  Q* = 30,
#                                                              P* = 60
#     Germany  P = 95 − 0.5Q   MR = 95 − Q   MR = MC = 45  ->  Q* = 50,
#                                                              P* = 70
#
#  which is a $60,000 car on 30,000 US units and a $70,000 car on 50,000
#  German units, at elasticities of −4.0 and −2.8 and Lerner indices
#  of 0.25 and 0.36.  The two demands share a slope and differ in
#  intercept, which keeps the arithmetic doable on camera and makes the
#  point that elasticity is a property of WHERE on the curve the firm ends
#  up, not of the slope alone.
#
#  CAVEAT, flagged to Nico 2026-09-14: the swap also makes Germany the
#  LARGER market, 50,000 cars against 30,000.  BMW in fact sells roughly
#  1.5 times as many cars in the US as at home, so if that matters
#  pedagogically the fix is to re-slope US demand (a flatter b), not to
#  swap the intercepts back.
#
#  Uniform pricing sells one price into the horizontal SUM of the two
#  demands.  Q(P) = (150 − 2P) + (190 − 2P) = 340 − 4P for P <= 75, and
#  only GERMANY buys above that, so the joint demand KINKS at P = 75,
#  Q = 40.  Below the kink P = 85 − 0.25Q, MR = 85 − 0.5Q, and MR = 45
#  gives Q = 80, P = 65.
#
#  Segment profit  = (70 − 45)·50 + (60 − 45)·30 = 1250 + 450 = 1700  (B + C)
#  Uniform profit  = (65 − 45)·80                              = 1600  (A)
#
#  So B + C > A, which is the claim the slide makes — verified, not asserted.
# --------------------------------------------------------------------------

BMW_MC = 45.0
# 2026-09-14 (Nico): "change the demand functions so that the price is
# reversed: 60k in the US and 70k in Germany."  The two demands simply
# swap -- the least disruptive route, and it keeps every number in the
# practice video's algebra.  It is also the right way round for the list
# prices: German ones carry 19 % VAT and run above US transaction prices.
# The block above carries the arithmetic and the market-size caveat.
BMW_US = (75.0, 0.5)              # P = a − bQ
BMW_DE = (95.0, 0.5)


def _bmw_opt(a, b, mc=BMW_MC):
    """Profit-maximising (Q*, P*) for P = a − bQ against constant mc."""
    q = (a - mc) / (2.0 * b)
    return q, a - b * q


BMW_QUS, BMW_PUS = _bmw_opt(*BMW_US)          # 15, 100
BMW_QDE, BMW_PDE = _bmw_opt(*BMW_DE)          # 10, 90
# joint (uniform) market, below the kink
# the horizontal sum, derived rather than typed: 1/b summed over markets
_BSUM = 1.0 / BMW_US[1] + 1.0 / BMW_DE[1]
_JA = (BMW_US[0] / BMW_US[1] + BMW_DE[0] / BMW_DE[1]) / _BSUM
_JB = 1.0 / _BSUM                             # P = 90 − 0.375Q
BMW_QJ, BMW_PJ = _bmw_opt(_JA, _JB)           # 60, 67.5
# the kink is where the SMALLER market drops out
BMW_KINK_P = min(BMW_US[0], BMW_DE[0])        # 75
BMW_KINK_Q = (_JA - BMW_KINK_P) / _JB         # 40
# ABOVE the kink only one market still buys -- the one with the higher
# choke price -- so the joint demand's upper segment IS that market's own
# demand curve.  Derived, never named: the two demands were swapped on
# 2026-09-14 and the call site still handed the panel BMW_US, which drew
# the upper segment FLAT from (0, 75) to (40, 75).  A horizontal top is
# not a demand curve at all, and it is not the horizontal sum of anything
# (2026-09-15, Nico: "the D curve on the right is not the aggregate
# demand curve.  You need to use horizontal aggregation.").
BMW_TOP = BMW_DE if BMW_DE[0] > BMW_US[0] else BMW_US

BMW_PROFIT_SEG = (BMW_PUS - BMW_MC) * BMW_QUS + (BMW_PDE - BMW_MC) * BMW_QDE
BMW_PROFIT_UNI = (BMW_PJ - BMW_MC) * BMW_QJ
assert BMW_PROFIT_SEG > BMW_PROFIT_UNI, "segment pricing must beat uniform"


def _bmw_panel(slide, left_in, width_in, *, a, b, qstar, pstar, dlabel,
               mrlabel, region_letter, heading, qmax, pmax, key="",
               top_in=2.40,
               height_in=3.15, kink=None, kink_a=None, kink_b=None,
               dend=None):
    """One market panel of the segment-vs-uniform figure: demand, its MR,
    the MC line, the profit rectangle, and the letter naming it."""
    fig = SimpleFig(left_in, top_in + height_in, width_in, height_in,
                    xmax=qmax, ymax=pmax)
    # `region_letter=None` drops the parenthetical (2026-09-15, Nico, on
    # the practice video: "Slide 11: don't use (A)" / "no idea where you
    # use B and C there").  The letters name AREAS being compared, which
    # is what this slide does and the practice video does not -- there
    # each panel stands alone, so the letter labels nothing.  The main
    # deck keeps passing B / C / A.
    if region_letter:
        heading = "%s  (%s)" % (heading, region_letter)
    _add_text(slide, int(Inches(left_in - 0.25)), int(Inches(top_in - 0.42)),
              int(Inches(width_in + 0.5)), Inches(0.34), heading, size=17,
              bold=True, color=NAVY, font="Calibri", align=PP_ALIGN.CENTER)
    _fig_axes(slide, fig, x_title="Q", y_title="P", label_size=14)
    # the profit rectangle: price down to MC, out to Q*
    _fig_poly(slide, fig, [(0, pstar), (0, BMW_MC), (qstar, BMW_MC),
                           (qstar, pstar)],
              fill=DARKRED, line=None, alpha=24000)
    if kink is None:
        _fig_line(slide, fig, (0, a), (a / b, 0), color=DARKRED,
                  weight_pt=2.25)
    else:
        # The joint demand is the HORIZONTAL SUM of the two market demands,
        # so it kinks where the smaller market drops out.  Above the kink
        # only the market with the higher choke price is still buying, and
        # the upper segment IS that market's demand curve -- `kink_a` and
        # `kink_b` are its intercept and slope, so the segment runs from
        # (0, kink_a) down to the kink at that market's own gradient.
        qk, pk = kink
        _fig_line(slide, fig, (0, kink_a), (qk, pk), color=DARKRED,
                  weight_pt=2.25)
        _fig_line(slide, fig, (qk, pk), (a / b, 0), color=DARKRED,
                  weight_pt=2.25)
    # MR: dashed, and the same dark red as the demand it comes from
    _fig_line(slide, fig, (0, a), (a / (2.0 * b), 0), color=DARKRED,
              weight_pt=1.5, dash="dash", curve="MR" + key)
    _fig_line(slide, fig, (0, BMW_MC), (qmax * 0.97, BMW_MC), color=NAVY,
              weight_pt=1.75)
    _fig_curve_label(slide, fig, qmax * 0.80, BMW_MC + pmax * 0.05, "MC",
                     color=NAVY, size=13)
    # 2026-09-10 (Nico): "MR and D sit on the wrong side of their curve.
    # Move them to the right, as by our guidelines."  Each label now sits
    # just PAST its own line's right end -- demand reaches the Q axis at
    # a/b, its MR at a/2b -- and a shade above the axis.  Beyond the
    # intercept there is no line at all, so nothing can cross the words,
    # and the two cannot collide with each other because the intercepts
    # are a/2b apart.  `dend` overrides where the demand intercept sits
    # too near the panel's own right edge to put a label past it.
    _fig_curve_label(slide, fig, a / (2.0 * b) + qmax * 0.014,
                     pmax * 0.07, mrlabel, color=DARKRED, size=13,
                     curve="MR" + key)
    if dend is None:
        _fig_curve_label(slide, fig, a / b + qmax * 0.014, pmax * 0.07,
                         dlabel, color=DARKRED, size=13)
    else:
        _fig_curve_label(slide, fig, dend[0], dend[1], dlabel,
                         color=DARKRED, size=13,
                         clear_of=((0, a), (a / b, 0)))
    _fig_guide(slide, fig, (qstar, pstar))
    # 2026-09-14: the letter moved OUT of the profit rectangle and into the
    # panel heading.  With the realistic numbers the German rectangle is
    # only 15 price units tall and 30 quantity units wide against a shared
    # 360-unit axis, and MR cuts it diagonally -- measured, there is no
    # position inside it that a readable letter can occupy without a line
    # through it (_check_labels.py, ONLINE).  Naming the area in the heading
    # keeps the B + C > A story legible and costs nothing.
    return fig


def s_bmw_graphical(prs, page_num, *, tag=None):
    """A7 [NEW – PG1 40] — segment vs. uniform pricing, three panels.

    Adopted with Nico's approval as the ONE piece of the BMW block that
    belongs in the main deck: it shows the RESULT of the practice video
    without repeating its algebra.  Also slide 9 of the practice-video deck.
    """
    def draw(slide):
        # 2026-09-15 (Nico): "move the comparison of profits from top to
        # bottom of the slide and make it a dark blue box."  It used to
        # sit above the three panels, which told the class the answer
        # before the first curve was drawn; it is now the closing bar.
        # the two segmented markets on the left, the joint market on the
        # right, separated by a thin rule so the comparison reads as 2 v 1
        # ONE quantity scale across all three panels, so the rectangles
        # really are comparable by eye and the "B + C > A" claim the title
        # makes is something the reader can see rather than take on trust.
        QM, PM = 360.0, 105.0
        _bmw_panel(slide, 1.05, 3.05, a=BMW_US[0], b=BMW_US[1],
                   qstar=BMW_QUS, pstar=BMW_PUS, dlabel="D", mrlabel="MR",
                   key=":us",
                   region_letter="B", heading="US market",
                   qmax=QM, pmax=PM)
        _bmw_panel(slide, 4.85, 3.05, a=BMW_DE[0], b=BMW_DE[1],
                   qstar=BMW_QDE, pstar=BMW_PDE, dlabel="D", mrlabel="MR",
                   key=":de",
                   region_letter="C", heading="German market",
                   qmax=QM, pmax=PM)
        _add_rect(slide, int(Inches(8.55)), int(Inches(1.82)), 12700,
                  int(Inches(4.30)), RULE)
        fig_j = _bmw_panel(slide, 9.35, 3.05, a=_JA, b=_JB, qstar=BMW_QJ,
                   pstar=BMW_PJ, dlabel="D", mrlabel="MR", region_letter="A",
                   key=":joint",
                   heading="Joint market (US + Germany)",
                   qmax=QM, pmax=PM,
                   kink=(BMW_KINK_Q, BMW_KINK_P), kink_a=BMW_TOP[0],
                   kink_b=BMW_TOP[1],
                   # the joint demand reaches the axis at Q = 68 of this
                   # panel's 72, so there is no room to put its label PAST
                   # the intercept: it goes immediately ABOVE the right
                   # end instead, in the clear band between demand (P = 9
                   # there) and the MC line at 40
                   dend=(_JA / _JB - 4.0, 22.0))
        # 2026-09-15 (Nico): "for the kink, add a small textbox with arrow
        # that says 'kink results from demand aggregation (beyond the
        # scope of this class)'".  Set at 14 pt rather than the 16 pt
        # chart-label floor: three panels share this slide, so the panel is
        # 3.05" wide and the box has to live in the clear air above the
        # demand line -- the narrow-object exception in the Teaching
        # CLAUDE.md.  Gray italic, because it is an aside rather than
        # something the class is asked to learn.
        _kb = (Inches(10.05), Inches(2.45))
        _kw = (Inches(2.98), Inches(0.60))
        _add_text(slide, _kb[0], _kb[1], _kw[0], _kw[1],
                  "Kink results from demand aggregation "
                  "(beyond the scope of this class)",
                  size=14, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)
        _arrow_from_box(slide, _kb, _kw,
                        (fig_j.x(BMW_KINK_Q), fig_j.y(BMW_KINK_P)),
                        color=GRAY, head=True)
        # 2026-09-14 (Nico): "just mention the different profits on slide
        # 35 and say that those are from the practice video."  The algebra
        # that produces them stays in the practice video.
        #
        # 2026-09-15 (Nico): "mention the profit for each market, and then
        # the total profit.  On the right, also list the total profit for
        # the joint market."  So each panel now carries its OWN rectangle
        # in numbers, directly under it, and the two totals sit on one
        # line so the comparison the slide is making is a single
        # left-to-right read.  Every figure is derived from the BMW
        # constants.
        for left_in, pr, qty in ((0.80, BMW_PUS, BMW_QUS),
                                 (4.60, BMW_PDE, BMW_QDE),
                                 (9.10, BMW_PJ, BMW_QJ)):
            _add_text(slide, Inches(left_in), Inches(5.95), Inches(3.55),
                      Inches(0.34),
                      # 2026-09-15 (Nico): "show only the profit.  Not
                      # the computation."  The algebra is the practice
                      # video's job; this slide compares three numbers.
                      "Profit = $%s m"
                      % format(int(round((pr - BMW_MC) * qty)), ","),
                      size=16, bold=True, color=NAVY, font="Calibri",
                      align=PP_ALIGN.CENTER)
        # The comparison and the two totals are ONE bar: three stacked
        # rows of numbers under the panels would not fit above the
        # practice-video box, and the comparison IS the two totals set
        # against each other.  Navy fill, white text, the treatment his
        # slide 57 bar already uses.
        _add_rounded_filled_box(
            slide, Inches(0.40), Inches(6.32), Inches(12.53), Inches(0.46),
            "Segment pricing:  $%s m   >   uniform pricing:  $%s m"
            % (format(int(round(BMW_PROFIT_SEG)), ","),
               format(int(round(BMW_PROFIT_UNI)), ",")),
            fill=NAVY, text_color=WHITE, size=19, bold=True,
            corner_pct=0.16)
        # the "worked out in the practice video" line is gone: the
        # practice-video box sits right beside where it was and says the
        # same thing, and the row it occupied is now the profit totals'
    sl = make_diagram_slide(prs, page_num, tag or TAG_SEGMENT,
                            "Segment vs. Uniform Pricing: The Graphical "
                            "Answer", draw)
    # after the slide is BUILT, so it sits in front of the footer rule
    _add_practice_video_box(sl)
    # Every figure here is DERIVED (2026-09-14).  The narration had gone
    # stale twice already -- once when the numbers were made realistic and
    # again when the two markets were swapped -- and a speaker note that
    # contradicts the figure beside it is worse than no note.
    _set_notes(sl, (
        "Here is why splitting the market pays. On the left the firm "
        "treats the US and Germany as two markets and sets marginal "
        "revenue equal to marginal cost in each one, so it sells "
        "%g thousand cars at $%g thousand in the US and %g thousand at "
        "$%g thousand in Germany. The two profit rectangles are B and C. "
        "On the right it charges one price to both countries together; "
        "the best it can do is $%g thousand on %g thousand cars, which is "
        "rectangle A. B plus C comes to $%.2f billion, A comes to "
        "$%.2f billion. The gap is about %.0f percent, and it is always in "
        "the same direction: charging one price to two different demand "
        "curves can never beat charging each of them its own price."
        % (BMW_QUS, BMW_PUS, BMW_QDE, BMW_PDE, round(BMW_PJ, 2), BMW_QJ,
           BMW_PROFIT_SEG / 1000.0, BMW_PROFIT_UNI / 1000.0,
           (BMW_PROFIT_SEG - BMW_PROFIT_UNI) / BMW_PROFIT_UNI * 100)))
    return sl


def s_bmw_challenges(prs, page_num, *, tag=None):
    """[NV 30] · [V3 s10] — arbitrage, backlash, and (A8) the legal limit.

    A8 [NEW – PG1 43] added the Robinson-Patman Act and the closing line
    that versioning is how firms get round these problems, which hands off
    to Video 4.  2026-09-14 (Nico): "delete Legal Impediment..." -- so the
    Robinson-Patman half of A8 is WITHDRAWN and only the closing line
    stays.  US antitrust has in any case stopped enforcing Robinson-Patman
    in practice, so it was the weakest of the three obstacles on the slide.
    """
    # 2026-09-14 (Nico): "adopt exactly the wording from my original slide
    # 30" -- "Germany vs. U.S." without the article, "miles vs kilometers"
    # without the stop.  The two PG1-43 lines below are KEPT, since they
    # were adopted deliberately (A8) and the ask was about wording.
    bullets = [
        ("Suppose BMW sets very different prices in Germany vs. U.S.", 0),
        ("Do you anticipate any problems?", 0),
        ("Arbitrage: re-importation from the cheaper market", 1),
        ("Difficult in this particular case (e.g., miles vs kilometers; "
         "taxation)", 2),
        ("However, easier in other cases (e.g., buying textbooks in India "
         "and reselling them in the US)", 2),
        ("Consumer backlash", 1),
        ("Way around these problems: versioning — second degree price "
         "discrimination", 0),
    ]

    def extras(slide):
        # his own illustration of the arbitrage, added by hand
        # 2026-09-15: Germany at the high price, the US at the low one,
        # and the barriers in between (km/h, DEKRA, duty)
        _add_media_image(slide, "NV_s41_bmw_arbitrage.png",
                         left=Inches(8.00), top=Inches(1.82),
                         width=Inches(4.72))

    return content_slide(prs, page_num, tag or TAG_SEGMENT,
                         "Challenges with Segment Pricing for BMW", bullets,
                         size=23, sub_size=21, line_spacing_pts=9,
                         bullets_width=Inches(6.98), extras=extras,
                         notes=NOTES_NV.get(30))


# ==========================================================================
#  4 · Second Degree: Versioning and Coupons — Video 4
#     [V4 = NV 31 (agenda), 32, 33, 34, 59 (Costco), 37, 42]
# ==========================================================================

def s_versioning_concept(prs, page_num, *, tag=None):
    """[NV 32] · [V4 s3] — conditions for versioning."""
    bullets = [
        ("Indirect price discrimination", 0),
        ("Charge a different price for different versions of the product", 0),
        ("Conditions:", 0),
        ("Firm has market power", 1),
        # his wording, 2026-09-13: "airplane cabins" is the versioning
        # example the next slides actually use, and "directly" is the word
        # that separates second degree from third
        ("Consumer groups with different willingness to pay for the "
         "product category (e.g., airplane cabins)", 1),
        ("Firm cannot directly identify the different groups of consumers",
         1),
    ]

    def extras(slide):
        _add_media_image(slide, "NV_s33_versioning.png",
                         left=Inches(7.475), top=Inches(1.542),
                         width=Inches(5.350))

    return content_slide(prs, page_num, tag or TAG_VERSION,
                         "Second Degree Price Discrimination: Versioning",
                         bullets, size=25, sub_size=23, line_spacing_pts=13,
                         bullets_width=Inches(6.20), extras=extras)


def s_versioning_chart(prs, page_num, *, tag=None):
    """[NV 33] · [V4 s4] — the version staircase under one demand curve.

    Five versions, each a block under P = 10 − Q: version k is priced at
    P_k and takes the units between Q_{k-1} and Q_k, so every block sits on
    the demand line by construction rather than by eye.
    """
    def draw(slide):
        fig = SimpleFig(1.45, 6.55, 5.5, 4.85, xmax=11.8, ymax=11.8)
        _fig_axes(slide, fig, x_title="Q", y_title="P / MPV",
                  label_size=18)
        prices = [8.5, 7.0, 5.5, 4.0, 2.5]
        q_prev = 0.0
        for k, p in enumerate(prices, start=1):
            q = SEG_QMAX - p                     # the point ON the demand line
            _fig_poly(slide, fig,
                      [(q_prev, p), (q_prev, 0), (q, 0), (q, p)],
                      fill=GOLD if k % 2 else NAVY,
                      line=GOLD if k % 2 else NAVY,
                      alpha=24000 if k % 2 else 16000, line_w=1.0)
            _region_label(slide, fig,
                          [(q_prev, p), (q_prev, 0), (q, 0), (q, p)],
                          "Version %d" % k, size=13,
                          dy=(p / 2.0) - 0.55 if p > 3 else 0.0)
            _fig_ylab(slide, fig, p, "P%s" % "₁₂₃₄₅"[k - 1])
            q_prev = q
        _fig_line(slide, fig, (0, SEG_PMAX), (SEG_QMAX, 0), color=DARKRED,
                  weight_pt=2.75)
        _fig_curve_label(slide, fig, SEG_QMAX - 0.85, 1.20, "D",
                         color=DARKRED, size=17,
                         clear_of=((0, SEG_PMAX), (SEG_QMAX, 0)))
        # his own image, added by hand on 2026-09-14; the text drops below
        # it rather than sharing the column
        _add_media_image(slide, "NV_s48_new.png", left=Inches(6.71),
                         top=Inches(1.39), width=Inches(5.56))
        _add_hierarchical_bullets(
            slide, Inches(7.00), Inches(4.25), Inches(5.9), Inches(2.2),
            [([("Create different “versions” of your product", {})], 0, {}),
             ([("Design them so that customers with different MPV "
                "“self-select” into the version you intend for them", {})],
              0, {})],
            size=21, line_spacing_pts=18)
    sl = make_diagram_slide(prs, page_num, tag or TAG_VERSION, "Versioning", draw)
    _set_notes(sl, NOTES_NV.get(33, "") or NOTES_NV.get(34, ""))
    return sl


def s_product_line(prs, page_num, *, tag=None):
    """[NV 34] · [V4 s5] — offer a product line and let users sort."""
    def draw(slide):
        _add_text(slide, Inches(0.70), Inches(1.60), Inches(11.9),
                  Inches(0.5),
                  "Offer a product line and let users sort themselves",
                  size=26, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.LEFT)
        # 2026-09-14 (Nico): "I replaced the windows image."  His own shot
        # of the four Windows 11 editions, at the size and place he set.
        # 1920 x 544 -> 2.23" tall at 7.88".  The two photos are 4:3, so
        # 2.25" tall at 3.0".  Everything ends above the footer rule.
        _add_media_image(slide, "NV_s46_windows.jpg", left=Inches(0.28),
                         top=Inches(3.06), width=Inches(7.88))
        _add_media_image(slide, "NV_s34_5_0ee3b24c.jpg", left=Inches(8.75),
                         top=Inches(2.35), width=Inches(3.0))
        _add_media_image(slide, "NV_s34_6_c3db5332.jpg", left=Inches(8.75),
                         top=Inches(4.80), width=Inches(3.0))
    sl = make_diagram_slide(prs, page_num, tag or TAG_VERSION, "Versioning", draw)
    _set_notes(sl, NOTES_NV.get(34, ""))
    return sl


def s_costco_versioning(prs, page_num, *, tag=None, title=None):
    """[NV 59] · [V4 s6] — "Versioning: Costco".

    2026-09-09 (Nico): Costco appears TWICE.  The tape has it here, inside
    Video 4, titled "Versioning: Costco"; my 2024 main deck retitled it
    "Costco: Combination of Versioning and Two-Part Pricing" and moved it
    after the two-part tariff.  Nico chose "both", so this is the Video 4
    copy with the tape's title and the in-class block carries the 2024 one.

    A10 [NEW – PG2 7]: the picture placeholder on this slide is EMPTY in
    both of my decks; PG's Costco membership-tier image fills it exactly.
    """
    def draw(slide):
        # a tall image -- size it by HEIGHT so it cannot run past the
        # footer rule, and let the width follow
        _add_media_image(slide, "PG2_s07_rId3.png", left=Inches(9.05),
                         top=Inches(1.70), height=Inches(4.85),
                         rounded=False)
        _add_hierarchical_bullets(
            slide, Inches(0.70), Inches(2.20), Inches(7.4), Inches(3.4),
            [([("Business model", {})], 0, {}),
             ([("Make money on the memberships", {})], 1, {}),
             ([("Sell goods at the cheapest price possible", {})], 1, {}),
             ([("Two membership versions, and the customer picks", {})], 0,
              {}),
             ([("WSJ video on the strategic layout of Costco stores", {})],
              0, {})],
            size=23, sub_size=21, line_spacing_pts=13)
    return make_diagram_slide(prs, page_num, tag or TAG_VERSION,
                              title or "Versioning: Costco", draw)


def s_more_versioning(prs, page_num, *, tag=None):
    """[NV 37] · [V4 s7] — more examples, and incentive compatibility."""
    bullets = [
        ("Internet service: higher speed at a higher price", 0),
        ("FedEx/UPS: ground vs. same-day delivery", 0),
        ("Airlines: business / premium economy / economy", 0),
        ("Important note on versioning: keep the low-quality option "
         "inferior enough to prevent premium customers from being tempted",
         0),
        ("This ensures “incentive-compatibility”", 1),
        ("So you don’t need to worry about resales", 1),
    ]
    def extras(slide):
        # his own images, added by hand on 2026-09-14
        _add_media_image(slide, "NV_s52_a.png", left=Inches(7.56),
                         top=Inches(1.42), width=Inches(5.26))
        _add_media_image(slide, "NV_s52_b.png", left=Inches(9.28),
                         top=Inches(3.30), width=Inches(3.38))

    return content_slide(prs, page_num, tag or TAG_VERSION,
                         "More Examples for Versioning", bullets,
                         bullets_width=Inches(7.18), extras=extras,
                         size=25, sub_size=23, line_spacing_pts=13)


def s_coupons(prs, page_num, *, tag=None):
    """[NV 42] · [V4 s8] — coupons as second-degree price discrimination."""
    def draw(slide):
        # 292 x 166 -> 2.50" tall at 4.4"; the two small coupons are
        # 177 x 130 and 160 x 121, so ~1.8" tall at 2.4".
        _add_media_image(slide, "NV_s42_2_9aea62c3.png", left=Inches(1.05),
                         top=Inches(1.65), width=Inches(4.4))
        _add_media_image(slide, "NV_s42_3_93673a62.png", left=Inches(6.35),
                         top=Inches(1.65), width=Inches(2.4))
        _add_media_image(slide, "NV_s42_4_94c45bea.png", left=Inches(9.35),
                         top=Inches(1.65), width=Inches(2.4))
        _add_hierarchical_bullets(
            slide, Inches(0.85), Inches(4.55), Inches(11.6), Inches(1.5),
            [([("Clipping a coupon costs time, and time is worth least to "
                "the most price-sensitive customers", {})], 0, {}),
             ([("So the discount reaches exactly the customers who would "
                "not have paid full price", {})], 0, {})],
            size=21, line_spacing_pts=16)
        # 2026-09-14 (Nico): "make the box in the bottom dark blue."
        _add_rounded_filled_box(
            slide, Inches(0.90), Inches(6.30), Inches(11.53), Inches(0.52),
            "Coupons ensure incentive-compatibility",
            fill=NAVY, text_color=WHITE, size=20, bold=True,
            corner_pct=0.16)
    sl = make_diagram_slide(
        prs, page_num, tag or TAG_VERSION,
        "Second-Degree Price Discrimination: Coupons", draw)
    _set_notes(sl, NOTES_NV.get(42, ""))
    return sl


# ==========================================================================
#  5a · Flat Fee Pricing — Video 5
#      [V5 = NV 43 (agenda), 44, 45, 46, 47, 49]
# ==========================================================================

def s_context(prs, page_num, *, tag=None):
    """[NV 44] · [V5 s3]."""
    bullets = [
        ("So far: we had different consumers with different demand curves "
         "(marginal personal value)", 0),
        ("Extract surplus via direct / indirect price discrimination", 1),
        # "multiple units" in dark red (2026-09-14, Nico) -- it is the
        # switch from the earlier degrees to the advanced strategies
        ([("Now: consumers have identical demand curves, but each of "
           "them purchases ", {}),
          ("multiple units ", {"color": DARKRED, "bold": True}),
          ("of the product", {})], 0),
        # the phrase in dark red (2026-09-14, Nico)
        ([("We will focus on a “", {}),
          ("representative consumer", {"color": DARKRED, "bold": True}),
          ("”", {})], 1),
    ]
    def extras(slide):
        # his own image, added by hand on 2026-09-14
        # he replaced the image on 2026-09-14
        _add_media_image(slide, "NV_s58_img2.png", left=Inches(6.16),
                         top=Inches(2.32), width=Inches(7.06))

    return content_slide(prs, page_num, tag or TAG_FLAT, "Context", bullets,
                         size=26, sub_size=24, line_spacing_pts=18,
                         bullets_width=Inches(5.52), extras=extras,
                         notes=NOTES_NV.get(44))


def s_advanced_pricing(prs, page_num, *, tag=None):
    """[NV 45] · [V5 s4]."""
    bullets = [
        ("How to extract consumer surplus from a given consumer who is "
         "prepared to buy multiple units", 0),
        # he set this line in dark red by hand (2026-09-15): it is the one
        # idea on the slide that the three variants below all turn on
        ([("Diminishing marginal personal value comes into play",
           {"color": DARKRED})], 1, {}),
        ("We will study three variants of advanced pricing", 0),
        ("All-or-none offers (flat fee)", 1),
        ("Two-part tariffs (flat fee + usage fee)", 1),
        ("Block pricing", 1),
    ]
    def extras(slide):
        # 2026-09-14 (Nico): "add the video references (with our video box
        # from the agenda slides) to the last three bullet points."  Same
        # pill as the agenda: 1.14 x 0.36", gold, navy 13 pt bold, corner
        # 0.30, right-aligned to the agenda's own 12.85" edge.  The three
        # y positions are MEASURED off a render of this slide (the bullet
        # helper centres the block, so they cannot be derived), and the
        # labels come from the outline's item->video map, so a renaming
        # there cannot leave these pointing at the wrong video.
        # 2026-09-15: his own image on the right, so the pill column
        # comes in from 12.85" to 7.24" -- it now sits just past the
        # bullets rather than at the slide edge under the picture.
        _add_media_image(slide, "NV_s59_advanced.png", left=Inches(8.06),
                         top=Inches(2.62), width=Inches(5.04))
        for topic, y in ((4, 4.64), (5, 5.07), (6, 5.49)):
            pill = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                int(Inches(7.24) - Inches(1.14)), int(Inches(y)),
                int(Inches(1.14)), int(Inches(0.36)))
            try:
                pill.adjustments[0] = 0.30
            except Exception:
                pass
            pill.fill.solid()
            pill.fill.fore_color.rgb = GOLD
            pill.line.fill.background()
            pill.shadow.inherit = False
            _add_drop_shadow(pill)
            tf = pill.text_frame
            tf.margin_left = tf.margin_right = 0
            tf.margin_top = tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            para = tf.paragraphs[0]
            para.alignment = PP_ALIGN.CENTER
            run = para.add_run()
            run.text = COVERAGE_LABEL[topic]
            run.font.size = Pt(13)
            run.font.bold = True
            run.font.name = "Calibri"
            run.font.color.rgb = NAVY

    return content_slide(prs, page_num, tag or TAG_FLAT, "Advanced Pricing",
                         bullets, size=25, sub_size=23, extras=extras,
                         bullets_width=Inches(7.78),
                         line_spacing_pts=13, notes=NOTES_NV.get(45))


def s_netflix_flatfee(prs, page_num, *, tag=None):
    """[NV 46] · [V5 s5] — the flat fee takes the WHOLE area under demand.

    Same demand as Video 1 (P = 4 − Q/3), so the two slides are directly
    comparable: simple pricing raised $12 a month, the flat fee raises
    (12 × 4)/2 = $24, exactly twice as much.
    """
    area = NFX_FLAT

    def draw(slide):
        fig = SimpleFig(1.05, 6.55, 5.5, 4.9,
                        xmax=NFX_XMAX, ymax=NFX_YMAX)
        _fig_axes(slide, fig, x_title=NFX_AXIS_Q, y_title="P",
                  label_size=18)
        _fig_poly(slide, fig, [(0, NFX_PMAX), (0, 0), (NFX_QMAX, 0)],
                  fill=GOLD, line=GOLD, alpha=26000, line_w=1.0)
        _fig_line(slide, fig, (0, NFX_PMAX), (NFX_QMAX, 0), color=DARKRED,
                  weight_pt=2.75)
        _fig_curve_label(slide, fig, NFX_QMAX - 1.05, 0.56, "D",
                         color=DARKRED, size=20)
        _region_label(slide, fig,
                      [(0, NFX_PMAX), (0, 0), (NFX_QMAX, 0)],
                      "Revenues = $%d" % int(area), size=19, dy=-0.20)
        _fig_ylab(slide, fig, NFX_PMAX, "$%g" % NFX_PMAX)
        _fig_xlab(slide, fig, NFX_QMAX, "%g" % NFX_QMAX)
        _add_math_equation(
            slide, int(fig.x(0.975)), int(fig.y(4.56)), Inches(2.3),
            Inches(0.5),
            _omml_run("P") + _omml_text(" = %g − " % NFX_PMAX)
            + _omml_frac(_omml_run("Q"), _omml_text("%g" % NFX_SLOPE_DEN)),
            size_pt=22)
        # 2026-09-14 (Nico): "first bullet -- make this the triangle
        # shape.  Then say Triangle = ..."  The mark IS the reference, so
        # the line names the shape instead of its colour, and the shape
        # copies the region: a right triangle in the same gold wash, per
        # the legend-mark rule.  This box is top-anchored (it is drawn
        # inside draw(), not through content_slide), so the first line
        # starts at its own top and the mark can be placed against it.
        _tri = slide.shapes.add_shape(
            MSO_SHAPE.RIGHT_TRIANGLE, int(Inches(6.80)), int(Inches(1.99)),
            int(Inches(0.34)), int(Inches(0.24)))
        _tri.fill.solid()
        _tri.fill.fore_color.rgb = GOLD
        _set_fill_alpha(_tri, 26)
        _tri.line.color.rgb = NAVY
        _tri.line.width = Pt(0.75)
        _tri.shadow.inherit = False
        _add_hierarchical_bullets(
            slide, Inches(7.28), Inches(1.90), Inches(5.57), Inches(4.2),
            [([("Triangle = total value of unlimited movies for the "
                "consumer = the maximum fee Netflix can charge", {})], 0,
              {"bullet_style": "none"}),
             ([("Area = (%g × %g)/2 = %g"
                % (NFX_QMAX, NFX_PMAX, NFX_FLAT), {})], 0, {}),
             ([("With a $%g flat fee for unlimited access, Netflix can "
                "attain $%g revenue — twice as much as under simple "
                "pricing" % (NFX_FLAT, NFX_FLAT), {})], 0, {}),
             ([("The consumer will watch %g movies a month under this "
                "fixed-fee-unlimited-access system" % NFX_QMAX, {})],
              0, {}),
             ([("Consumer surplus in this case = 0", {})], 0, {}),
             ([("Note: only makes sense when MC = 0", {})], 0, {})],
            size=19, line_spacing_pts=13)
    sl = make_diagram_slide(prs, page_num, tag or TAG_FLAT,
                            "Example: Netflix Flat Fee (Assume MC = 0)",
                            draw)
    _set_notes(sl, NOTES_NV.get(46, ""))
    return sl


# Netflix's US plans as of September 2026.  2026-09-14 (Nico): "can you
# update the table to reflect the latest prices?  I've copied those as a
# small image, which also shows the correct options.  Check online if you
# need to update the various rows."  His image gives the three plans and
# their prices; the feature rows were checked against Netflix's own plan
# comparison as reported in September 2026 (the Basic plan is gone, and
# Netflix raised all three prices in March 2026).  Built as a NATIVE
# table, not a screenshot, so the next price change is a one-line edit.
NFX_PLANS = ("Standard with Ads", "Standard", "Premium")
NFX_PLAN_ROWS = [
    ("Monthly cost (US$)", ("$8.99", "$19.99", "$26.99")),
    ("Screens you can watch at the same time", ("2", "2", "4")),
    ("Devices you can download on", ("2", "2", "6")),
    ("Unlimited movies, TV shows and mobile games",
     ("\u2713", "\u2713", "\u2713")),
    ("Watch on laptop, TV, phone and tablet",
     ("\u2713", "\u2713", "\u2713")),
    ("HD (1080p) available", ("\u2713", "\u2713", "\u2713")),
    ("Ultra HD (4K) and HDR available", ("", "", "\u2713")),
    ("Ads", ("yes", "\u2014", "\u2014")),
    ("Extra member slots", ("\u2014", "1", "2")),
]


def s_netflix_strategy(prs, page_num, *, tag=None):
    """[NV 47] · [V5 s6] — the plan menu, rebuilt natively."""
    def draw(slide):
        _add_styled_table(
            slide, Inches(0.62), Inches(1.58), Inches(9.10), Inches(4.40),
            [("",) + NFX_PLANS]
            + [(lbl,) + vals for lbl, vals in NFX_PLAN_ROWS],
            col_widths=[Inches(4.30), Inches(1.60), Inches(1.60),
                        Inches(1.60)],
            font_size=14, header_size=15)
        # his own Netflix image, copied in on 2026-09-14
        _add_media_image(slide, "NV_s64_netflix.png", left=Inches(10.05),
                         top=Inches(1.58), width=Inches(2.73))
        # his own size for it (2026-09-14): sized to the sentence rather
        # than spanning the slide
        _add_rounded_filled_box(
            slide, Inches(3.46), Inches(6.44), Inches(6.59), Inches(0.45),
            "This pricing strategy is a mix of flat fee and versioning",
            fill=GOLD, text_color=NAVY, size=17, bold=True,
            corner_pct=0.16)
    sl = make_diagram_slide(prs, page_num, tag or TAG_FLAT,
                            "Netflix Pricing Strategy", draw)
    _set_notes(sl, NOTES_NV.get(47, ""))
    return sl

def s_classpass(prs, page_num, *, tag=None):
    """A11 [NEW – PG2 16] — ClassPass credits.

    A current example that is volume pricing and segment pricing at once:
    the plans are a credit bundle, and a class costs more credits when it
    is popular or at a peak hour.
    """
    def draw(slide):
        # 2026-09-14 (Nico): "make the image corners round" -- it was
        # flat as a screenshot; it reads as a product shot, so it takes
        # the deck's ordinary picture treatment.
        _add_media_image(slide, "PG2_s16_rId3.png", left=Inches(0.85),
                         top=Inches(1.65), width=Inches(7.7))
        _add_hierarchical_bullets(
            slide, Inches(9.00), Inches(2.10), Inches(3.85), Inches(3.8),
            [([("Buy a bundle of credits per month, not a class", {})], 0,
              {}),
             ([("More credits per month costs less per credit — volume "
                "pricing", {})], 0, {}),
             ([("A popular class, or a peak hour, costs more credits — "
                "segment pricing by time", {})], 0, {})],
            size=19, line_spacing_pts=16)
    sl = make_diagram_slide(prs, page_num, tag or TAG_FLAT,
                            "ClassPass: Credits as Volume Pricing", draw)
    _set_notes(sl, (
        "ClassPass sells credits rather than classes. You pick how many "
        "credits you want each month, and the more you buy the less each "
        "one costs you — that is volume pricing, the same idea as the flat "
        "fee we just did, only in steps. Then the second half: a class does "
        "not cost a fixed number of credits. A popular studio, or a 6pm "
        "slot, costs more credits than an off-peak one. So the same "
        "subscription is doing volume pricing and segment pricing by time "
        "of day at the same time."))
    return sl


def s_flat_rate_wrong(prs, page_num, *, tag=None):
    """[NV 49] · [V5 s7] — what can go wrong with flat rates.

    2026-09-09 CORRECTION.  My first pass invented two bullets here because
    the python-pptx inventory showed the slide as a headline and nothing
    else.  It is not: the real content sits inside an mc:AlternateContent
    block (OMML "MC > 0"), which python-pptx does not enumerate.  Nico's
    own wording is restored below — the AAirpass quote, his two questions,
    and his take-away.  "Only 65 customers bought it" is the ANSWER to the
    first question, so it stays in the speaker notes where he had it.
    """
    def draw(slide):
        _add_media_image(slide, "NV_s49_2_2607387e.png", left=Inches(3.35),
                         top=Inches(1.45), width=Inches(6.6),
                         rounded=False)
        _add_convention_box(
            slide, Inches(0.80), Inches(2.95), Inches(11.7), Inches(1.42),
            prefix="Lifetime first class flat rate sold in the 1980-90s:  ",
            body="“Each had paid American more than $350,000 for an "
                 "unlimited AAirpass and a companion ticket that allowed "
                 "them to take someone along on their adventures. Both "
                 "agree it was the best purchase they ever made, one that "
                 "completely redefined their lives.”",
            size=17, anchor=MSO_ANCHOR.MIDDLE)
        _add_hierarchical_bullets(
            slide, Inches(0.80), Inches(4.50), Inches(11.7), Inches(1.10),
            [([("How many customers bought it?", {})], 0, {}),
             ([("How did they behave?", {})], 0, {})],
            size=22, line_spacing_pts=12)
        _add_rounded_filled_box(
            slide, Inches(0.80), Inches(5.75), Inches(11.7), Inches(1.10),
            "If MC > 0, offering a flat rate with unlimited access is a bad "
            "idea:\nconsumers will use ‘too much’ of the service — this can "
            "be solved by two-part tariffs",
            fill=GOLD, text_color=NAVY, size=18, bold=True, corner_pct=0.14)
    sl = make_diagram_slide(prs, page_num, tag or TAG_FLAT,
                            "What Can Go Wrong with Flat Rates…", draw)
    _set_notes(sl, (NOTES_NV.get(49, "") or "").strip() or
               "Only 65 customers bought the flat-rate pass.")
    return sl


# ==========================================================================
#  5b · Two-Part Tariff — Video 6
#      [V6 = NV 50 (agenda), 51, 52, 53]
# ==========================================================================

def s_two_part_chart(prs, page_num, *, tag=None):
    """[NV 51] · [V6 s3] — the two-part tariff once MC > 0.

    Demand P = 4 − Q/3 again; MC = $1.50, so the consumer buys where
    demand meets MC — Q = 7.5, computed from the two lines.  F* is the
    whole consumer surplus above P* = MC.

    The MC is the same FRACTION of the choke price it always was (0.375),
    so the picture keeps its proportions through the 2026-09-13 rescale.
    Neither number is ever printed here — this slide labels F* and P*
    symbolically — so only the geometry has to come out right.
    """
    mc = 0.375 * NFX_PMAX                              # $1.50
    q_mc = (NFX_PMAX - mc) * NFX_SLOPE_DEN             # = 7.5

    def draw(slide):
        fig = SimpleFig(1.05, 6.55, 5.5, 4.9,
                        xmax=NFX_XMAX, ymax=NFX_YMAX)
        _fig_axes(slide, fig, x_title="Q", y_title="P", label_size=18)
        _fig_poly(slide, fig, [(0, NFX_PMAX), (0, mc), (q_mc, mc)],
                  fill=GOLD, line=GOLD, alpha=26000, line_w=1.0)
        # 2026-09-15 (Nico): the area UNDER the usage fee is the revenue
        # it raises, in the same wash slide 69 uses for it
        _fig_poly(slide, fig, [(0, mc), (0, 0), (q_mc, 0), (q_mc, mc)],
                  fill=NAVY, line=NAVY, alpha=16000, line_w=1.0)
        _fig_line(slide, fig, (0, NFX_PMAX), (NFX_QMAX, 0), color=DARKRED,
                  weight_pt=2.75)
        _fig_line(slide, fig, (0, mc), (NFX_QMAX + 1.2, mc), color=NAVY,
                  weight_pt=2.25)
        _fig_curve_label(slide, fig, NFX_QMAX + 1.35, mc + 0.28, "MC",
                         color=NAVY, size=19)
        _fig_curve_label(slide, fig, NFX_QMAX - 1.05, 0.56, "D",
                         color=DARKRED, size=20)
        _region_label(slide, fig, [(0, NFX_PMAX), (0, mc), (q_mc, mc)],
                      "F*", size=22)
        # 2026-09-15 (Nico): the grey box says what it is, like the
        # legend row beside it
        _region_label(slide, fig, [(0, mc), (0, 0), (q_mc, 0), (q_mc, mc)],
                      "Revenues from usage fee", size=17)
        _fig_ylab(slide, fig, mc, "P*")
        _fig_guide(slide, fig, (q_mc, mc))
        # The two lines name things that are DRAWN on the chart, so each
        # carries a mark in that thing's colour and shape (2026-09-10,
        # Nico -- the legend-mark rule Module 4 established): a gold right
        # triangle for the flat-fee area, whose right angle is at the
        # bottom left exactly as the region's is, and a short navy bar for
        # the MC line.
        _welfare_rows(
            slide, Inches(7.95), 2.62, Inches(4.9),
            [("h", [(GOLD, "tri", 26)], "Flat fee F* = the entire "
                                        "consumer surplus"),
             ("h", [(NAVY, "line")], "Usage fee P* = marginal cost"),
             ("h", [(NAVY, "sq", 16)], "Revenue from usage fee")],
            size=22, pitch=0.62, swatch=0.24, text_dx=0.44)
    sl = make_diagram_slide(prs, page_num, tag or TAG_TWOPART,
                            "Two-Part Tariff (Now MC > 0)", draw)
    _set_notes(sl, NOTES_NV.get(51, ""))
    return sl


def s_two_part_summary(prs, page_num, *, tag=None):
    """[NV 52] · [V6 s4] — Nico's own summary, restored.

    2026-09-09 CORRECTION.  I first recorded this slide as "title-only in
    every deck and in the tape" and filled it with PG2 18 (adoption A9).
    Both halves of that were wrong: the slide carries a six-paragraph
    content placeholder wrapped in mc:AlternateContent (it holds OMML for
    F, P, P* and F*), which python-pptx does not enumerate — and the tape
    animates those six paragraphs one click at a time.  PG's slide says
    the same things because it came from this one.  A9 is WITHDRAWN and
    Nico's wording is used verbatim.

    His paragraph order, and the tape's click grouping, are preserved: the
    usage-fee rule and "avoids excessive use" share a click.
    """
    ital = {'italic': True}

    def draw(slide):
        # 2026-09-14 (Nico): "image added and textbox moved" -- the text
        # column narrows to 6.15" so the Zipcar photo has the right half.
        _add_hierarchical_bullets(
            slide, Inches(0.80), Inches(1.95), Inches(6.15), Inches(4.5),
            [([("Pricing decision is setting the flat fee (", {}),
               ("F", ital), (") and the per-unit usage fee (", {}),
               ("P", ital), (")", {})], 0, {}),
             ([("Zipcar: annual membership (fixed fee), hourly rental fee "
                "(per-unit price)", {})], 1, {}),
             ([("Amazon Prime: yearly charge (fixed fee), any purchase "
                "(per-unit price)", {})], 1, {}),
             ([("Set the usage fee  ", {}), ("P*", ital),
               (" = MC", {})], 0, {}),
             ([("Avoids excessive use", {})], 1, {}),
             ([("Set the flat fee  ", {}), ("F*", ital),
               (" = entire consumer surplus", {})], 0, {})],
            size=25, sub_size=23, line_spacing_pts=15)
        # 565 x 271 -> 2.49" tall at 5.18", at the place he set
        _add_media_image(slide, "NV_s68_twopart.png", left=Inches(7.49),
                         top=Inches(2.94), width=Inches(5.18))
    return make_diagram_slide(prs, page_num, tag or TAG_TWOPART,
                              "Two-Part Tariff: Summary", draw)


def s_zipcar(prs, page_num, *, tag=None):
    """[NV 53] · [V6 s5] — ZipCar as a worked two-part tariff.

    Demand P = 1.50 − Q/100 (so P = $0.50 at Q = 100 miles), MC = $0.50.
    Flat fee = (1.50 − 0.50) × 100 / 2 = $50; usage revenue = 0.50 × 100 =
    $50.  Both are computed here, so the labels cannot drift from the
    geometry.
    """
    p0, slope, mc = 1.50, 1.0 / 100.0, 0.50
    q_mc = (p0 - mc) / slope                            # 100 miles
    flat = (p0 - mc) * q_mc / 2.0                       # $50
    usage = mc * q_mc                                   # $50
    assert abs(flat - 50.0) < 1e-9 and abs(usage - 50.0) < 1e-9

    def draw(slide):
        fig = SimpleFig(1.15, 6.55, 5.6, 4.9, xmax=175.0, ymax=1.85)
        _fig_axes(slide, fig, x_title="Q (miles per year)", y_title="P",
                  label_size=17)
        _fig_poly(slide, fig, [(0, p0), (0, mc), (q_mc, mc)],
                  fill=GOLD, line=GOLD, alpha=26000, line_w=1.0)
        _fig_poly(slide, fig, [(0, mc), (0, 0), (q_mc, 0), (q_mc, mc)],
                  fill=NAVY, line=NAVY, alpha=16000, line_w=1.0)
        _fig_line(slide, fig, (0, p0), (p0 / slope, 0), color=DARKRED,
                  weight_pt=2.75)
        _fig_line(slide, fig, (0, mc), (165.0, mc), color=NAVY,
                  weight_pt=2.25)
        _fig_curve_label(slide, fig, 148.0, mc + 0.11, "MC", color=NAVY,
                         size=18)
        _fig_curve_label(slide, fig, 140.0, 0.20, "D", color=DARKRED,
                         size=19)
        # Nico's own demand equation, which the inventory hid inside an
        # mc:AlternateContent block: P = 1.5 - 0.01Q.  It is exactly the
        # line drawn above, so the figure and the algebra agree.
        # 2026-09-15 (Nico): dark red, at the height he set -- the same
        # treatment the zoo slides' formulas now carry
        _add_math_equation(
            slide, Inches(1.50), Inches(2.17), Inches(2.70), Inches(0.45),
            _omml_run("P") + _omml_text(" = 1.5 − 0.01") + _omml_run("Q"),
            size_pt=21, color=DARKRED)
        _fig_ylab(slide, fig, p0, "$1.50")
        _fig_ylab(slide, fig, mc, "$0.50")
        _fig_xlab(slide, fig, q_mc, "100")
        _fig_guide(slide, fig, (q_mc, mc))
        _region_label(slide, fig, [(0, p0), (0, mc), (q_mc, mc)],
                      "Flat fee = $%d" % int(flat), size=16, dy=0.06)
        _region_label(slide, fig, [(0, mc), (0, 0), (q_mc, 0), (q_mc, mc)],
                      # 2026-09-15 (Nico): the rectangle is the REVENUE
                      # the usage fee raises, not the fee itself
                      "Revenues from Usage Fee = $%d" % int(usage),
                      size=16)
        # 2026-09-15 (Nico): "in the legend, use the symbols for triangle
        # and grey box" -- the same vocabulary slide 67 now carries, so a
        # student meets one set of marks across every two-part slide.
        # "Two-part tariff:" is a HEADER, not a bullet -- his own
        # formatting, adopted on 72/73/74 first and here on 2026-09-15
        _add_text(slide, Inches(7.84), Inches(2.21), Inches(4.68),
                  Inches(0.40), "Two-part tariff:", size=20, bold=True,
                  color=NAVY, font="Calibri", align=PP_ALIGN.LEFT)
        _region_bullets(
            slide, Inches(8.05), 2.62, Inches(4.8), 3.05,
            [(("zfee",), "Flat fee  F = $%d" % int(flat), 1),
             (("zline",), "Usage fee  P = $0.50 / mile", 1),
             (("zuse",), "Revenue from usage fee = $%d" % int(usage), 1),
             (None, "Would a flat fee with no usage fee (P = 0) be a "
              "good idea?", 0),
             (("rarr",), "No — it would lead to excessive use", 1)],
            size=21, sub_size=19, mark_w=0.50, mark_h=0.46, gap=0.12)
    sl = make_diagram_slide(prs, page_num, tag or TAG_TWOPART,
                            "Two-Part Tariff: ZipCar", draw)
    _set_notes(sl, NOTES_NV.get(53, ""))
    return sl


# ==========================================================================
#  5c · Block Pricing — Video 7        [V7 = NV 60 (agenda), 61, 62]
# ==========================================================================

def s_block_concept(prs, page_num, *, tag=None):
    """[NV 61] · [V7 s3] — the definition as a hero concept box."""
    def draw(slide):
        _add_rounded_filled_box(
            slide, Inches(1.35), Inches(2.60), Inches(10.6), Inches(1.85),
            "Block pricing is the practice of reducing the price of a good\n"
            "when the same customer buys more of it",
            fill=NAVY, text_color=WHITE, size=26, bold=True,
            corner_pct=0.10)
        _add_text(slide, Inches(2.15), Inches(4.90), Inches(9.0),
                  Inches(0.9),
                  "The discount is tied to the QUANTITY one customer takes,\n"
                  "not to who the customer is",
                  size=21, bold=False, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)
    return make_diagram_slide(prs, page_num, tag or TAG_BLOCK, "Block Pricing",
                              draw)


# Walmart photo cards.  The three purchase options on my slide are 100 at
# $25, 125 at $30 and 175 at $35, i.e. marginal prices of $0.25, $0.20 and
# $0.10.  Those three points are COLLINEAR — P = 0.45 − 0.002Q — so the
# blocks really do sit under one demand curve, and MC = $0.05 meets it at
# Q = 200.
WM_A, WM_B_ = 0.45, 0.002
WM_MC = 0.05
WM_STEPS = [(100.0, 0.25), (125.0, 0.20), (175.0, 0.10)]
for _q, _p in WM_STEPS:
    assert abs((WM_A - WM_B_ * _q) - _p) < 1e-9, "block step off the curve"


def s_walmart_blocks(prs, page_num, *, tag=None):
    """[NV 62] · [V7 s4] — block pricing at Walmart, as native shapes."""
    def draw(slide):
        fig = SimpleFig(1.35, 6.55, 5.3, 4.85, xmax=225.0, ymax=0.50)
        # a short y title: "Price ($/photo card)" overflowed its
        # label-width box and rendered clipped
        _fig_axes(slide, fig, x_title="Q (photo cards)",
                  y_title="P ($/card)", label_size=16)
        q_prev = 0.0
        for letter, (q, p) in zip("ABC", WM_STEPS):
            _fig_poly(slide, fig,
                      [(q_prev, p), (q_prev, WM_MC), (q, WM_MC), (q, p)],
                      fill=GOLD, line=GOLD, alpha=26000, line_w=1.0)
            _region_label(slide, fig,
                          [(q_prev, p), (q_prev, WM_MC), (q, WM_MC),
                           (q, p)], letter, size=19)
            _fig_ylab(slide, fig, p, "$%.2f" % p)
            _fig_xlab(slide, fig, q, "%d" % int(q))
            q_prev = q
        _fig_line(slide, fig, (0, WM_A), (WM_A / WM_B_, 0), color=DARKRED,
                  weight_pt=2.75)
        _fig_line(slide, fig, (0, WM_MC), (215.0, WM_MC), color=NAVY,
                  weight_pt=2.25)
        _fig_curve_label(slide, fig, 202.0, WM_MC + 0.022, "MC",
                         color=NAVY, size=17)
        _fig_curve_label(slide, fig, 190.0, 0.135, "D", color=DARKRED,
                         size=19)
        _fig_ylab(slide, fig, WM_MC, "$0.05")
        _add_convention_box(
            slide, Inches(7.60), Inches(1.62), Inches(5.25), Inches(1.35),
            prefix="Three purchase options:  ",
            body="100 cards at $25 · 125 cards at $30 · 175 cards at $35",
            size=18, anchor=MSO_ANCHOR.MIDDLE)
        _add_hierarchical_bullets(
            slide, Inches(7.60), Inches(3.20), Inches(5.25), Inches(3.0),
            [([("Marginal cost of printing photo cards: MC = $0.05", {})],
              0, {}),
             ([("Assume zero fixed costs", {})], 0, {}),
             ([("The first option is 100 cards at $0.25", {})], 0, {}),
             ([("Profits are A (since TFC = 0)", {})], 1, {}),
             ([("Sell additional 25 cards at $0.20 ($5 extra for 25 cards)",
                {})], 0, {}),
             ([("Additional profits B", {})], 1, {}),
             ([("Sell yet additional 50 cards at $0.10 ($5 extra for 50 "
                "cards)", {})], 0, {}),
             ([("Additional profits C", {})], 1, {})],
            size=17, sub_size=16, line_spacing_pts=7)
    sl = make_diagram_slide(prs, page_num, tag or TAG_BLOCK,
                            "Block Pricing: Photo Cards at Walmart", draw)
    _set_notes(sl, NOTES_NV.get(62, ""))
    return sl


# ==========================================================================
#  6 · Summary of Pricing Strategies — Video 8     [V8 = NV 64 only]
# ==========================================================================

def s_summary_tree(prs, page_num, *, tag=None):
    """[NV 64] / [NVapp 44] — the decision tree, native.

    2026-09-09 (Nico): "stay much closer to my original.  Location of
    boxes, spacing, color scheme (only adjust colors to fit our broad
    scheme, but keep the different color groups)."

    HIS GEOMETRY, scaled from his 10 x 7.5" canvas to 13.33 x 7.5" (x is
    multiplied by 1.333; y is shifted down 0.43" to clear our top bar and
    action title, which his deck did not have -- his title sat in the
    upper LEFT instead).  His arrangement is what is reproduced: the four
    questions in a centre column, their "No" outcomes in a right column,
    and the "Yes" chain dropping to a direct-price-discrimination box at
    the BOTTOM LEFT with its two terminals side by side beneath it.

    HIS COLOUR GROUPS, re-expressed in our palette, as revised
    2026-09-15:
        questions                          -> NAVY filled, white text
        "perfect competition"              -> DARK GREEN (no market power
                                              at all -- the one branch
                                              where price is not a choice)
        "simple pricing"                   -> CREAM (a price setter, but
                                              one price for everyone)
        advanced / indirect pricing        -> GOLD
        DIRECT price discrimination        -> DARK RED, and its two
                                              terminals are outlined in
                                              dark red to match
    The gold family was one colour because it is all complex pricing.
    Direct price discrimination is now split out of it in dark red: it is
    the branch that needs the firm to identify customers one by one, and
    the two cards under it are its two degrees.
    """
    def draw(slide):
        # His columns, scaled -- and then opened up (2026-09-10, Nico:
        # "Make the spaces between boxes larger, so the arrows and yes /
        # no are larger.  Stay very close to my original slides on this
        # formatting").  Both column ORIGINS are still his, `qx` 5.89 and
        # `ox` 9.57; the room comes from the boxes, which were far taller
        # and wider than their text needs:
        #   * every question wraps to 2 lines at 16 pt, measuring 0.64"
        #     plus the box's 0.10" of vertical inset -- so 0.78", not the
        #     0.92" they had.  That buys 0.14" per row.
        #   * the "No" arrow gap, 0.53", is opened to 0.71" by taking
        #     0.10" off the question boxes and 0.08" off the outcome
        #     boxes.  Both columns keep at least 0.20" of WRAP SLACK on
        #     every line -- PowerPoint needs a little more width than PIL
        #     measures, so `qw` 2.85 (0.01" of slack on question 3) spilled
        #     two questions onto 3 lines and burst their boxes.
        # Result: uniform 0.34" gaps between the question boxes, up from
        # 0.17-0.23", and Yes / No up from 13 pt to 15.
        #
        # 0.34" is what the BOTTOM of the tree allows, not a preference.
        # Question 4 sits above the two terminal cards, which reach across
        # to x 6.71 and cannot move (their text needs every bit of the
        # 1.18" they have, and they already stop 0.05" short of the footer
        # rule).  The final "Yes" hangs below question 4 in that band, so
        # every 0.01" added to a row gap is taken off that label's room.
        qx, qw = Inches(5.89), Inches(3.05)
        ox, ow = Inches(9.65), Inches(3.40)   # right edge still 13.05
        # 2026-09-15 (Nico): "make the boxes tighter around the text, so
        # there is more space between them."  A question wraps to 2 lines
        # at 16 pt: 2 x 0.27" of rendered line plus 0.14" of vertical
        # inset is 0.68", not the 0.78" they carried.  The 0.10" saved on
        # each of four rows plus the room that frees goes into the gaps,
        # which go 0.34" -> 0.44".  The tree still ends above the
        # terminal cards: 4 x 0.68 + 3 x 0.44 from 1.45" ends at 5.49",
        # which leaves the final "Yes" MORE room than before, not less.
        q_h = Inches(0.68)
        q_gap = Inches(0.44)
        q_tops = [int(Inches(1.45) + i * (q_h + q_gap)) for i in range(4)]
        # Outcome boxes carry 2 or 3 lines, so each is sized to its own
        # text and centred on its question box -- his layout let them
        # overhang the row for the same reason.  They cannot be widened
        # into 2 lines: even at ow 3.75 the 3-line ones stay 3 lines, and
        # 3.90 runs off the canvas.
        o_h = [Inches(0.68), Inches(0.95), Inches(0.95), Inches(0.95)]
        o_tops = [int(q_tops[i] + (q_h - o_h[i]) / 2) for i in range(4)]
        yn_size = 15
        yn_w, yn_h = Inches(0.55), Inches(0.26)

        questions = [
            "Does the firm have market power?",
            "Can the firm prevent resale and arbitrage?",
            "Do the firm’s customers have different demand curves?",
            "Can the firm directly identify different customers?",
        ]
        # (text, fill, text colour, border) -- his colour groups
        outcomes = [
            ("Perfect competition produces quantity at which MR = P = MC",
             GREEN_DK, WHITE, None),
            # his own wording, 2026-09-15
            ("Simple pricing under monopolistic competition: Set price "
             "such that MR = MC at optimal quantity Q*", CREAM, NAVY,
             NAVY),
            # a no-break space glues each "•" to its own item, so a wrap
            # can only fall BEFORE a bullet, never leaving one orphaned
            # at the end of a line (2026-09-10)
            ("Advanced pricing strategies\n• Flat-fee  • Two-part "
             "tariff  • Block pricing", GOLD, NAVY, None),
            ("Indirect (second-degree) price discrimination\n"
             "• Versioning  • Coupons", GOLD, NAVY, None),
        ]
        for i, q in enumerate(questions):
            _add_rounded_filled_box(slide, qx, q_tops[i], qw, q_h, q,
                                    fill=NAVY, text_color=WHITE, size=16,
                                    bold=True, corner_pct=0.12)
            label, fill, tc, border = outcomes[i]
            _add_outlined_box(slide, ox, o_tops[i], ow, o_h[i], label,
                              line=border or fill, text_color=tc, fill=fill,
                              size=16, bold=False, rounded=True,
                              shadow=True, corner_pct=0.12)
            mid = int(q_tops[i] + q_h / 2)
            _add_arrow(slide, (int(qx + qw), mid), (int(ox), mid),
                       color=NAVY, weight_pt=1.75)
            # "No" at the arrow's ORIGIN end, sitting just clear ABOVE the
            # line rather than on it (the arrow-origin rule)
            _add_text(slide, int(qx + qw + Inches(0.06)),
                      int(mid - yn_h - Inches(0.04)), yn_w, yn_h,
                      "No", size=yn_size, bold=True, color=GRAY,
                      font="Calibri", align=PP_ALIGN.CENTER)
            if i < 3:
                _add_arrow(slide, (int(qx + qw / 2), int(q_tops[i] + q_h)),
                           (int(qx + qw / 2), int(q_tops[i + 1])),
                           color=NAVY, weight_pt=1.75)
                _add_text(slide, int(qx + qw / 2 + Inches(0.09)),
                          int(q_tops[i] + q_h + Inches(0.04)), yn_w, yn_h,
                          "Yes", size=yn_size, bold=True, color=GRAY,
                          font="Calibri", align=PP_ALIGN.LEFT)

        # the final Yes drops LEFT and DOWN to the direct-PD box, exactly
        # as in his layout
        dx, dw = Inches(2.85), Inches(2.20)
        dy, dh = Inches(5.22), Inches(0.55)
        _add_arrow(slide, (int(qx + Inches(0.30)), int(q_tops[3] + q_h)),
                   (int(dx + dw), int(dy + Inches(0.10))), color=NAVY,
                   weight_pt=1.75)
        # the label goes at the arrow's ORIGIN -- under question 4, where
        # the branch leaves -- not down at the head next to the box it
        # points to (2026-09-09).  It hangs in the band between question
        # 4 and the terminal cards, which is why the row gaps above are
        # 0.34" and not more.
        _add_text(slide, int(qx + Inches(0.36)),
                  int(q_tops[3] + q_h + Inches(0.03)), yn_w, yn_h,
                  "Yes", size=yn_size, bold=True, color=GRAY,
                  font="Calibri", align=PP_ALIGN.LEFT)
        _add_rounded_filled_box(slide, dx, dy, dw, dh,
                                "Direct price discrimination",
                                fill=DARKRED, text_color=WHITE, size=16,
                                bold=True, corner_pct=0.16)
        # his two terminals, side by side under it
        tw, th = Inches(2.95), Inches(1.06)
        ty = Inches(5.98)
        for left, lead, tail in (
                (Inches(0.56), "Complete information on every customer",
                 "perfect (first-degree) price discrimination"),
                (Inches(3.66), "Information on groups of customers",
                 "segmenting (third-degree) price discrimination")):
            _add_hierarchical_bullets(
                slide, left, ty, tw, th,
                [([(lead, {'bold': True})], 0,
                  {'bullet_style': 'none', 'space_before_pts': 0}),
                 ([("→ " + tail, {})], 0,
                  {'bullet_style': 'none', 'space_before_pts': 2})],
                size=16, line_spacing_pts=0)
            box = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, int(left - Inches(0.10)),
                int(ty - Inches(0.06)), int(tw + Inches(0.20)),
                int(th + Inches(0.12)))
            try:
                box.adjustments[0] = 0.12
            except Exception:
                pass
            box.fill.background()
            # dark red, to match the box they hang from
            box.line.color.rgb = DARKRED
            box.line.width = Pt(1.25)
            box.shadow.inherit = False
            # the outline sits BEHIND the text it frames
            box._element.getparent().remove(box._element)
            slide.shapes._spTree.insert(2, box._element)
        _add_arrow(slide, (int(dx + dw / 2), int(dy + dh)),
                   (int(dx + dw / 2), int(ty)), color=NAVY, weight_pt=1.75)
        _add_text(slide, ox, Inches(6.58), ow,
                  Inches(0.30), "Source: Goolsbee, Levitt, Syverson",
                  size=12, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)
    return make_diagram_slide(
        prs, page_num, tag or TAG_SUMTAG,
        "How to Extract Consumer Surplus? A Summary of Pricing Strategies",
        draw)


def s_module_summary(prs, page_num, *, tag=None):
    """[NV 74] — the summary closer, outside every video block."""
    bullets = [
        ("Limitations of simple pricing", 0),
        ("Direct price discrimination", 0),
        ("Personalized pricing (first degree) — (still) hypothetical", 1),
        ("Segment pricing (third degree) — higher price to less elastic "
         "groups", 1),
        ("Indirect price discrimination", 0),
        ("Versioning and coupons (second degree)", 1),
        ("Advanced pricing strategies", 0),
        ("If MC = 0: flat fee and unlimited access (e.g., Netflix)", 1),
        ("If MC > 0: two-part tariff — usage fee = MC to avoid excessive "
         "use, flat fee = the area under demand above the per-unit fee", 1),
        ("Block pricing", 1),
    ]
    return content_slide(prs, page_num, TAG_SUMMARY, "Module 6: Summary",
                         bullets, size=23, sub_size=21,
                         line_spacing_pts=8, notes=NOTES_NV.get(74))


# ==========================================================================
#  In-Class Examples block
#
#  The applications-class deck's own sequence (its slides 7 – 52), with the
#  recap slides KEPT so the block runs standalone the way Nico's 2022 deck
#  did, plus the PG case additions and the second Costco slide.
#
#  Every slide here is tagged "Module 6 · In Class · Examples · <topic>",
#  the topic spliced in from the video topic it applies (never retyped), per
#  the four-level tag rule.  The block's two agenda slides read
#  "Module 6 · In Class · Agenda".
# ==========================================================================

def s_ic_divider(prs):
    return divider(prs, "Slides for the On-Campus Applications Class",
                   "Applications of the material in the eight videos")


def s_ic_outline(prs, page_num):
    """[NVapp 7] — what was on video, and what we do in class today."""
    bullets = [
        ("In the Videos:", 0),
        ("Simple vs. complex pricing", 1),
        ("Direct price discrimination", 1),
        ("Perfect price discrimination (first degree)", 2),
        ("Segment pricing (third degree)", 2),
        ("Indirect price discrimination (second degree)", 1),
        ("Versioning", 2),
        ("Coupons", 2),
        ("In class today:", 0),
        ("Advanced pricing strategies", 1),
        ("Flat fee pricing", 2),
        ("Two-part tariff", 2),
        ("Block pricing", 2),
    ]
    return content_slide(prs, page_num, TAG_IC_AGENDA, "Outline of Module 6",
                         bullets, size=21, sub_size=19, line_spacing_pts=4)


def s_ic_podcast_dilemma(prs, page_num):
    """[NVapp 8] — the dilemma, as the podcast tells it."""
    bullets = [
        ("MC of a burger is $4.50", 0),
        ("NYC restaurant sells it for $18", 0),
        ("Lose customers with MPV = $15, $11, etc", 0),
        ("What if you charge $10?", 0),
        ("Give it for too little to those with MPV > $18", 1),
    ]
    def extras(slide):
        # his own image, added by hand on 2026-09-14
        _add_media_image(slide, "NV_s08_new.png", left=Inches(6.78),
                         top=Inches(1.42), width=Inches(6.27))

    return content_slide(prs, page_num, TAG_IC[0],
                         "Dilemma of Simple Pricing, Explained by the "
                         "Podcast", bullets, size=26, sub_size=24,
                         line_spacing_pts=16, bullets_width=Inches(6.10),
                         extras=extras)


def s_ic_news(prs, page_num, *, title, image, questions, notes=None):
    """A13 / A14 [NEW – PG1 3, PG1 4] — an "In the News" discussion opener:
    the headline, then the questions to put to the class."""
    def draw(slide):
        _add_media_image(slide, image, left=Inches(1.55), top=Inches(1.60),
                         width=Inches(10.2), rounded=False, shadow=True)
        _add_hierarchical_bullets(
            slide, Inches(1.55), Inches(4.85), Inches(10.2), Inches(1.5),
            [([(q, {})], 0, {}) for q in questions],
            size=21, line_spacing_pts=15)
        _add_discussion_badge(slide)
    sl = make_diagram_slide(prs, page_num, TAG_IC[0], title, draw)
    if notes:
        _set_notes(sl, notes)
    return sl


def _add_discussion_badge(slide, label="Discussion"):
    """The gold badge, relabelled for a DISCUSSION rather than a poll.

    Same geometry as the Poll Break parallelogram (Teaching CLAUDE.md gives
    one fixed geometry for it), with the label swapped — a discussion slide
    reading "Poll Break" would promise a poll that never comes.  Never
    animated, and drawn last so it sits in front of the footer rule.
    """
    grp = _add_pollbreak_badge(slide)
    for tnode in grp.iter(qn('a:t')):
        if tnode.text == "Poll Break":
            tnode.text = label
    return grp


def s_ic_first_degree(prs, page_num):
    """[NVapp 11] — the in-class version ASKS whether it is hypothetical."""
    bullets = [
        ("Charge each customer a different price", 0),
        ("Also known as perfect price discrimination", 0),
        ("Is this a hypothetical scenario?", 0),
    ]
    return content_slide(prs, page_num, TAG_IC[1],
                         "First Degree Price Discrimination", bullets,
                         size=28, line_spacing_pts=18)


def s_ic_dystopian_q(prs, page_num):
    """[NVapp 12] · [NV 13] — the QUESTION version, for class.  The video
    states the point; here Nico asks it."""
    def draw(slide):
        _add_hierarchical_bullets(
            slide, Inches(0.70), Inches(1.95), Inches(5.1), Inches(4.2),
            [([("Why are companies increasingly able to use first-degree "
                "price discrimination (personally tailored prices)? Any "
                "examples?", {})], 0, {}),
             ([("Is the trend towards more price discrimination good or bad "
                "for consumers?", {})], 0, {})],
            size=22, line_spacing_pts=20)
        _add_media_image(slide, "NV_s13_g3.2_4f2ec7c5.png", left=Inches(6.20),
                         top=Inches(1.58), width=Inches(6.5), rounded=False,
                         shadow=False)
        # 2026-09-14: he moved this one down-right to overlap the banner
        # above it the way he wants the pair read, and grouped the two
        _add_media_image(slide, "NV_s13_g3.1_6471d638.png", left=Inches(6.90),
                         top=Inches(2.80), width=Inches(5.8))
        _add_discussion_badge(slide)
    return make_diagram_slide(
        prs, page_num, TAG_IC[1],
        "The Dystopian Future of Price Discrimination", draw)


def s_ic_uber(prs, page_num):
    """[NVapp 13] · [NV 16] — the live volunteer exercise.  Confirmed absent
    from Video 2, so it exists only here."""
    def draw(slide):
        _add_hierarchical_bullets(
            slide, Inches(0.70), Inches(2.05), Inches(6.6), Inches(4.0),
            [([("I need 10 volunteers (5 for Uber, 5 for Lyft)", {})], 0,
              {}),
             ([("Open your Uber/Lyft app", {})], 0, {}),
             ([("Enter “LAX” in the destination address", {})], 0, {}),
             ([("Do NOT hit request", {})], 0, {}),
             ([("Just tell me the UberX/Lyft price", {})], 0, {})],
            size=24, line_spacing_pts=18)
        # 2026-09-14: he swapped in his own applications-deck image (the
        # volunteer set-up) and marked the unknown price with a "?"
        _add_media_image(slide, "NVapp_s13_uber.png", left=Inches(7.17),
                         top=Inches(1.49), height=Inches(5.40))
        # 2026-09-14 (Nico): "you need to show also the text box that
        # covers the prices in the image."  The three fares are printed in
        # the screenshot, so the class can read the answer off the slide
        # unless they are masked -- the same device as the doll prices.
        # Two white blocks: one over the UberX fare, carrying the "?", and
        # one over the two fares below it.
        for mx, my, mw, mh in ((9.58, 4.98, 0.72, 0.32),
                               (9.58, 5.74, 0.72, 1.04)):
            _add_rect(slide, int(Inches(mx)), int(Inches(my)),
                      int(Inches(mw)), int(Inches(mh)), WHITE)
        _add_text(slide, int(Inches(9.62)), int(Inches(4.98)),
                  int(Inches(0.62)), int(Inches(0.32)), "?", size=24,
                  bold=True, color=DARKRED, font="Calibri",
                  align=PP_ALIGN.CENTER)
        _add_discussion_badge(slide)
    return make_diagram_slide(
        prs, page_num, TAG_IC[1],
        # his header, 2026-09-14
        "Is Uber Charging What You’re Willing to Pay?", draw)


def s_ic_lodine_setup(prs, page_num, *, tag=None, poll=True):
    """[NVapp 15] · [NV 26] — the Lodine "dog slide".

    Called twice since 2026-09-14: once inside Video 3, where he asked for
    his original 24 – 26 to be introduced, and once in the in-class block as
    the set-up for the PollEverywhere activity.  The video copy passes
    ``poll=False`` — there is no poll to break for in a taped video, and the
    Poll Break badge is a promise to the room.
    """
    def draw(slide):
        _add_hierarchical_bullets(
            slide, Inches(0.70), Inches(1.95), Inches(7.2), Inches(2.0),
            [([("Lodine, a medication for arthritis, is used for both "
                "humans and dogs", {})], 0, {}),
             ([("Which segment ‘pays’ a higher price?", {})], 0, {})],
            size=25, line_spacing_pts=20)
        _add_media_image(slide, "NV_s26_4_3b7dbd94.png", left=Inches(8.35),
                         top=Inches(1.95), height=Inches(4.3))
        if poll:
            _add_pollbreak_badge(slide)
    return make_diagram_slide(prs, page_num, tag or TAG_IC[2],
                              "Consumer Segments in Medication", draw)


def s_ic_lodine_solution(prs, page_num):
    """[NVapp 17] · [NV 28] — the answer, with the 3x figure from my notes."""
    def draw(slide):
        _add_rounded_filled_box(
            slide, Inches(1.20), Inches(1.85), Inches(6.4), Inches(1.05),
            "Higher price for humans!", fill=GOLD, text_color=NAVY,
            size=28, bold=True, corner_pct=0.14)
        _add_hierarchical_bullets(
            slide, Inches(1.20), Inches(3.25), Inches(6.4), Inches(2.2),
            [([("The price of Lodine for humans is three times the price "
                "for animals", {})], 0, {}),
             ([("Same molecule, same marginal cost — two segments the seller "
                "can tell apart, and can keep apart", {})], 0, {})],
            size=21, line_spacing_pts=18)
        # 597 x 442 -> 3.40" tall at 4.60" wide, so it ends at 12.80"
        _add_media_image(slide, "NV_s28_4_02cd023d.png", left=Inches(8.20),
                         top=Inches(2.15), width=Inches(4.60))
        _add_pollbreak_badge(slide)
    sl = make_diagram_slide(prs, page_num, TAG_IC[2], "Solution", draw)
    _set_notes(sl, NOTES_NV.get(28, ""))
    return sl


def s_ic_segment_analytics(prs, page_num):
    """[NVapp 18] · [NV 29] — the heavier analytics slide, which the tape
    does not use (Video 3 carries the lighter version)."""
    def draw(slide):
        _add_hierarchical_bullets(
            slide, Inches(0.75), Inches(1.75), Inches(8.4), Inches(1.7),
            [([("A firm produces the same product for two markets", {})], 0,
              {}),
             ([("Math:", {})], 0, {})],
            size=25, line_spacing_pts=14)
        _add_math_equation(
            slide, Inches(1.30), Inches(3.35), Inches(6.0), Inches(0.7),
            _omml_text("Set ") + _omml_sub(_omml_run("MR"), _omml_run("i"))
            + _omml_text(" = ") + _omml_sub(_omml_run("MC"), _omml_run("i"))
            + _omml_text("  in each market ") + _omml_run("i"),
            size_pt=26)
        _add_hierarchical_bullets(
            slide, Inches(0.75), Inches(4.45), Inches(8.4), Inches(1.9),
            [([("Prices different if demand (MR) and/or costs (MC) are "
                "different", {})], 0, {}),
             ([("Rule: charge lower price in more elastic markets", {})], 0,
              {})],
            size=22, line_spacing_pts=16)
    sl = make_diagram_slide(prs, page_num, TAG_IC[2],
                            "Important: Segment Pricing Analytically",
                            draw)
    _add_practice_video_box(sl)          # after the footer, see above
    return sl


# --------------------------------------------------------------------------
#  The week-1 airline puzzle and its answer  [NV 24, NV 25]
#
#  Both panels share ONE marginal cost -- it is the same LAX-FRA aircraft.
#  What differs is the demand: the DIRECT flight is bought by travellers
#  with no good substitute, so its demand is less elastic and its price
#  higher, even though it is the shorter trip.  Numbers chosen so that
#  falls out of MR = MC rather than being asserted:
#
#      connecting  P = 10 - Q     MR = 10 - 2Q   MC = 2  ->  Q* 4, P* 6
#      direct      P = 14 - 2Q    MR = 14 - 4Q   MC = 2  ->  Q* 3, P* 8
# --------------------------------------------------------------------------

# 2026-09-14 (Nico): "adopt the location of all curves exactly as in my
# original slides (eg different MC)."  Read off his slide 25 as fractions of
# each panel's own axes, then turned back into demand and cost parameters:
#
#                        his P-intercept   his MC    his P*
#   connecting (2 legs)      0.57              0.386     0.473
#   direct     (1 leg)       1.08 (off-panel)  0.285     0.702
#
# The connecting flight carries the HIGHER marginal cost, which is right --
# it is two legs, not one -- and his drawing says so even though the caption
# this deck had claimed the two shared an MC.  The numbers below reproduce
# those proportions while keeping the figure economically exact (MR bisects
# demand, Q* is the true MR = MC crossing), per the exactness rule.
LH_PMAX, LH_QMAX = 17.0, 15.0
LH_MC_CONN = 6.18              # two legs: the costlier itinerary
LH_MC_DIR = 4.56               # one leg
LH_MC = LH_MC_CONN             # kept for anything still reading one MC
LH_CONN = (9.1, 0.652)         # P = a - bQ, the connecting flight
LH_DIR = (16.8, 1.179)         # the direct flight


def _lh_opt(a, b, mc=None):
    mc = LH_MC if mc is None else mc
    q = (a - mc) / (2.0 * b)
    return q, a - b * q


LH_QC, LH_PC = _lh_opt(*LH_CONN, mc=LH_MC_CONN)   # 2.24, 7.64
LH_QD, LH_PD = _lh_opt(*LH_DIR, mc=LH_MC_DIR)     # 5.19, 10.68
assert LH_PD > LH_PC, "the direct flight must come out dearer"


# His own slide 24 rings ten details on the two itineraries in red -- the
# dates, the durations, the two "Business Basic" fares and, the point of the
# slide, the two totals.  2026-09-14 (Nico): "adopt all round shapes exactly
# in the right position from my original slide."
#
# "Exactly" means the same positions RELATIVE TO THE SCREENSHOTS, so the
# whole slide is placed through ONE similarity transform read off his 4:3
# canvas: X = 2.882 + 0.76x, Y = 1.46 + 0.76(y - 1.22).  The scale is what
# lets the two screenshots and the takeaway bar share the content area; the
# rings then land where he put them without any of them being re-measured.
_LH_PUZ_S = 0.76


def _puz_x(x):
    return Inches(2.882 + _LH_PUZ_S * x)


def _puz_y(y):
    return Inches(1.46 + _LH_PUZ_S * (y - 1.22))


# (x, y, w, h) on his canvas
_LH_PUZ_RINGS = [
    (0.48, 1.44, 1.78, 0.22),      # outbound date
    (0.40, 1.61, 1.78, 0.22),      # outbound route
    (2.75, 2.10, 0.33, 0.22),      # outbound stop
    (0.41, 2.98, 1.78, 0.22),      # return route
    (2.75, 3.49, 0.33, 0.22),      # return stop
    (8.03, 3.91, 1.78, 0.43),      # TOTAL, connecting itinerary
    (0.36, 4.35, 2.05, 0.37),      # second itinerary's heading
    (0.44, 4.65, 1.78, 0.22),
    (0.45, 6.08, 1.78, 0.22),
    (8.22, 7.04, 1.78, 0.43),      # TOTAL, direct itinerary
]


def s_ic_airline_puzzle(prs, page_num, *, tag=None):
    """[NV 24] — the two Lufthansa fares, as he screenshotted them."""
    def draw(slide):
        _add_media_image(slide, "NV_s24_4_257493da.png",
                         left=_puz_x(0.02), top=_puz_y(1.22),
                         width=Inches(9.64 * _LH_PUZ_S), rounded=False)
        _add_media_image(slide, "NV_s24_2_750d0262.png",
                         left=_puz_x(-0.02), top=_puz_y(4.22),
                         width=Inches(10.00 * _LH_PUZ_S), rounded=False)
        for x, y, w, h in _LH_PUZ_RINGS:
            ring = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, int(_puz_x(x)), int(_puz_y(y)),
                int(Inches(w * _LH_PUZ_S)), int(Inches(h * _LH_PUZ_S)))
            ring.fill.background()
            ring.line.color.rgb = RGBColor(0xFF, 0x00, 0x00)
            ring.line.width = Pt(2.0)
            ring.shadow.inherit = False
        _add_takeaway_bar(
            slide,
            "The longer trip is the cheaper one \u2014 why?",
            top=Inches(6.42), text_color=NAVY, size=20, rounded=True,
            shadow=True)
    sl = make_diagram_slide(
        prs, page_num, tag or TAG_IC[2],
        "Airline Pricing: We Can Now Solve the \u2018Puzzle\u2019 from Week 1",
        draw)
    _set_notes(sl, (
        "Back in week 1 we looked at these two Lufthansa fares and could "
        "not explain them. The connecting flight from LAX to Istanbul via "
        "Frankfurt is CHEAPER than the direct LAX to Frankfurt flight, even "
        "though the connecting passenger occupies the same seat on the same "
        "LAX-Frankfurt aircraft and then flies a second leg on top. On cost "
        "grounds that makes no sense. Now we have the tool: this is segment "
        "pricing. The two passengers are in different segments with "
        "different elasticities, and the airline can tell them apart and "
        "keep them apart."))
    return sl


def _lh_panel(slide, left_in, width_in, *, a, b, qstar, pstar, heading,
              key="", mc=None,
              qmax, pmax, top_in=2.30, height_in=3.85):
    """One market panel of the Lufthansa answer."""
    fig = SimpleFig(left_in, top_in + height_in, width_in, height_in,
                    xmax=qmax, ymax=pmax)
    _add_text(slide, int(Inches(left_in - 0.30)), int(Inches(top_in - 0.46)),
              int(Inches(width_in + 0.6)), Inches(0.34), heading, size=18,
              bold=True, color=NAVY, font="Calibri", align=PP_ALIGN.CENTER)
    _fig_axes(slide, fig, x_title="Q", y_title="P", label_size=16)
    _fig_line(slide, fig, (0, a), (a / b, 0), color=DARKRED, weight_pt=2.5,
              curve="D" + key)
    # MR: dashed, and the same dark red as the demand it comes from
    _fig_line(slide, fig, (0, a), (a / (2.0 * b), 0), color=DARKRED,
              weight_pt=2.0, dash="dash", curve="MR" + key)
    mc = LH_MC if mc is None else mc
    _fig_line(slide, fig, (0, mc), (qmax * 0.95, mc), color=NAVY,
              weight_pt=2.0)
    _fig_curve_label(slide, fig, qmax * 0.80, mc + pmax * 0.05, "MC",
                     color=NAVY, size=16)
    # just BELOW demand at its right end.  A fixed fraction of pmax put the
    # label wherever the panel's own MC happened to be once the two panels
    # stopped sharing one MC, and the MC line ran straight through it.
    # just above demand near its Q intercept, which is the only band clear
    # of BOTH the demand line and this panel's own MC (they differ per
    # panel since 2026-09-14, and a fixed fraction of pmax sat on one)
    _fig_curve_label(slide, fig, a / b * 0.92, a * 0.08 + pmax * 0.06, "D",
                     color=DARKRED, size=17, curve="D" + key,
                     clear_of=((0, a), (a / b, 0)))
    _fig_curve_label(slide, fig, a / (2.0 * b) * 0.40, pmax * 0.06, "MR",
                     color=DARKRED, size=16, curve="MR" + key,
                     clear_of=((0, a), (a / (2.0 * b), 0)))
    _fig_guide(slide, fig, (qstar, pstar))
    _fig_ylab(slide, fig, pstar, "P*")
    _fig_xlab(slide, fig, qstar, "Q*")
    return fig


def s_ic_lufthansa(prs, page_num, *, tag=None):
    """[NV 25] — the answer, as two panels sharing one MC."""
    def draw(slide):
        QM, PM = LH_QMAX, LH_PMAX
        _lh_panel(slide, 1.00, 4.55, key=":conn", mc=LH_MC_CONN,
                  a=LH_CONN[0], b=LH_CONN[1],
                  qstar=LH_QC, pstar=LH_PC, qmax=QM, pmax=PM,
                  heading="Flight LAX \u2013 IST via FRA")
        _lh_panel(slide, 7.10, 4.55, key=":dir", mc=LH_MC_DIR,
                  a=LH_DIR[0], b=LH_DIR[1],
                  qstar=LH_QD, pstar=LH_PD, qmax=QM, pmax=PM,
                  heading="Direct flight LAX \u2013 FRA")
        # 2026-09-14 (Nico): "make the box in the bottom dark blue.  Make
        # sure it's not overlapping with the graphs/labels."  The panels
        # bottom out at an axis at y 6.20", and their Q* / Q labels hang to
        # 6.40", so the box starts at 6.55" and is 0.50" tall: clear of the
        # labels above and of the footer rule at 7.15" below.
        #
        # The caption also drops its "same MC" clause: his own drawing gives
        # the two panels DIFFERENT marginal costs, and the connecting
        # itinerary is two legs, so the claim was never right.
        _add_rounded_filled_box(
            slide, Inches(1.00), Inches(6.55), Inches(11.33), Inches(0.50),
            "Higher price for the shorter direct flight \u2014 less elastic "
            "demand, despite its lower MC",
            fill=NAVY, text_color=WHITE, size=18, bold=True,
            corner_pct=0.16)
    sl = make_diagram_slide(prs, page_num, tag or TAG_IC[2],
                            "Lufthansa Pricing: Solution", draw)
    _set_notes(sl, (
        "Here is the answer. The marginal cost is the same in both panels, "
        "because it is the same aircraft on the same LAX to Frankfurt leg. "
        "What differs is the demand curve. The direct flight is bought "
        "largely by people who need to be in Frankfurt, and they have no "
        "good substitute, so their demand is steep \u2014 less elastic. The "
        "connecting passenger is going to Istanbul and has a dozen ways to "
        "get there, so that demand is flatter. Set marginal revenue equal "
        "to marginal cost in each panel and the less elastic market comes "
        "out with the higher price. That is why the shorter trip costs "
        "more, and it is exactly the rule from the practice video: charge "
        "the higher price where demand is less elastic."))
    return sl


def s_ic_movies(prs, page_num):
    """[NVapp 19] — higher prices for new releases and blockbusters.

    2026-09-14, his rearrangement: his own two clippings stacked on the
    left, and the Odyssey card on the right given the film still he
    copied in as its BACKGROUND rather than the cream fill -- so the card
    reads as being about that film rather than as a generic callout.

    Every figure is press-reported and dated on the slide, per the
    provenance rule.  Sources: Variety, "'The Odyssey' Tickets Bonanza",
    19 July 2025; Washington Post, "To see 'The Odyssey' on Imax, fans pay
    premiums and brave 3 a.m. showings", 16 July 2026.
    """
    def draw(slide):
        _add_text(slide, Inches(0.70), Inches(1.58), Inches(11.90),
                  Inches(0.44),
                  "Higher prices for new movies and blockbusters",
                  size=26, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.LEFT)
        _add_media_image(slide, "NV_s42_a.png", left=Inches(0.42),
                         top=Inches(2.26), width=Inches(6.60))
        _add_media_image(slide, "NV_s42_b.png", left=Inches(0.28),
                         top=Inches(4.90), width=Inches(7.17))
        # the still is CROPPED to the card own aspect (940x659 px at
        # 5.55 x 3.30 in needs 7.6 % off the top and the bottom), then a
        # white scrim over it keeps the navy text readable
        # BOTH dimensions are set: a crop does not resize the frame, so
        # with width alone the still kept its uncropped 3.89" height and
        # hung 0.6" below the scrim, with the source line stranded on it.
        # 940 x 659 px less 7.6 % top and bottom is exactly 5.55 x 3.30.
        _add_media_image(slide, "NV_s42_bg.jpg", left=Inches(7.30),
                         top=Inches(2.26), width=Inches(5.55),
                         height=Inches(3.30),
                         crop=(0.0, 0.0, 0.076, 0.076))
        _scrim = _add_rounded_filled_box(
            slide, Inches(7.30), Inches(2.26), Inches(5.55), Inches(3.30),
            "", fill=WHITE, corner_pct=0.06, shadow=False)
        # _set_fill_alpha takes a PERCENT, not OOXML thousandths --
        # 78000 wrote an alpha 780x over the legal maximum and
        # PowerPoint then refused to open the deck at all
        _set_fill_alpha(_scrim, 78)
        _add_hierarchical_bullets(
            slide, Inches(7.50), Inches(2.42), Inches(5.15), Inches(3.00),
            [([("The Odyssey", {"bold": True}),
               ("  (Nolan, released 17 July 2026)", {})], 0,
              {"bullet_style": "none"}),
             ([("Universal sold IMAX 70 mm seats a YEAR ahead, on "
                "18 July 2025", {})], 0, {}),
             ([("95 % of them went in under 24 hours \u2014 most in the "
                "first hour", {})], 0, {}),
             ([("At the counter: $33 a seat at AMC Lincoln Square, "
                "against a $16.30 average US ticket in 2026", {})], 0, {}),
             ([("Resale reported at 400 % of face value, averaging $165, "
                "with one seat at $600", {})], 0, {})],
            size=15, line_spacing_pts=7)
        _add_text(slide, Inches(7.30), Inches(5.68), Inches(5.55),
                  Inches(0.30),
                  "Reported figures \u2014 Variety (19 July 2025) and the "
                  "Washington Post (16 July 2026)",
                  size=12, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)
        # his own closing line, with the interpuncts he typed as \cdot
        _add_rounded_filled_box(
            slide, Inches(0.62), Inches(6.45), Inches(12.13), Inches(0.52),
            "Premium for:  \u00b7 Format (IMAX 70 mm)   \u00b7 Date   "
            "\u00b7 Seat location (in many theatres)",
            fill=GOLD, text_color=NAVY, size=17, bold=True,
            corner_pct=0.16)
    return make_diagram_slide(prs, page_num, TAG_IC[2],
                              "Segment Pricing in Movies", draw)

def s_ic_versioning_concept(prs, page_num):
    """[NVapp 21] — the short in-class version."""
    bullets = [
        ("Indirect price discrimination", 0),
        ("Charge different price for different versions of the product", 0),
    ]
    return content_slide(prs, page_num, TAG_IC[3],
                         "Second Degree Price Discrimination: Versioning",
                         bullets, size=28, line_spacing_pts=20)


def s_ic_amc(prs, page_num):
    """[NV 35] — AMC's ticket tiers.  My originals are two .emf price
    tables; rebuilt as native PowerPoint tables."""
    def draw(slide):
        # HIS two price tables, placed as the .wmf images they are.  An
        # earlier pass rebuilt them as a native table with invented tiers
        # and prices, which is not acceptable: the .wmf text is not
        # machine-readable, so there is nothing to reproduce from.  Convert
        # them natively only once the real figures are to hand.
        _add_media_image(slide, "NV_s35_2_4b50cc3d.wmf", left=Inches(0.70),
                         top=Inches(1.70), width=Inches(11.9),
                         rounded=False, shadow=False)
        _add_media_image(slide, "NV_s35_3_185a113a.wmf", left=Inches(1.05),
                         top=Inches(5.05), width=Inches(11.2),
                         rounded=False, shadow=False)
    sl = make_diagram_slide(prs, page_num, TAG_IC[3], "Versioning: AMC", draw)
    _set_notes(sl, (
        "AMC sells the same seat in the same auditorium under several "
        "ticket types and lets the customer sort herself. What differs "
        "across the versions is not the film, the screen or the seat — it "
        "is how much the customer commits up front, and that is what "
        "separates the frequent movie-goer from the occasional one.\n"
        "NOTE FOR THE REBUILD: the two price tables on this slide are .wmf "
        "images carried over from the original deck, so they are not yet "
        "native PowerPoint tables. Converting them needs the current AMC "
        "prices, which change most years."))
    return sl


def s_ic_dolls(prs, page_num):
    """[NVapp 23] · [NV 36] — "Versioning: dolls".

    2026-09-09 (Nico): "keep this as was, including the three empty text
    boxes initially covering the prices and then being animated out."

    Those boxes are MASKS, not labels: white (bg1) 1.25 x 0.39" rectangles
    sitting exactly over the three Mattel price screenshots, removed by exit
    fades one per click.  `_extract_anim.py` gives the original reveal
    order, which is kept: $9.99 (Farmer) -> $13.99 (Robotics Engineer) ->
    $150.00 (Yves Saint Laurent), so the designer doll lands last.

    The masks are named MASK-1/2/3 in reveal order so `_animate.py` can
    select them, and they must be emitted AFTER the price images so they
    cover them.
    """
    def draw(slide):
        _add_media_image(slide, "NV_s36_3_5f2b3183.png", left=Inches(1.20),
                         top=Inches(1.85), width=Inches(10.9),
                         rounded=False, shadow=False)
        # the three price screenshots, under their dolls
        prices = [("NV_s36_4_beb5f5c8.png", 2.55, 0.98),
                  ("NV_s36_5_813fab59.png", 6.05, 1.43),
                  ("NV_s36_6_9e761089.png", 9.95, 1.13)]
        # 6.00, not 5.55: each screenshot carries an "IN STOCK" line above
        # its price, and at 5.55 that line was pasted across the doll NAMES
        # in the big image behind, cutting them into fragments.  Dropping
        # the three by 0.45" clears the name row (2026-09-14).
        for fname, left, w in prices:
            _add_media_image(slide, fname, left=Inches(left),
                             top=Inches(6.00), width=Inches(w),
                             rounded=False, shadow=False)
        # The masks, in Nico's reveal order: $9.99, $13.99, $150.00.
        #
        # 2026-09-14 (Nico): "Move the boxes that cover the prices and then
        # fade out upon click right over the prices -- that's the point
        # here!"  They were a row too high, sitting across the doll NAMES
        # while the prices below stayed visible the whole time.
        #
        # The boxes are now derived from where the PRICE actually renders:
        # each screenshot holds two bands of ink, a name fragment at the top
        # and the price underneath, and a row-profile of the three files
        # puts the price bands at y 6.10-6.31, 6.13-6.34 and 6.08-6.29, with
        # x 2.69-3.36, 6.29-7.27 and 10.12-10.92.  A common 6.03 / 0.38"
        # band covers all three with margin, so they read as one device.
        for i, (left, w) in enumerate(
                [(2.59, 0.87), (10.02, 1.00), (6.19, 1.18)], start=1):
            box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, int(Inches(left)), int(Inches(6.48)),
                int(Inches(w)), int(Inches(0.38)))
            box.fill.solid()
            box.fill.fore_color.rgb = WHITE
            box.line.fill.background()
            box.shadow.inherit = False
            box.name = "MASK-%d" % i
    return make_diagram_slide(prs, page_num, TAG_IC[3], "Versioning: Dolls",
                              draw)


def s_ic_wendys(prs, page_num):
    """[NV 38] — Wendy's surge-pricing backlash."""
    def draw(slide):
        # 580 x 691 is PORTRAIT: 6.0" wide would be 7.15" tall.  Size it
        # by the height the content area actually has.
        _add_media_image(slide, "NV_s38_3_fb8bfd15.png", left=Inches(1.05),
                         top=Inches(1.60), width=Inches(4.20))
        _add_hierarchical_bullets(
            slide, Inches(7.35), Inches(1.90), Inches(5.45), Inches(3.3),
            [([("Created immediate consumer backlash and a PR disaster: "
                "“charge higher prices when consumers are hungry”", {})], 0,
              {}),
             ([("Wendy’s issued a correction: now planning to introduce "
                "versioning (different menu options and discounts)", {})], 0,
              {}),
             ([("Direct price discrimination creates more backlash than "
                "indirect price discrimination", {})], 0, {})],
            size=19, line_spacing_pts=14)
        _add_convention_box(
            slide, Inches(7.35), Inches(5.35), Inches(5.45), Inches(1.35),
            prefix="Sen. Elizabeth Warren:  ",
            body="“It’s price gouging plain and simple, and American "
                 "families have had enough.”",
            size=16, anchor=MSO_ANCHOR.MIDDLE)
    sl = make_diagram_slide(prs, page_num, TAG_IC[3],
                            "Wendy’s Failed Attempt at Dynamic Pricing",
                            draw)
    _set_notes(sl, NOTES_NV.get(38, ""))
    return sl


def s_ic_airline_setup(prs, page_num):
    """[NVapp 24] · [NV 39] — the airline poll set-up."""
    def draw(slide):
        _add_text(slide, Inches(0.70), Inches(1.85), Inches(11.9),
                  Inches(1.0),
                  "How should United respond to Southwest starting to offer "
                  "flights on one of its routes?", size=26, bold=False,
                  color=NAVY, font="Calibri", align=PP_ALIGN.LEFT)
        # 558 x 405 -> 3.63" tall at 5.0" wide, ending at 6.78"
        _add_media_image(slide, "NV_s39_1_629f2784.gif", left=Inches(4.15),
                         top=Inches(3.15), width=Inches(5.0))
        _add_pollbreak_badge(slide)
    return make_diagram_slide(prs, page_num, TAG_IC[3],
                              "Airline Response to Low-Cost Carriers", draw)


def s_ic_airline_solution(prs, page_num):
    """[NVapp 26] · [NV 41] — the answer, and the backup link."""
    bullets = [
        ("Response by major US airlines to low-cost carriers (Southwest, "
         "Spirit):", 0),
        ("No “free” carry-on bag", 1),
        ("No advance seat assignments", 1),
        ("Don’t collect miles…", 1),
        ("Compete with low-cost competitors by offering a low-cost / "
         "low-quality version of your own product", 0),
        ("Not a new strategy. E.g.: Kodak Gold vs. FunFilm in the 1990s", 0),
    ]

    def extras(slide):
        # his own banner, at his position (2026-09-14).  It was written
        # last round but never wired in -- the call below did not pass
        # extras, so the image was silently dropped.
        # he nudged it up 0.09" by hand on 2026-09-15:
        # "adopt EXACTLY the slide as is now"
        _add_media_image(slide, "NV_s41_airline_banner.png",
                         left=Inches(0.68), top=Inches(2.44),
                         width=Inches(7.48), rounded=False)

    sl = content_slide(prs, page_num, TAG_IC[3],
                       "Versioning in Response to Low-Price Competition",
                       bullets, size=23, sub_size=21, line_spacing_pts=11,
                       extras=extras, notes=NOTES_NV.get(41))
    _add_pollbreak_badge(sl)
    return sl


def s_ic_podcast_pd(prs, page_num):
    """[NVapp 27] — the podcast questions on price discrimination."""
    bullets = [
        ("Historically: price discrimination by haggling", 0),
        ("Can it lead to perfect price discrimination?", 1),
        ("What’s the problem with haggling?", 1),
        ("How does the NYC restaurant try to price discriminate?", 0),
        ("Regular price: $18", 1),
        ("MC of a burger + drinks is $4.50", 1),
        ("How much is the coupon? Does that make sense?", 1),
        ("How do coupons relate to price discrimination?", 0),
    ]
    def extras(slide):
        # his own image, added by hand on 2026-09-14
        _add_media_image(slide, "NV_s58_img.png", left=Inches(6.87),
                         top=Inches(1.84), width=Inches(5.59))

    sl = content_slide(prs, page_num, TAG_IC[3],
                       "Podcast on Price Discrimination", bullets,
                         bullets_width=Inches(5.87), extras=extras,
                       size=23, sub_size=21, line_spacing_pts=10,
                       notes=NOTES_APP.get(27))
    _add_discussion_badge(sl)
    return sl


def s_ic_wsj_netflix(prs, page_num):
    """[NVapp 33] · [NV 48] — what is wrong with the statement?

    2026-09-14, his layout.  The cream box that paraphrased the sentence
    is gone: his headline clipping carries the WSJ wording itself, which
    is the whole point of the slide -- the class has to find the error in
    what the paper actually printed, not in a retyping of it.  (The
    paraphrase was there because the clipping is a .wmf and the first pass
    read it as unusable; it renders fine.)
    """
    def draw(slide):
        _add_media_image(slide, "NV_s48_wsj_article.wmf", left=Inches(0.60),
                         top=Inches(1.45), width=Inches(4.69))
        _add_text(slide, Inches(0.99), Inches(5.01), Inches(11.10),
                  Inches(0.60), "What\u2019s wrong with the following statement?",
                  size=26, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)
        _add_media_image(slide, "NV_s48_wsj_headline.wmf", left=Inches(2.30),
                         top=Inches(5.52), width=Inches(8.74),
                         rounded=False, shadow=False)
        _add_discussion_badge(slide)
    sl = make_diagram_slide(prs, page_num, TAG_IC[4],
                            "WSJ Article on Netflix", draw)
    _set_notes(sl, (
        "The flat fee does not depend on how much you watch, so a heavy "
        "viewer and a light viewer pay exactly the same increase in "
        "dollars. If anything the heavy viewer is getting the better deal "
        "per film. The statement confuses the total bill with the price "
        "per unit \u2014 under a flat fee there is no price per unit."))
    return sl

def s_ic_zoo_setup(prs, page_num):
    """[NVapp 37] — the zoo poll set-up.  MC = $4 here (the applications
    version Nico chose), not the main deck's MC = 0."""
    def draw(slide):
        _add_hierarchical_bullets(
            slide, Inches(0.80), Inches(1.90), Inches(11.7), Inches(1.1),
            [([("Demand for the average customer is:", {})], 0, {})],
            size=25, line_spacing_pts=0)
        _add_math_equation(
            slide, Inches(1.40), Inches(2.65), Inches(5.0), Inches(0.7),
            _omml_run("P") + _omml_text(" = %g − %g" % (ZOO_A, ZOO_B))
            + _omml_run("Q"), size_pt=28)
        # his own zoo photo, added by hand on 2026-09-14
        _add_media_image(slide, "NV_s54_zoo.jpg", left=Inches(7.56),
                         top=Inches(1.95), width=Inches(4.83))
        _add_hierarchical_bullets(
            slide, Inches(0.80), Inches(3.60), Inches(11.7), Inches(2.4),
            [([("Q: visits per year", {})], 0, {}),
             ([("Marginal cost per visit: MC = $%g" % ZOO_MC, {})], 0, {}),
             ([("What is the optimal pricing strategy?", {})], 0, {})],
            size=25, line_spacing_pts=18)
        _add_pollbreak_badge(slide)
    return make_diagram_slide(prs, page_num, TAG_IC[5],
                              "Pricing at a Local Zoo", draw)


# The zoo.  2026-09-15 (Nico): "make the numbers for zoo pricing more
# realistic."  His applications deck had P = 20 - 4Q with MC = $4, which
# put the first visit at $20 and a full annual pass at $50 -- roughly half
# what a US city zoo charges.  Doubling the price axis to P = 40 - 8Q with
# MC = $8 lands on real figures: a simple-pricing ticket of $24 (LA Zoo is
# $28 for an adult, the Bronx $41), a pure annual pass at $100 (LA Zoo
# membership is about $109, San Diego $129), and a two-part tariff of $64
# a year plus $8 a visit.
#
# The SLOPE doubles with the intercept on purpose.  Every quantity in the
# example is then unchanged -- 4 visits at MC, 5 at a zero price, 2 under
# simple pricing -- so the three charts keep their shape and all the
# arithmetic stays whole-dollar on camera.  Only the price axis rescales.
ZOO_A, ZOO_B = 40.0, 8.0
ZOO_MC = 8.0
ZOO_QMC = (ZOO_A - ZOO_MC) / ZOO_B                     # 4 visits
ZOO_FLAT = (ZOO_A - ZOO_MC) * ZOO_QMC / 2.0            # $64
ZOO_USAGE = ZOO_MC * ZOO_QMC                           # $32
ZOO_QMAX = ZOO_A / ZOO_B                               # 5 visits
# 2026-09-14 (Nico): "adopt all equations from my original in-class
# material slides for the zoo example."  His NVapp 40 sets MR = MC = 4, not
# MR = 0 -- and he is right: comparing simple pricing at MC = 0 against a
# two-part tariff and a flat fee that both charge MC = $4 compares three
# different worlds.  With his MC the three profits are $32, $16 and $30, and
# the comparison the three slides are making actually holds.
ZOO_SIMPLE_Q = (ZOO_A - ZOO_MC) / (2.0 * ZOO_B)        # 2 visits
ZOO_SIMPLE_P = ZOO_A - ZOO_B * ZOO_SIMPLE_Q            # $24
ZOO_SIMPLE_PROFIT = (ZOO_SIMPLE_P - ZOO_MC) * ZOO_SIMPLE_Q   # $32
ZOO_FF_FLAT = ZOO_A * ZOO_QMAX / 2.0                   # $100
ZOO_FF_COST = ZOO_MC * ZOO_QMAX                        # $40
assert (ZOO_FLAT, ZOO_USAGE) == (64.0, 32.0)
# his three profits: two-part $64 > pure flat fee $60 > simple $32
assert (ZOO_SIMPLE_PROFIT, ZOO_FF_FLAT - ZOO_FF_COST) == (32.0, 60.0)
assert ZOO_FLAT > ZOO_FF_FLAT - ZOO_FF_COST > ZOO_SIMPLE_PROFIT
# every quantity must stay a whole number of visits, or the charts stop
# landing on their own tick marks
assert all(float(v).is_integer() for v in
           (ZOO_QMC, ZOO_QMAX, ZOO_SIMPLE_Q))


def _zoo_fig(slide):
    # the price axis follows ZOO_A, so a rescale of the example does not
    # leave the demand line running off the top of the panel
    fig = SimpleFig(1.25, 6.55, 5.2, 4.9, xmax=ZOO_QMAX * 1.24,
                    ymax=ZOO_A * 1.15)
    _fig_axes(slide, fig, x_title="Q (visits per year)", y_title="P",
              label_size=17)
    return fig


def s_ic_zoo_solution(prs, page_num):
    """[NVapp 39] — the two-part tariff answer."""
    def draw(slide):
        fig = _zoo_fig(slide)
        _fig_poly(slide, fig, [(0, ZOO_A), (0, ZOO_MC), (ZOO_QMC, ZOO_MC)],
                  fill=GOLD, line=GOLD, alpha=26000, line_w=1.0)
        _fig_poly(slide, fig, [(0, ZOO_MC), (0, 0), (ZOO_QMC, 0),
                               (ZOO_QMC, ZOO_MC)],
                  fill=NAVY, line=NAVY, alpha=16000, line_w=1.0)
        _fig_line(slide, fig, (0, ZOO_A), (ZOO_QMAX, 0), color=DARKRED,
                  weight_pt=2.75)
        _fig_line(slide, fig, (0, ZOO_MC), (ZOO_QMAX * 1.18, ZOO_MC),
                  color=NAVY, weight_pt=2.25)
        _fig_curve_label(slide, fig, ZOO_QMAX * 1.07,
                         ZOO_MC + ZOO_A * 0.065, "MC", color=NAVY,
                         size=18)
        _fig_curve_label(slide, fig, 4.55, ZOO_A * 0.10, "D", color=DARKRED,
                         size=19, clear_of=((0, ZOO_A), (ZOO_QMAX, 0)))
        _fig_ylab(slide, fig, ZOO_A, "$%g" % ZOO_A)
        _fig_ylab(slide, fig, ZOO_MC, "$%g" % ZOO_MC)
        _fig_xlab(slide, fig, ZOO_QMC, "%g" % ZOO_QMC)
        _fig_xlab(slide, fig, ZOO_QMAX, "%g" % ZOO_QMAX)
        _fig_guide(slide, fig, (ZOO_QMC, ZOO_MC))
        _region_label(slide, fig, [(0, ZOO_A), (0, ZOO_MC),
                                   (ZOO_QMC, ZOO_MC)],
                      "Flat fee\n= $(%g − %g)·%g/2 = $%d"
                      % (ZOO_A, ZOO_MC, ZOO_QMC, int(ZOO_FLAT)),
                      size=15, dy=0.6)
        _region_label(slide, fig, [(0, ZOO_MC), (0, 0), (ZOO_QMC, 0),
                                   (ZOO_QMC, ZOO_MC)],
                      # his wording, 2026-09-15
                      "Revenues from usage fee = $%g·%g = $%d"
                      % (ZOO_MC, ZOO_QMC, int(ZOO_USAGE)),
                      size=14)
        # his own algebra, from NVapp 39 (2026-09-14)
        # 2026-09-15 (Nico): "move the formula to the left, where I put
        # it, on two lines and dark red."  His box is 3.20" wide, so the
        # implication wraps to a second line, in the chart's own empty
        # upper-left corner.
        _add_math_equation(
            slide, Inches(1.62), Inches(1.92), Inches(3.20), Inches(0.77),
            _omml_run("P") + _omml_text(" = %g \u2212 %g" % (ZOO_A, ZOO_B))
            + _omml_run("Q") + _omml_text("   \u21d2   ") + _omml_run("Q")
            + _omml_text(" = %g \u2212 %g" % (ZOO_QMAX, 1.0 / ZOO_B))
            + _omml_run("P"), size_pt=20, color=DARKRED)
            # 2026-09-14 (Nico): "for the agenda, use the shape-symbols."
            # Each line that names a shaded area carries that area's own
            # mark instead of a bullet dot, the way slide 10 does it.
        # "Two-part tariff" is a HEADER, not a bullet (2026-09-15, Nico:
        # "adopt my formatting"): it sits at the list's left edge with no
        # mark, and the marked rows hang under it.
        _add_text(slide, Inches(7.40), Inches(2.29), Inches(4.68),
                  Inches(0.40), "Two-part tariff:", size=20, bold=True,
                  color=NAVY, font="Calibri", align=PP_ALIGN.LEFT)
        _region_bullets(
            slide, Inches(7.55), 2.84, Inches(5.3), 3.00,
            [
             (("zfee",), "Flat fee  F = $%d" % int(ZOO_FLAT), 1),
             # the fee IS the MC line; the box under it is its revenue
             (("zline",), "Usage fee  P = $%d/visit" % int(ZOO_MC), 1),
             (("zuse",), "Revenue from usage fee = $%d" % int(ZOO_USAGE),
              1),
             (None, "Revenues from usage fee recover the variable costs "
              "of $%d" % int(ZOO_USAGE), 0),
             (None, "Profit: $%d per customer (assuming zero fixed cost)"
              % int(ZOO_FLAT), 0)],
            size=20, sub_size=19, mark_w=0.50, mark_h=0.46, gap=0.12)
    return make_diagram_slide(prs, page_num, TAG_IC[5],
                              "Zoo Pricing – Solution", draw)


def s_ic_zoo_simple(prs, page_num):
    """[NVapp 40] · [NV 58] — compared with simple pricing.

    2026-09-14: his own positive MC, not the MC = 0 this had, so that the
    three zoo slides compare a two-part tariff, simple pricing and a pure
    flat fee in ONE cost world -- which is the comparison they claim to be
    making.  At the 2026-09-15 figures MR = 40 − 16Q against MC = 8 gives
    Q* = 2, P* = $24 and a profit of $32, against $64 for the two-part
    tariff and $60 for the pure flat fee.
    """
    def draw(slide):
        fig = _zoo_fig(slide)
        # profit is the rectangle between P* and MC, not down to the axis
        _fig_poly(slide, fig, [(0, ZOO_SIMPLE_P), (0, ZOO_MC),
                               (ZOO_SIMPLE_Q, ZOO_MC),
                               (ZOO_SIMPLE_Q, ZOO_SIMPLE_P)],
                  fill=GOLD, line=GOLD, alpha=26000, line_w=1.0)
        _fig_line(slide, fig, (0, ZOO_MC), (ZOO_QMAX * 1.02, ZOO_MC),
                  color=NAVY, weight_pt=2.25, curve="MC")
        _fig_line(slide, fig, (0, ZOO_A), (ZOO_QMAX, 0), color=DARKRED,
                  weight_pt=2.75, curve="D")
        # MR: dashed, and the same dark red as the demand it comes from
        _fig_line(slide, fig, (0, ZOO_A), (ZOO_QMAX / 2.0, 0), color=DARKRED,
                  weight_pt=2.25, dash="dash", curve="MR")
        # these three heights are FRACTIONS of the price axis, not the
        # dollar figures they used to be: at the 2026-09-15 rescale a
        # label pinned to "2.0" sat on the x axis with the demand line
        # running straight through it (_check_labels, ONLINE).
        _fig_curve_label(slide, fig, 4.55, ZOO_A * 0.10, "D", color=DARKRED,
                         size=19, curve="D",
                         clear_of=((0, ZOO_A), (ZOO_QMAX, 0)))
        # past the MR line's OWN intercept at Q = a/2b, the way every
        # other curve label in the deck sits: at Q = 2.05 it straddled
        # the dashed drop line at Q* = 2, which struck through it.
        _fig_curve_label(slide, fig, ZOO_QMAX / 2.0 + 0.12,
                         ZOO_A * 0.06, "MR",
                         color=DARKRED, size=18, curve="MR",
                         clear_of=((0, ZOO_A), (ZOO_QMAX / 2.0, 0)))
        _fig_curve_label(slide, fig, ZOO_QMAX * 1.04,
                         ZOO_MC + ZOO_A * 0.045, "MC",
                         color=NAVY, size=16, curve="MC")
        _fig_ylab(slide, fig, ZOO_A, "$%g" % ZOO_A)
        _fig_ylab(slide, fig, ZOO_SIMPLE_P, "$%g" % ZOO_SIMPLE_P)
        _fig_ylab(slide, fig, ZOO_MC, "$%g" % ZOO_MC)
        _fig_xlab(slide, fig, ZOO_SIMPLE_Q, "%g" % ZOO_SIMPLE_Q)
        _fig_xlab(slide, fig, ZOO_QMAX, "%g" % ZOO_QMAX)
        _fig_guide(slide, fig, (ZOO_SIMPLE_Q, ZOO_SIMPLE_P))
        # MR cuts this rectangle diagonally, from (1, 12) to (2, 4), now
        # that MC is $4, and the rectangle is only 8 price units tall --
        # measured, the two-line caption cannot sit in it without a line
        # through it.  The area carries its AMOUNT and the list beside the
        # chart says what the amount is, which it already did.
        _region_label(slide, fig, [(0, ZOO_SIMPLE_P), (0, ZOO_MC),
                                   (ZOO_SIMPLE_Q * 0.60, ZOO_MC),
                                   (ZOO_SIMPLE_Q * 0.60, ZOO_SIMPLE_P)],
                      "$%d" % int(ZOO_SIMPLE_PROFIT), size=16)
        # his own placement, 2026-09-15: the demand / MR line moves to
        # the chart's empty upper-left corner, two lines, dark red
        _add_math_equation(
            slide, Inches(1.51), Inches(1.90), Inches(3.55), Inches(0.77),
            _omml_run("P")
            + _omml_text(" = %g − %g" % (ZOO_A, ZOO_B)) + _omml_run("Q")
            + _omml_text("   ⇒   ") + _omml_run("MR")
            + _omml_text(" = %g − %g" % (ZOO_A, 2 * ZOO_B))
            + _omml_run("Q"), size_pt=20, color=DARKRED)
        _add_math_equation(
            slide, Inches(7.44), Inches(3.11), Inches(5.3), Inches(0.60),
            _omml_run("MR") + _omml_text(" = ") + _omml_run("MC")
            + _omml_text(" = %g  ⇒  " % ZOO_MC)
            + _omml_sup(_omml_run("Q"), _omml_text("*"))
            + _omml_text(" = %g;  " % ZOO_SIMPLE_Q)
            + _omml_sup(_omml_run("P"), _omml_text("*"))
            + _omml_text(" = %g" % ZOO_SIMPLE_P), size_pt=20)
            # 2026-09-14 (Nico): "for the agenda, use the shape-symbols."
            # Each line that names a shaded area carries that area's own
            # mark instead of a bullet dot, the way slide 10 does it.
        # 2026-09-15 (Nico): a header above the list, the way slides 72
        # and 74 now carry one, and the marked row renamed to say what
        # the shaded rectangle IS
        _add_text(slide, Inches(7.40), Inches(2.57), Inches(4.68),
                  Inches(0.40), "Simple Pricing:", size=20, bold=True,
                  color=NAVY, font="Calibri", align=PP_ALIGN.LEFT)
        _region_bullets(
            slide, Inches(7.55), 3.99, Inches(5.3), 1.8,
            [(("zrev",), "Simple pricing profits:", 0),
             (None, "Profit: $%d per customer (assuming zero fixed cost)"
              % int(ZOO_SIMPLE_PROFIT), 1)],
            size=20, sub_size=19, mark_w=0.50, mark_h=0.46, gap=0.12)
        _add_ps_pointer(slide, top=Inches(6.53), label=PS_ZOO)
    return make_diagram_slide(prs, page_num, TAG_IC[5],
                              "Zoo Pricing – Comparison to Simple Pricing",
                              draw)


def s_ic_zoo_flatfee(prs, page_num):
    """[NVapp 41] — compared with a pure flat fee.

    With no usage fee the customer takes all 5 visits, which costs the zoo
    $8 x 5 = $40, so the $100 flat fee nets $60 — less than the $64 the
    two-part tariff earns.  That is the whole point of the comparison.
    """
    def draw(slide):
        fig = _zoo_fig(slide)
        _fig_poly(slide, fig, [(0, ZOO_A), (0, 0), (ZOO_QMAX, 0)],
                  fill=GOLD, line=GOLD, alpha=26000, line_w=1.0)
        # 2026-09-15 (Nico): "make the variable cost box dark red and
        # transparent."  It is cost, and on this slide part of it is not
        # covered -- the red says so before a word is read.
        _fig_poly(slide, fig, [(0, ZOO_MC), (0, 0), (ZOO_QMAX, 0),
                               (ZOO_QMAX, ZOO_MC)],
                  fill=DARKRED, line=DARKRED, alpha=16000, line_w=1.0)
        # "Make the triangle of the variable cost box that sits to the
        # right of D have a darker red."  Demand meets MC at ZOO_QMC, so
        # beyond that the customer values a visit at LESS than it costs:
        # the triangle (QMC, MC) - (QMAX, MC) - (QMAX, 0) is variable
        # cost the flat fee never recovers.  Its area is computed, not
        # eyeballed: (QMAX - QMC) x MC / 2.
        _fig_poly(slide, fig, [(ZOO_QMC, ZOO_MC), (ZOO_QMAX, ZOO_MC),
                               (ZOO_QMAX, 0)],
                  fill=DARKRED, line=DARKRED, alpha=42000, line_w=1.0)
        _fig_line(slide, fig, (0, ZOO_A), (ZOO_QMAX, 0), color=DARKRED,
                  weight_pt=2.75)
        _fig_line(slide, fig, (0, ZOO_MC), (ZOO_QMAX * 1.18, ZOO_MC),
                  color=NAVY, weight_pt=2.25)
        _fig_curve_label(slide, fig, ZOO_QMAX * 1.07,
                         ZOO_MC + ZOO_A * 0.065, "MC", color=NAVY,
                         size=18)
        # he moved D up the line on 2026-09-15, clear of the cost box
        _fig_curve_label(slide, fig, 4.02, ZOO_MC * 1.32, "D",
                         color=DARKRED, size=19,
                         clear_of=((0, ZOO_A), (ZOO_QMAX, 0)))
        _fig_ylab(slide, fig, ZOO_A, "$%g" % ZOO_A)
        _fig_ylab(slide, fig, ZOO_MC, "$%g" % ZOO_MC)
        _fig_xlab(slide, fig, ZOO_QMAX, "%g" % ZOO_QMAX)
        _region_label(slide, fig, [(0, ZOO_A), (0, 0), (ZOO_QMAX, 0)],
                      "Flat fee\n= $%g·%g/2 = $%d"
                      % (ZOO_A, ZOO_QMAX, int(ZOO_FF_FLAT)),
                      size=15, dy=2.6)
        _region_label(slide, fig, [(0, ZOO_MC), (0, 0), (ZOO_QMAX, 0),
                                   (ZOO_QMAX, ZOO_MC)],
                      "Variable cost = $%g·%g = $%d"
                      % (ZOO_MC, ZOO_QMAX, int(ZOO_FF_COST)),
                      size=14)
        # his own algebra, from NVapp 41 (2026-09-14), moved to the left
        # of the chart and set in dark red on 2026-09-15
        _add_math_equation(
            slide, Inches(1.53), Inches(2.15), Inches(2.45), Inches(0.44),
            _omml_run("P") + _omml_text(" = %g \u2212 %g" % (ZOO_A, ZOO_B))
            + _omml_run("Q"), size_pt=20, color=DARKRED)
            # 2026-09-14 (Nico): "for the agenda, use the shape-symbols."
            # Each line that names a shaded area carries that area's own
            # mark instead of a bullet dot, the way slide 10 does it.
        # his header, 2026-09-15 -- the same treatment as slides 72 / 73
        _add_text(slide, Inches(7.40), Inches(2.29), Inches(4.68),
                  Inches(0.40), "Flat-Fee Pricing:", size=20, bold=True,
                  color=NAVY, font="Calibri", align=PP_ALIGN.LEFT)
        _region_bullets(
            slide, Inches(7.55), 2.84, Inches(5.3), 2.90,
            [
             (("zfee",), "Flat fee  F = $%d" % int(ZOO_FF_FLAT), 1),
             (("zcost",), "Costs due to visits: $%d" % int(ZOO_FF_COST), 1),
             (None, "Profit: $%d per customer (assuming zero fixed cost)"
              % int(ZOO_FF_FLAT - ZOO_FF_COST), 0),
             (None, "Overall %g visits, since the price paid per visit is "
              % ZOO_QMAX +
              "now zero", 0)],
            size=20, sub_size=19, mark_w=0.50, mark_h=0.46, gap=0.12)
        # the callout for that triangle.  Text box in the clear air above
        # the demand line, arrow from its EDGE down to the wedge -- the
        # arrow-origin rule, so the leader cannot cross its own words.
        _xb = (Inches(4.30), Inches(4.52))
        _xw = (Inches(2.55), Inches(0.62))
        _add_text(slide, _xb[0], _xb[1], _xw[0], _xw[1],
                  "Excess variable cost not covered by the flat fee",
                  size=14, italic=True, bold=True, color=DARKRED,
                  font="Calibri", align=PP_ALIGN.CENTER)
        _arrow_from_box(
            slide, _xb, _xw,
            (fig.x((2 * ZOO_QMAX + ZOO_QMC) / 3.0),
             fig.y(ZOO_MC / 3.0)),
            color=DARKRED, weight_pt=1.75, head=True)
        # in the right-hand column, under the bullets -- a full-width bar
        # here would lie across the chart's own x axis
        _add_takeaway_bar(
            slide, "Two-part tariff: $%d   ·   pure flat fee: $%d"
            % (int(ZOO_FLAT), int(ZOO_FF_FLAT - ZOO_FF_COST)),
            left=Inches(7.55), width=Inches(5.3), top=Inches(5.95),
            text_color=NAVY, size=19, rounded=True, shadow=True)
    return make_diagram_slide(prs, page_num, TAG_IC[5],
                              "Zoo Pricing – Comparison to Flat Fee", draw)


def s_ic_ice_cream(prs, page_num):
    """[NVapp 43] · [NV 63] — block pricing on ice cream.

    1 scoop $4, 2 for $7, 3 for $8.50, so the marginal price of each
    successive scoop is $4.00, $3.00 and $1.50.

    Unlike the Walmart photo cards, those three points are NOT collinear —
    the drop is $1.00 and then $1.50.  So demand is drawn PIECEWISE through
    his own three points rather than as a fitted straight line: every block
    corner then sits exactly on the curve, which is what the exactness rule
    requires.  (A straight line through the first and last point would put
    the 2-scoop corner $0.25 off it.)
    """
    mc = 1.0
    steps = [(1.0, 4.0), (2.0, 3.0), (3.0, 1.5)]
    # the demand path: the choke price, his three points, then down to MC
    dpath = [(0.0, 5.0)] + steps + [(3.6, 0.55)]

    def draw(slide):
        fig = SimpleFig(1.35, 6.55, 5.2, 4.85, xmax=4.3, ymax=6.2)
        _fig_axes(slide, fig, x_title="Q (scoops per customer)",
                  y_title="P ($/scoop)", label_size=16)
        q_prev = 0.0
        for letter, (q, p) in zip("ABC", steps):
            _fig_poly(slide, fig, [(q_prev, p), (q_prev, mc), (q, mc),
                                   (q, p)],
                      fill=GOLD, line=GOLD, alpha=26000, line_w=1.0)
            _region_label(slide, fig, [(q_prev, p), (q_prev, mc), (q, mc),
                                       (q, p)], letter, size=18)
            _fig_ylab(slide, fig, p, "$%.2f" % p)
            _fig_xlab(slide, fig, q, "%d" % int(q))
            q_prev = q
        for p0, p1 in zip(dpath, dpath[1:]):
            _fig_line(slide, fig, p0, p1, color=DARKRED, weight_pt=2.75)
        _fig_line(slide, fig, (0, mc), (4.1, mc), color=NAVY, weight_pt=2.25)
        _fig_curve_label(slide, fig, 3.72, mc + 0.42, "MC", color=NAVY,
                         size=17)
        _fig_curve_label(slide, fig, 3.30, 1.25, "D", color=DARKRED, size=19)
        _fig_ylab(slide, fig, mc, "$1.00")
        _add_convention_box(
            slide, Inches(7.45), Inches(1.62), Inches(5.4), Inches(1.35),
            prefix="Purchase options:  ",
            body="1 scoop for $4 · 2 scoops for $7 · 3 scoops for $8.50",
            size=18, anchor=MSO_ANCHOR.MIDDLE)
        # 2026-09-14 (Nico): "adopt image and exact wording from my
        # slide 43 in in-class material."  His slide carries the
        # purchase options and the photo and nothing else, so the three
        # lines this had added are gone and his cone takes their place.
        _add_media_image(slide, "NVapp_s43_icecream.png",
                         left=Inches(8.10), top=Inches(3.30),
                         width=Inches(4.10))
    sl = make_diagram_slide(prs, page_num, TAG_IC[6],
                            "Block Pricing: Ice Cream", draw)
    _set_notes(sl, NOTES_NV.get(63, ""))
    return sl


def s_ic_costco_combined(prs, page_num):
    """[NV 59] — the SECOND Costco slide Nico asked for, with his 2024
    title and framing (the Video 4 copy keeps the tape's title)."""
    return s_costco_versioning(
        prs, page_num, tag=TAG_IC[5],
        title="Costco: Combination of Versioning and Two-Part Pricing")


# --------------------------------------------------------------------------
#  A12 [NEW – PG2 25–30] — the Disneyland integrative case
# --------------------------------------------------------------------------

def s_dis_open(prs, page_num):
    def draw(slide):
        _add_media_image(slide, "PG2_s25_rId2.jpeg", left=Inches(0.55),
                         top=Inches(2.55), width=Inches(12.2))
        _add_text(slide, Inches(0.70), Inches(1.70), Inches(11.9),
                  Inches(0.6), "Explain Disneyland’s pricing strategy",
                  size=28, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.LEFT)
        _add_discussion_badge(slide)
    return make_diagram_slide(prs, page_num, TAG_IC[2],
                              "Disneyland: Pricing Strategy", draw)


def _dis_slide(prs, page_num, tag, title, image, bullets, *, iw=7.4,
               ix=0.70, iy=1.75, tx=None):
    """`ix` / `iy` let one slide place its own picture where Nico put it
    (slide 99, 2026-09-15); `tx` pins the text column when the picture is
    narrower than the frame the column was sized against."""
    def draw(slide):
        _add_media_image(slide, image, left=Inches(ix), top=Inches(iy),
                         width=Inches(iw), rounded=False)
        _tl = Inches(tx) if tx is not None else (
            Inches(0.70) + Inches(iw) + Inches(0.35))
        _add_hierarchical_bullets(
            slide, _tl, Inches(2.15),
            Inches(12.85) - _tl,
            Inches(3.6), [([(b, {})], 0, {}) for b in bullets],
            size=19, line_spacing_pts=16)
    return make_diagram_slide(prs, page_num, tag, title, draw)


def s_dis_third_age(prs, page_num):
    return _dis_slide(
        prs, page_num, TAG_IC[2],
        "Disneyland: Third-Degree Price Discrimination",
        "PG2_s27_rId3.png",
        ["A child’s ticket is cheaper than an adult’s",
         "The park can verify age at the gate, so the segments cannot swap",
         "Under 3 is free — the segment that would not have come otherwise"])


def s_dis_third_date(prs, page_num):
    return _dis_slide(
        prs, page_num, TAG_IC[2],
        "Disneyland: Third-Degree Price Discrimination",
        "PG2_s28_rId2.png",
        ["The same ticket runs from about $104 to $224 by DATE",
         "Peak dates are when high-willingness-to-pay visitors can come",
         "Segmenting by the calendar, with no need to identify anyone"])


def s_dis_second(prs, page_num):
    return _dis_slide(
        prs, page_num, TAG_IC[3],
        "Disneyland: Second-Degree Price Discrimination",
        "PG2_s29_rId2.png",
        ["Lightning Lane comes in three versions: Multi, Single, Premier",
         "Visitors sort themselves — which is what makes it second degree",
         "The sorting is incentive-compatible (customers with high "
         "opportunity cost of time buy the Multi Pass)"])


def s_dis_twopart(prs, page_num):
    return _dis_slide(
        prs, page_num, TAG_IC[5], "Disneyland: Two-Part Tariff",
        "PG2_s26_rId3.png",
        ["Park admission is the flat fee",
         "Food, merchandise, etc are the usage fees",
         "The gate price captures the surplus; the in-park prices track the "
         "variable cost of serving you"], iw=8.2)


def s_dis_volume(prs, page_num):
    return _dis_slide(
        prs, page_num, TAG_IC[4], "Disneyland: Volume Pricing",
        # his own image, swapped in by hand on 2026-09-15, at his frame
        "NV_s99_magickey.png",
        ["The more you commit to, the less each visit costs",
         "The frequent visitor pays a lower price per visit"],
        iw=6.56, ix=0.76, iy=1.99, tx=8.45)


def s_dis_summary(prs, page_num):
    """One card closing the case on the module's own framework."""
    def draw(slide):
        _add_styled_table(
            slide, Inches(1.15), Inches(1.85), Inches(11.0), Inches(3.5),
            [["What Disneyland does", "Which strategy it is"],
             ["Child vs. adult tickets", "Third degree — segment pricing"],
             ["Prices that vary by date", "Third degree — intertemporal"],
             ["Lightning Lane tiers", "Second degree — versioning"],
             ["Gate price + in-park spending", "Two-part tariff"],
             ["Bundles of day tickets", "Volume / block pricing"]],
            font_size=18, header_size=18)
        # 2026-09-15 (Nico): "deleted the last box" -- the table already
        # makes the point that every strategy in the module shows up in
        # one firm, so the bar underneath only said it again.
    return make_diagram_slide(prs, page_num, TAG_IC[7],
                              "Disneyland: What Strategy Is Where", draw)


# --------------------------------------------------------------------------
#  Innovative / behavioral pricing, and the backup
# --------------------------------------------------------------------------

def s_ic_mad_optimist(prs, page_num):
    """[NV 66] — kept, per Nico 2026-09-09.  Only the verbatim second copy
    at his slides 69–70 is not repeated."""
    def draw(slide):
        _add_text(slide, Inches(0.70), Inches(1.58), Inches(11.9),
                  Inches(0.5),
                  "An interesting way to extract consumer surplus",
                  size=24, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.LEFT)
        # 1574 x 829 -> 4.64" tall at 8.8" wide, centred, ending at 6.84"
        _add_media_image(slide, "NV_s66_3_00445408.png", left=Inches(2.27),
                         top=Inches(2.20), width=Inches(8.80))
    sl = make_diagram_slide(prs, page_num, TAG_IC[7], "The Mad Optimist",
                            draw)
    _set_notes(sl, NOTES_NV.get(66, ""))
    return sl


def s_ic_mad_optimist_price(prs, page_num):
    """[NV 67]."""
    def draw(slide):
        _add_media_image(slide, "NV_s67_2_d6badc94.png", left=Inches(1.20),
                         top=Inches(1.70), width=Inches(10.9))
    sl = make_diagram_slide(prs, page_num, TAG_IC[7],
                            "The Mad Optimist – “Choose Your Price”", draw)
    _set_notes(sl, NOTES_NV.get(67, ""))
    return sl


def s_ic_rent_runway(prs, page_num):
    """[NVapp 47] · [NV 68]."""
    def draw(slide):
        _add_text(slide, Inches(0.70), Inches(1.58), Inches(11.9),
                  Inches(0.5), "Combination of block pricing and versioning",
                  size=24, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.LEFT)
        # 1276 x 616 -> 4.64" tall at 9.6" wide, centred, ending at 6.84"
        _add_media_image(slide, "NV_s68_3_c94534f1.png", left=Inches(1.87),
                         top=Inches(2.20), width=Inches(9.60))
    sl = make_diagram_slide(prs, page_num, TAG_IC[7], "Rent the Runway",
                            draw)
    _set_notes(sl, NOTES_NV.get(68, ""))
    return sl


def s_ic_economist(prs, page_num):
    """[NVapp 48] · [NV 71] — Decision D.

    The picture placeholder on my slide 71 is EMPTY; only the 16% / 84%
    labels survive.  The three subscription options are rebuilt as a native
    table from the figures the slide's own labels imply.

    2026-09-14 (Nico): "adopt this exactly from my original slide."  So
    the share of buyers is NOT a column of the table: on his slide 71 the
    16% and the 84% are two loose labels sitting beside the rows they
    belong to, and the print-only row deliberately carries no label at
    all.  That is the whole point of the slide -- the option nobody buys
    is the one doing the work -- and a "0%" cell in a tidy third column
    states it instead of letting the class find it.  The second bullet
    ("nobody takes print-only, so why offer it") was invented in the
    rebuild and gave the answer away; it is gone, and the question he
    actually asks is the only text on the right.
    """
    def draw(slide):
        # 2026-09-15 (Nico): "I reintroduced my original graph and
        # numbers."  So the native table is gone and his own screenshot of
        # the subscription page is back, at the size and place he set --
        # which is his original slide 71 geometry carried across.
        _add_media_image(slide, "NV_s89_economist_graph.jpg",
                         left=Inches(2.11), top=Inches(1.75),
                         width=Inches(4.93))
        # "Make the numbers and arrows dark red."  Each share sits level
        # with the option it belongs to and points at it: the $59
        # web-only line is 39 % down the screenshot and the print-and-web
        # line 80 % down, measured off the image itself, so the two
        # targets move with the picture rather than being eyeballed.
        img_top, img_h, img_right = 1.75, 3.84, 2.11 + 4.93
        for y, share, frac in ((3.11, " 16%", 0.389),
                               (4.72, " 84%", 0.795)):
            box_xy = (Inches(7.35), Inches(y))
            box_wh = (Inches(1.44), Inches(0.50))
            _add_text(slide, box_xy[0], box_xy[1], box_wh[0], box_wh[1],
                      share, size=21, bold=True, color=DARKRED,
                      font="Calibri", align=PP_ALIGN.LEFT,
                      anchor=MSO_ANCHOR.MIDDLE)
            # the head stops just SHORT of the screenshot rather than
            # landing on it: an arrow drawn across the image reads as a
            # line through the page it is annotating, and the format
            # audit flags it as a curve through a card.  At this
            # distance -- his labels sit 0.31" from the edge -- the
            # arrow is short by construction; it is a pointer at a row,
            # not a leader across the slide.
            _arrow_from_box(slide, box_xy, box_wh,
                            (Inches(img_right + 0.04),
                             Inches(img_top + frac * img_h)),
                            color=DARKRED, weight_pt=1.75, head=True,
                            gap=0.03)
        _add_hierarchical_bullets(
            slide, Inches(9.10), Inches(2.10), Inches(3.80), Inches(1.60),
            [([("What do you think about this pricing strategy?", {})], 0,
              {})],
            size=21, line_spacing_pts=18)
        _add_takeaway_bar(
            slide,
            "Customer attention focused on the comparison between the two "
            "$125 options",
            left=Inches(0.80), width=Inches(8.9), top=Inches(6.10),
            text_color=NAVY, size=19, rounded=True, shadow=True)
        _add_discussion_badge(slide)
    sl = make_diagram_slide(prs, page_num, TAG_IC[7],
                            "The Economist Pricing Example", draw)
    _set_notes(sl, NOTES_NV.get(71, ""))
    return sl


def s_ic_ariely(prs, page_num):
    """[NVapp 49] · [NV 72].

    2026-09-14 (Nico): "adopt this exactly from my original slide."  His
    slide 72 sits the figure LEFT of centre -- x 1.61 of a 10" canvas,
    16% in -- rather than centred, so the x is carried across as the same
    fraction (2.15 of 13.33").  It is then made as wide as the takeaway
    bar allows: 7.85" gives 4.35" of height and ends at 6.03", clear of
    the bar at 6.15".  Scaling his 6.16 x 3.42" by the full 1.333 would
    be 4.56" tall and would run into the bar, because the canvas grew
    only sideways.
    """
    def draw(slide):
        _add_media_image(slide, "NV_s72_3_38737db3.png", left=Inches(2.15),
                         top=Inches(1.68), width=Inches(7.85))
        _add_takeaway_bar(
            slide,
            "Customer attention drawn to the comparison between $59 and "
            "$125",
            top=Inches(6.15), text_color=NAVY, size=19, rounded=True,
            shadow=True)
    sl = make_diagram_slide(prs, page_num, TAG_IC[7],
                            "Dan Ariely’s Experiment", draw)
    _set_notes(sl, (
        "With all three options on the table, 84% took print-plus-web and "
        "only 16% the cheap web-only option. Take the useless print-only "
        "option away and the split flips: just 32% pay for the expensive "
        "one and 68% take the $59. The decoy does not have to sell a single "
        "copy to earn its place on the page.\n"
        "https://cxl.com/blog/pricing-experiments-you-might-not-know-but-"
        "can-learn-from/"))
    return sl


def s_ic_framing(prs, page_num):
    """[NV 73] — three framing devices, in HIS arrangement.

    2026-09-14 (Nico): "put the images in the same position as in my
    original slide."  Three equal panels in a row had squeezed the airline
    baggage table down to 3.95" wide, where its own prices are illegible.
    His slide 73 stacks the two small figures on the left and gives the
    baggage table the whole right half, which is the only arrangement in
    which it can be read.  Geometry is his, carried across from the 4:3
    canvas: x scaled by 13.333/10, sizes and y kept (the canvas is 7.5"
    tall either way, so nothing has to move down).

    His slide carries TWO captions, not three -- one under the left pair
    and one above the baggage table -- so the middle caption ("a charm
    price reads as a whole dollar cheaper than it is") is dropped.  It
    restated the left caption's point, and there is no room for a third
    caption between two stacked figures.
    """
    def draw(slide):
        # left column: the sale tags above the iPhone price ladder
        _add_media_image(slide, "NV_s73_2_8f303734.png", left=Inches(0.67),
                         top=Inches(1.91), width=Inches(2.70))
        _add_media_image(slide, "NV_s73_3_1e380d25.png", left=Inches(0.52),
                         top=Inches(4.14), width=Inches(3.11))
        _add_text(slide, Inches(0.67), Inches(6.25), Inches(3.70),
                  Inches(0.71),
                  "Prices ending in 9 or .99 draw focus to the previous "
                  "digit", size=16, bold=False, color=NAVY,
                  font="Calibri", align=PP_ALIGN.CENTER)
        # right half: the baggage table, its caption ABOVE it as on his
        # slide, where the caption and figure were one group
        _add_text(slide, Inches(6.32), Inches(1.91), Inches(4.39),
                  Inches(0.71),
                  "Adding a highly priced option draws attention to the "
                  "intermediate prices", size=16, bold=False, color=NAVY,
                  font="Calibri", align=PP_ALIGN.CENTER)
        _add_media_image(slide, "NV_s73_g5.2_75af4d04.png",
                         left=Inches(5.40), top=Inches(2.79),
                         width=Inches(5.75))
    sl = make_diagram_slide(prs, page_num, TAG_IC[7],
                            "Behavioral Insight: Price Framing", draw)
    _set_notes(sl, NOTES_NV.get(73, ""))
    return sl


def s_backup_seating(prs, page_num):
    """[NV 76] — the airline seating cartoon, full-bleed.

    A full-bleed backup figure carries NO top bar and therefore no tag,
    per the backup-slide exception.  The navy "← Back" button returns to
    the versioning slide that links here.
    """
    slide = _blank_slide(prs)
    # 600 x 1314 -> at 3.33" wide it is 7.29" tall, which is the most the
    # 7.5" canvas holds
    _add_media_image(slide, "NV_s76_4_2e8228f0.gif",
                     left=int((SLIDE_W - Inches(3.33)) / 2), top=Inches(0.10),
                     width=Inches(3.33), rounded=False, shadow=False)
    _add_text(slide, Inches(0.50), Inches(0.20), Inches(4.0), Inches(0.4),
              "Airplane Seating Chart", size=20, bold=True, color=NAVY,
              font="Calibri", align=PP_ALIGN.LEFT)
    _add_text(slide, Inches(0.50), Inches(0.62), Inches(4.0), Inches(0.32),
              "TheCooperReview.com", size=13, italic=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.LEFT)
    return slide


def link_backup(source_slide, backup_slide):
    """Wire the versioning slide to the seating-chart backup and back.

    The backup link takes the lower-right corner on the source slide, and
    the backup carries the deck-standard navy "← Back" pill in the same
    corner.  Called from main() once both slides exist.
    """
    # the Poll Break badge owns the bottom-right corner on this slide, so
    # the backup pill moves UP to sit directly above it, keeping its own
    # right-hand alignment (Teaching CLAUDE.md, poll-mark precedence)
    # 2026-09-14: he moved it up, clear of the banner he added below it
    _add_jump_pill(source_slide, backup_slide, left=Inches(10.05),
                   top=Inches(4.39), width=Inches(2.75),
                   label="More seating versions", size=15)
    # Teaching CLAUDE.md: a back button is a NAVY rounded pill with white
    # bold text.  _add_jump_pill defaults to a white pill with a navy
    # label and back=True does NOT change that, so the colours are passed
    # explicitly (caught by _audit_format.py, 2026-09-09).
    _add_jump_pill(backup_slide, source_slide, left=Inches(11.72),
                   top=Inches(6.60), width=Inches(1.55), label="← Back",
                   height=Inches(0.46), back=True, size=15,
                   fill=NAVY, text_color=WHITE)


# ==========================================================================
#  Assembly
# ==========================================================================

def _apply_note_fixes(prs):
    """Overwrite notes that were captured verbatim from an old deck and
    have since gone stale against the slide (2026-09-15).

    Unlike _apply_written_notes this REPLACES whatever is there: the whole
    point is that the existing note is wrong.  The guard is a substring
    that must appear somewhere on the slide, checked before the write, so
    a renumbering fails loudly rather than attaching a note to the wrong
    slide.
    """
    slides = list(prs.slides)
    n = 0
    for disp, (guard, note) in sorted(NOTE_FIXES.items()):
        if disp > len(slides):
            raise ValueError("note fix %d is past the end of the deck" % disp)
        slide = slides[disp - 1]
        blob = []

        def _walk(shapes):
            for sh in shapes:
                if sh.shape_type == 6:              # MSO_SHAPE_TYPE.GROUP
                    _walk(sh.shapes)
                    continue
                if sh.has_text_frame:
                    blob.append(sh.text_frame.text)

        _walk(slide.shapes)
        if guard not in "\n".join(blob):
            raise ValueError(
                "note fix %d expects %r somewhere on the slide and it is "
                "not there - the deck has been renumbered or the slide "
                "reworded, so re-key the fix" % (disp, guard))
        _set_notes(slide, note)
        n += 1
    return n


def _apply_written_notes(prs):
    """Fill the slides that carry no substantive note.

    Never overwrites a longer existing note, so Nico's own notes win.  The
    title guard in WRITTEN_NOTES is checked here: keying by display number
    is exactly what went wrong when two slides were inserted mid-deck, and
    a silent mis-attachment would be worse than no note at all.
    """
    slides = list(prs.slides)
    applied, skipped = 0, 0
    for disp, (want_title, note) in sorted(WRITTEN_NOTES.items()):
        if disp > len(slides):
            raise ValueError("written note %d is past the end of the deck"
                             % disp)
        slide = slides[disp - 1]
        # the action title is the text box at 0.28 / 0.55
        title = ""
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            try:
                x = sh.left / 914400.0
                y = sh.top / 914400.0
            except TypeError:
                continue
            if abs(x - 0.28) < 0.03 and abs(y - 0.55) < 0.04:
                title = sh.text_frame.text.strip()
        if not title.startswith(want_title):
            raise ValueError(
                "written note %d expects a slide titled %r but found %r "
                "- the deck has been renumbered, so re-key the notes"
                % (disp, want_title, title[:60]))
        # the logistics note names the exam date, and the slide draws it
        # from the course calendar -- fill it from the same constants so
        # the two cannot disagree (2026-09-15)
        note = note.replace(
            "{final_date}",
            _FINAL_DATE.strftime("%A, %B %d").replace(" 0", " "))
        note = note.replace("{final_slot}", _FINAL_SLOT)
        existing = ""
        if slide.has_notes_slide:
            tf = slide.notes_slide.notes_text_frame
            existing = (tf.text or "").strip() if tf is not None else ""
        if len(existing) >= len(note):
            skipped += 1
            continue
        _set_notes(slide, note)
        applied += 1
    return applied, skipped


def _raise_corner_marks(prs):
    """Move every bottom-right corner mark to the END of its slide's spTree.

    2026-09-14 (Nico): "All slides that have a box in the bottom right (such
    as 'Discussion'): make sure this box is in the front, so it covers the
    line at the bottom of the slide."  The rule was already written down for
    the Poll Break badge and the practice-video box -- draw it LAST, after
    the footer -- but it was enforced one call site at a time, and eleven of
    them call the helper from inside the body callback, which runs BEFORE
    the footer.  The footer rule then cuts a line straight across the badge
    (visible on slide 22's "Discussion").

    Enforcing it here instead makes the call site irrelevant.  The test is
    geometric and deliberately narrow: a filled box or a group that STRADDLES
    the footer rule at y 7.15" is the corner-mark family and nothing else --
    the rule itself, the gold strip, the footer text and the page number all
    sit entirely on one side of it.
    """
    RULE_Y = Inches(7.15)
    moved = 0
    for slide in prs.slides:
        spTree = slide.shapes._spTree
        for sh in list(slide.shapes):
            tagname = sh._element.tag.rsplit("}", 1)[-1]
            if tagname not in ("grpSp", "sp"):
                continue
            if sh.top is None or sh.height is None:
                continue
            if sh.height < Inches(0.25):
                continue
            if not (sh.top < RULE_Y < sh.top + sh.height):
                continue
            el = sh._element
            spTree.remove(el)
            spTree.append(el)
            moved += 1
    return moved


def _strip_unused_layouts(prs):
    """Drop the python-pptx template's unused slide layouts.

    Course rule: ONE slide master for the deck, with the template defaults
    stripped.  Every slide here is built on the blank layout, so the other
    ten are dead weight that invite drift the moment a slide is inserted
    by hand.  Modules 3 and 4 both ship with exactly one layout; Module 6
    was still carrying eleven (2026-09-09).
    """
    used = {s.slide_layout.part.partname for s in prs.slides}
    dropped = 0
    for master in prs.slide_masters:
        id_lst = master.slide_layouts._sldLayoutIdLst
        for sldLayoutId in list(id_lst):
            rId = sldLayoutId.rId
            part = master.part.related_part(rId)
            if part.partname in used:
                continue
            id_lst.remove(sldLayoutId)
            master.part.drop_rel(rId)
            dropped += 1
    return dropped


def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    n = 1                      # page numbers are live fields; n is the cache

    # ---- Video 2 · Introduction to Module 6 --------------------------
    # 2026-09-23 (Nico): the introduction video's card opens the deck, and
    # the front matter plus the descriptive overview form its block.  The
    # overview used to open the Simple vs. Complex video.
    intro_video_card(prs); n += 1
    slide_title(prs)
    n += 1
    slide_logistics(prs, n); n += 1
    slide_roadmap(prs, n); n += 1
    make_m6_outline(prs, n, descriptions=True,
                    tag="Module 6 · Video 1 · Agenda"); n += 1

    # ======================================================================
    #  2026-09-14 (Nico): "Integrate those fully within the regular slide
    #  deck, so they don't show up twice anymore.  Put those where they were
    #  in my original Module 6 slides, or if they come from PG, put them
    #  where PG had them."
    #
    #  So the separate In-Class Examples block is GONE.  Each application
    #  slide now sits at its own place in the main flow, keyed to the NV
    #  number in its docstring (his original 76-slide deck's order), with the
    #  PG imports at PG's own relative positions.  Eleven slides that were
    #  built twice -- once on video, once in the applications block -- are
    #  built once.
    #
    #  A slide that is NOT in the taped video keeps the four-level
    #  `Module 6 · In Class · Examples · <topic>` tag even though it now sits
    #  inside a video's block, so a student can still tell at a glance what
    #  was taped and what is class-only.  Most of the s_ic_* builders carry
    #  that tag themselves; the rest are passed TAG_IC[k] here.
    # ======================================================================

    # ---- Video 2 · Simple vs. Complex Pricing ------------------------
    video_card(prs, 0); n += 1
    make_m6_outline(prs, n, highlight_idx=0); n += 1
    s_dilemma(prs, n); n += 1
    s_ic_podcast_dilemma(prs, n); n += 1          # [NVapp 8]
    s_netflix_simple(prs, n); n += 1              # [NV 7]
    s_netflix_lose(prs, n); n += 1                # [NV 8]
    # PG1 3 and PG1 4 sit at the FRONT of PG's deck and ahead of its
    # Dum-Dums slide, which is where they go here too
    s_ic_news(prs, n, title="In the News: Hospital Mergers",
              image="PG1_s03_rId2.png",
              questions=["What does recent research show about prices for "
                         "surgery, intensive care, and ER visits after "
                         "hospital mergers?"],
              notes="Larger hospital systems negotiate higher pay from "
                    "insurers, and that shows up as higher premiums for "
                    "workers. Market power is the first condition on our "
                    "list for complex pricing, and a merger is one way to "
                    "buy it.\nSource: Wall Street Journal."); n += 1
    s_ic_news(prs, n, title="In the News: Restaurant Dynamic Pricing",
              image="PG1_s04_rId2.png",
              questions=["Describe this pricing strategy. Is it more "
                         "profitable? Why do some chains reject it?",
                         "Do customers like dynamic pricing? Is there a "
                         "generational difference?"],
              notes="Surge pricing on a menu is the restaurant trade's "
                    "version of what we saw Wendy's attempt. Worth putting "
                    "the Wendy's backlash beside it.\nSource: Wall Street "
                    "Journal."); n += 1
    s_complex(prs, n); n += 1                     # [NV 9]
    s_dumdums(prs, n); n += 1                     # [PG1 20]
    s_three_degrees(prs, n); n += 1               # [NV 10]

    # ---- Video 3 · First-Degree Price Discrimination -----------------
    video_card(prs, 1); n += 1
    make_m6_outline(prs, n, highlight_idx=1); n += 1
    # 2026-09-14 (Nico): "slide 18: delete."  The question version and
    # the video slide below it carried the same title and nearly the same
    # bullets once the two blocks were merged.
    s_first_degree(prs, n); n += 1                # [NV 12] answers it
    s_fd_mc0(prs, n); n += 1                      # [NV 14]
    s_fd_mcpos(prs, n); n += 1                    # [NV 15]
    # 2026-09-14 (Nico): "Slide 23: Delete."  The video copy of the
    # dystopian slide carried the same Telegraph clipping and the same two
    # bullets as the in-class question version above it, so with the two
    # blocks merged it was a straight duplicate.
    s_ic_dystopian_q(prs, n); n += 1              # [NVapp 12] · [NV 13]
    s_ic_uber(prs, n); n += 1                     # [NV 16]

    # ---- Video 4 · Segment Pricing -----------------------------------
    video_card(prs, 2); n += 1
    make_m6_outline(prs, n, highlight_idx=2); n += 1
    s_segment_concept(prs, n); n += 1             # [NV 18]
    s_segment_chart(prs, n); n += 1               # [NV 19]
    s_student_discounts(prs, n); n += 1           # [NV 20]
    s_local_discounts(prs, n); n += 1             # [NV 21]
    s_timing(prs, n); n += 1                      # [NV 22]
    s_ways_to_segment(prs, n); n += 1             # [NV 23]
    s_recall_segment(prs, n); n += 1
    s_orbitz(prs, n); n += 1                      # [PG1 41]
    # NV 24 - 29, the Lufthansa pair and the Lodine sequence.  He asked for
    # 24-26 in this block on 2026-09-14, so those three keep the VIDEO tag;
    # the poll and its solution are class-only and keep the In-Class one.
    # There is now ONE Lodine set-up slide, and it carries the Poll Break
    # badge -- building a badgeless copy for the video and a badged copy for
    # the poll is exactly the duplication this pass removes.
    s_ic_airline_puzzle(prs, n, tag=TAG[2]); n += 1       # [NV 24]
    s_ic_lufthansa(prs, n, tag=TAG[2]); n += 1            # [NV 25]
    s_ic_lodine_setup(prs, n); n += 1                     # [NV 26]
    make_stub(prs, n, TAG_IC[2], "Lodine Poll",
              "PollEverywhere slide — spliced verbatim with its notes and "
              "tags by _splice_media.py."); n += 1        # [NV 27]
    s_ic_lodine_solution(prs, n); n += 1                  # [NV 28]
    # 2026-09-14 (Nico): "slide 42: delete that slide."  The analytical
    # treatment of segment pricing is now carried by slides 34 and 35 and
    # by the practice video, so this repeat of it went.
    s_ic_movies(prs, n); n += 1                           # [NVapp 19]

    # 2026-09-14 (Nico): "slides 34-36: move those just before slide 43."
    # The two-market analysis, its graphical answer and the BMW caveats
    # now close Video 3, after the examples rather than in front of them.
    s_two_markets(prs, n); n += 1
    s_bmw_graphical(prs, n); n += 1               # [PG1 40]
    s_bmw_challenges(prs, n); n += 1              # [NV 30]

    # ---- Video 5 · Versioning and Coupons ----------------------------
    video_card(prs, 3); n += 1
    make_m6_outline(prs, n, highlight_idx=3); n += 1
    # 2026-09-14 (Nico): "slide 45: delete" -- same duplication as 18.
    s_versioning_concept(prs, n); n += 1          # [NV 32] answers it
    s_versioning_chart(prs, n); n += 1            # [NV 33]
    s_product_line(prs, n); n += 1                # [NV 34]
    s_costco_versioning(prs, n); n += 1           # [NV 59], kept here
    # 2026-09-14 (Nico): "slide 50: delete."
    s_ic_dolls(prs, n); n += 1                    # [NV 36]
    s_more_versioning(prs, n); n += 1             # [NV 37]
    s_ic_wendys(prs, n); n += 1                   # [NV 38]
    s_ic_airline_setup(prs, n); n += 1            # [NV 39]
    make_stub(prs, n, TAG_IC[3], "Airline-Response Poll",
              "PollEverywhere slide — spliced verbatim by "
              "_splice_media.py."); n += 1                # [NV 40]
    _airline_solution = s_ic_airline_solution(prs, n); n += 1   # [NV 41]
    s_coupons(prs, n); n += 1                     # [NV 42]
    s_ic_podcast_pd(prs, n); n += 1               # [NVapp 27]

    # ---- Video 6 · Flat Fee Pricing ----------------------------------
    video_card(prs, 4); n += 1
    make_m6_outline(prs, n, highlight_idx=4); n += 1
    s_context(prs, n); n += 1                     # [NV 44]
    s_advanced_pricing(prs, n); n += 1            # [NV 45]
    s_netflix_flatfee(prs, n); n += 1             # [NV 46]
    s_netflix_strategy(prs, n); n += 1            # [NV 47]
    s_ic_wsj_netflix(prs, n); n += 1              # [NV 48]
    s_flat_rate_wrong(prs, n); n += 1             # [NV 49]
    s_classpass(prs, n); n += 1                   # [PG2 16]

    # ---- Video 7 · Two-Part Tariffs ----------------------------------
    video_card(prs, 5); n += 1
    make_m6_outline(prs, n, highlight_idx=5); n += 1
    s_two_part_chart(prs, n); n += 1              # [NV 51]
    s_two_part_summary(prs, n); n += 1            # [NV 52]
    s_zipcar(prs, n); n += 1                      # [NV 53]
    s_ic_zoo_setup(prs, n); n += 1                # [NVapp 37]
    make_stub(prs, n, TAG_IC[5], "Zoo Poll",
              "PollEverywhere slide — spliced verbatim by "
              "_splice_media.py."); n += 1                # [NVapp 38]
    s_ic_zoo_solution(prs, n); n += 1             # [NVapp 39]
    s_ic_zoo_simple(prs, n); n += 1               # [NV 58]
    s_ic_zoo_flatfee(prs, n); n += 1              # [NVapp 41]
    s_ic_costco_combined(prs, n); n += 1          # [NV 59], the second one

    # ---- Video 8 · Block Pricing -------------------------------------
    video_card(prs, 6); n += 1
    make_m6_outline(prs, n, highlight_idx=6); n += 1
    s_block_concept(prs, n); n += 1               # [NV 61]
    s_walmart_blocks(prs, n); n += 1              # [NV 62]
    s_ic_ice_cream(prs, n); n += 1                # [NV 63]

    # ---- Video 9 · Summary of Pricing Strategies ---------------------
    # Decision B: the tape has no agenda slide here, but one is added so
    # all eight blocks follow card -> agenda -> content.
    video_card(prs, 7); n += 1
    make_m6_outline(prs, n, highlight_idx=7); n += 1
    s_summary_tree(prs, n); n += 1                # [NV 64]

    # ---- summary closer, outside every video block -------------------
    s_module_summary(prs, n); n += 1              # [NV 74]

    # ---- innovative and behavioral pricing, in class only ------------
    divider(prs, "Some Examples for Innovative Pricing Strategies",
            "Only if we have time"); n += 1
    s_ic_mad_optimist(prs, n); n += 1             # [NV 66]
    s_ic_mad_optimist_price(prs, n); n += 1       # [NV 67]
    s_ic_rent_runway(prs, n); n += 1              # [NV 68]
    s_ic_economist(prs, n); n += 1                # [NV 71]
    s_ic_ariely(prs, n); n += 1                   # [NV 72]
    s_ic_framing(prs, n); n += 1                  # [NV 73]

    # ---- backup ------------------------------------------------------
    divider(prs, "Backup"); n += 1
    _backup = s_backup_seating(prs, n); n += 1    # [NV 76]
    link_backup(_airline_solution, _backup)

    # 2026-09-14 (Nico): "Slides 88-94: move those to the end, right after
    # 'Backup'."  The Disneyland integrative case [PG2 25-30] is the
    # only-if-there-is-time exercise, so it now closes the deck rather than
    # sitting between the summary tree and the module summary.
    s_dis_open(prs, n); n += 1
    s_dis_third_age(prs, n); n += 1
    s_dis_third_date(prs, n); n += 1
    s_dis_second(prs, n); n += 1
    s_dis_twopart(prs, n); n += 1
    s_dis_volume(prs, n); n += 1
    s_dis_summary(prs, n); n += 1

    apply_symbol_subscripts(prs)
    n_raised = _raise_corner_marks(prs)
    n_notes, n_kept = _apply_written_notes(prs)
    n_fixed = _apply_note_fixes(prs)
    n_dropped = _strip_unused_layouts(prs)
    prs.save(str(OUT))
    print("%s: %d slides (%d layouts dropped, %d notes written, "
          "%d of mine kept, %d corner marks raised)"
          % (OUT.name, len(prs.slides._sldIdLst), n_dropped, n_notes,
             n_kept, n_raised))


if __name__ == "__main__":
    main()
