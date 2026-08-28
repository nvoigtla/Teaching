# ==========================================================================
#  _build_Module4.py — build script for "Module 4 - Revised.pptx"
#
#  Competitive Markets and Market Interventions, 83 slides per the approved
#  outline in "Module 4 - Revised - outline.md".  The six PollEverywhere
#  slides are positional stubs until the _splice_media.py pass.
#
#  Primitives come from _m4_helpers.py (carved out of Module 1's build
#  script by _make_helpers.py); speaker notes of the original deck come
#  verbatim from _m4_notes.py.  This script is the SOURCE OF TRUTH for the
#  deck — hand edits in PowerPoint are ported back into it.
#
#  Pipeline (rerunnable, Module 7 pattern):
#      python _build_Module4.py            -> Module 4 - Revised.pptx
#      python _splice_media.py             -> polls, verbatim + notes + tags
#      python _group_pass.py               -> box+text / figure+shade groups
#      python _animate.py all apply        -> fade builds
# ==========================================================================

from pathlib import Path

from pptx import Presentation

import _m4_helpers as _H
from _m4_helpers import *                                        # noqa: F401,F403
from _m4_helpers import (
    CREAM, DARKRED, DIM, GOLD, GOLD_W, GRAY, MARGIN, NAVY, RULE, RULE_W,
    SLIDE_H, SLIDE_W, WHITE, FADED, BLUE_PED,
    Inches, Pt, PP_ALIGN, MSO_ANCHOR, MSO_SHAPE, RGBColor, qn,
    _add_arrow, _add_arrow_shape, _add_convention_box, _add_drop_shadow,
    _add_graphicframe_shadow, _add_hierarchical_bullets, _add_math_equation,
    _add_media_image, _add_mixed_textbox, _add_outlined_box,
    _add_pollbreak_badge, _add_ps_pointer, _add_rect,
    _add_rounded_filled_box, _add_slidenum_field, _add_styled_table,
    _add_takeaway_bar, _add_text, _apply_picture_style, _blank_slide,
    _draw_action_title, _draw_footer, _draw_poll_pill, _draw_top_bar_tc,
    _fig_axes, _fig_curve_label, _fig_guide, _fig_line, _fig_xlab, _fig_ylab,
    _omml_frac, _omml_run, _omml_sub, _omml_sup, _omml_text,
    _set_notes, _title_case, apply_symbol_subscripts,
    content_slide, make_diagram_slide, make_stub, SimpleFig,
)
from _m4_notes import NOTES

# Two of my original notes quote the OLD figures, so they would contradict
# the slide after the 2026-08-28 switch.  They are overridden here rather
# than edited in _m4_notes.py, which stays a verbatim capture.
NOTES = dict(NOTES)
NOTES[41] = (
    "A price taker\u2019s supply curve. We can trace out the supply curve "
    "of the firm.\nAt P = 210, you maximize profits by producing 187.5 "
    "tons.\nAt P = 400, you expand more.\nBottom line: the MC curve is "
    "the firm\u2019s short-run supply curve. We can draw the supply curve."
)
NOTES[19] = (
    "Market price: $400 per ton is about $20 per cwt, which is the USDA "
    "season-average price for fresh-market cabbage in 2020. Recent years "
    "run higher \u2014 $31.10 per cwt in 2024.\n"
    "https://www.agmrc.org/commodities-products/vegetables/cabbage"
)

OUT = Path(__file__).parent / "Module 4 - Revised.pptx"

# --------------------------------------------------------------------------
#  The OMML helpers build raw XML by string concatenation and do NOT escape
#  their argument, so a literal "<" in a comparison ("P < AVC") silently
#  produces an invalid slide part.  Module 4 is full of such comparisons, so
#  the two text-producing helpers are wrapped here rather than fixing them
#  at every call site.
# --------------------------------------------------------------------------

_omml_text_raw = _omml_text
_omml_run_raw = _omml_run


def _xml_escape(text):
    return (text.replace('&', '&amp;').replace('<', '&lt;')
                .replace('>', '&gt;'))


def _omml_text(text, **kw):                                  # noqa: F811
    return _omml_text_raw(_xml_escape(text), **kw)


def _omml_run(text, **kw):                                   # noqa: F811
    return _omml_run_raw(_xml_escape(text), **kw)

# --------------------------------------------------------------------------
#  Deck constants
# --------------------------------------------------------------------------

_H.FOOTER_TEXT = ("Management 405  ·  Module 4  ·  "
                  "Competitive Markets and Market Interventions")
FOOTER_TEXT = _H.FOOTER_TEXT

DECK_TITLE = "Competitive Markets and Market Interventions"
DECK_SUB = "Module 4"

# Top-bar tags.  One constant per outline item, derived from the item
# titles below, so renaming an item renames every tag that uses it.
TAG_LOG = "Module 4 · Logistics"
TAG_ROADMAP = "Module 4 · Course Roadmap"
TAG_OUTLINE = "Module 4 · Outline"
TAG_SUMMARY = "Module 4 · Summary"

# The poll slides are PLACEHOLDERS for now (2026-08-28, Nico): the two
# cabbage activities still offer the retired answers, so the live slides
# are deliberately not spliced in.  Each placeholder title carries
# "++UPDATE" so it is obvious at a glance which slides are still pending.
POLL_MARK = "++UPDATE"
STUB_POLL = ("PollEverywhere placeholder — re-key the activity in PollEv, "
             "then run _splice_media.py")

# --------------------------------------------------------------------------
#  The Yi-family cost function.
#
#  These five constants drive the cost table, the TC chart, both worked
#  solutions, the profit rectangle and the supply-curve traces.  EVERY
#  printed figure below is derived from them through _num() and the
#  TC/MC/AVC omml builders — nothing is typed twice.
#
#  2026-08-28 (Nico): switched from my original 30,000 + 40Q + 0.2Q² at
#  P = 230/160 to these figures, which are the ones in MW's deck.  They are
#  the more defensible pair: ATC at Q* is $358/ton against the $348/ton in
#  the UC Ventura County cost study my speaker notes already cite, and a
#  $400/ton price is USDA's 2020 season average.  See
#  _cabbage_numbers_check.md.  The two PollEverywhere activities are keyed
#  to the OLD answers and have to be re-entered in the PollEv account.
# --------------------------------------------------------------------------

TFC = 60000          # total fixed cost, $
B_LIN = 135          # linear term of TC, $ per ton
B_QUAD = 0.2         # quadratic term of TC
P_HIGH = 400         # first market price, $ per ton
P_LOW = 210          # market price after the Chinese imports, $ per ton


def tc(q):
    return TFC + B_LIN * q + B_QUAD * q * q


def mc(q):
    return B_LIN + 2 * B_QUAD * q


def avc(q):
    return B_LIN + B_QUAD * q


def q_star(p):
    return (p - B_LIN) / (2 * B_QUAD)


Q_HIGH = q_star(P_HIGH)          # 662.5
Q_LOW = q_star(P_LOW)            # 187.5


def _num(x, dp=None):
    """Money / quantity as the slides print it: thousands separated, and
    only as many decimals as the value actually has (max 2)."""
    if dp is None:
        if abs(x - round(x)) < 1e-9:
            dp = 0
        elif abs(x * 10 - round(x * 10)) < 1e-9:
            dp = 1
        else:
            dp = 2
    return "{:,.{}f}".format(x, dp)


def _omml_tc():
    """TC = F + bQ + cQ²  as OMML."""
    return (_omml_text('TC') + _omml_text(' = %s + %s' % (_num(TFC),
                                                          _num(B_LIN)))
            + _omml_run('Q') + _omml_text(' + %s' % _num(B_QUAD))
            + _omml_sup(_omml_run('Q'), _omml_text('2')))


def _omml_mc_expr():
    """b + 2cQ  as OMML (the marginal-cost expression, no left-hand side)."""
    return (_omml_text('%s + %s' % (_num(B_LIN), _num(2 * B_QUAD)))
            + _omml_run('Q'))


def _omml_avc_expr():
    """b + cQ  as OMML."""
    return (_omml_text('%s + %s' % (_num(B_LIN), _num(B_QUAD)))
            + _omml_run('Q'))

# --------------------------------------------------------------------------
#  Module outline — my original structure, with the three sub-topics kept
#  under "Perfect Competition".  Each row is (label, title, description,
#  is_sub).  The label is what goes in the gold circle.
#
#  No coverage pills and no video references in this round (2026-08-28,
#  Nico): the deck is built clean first, and the split into videos and
#  in-class material is a later pass.
# --------------------------------------------------------------------------

M4_OUTLINE = [
    ("1", "Introduction to market structures",
     "Where a firm sits between price taker and price setter", False),
    ("2", "Perfect competition",
     "Many small sellers, an identical product, and free entry", False),
    ("2a", "Profit maximization of a price taker in the short run",
     "How much to produce, and when to shut down", True),
    ("2b", "Firm-level and market supply",
     "Why the marginal cost curve is the supply curve", True),
    ("2c", "Long-run competitive equilibrium",
     "Entry and exit drive economic profits to zero", True),
    ("3", "Market distortions and regulations",
     "Who wins, who loses, and what is lost outright", False),
    ("4", "Externalities",
     "Costs and benefits that land on people outside the deal", False),
]

TAG = {i: "Module 4 · " + _title_case(row[1][0].upper() + row[1][1:])
       for i, row in enumerate(M4_OUTLINE)}

TAG_INTRO = TAG[0]      # Introduction to Market Structures
TAG_PC = TAG[1]         # Perfect Competition
TAG_SR = TAG[2]         # Profit Maximization of a Price Taker in the Short Run
TAG_SUPPLY = TAG[3]     # Firm-Level and Market Supply
TAG_LR = TAG[4]         # Long-Run Competitive Equilibrium
TAG_DIST = TAG[5]       # Market Distortions and Regulations
TAG_EXT = TAG[6]        # Externalities

# how far a shaded (one-line) row is nudged down inside its reserved
# two-row box, so single-line rows sit centred — see Teaching CLAUDE.md
DIM_DROP = Inches(0.19)


# ==========================================================================
#  Outline / agenda slides
# ==========================================================================

def make_m4_outline(prs, page_num, *, tag=None, title="Outline of Module 4",
                    descriptions=False, highlight_idx=None,
                    highlight_set=None, ps_pointer=False):
    """Module outline in the format converged on for Modules 1 – 3: a gold
    circle carrying the item label, the item title beside it in bold navy,
    and a one-line grey description underneath.

    Every item RESERVES the description row, so item positions are
    pixel-identical on every agenda slide.  The description shows only for
    the current topic(s), or for all of them when ``descriptions=True``.
    Section agendas band the current item in cream and shade the rest.

    Module 4 adds sub-items ("2a", "2b", "2c"): they are indented, set one
    step smaller, and their circle is a touch narrower, so the hierarchy of
    my original outline survives.
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, tag or TAG_OUTLINE)
    _draw_action_title(slide, title)

    hi = set()
    if highlight_idx is not None:
        hi.add(highlight_idx)
    if highlight_set:
        hi.update(highlight_set)
    if descriptions:
        hi = set(range(len(M4_OUTLINE)))

    n_items = len(M4_OUTLINE)
    top = Inches(1.42)
    bottom = Inches(7.02)
    gap = Inches(0.11) if n_items <= 6 else Inches(0.07)
    PITCH_MAX = Inches(0.91)          # the Module 2 pitch
    LAST_ROW_MAX = Inches(6.22)

    def _row_y(pitch, i):
        block = pitch * n_items - gap
        y0 = top + max(0, (bottom - top - block) // 2)
        return y0 + i * pitch

    pitch = PITCH_MAX
    while pitch > Inches(0.60) and (
            _row_y(pitch, n_items - 1) > LAST_ROW_MAX
            or pitch * n_items - gap > bottom - top):
        pitch -= 4572                  # 0.005" steps
    content = pitch - gap
    title_h = int(content * 0.525)
    desc_h = content - title_h
    y = int(_row_y(pitch, 0))

    for i, (label, item, desc, is_sub) in enumerate(M4_OUTLINE):
        lit = descriptions or i in hi
        circ_d = Inches(0.50) if is_sub else Inches(0.58)
        circ_x = Inches(1.60) if is_sub else Inches(1.15)
        text_x = Inches(2.35) if is_sub else Inches(2.05)
        t_size = 22 if is_sub else 25
        d_size = 20 if is_sub else 22

        if not descriptions and i in hi:
            band = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, int(Inches(0.90)),
                int(y - Inches(0.06)), int(Inches(12.15)),
                int(title_h + desc_h + Inches(0.10)))
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
            MSO_SHAPE.OVAL, int(circ_x), int(y + Inches(0.02)),
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
        run.font.size = Pt(19 if len(label) > 1 else 25)
        run.font.bold = True
        run.font.color.rgb = NAVY if lit else DIM
        run.font.name = "Calibri"

        rows = [([(_title_case(item[0].upper() + item[1:]),
                   {'bold': True, 'size': t_size,
                    'color': NAVY if lit else DIM})], 0,
                 {'bullet_style': 'none', 'space_before_pts': 0})]
        if i in hi:
            rows.append(([(desc, {'size': d_size, 'color': GRAY})], 0,
                         {'bullet_style': 'none', 'space_before_pts': 0}))
        _add_hierarchical_bullets(
            slide, text_x, y if lit else int(y + DIM_DROP),
            int(Inches(12.85) - text_x), title_h + desc_h,
            rows, size=t_size, line_spacing_pts=0)
        y = int(y + pitch)

    _draw_footer(slide, FOOTER_TEXT, page_num)
    if ps_pointer:
        _add_ps_pointer(slide, top=Inches(6.68), label="Problem Set 3")
    return slide


# ==========================================================================
#  FRONT MATTER — slides 1 – 4
# ==========================================================================

def slide_01_title(prs):
    """Deck title slide: centred, no top bar, no page number."""
    slide = _blank_slide(prs)
    # the deck name runs to two lines, so the box is tall enough for both
    _add_text(slide, MARGIN, Inches(1.95), RULE_W, Inches(1.70),
              DECK_TITLE, size=50, bold=True, color=NAVY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _add_text(slide, MARGIN, Inches(3.72), RULE_W, Inches(0.80),
              DECK_SUB, size=40, bold=True, color=GOLD,
              font="Calibri", align=PP_ALIGN.CENTER)
    _add_rect(slide, (SLIDE_W - Inches(4.0)) // 2, Inches(4.62),
              Inches(4.0), Inches(0.06), GOLD)
    _add_text(slide, MARGIN, Inches(4.98), RULE_W, Inches(0.45),
              "Management 405", size=26, bold=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _add_text(slide, MARGIN, Inches(5.52), RULE_W, Inches(0.45),
              "Prof. Nico Voigtländer  ·  UCLA Anderson", size=22,
              color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
    _add_rect(slide, MARGIN, Inches(7.15), RULE_W, Inches(0.02), RULE)
    _add_rect(slide, MARGIN, Inches(7.135), GOLD_W, Inches(0.05), GOLD)
    _add_media_image(slide, "NV_s01_4_67f56784.jpg",
                     left=Inches(0.60), top=Inches(6.42), width=Inches(2.20),
                     rounded=False, shadow=False)
    return slide


def slide_02_logistics(prs, page_num):
    """Logistics.  Exam dates refreshed from the course calendar
    (2026-08-28, Nico): the Fall 2026 exam period is 11 – 13 December, and
    the slide must make clear that there is ONE 3.5-hour window, the same
    for everyone, with the exact date announced later."""
    bullets = [
        ("Coffee & Econ office hours restart next week", 0),
        ("Final exam period: December 11 – 13, 2026", 0),
        ("One 3.5-hour window — the same window for everyone", 1),
        ("The exact date and time will be announced later", 1),
        ("Online, open book, open notes; covers all material", 1),
    ]
    return content_slide(prs, page_num, TAG_LOG, "Some Logistics", bullets,
                         size=28, sub_size=24, line_spacing_pts=18)


def make_m4_roadmap(prs, page_num, *, tag=None):
    """Course roadmap in the Module-3 standard diamond format.  Module 4
    sits in box 4 ("Markets, Pricing, and Strategy"), so that box is navy
    and carries the gold 'we are here' arrow."""

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

        arrow_w, arrow_h = Inches(0.6), Inches(0.5)
        arrow_left = bot_x - arrow_w - Inches(0.12)
        arrow_top = bot_y + (box_h - arrow_h) // 2
        _add_arrow_shape(slide, arrow_left, arrow_top, arrow_w, arrow_h,
                         direction="right", fill=GOLD)
        _add_text(slide, arrow_left - Inches(1.55),
                  bot_y + (box_h - Inches(0.32)) // 2,
                  Inches(1.45), Inches(0.32), "we are here",
                  size=16, italic=True, bold=True, color=GOLD,
                  font="Calibri", align=PP_ALIGN.RIGHT)

    return make_diagram_slide(prs, page_num, tag or TAG_ROADMAP,
                              "Agenda for the Class", draw)


# ==========================================================================
#  1 · INTRODUCTION TO MARKET STRUCTURES — slides 5 – 7
# ==========================================================================

def slide_06_market_structures(prs, page_num):
    """The four basic market structures.  My original is six separate
    one-column tables butted together; this is ONE native table.  Monopoly
    products read "Unique" in quotation marks (2026-08-28, Nico)."""
    rows = [
        ["", "Perfect\nCompetition", "Monopolistic\nCompetition",
         "Oligopoly", "Monopoly"],
        ["Number of firms", "Many", "Many", "Few", "One"],
        ["Type of products sold", "Identical", "Differentiated",
         "Identical or\nDifferentiated", "“Unique”"],
        ["Barriers to entry", "None", "None", "Some", "Many"],
    ]
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_INTRO)
    _draw_action_title(slide, "The Four Basic Market Structures")
    tbl_w = Inches(12.2)
    first_col = Inches(3.0)
    other = (tbl_w - first_col) // 4
    _add_styled_table(
        slide, (SLIDE_W - tbl_w) // 2, Inches(2.10), tbl_w, Inches(3.40),
        rows, col_widths=[first_col] + [other] * 4,
        row_heights=[Inches(1.00)] + [Inches(0.80)] * 3,
        font_size=20, header_size=20)
    _add_takeaway_bar(
        slide, "More firms, more similar products, easier entry "
               "→ more competition",
        top=Inches(6.10), width=Inches(9.6), height=Inches(0.55),
        left=(SLIDE_W - Inches(9.6)) // 2, fill=GOLD, text_color=NAVY,
        size=19, bold=True, rounded=True, shadow=True)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[6])
    return slide


def slide_07_market_power(prs, page_num):
    """The price-taker → price-searcher spectrum, with one illustration per
    market structure and the 'TODAY' band over the left end."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_INTRO)
    _draw_action_title(
        slide, "Extent of Market Power: Price Takers vs. Price Searchers")

    x0, x1 = Inches(0.55), Inches(12.78)
    cols = [Inches(0.55), Inches(3.60), Inches(6.70), Inches(9.75)]
    col_w = Inches(2.95)

    _add_text(slide, x0, Inches(1.44), Inches(3.6), Inches(0.40),
              "Least market power", size=17, italic=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.LEFT)
    _add_text(slide, Inches(9.2), Inches(1.44), Inches(3.6), Inches(0.40),
              "Most market power", size=17, italic=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.RIGHT)
    _add_arrow(slide, (x0, Inches(1.98)), (x1, Inches(1.98)),
               color=NAVY, weight_pt=2.5, head=True)

    labels = ["Perfect\nCompetition", "Monopolistic\nCompetition",
              "Oligopoly", "Monopoly"]
    for cx, lab in zip(cols, labels):
        _add_text(slide, cx, Inches(2.16), col_w, Inches(0.75), lab,
                  size=22, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)
    for sx in (Inches(3.45), Inches(6.55), Inches(9.60)):
        _add_rect(slide, sx, Inches(2.14), Inches(0.03), Inches(0.80), RULE)

    _add_text(slide, cols[0], Inches(3.02), col_w, Inches(0.35),
              "(Price takers)", size=18, italic=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _add_text(slide, cols[1], Inches(3.02),
              cols[3] + col_w - cols[1], Inches(0.35),
              "(Price searchers)", size=18, italic=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.CENTER)

    # the block we are in today: a gold outline round the left column, with
    # the TODAY pill straddling its bottom edge
    band = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, int(Inches(0.40)), int(Inches(1.90)),
        int(Inches(3.20)), int(Inches(4.00)))
    try:
        band.adjustments[0] = 0.04
    except Exception:
        pass
    band.fill.background()
    band.line.color.rgb = GOLD
    band.line.width = Pt(2.0)
    band.shadow.inherit = False

    # one illustration per structure, all on a common baseline
    ph_top = Inches(3.60)
    _add_media_image(slide, "NV_s07_10_f9f010a8.jpg",
                     left=Inches(1.12), top=ph_top, height=Inches(1.70))
    _add_media_image(slide, "NV_s07_g1.1_1f7061fb.png",
                     left=Inches(4.14), top=ph_top, height=Inches(1.70))
    _add_media_image(slide, "NV_s07_g12.1_a0b89503.jpg",
                     left=Inches(7.60), top=Inches(3.65), width=Inches(1.15),
                     rounded=False, shadow=False)
    _add_media_image(slide, "NV_s07_g12.2_3444df7e.png",
                     left=Inches(7.68), top=Inches(4.52), width=Inches(1.00),
                     rounded=False, shadow=False)
    _add_media_image(slide, "NV_s07_13_4dc4c7c4.jpg",
                     left=Inches(10.42), top=ph_top,
                     height=Inches(1.70), rounded=False, shadow=False)

    captions = ["Commodity markets", "Restaurants in WeHo",
                "Wide-body aircraft", "Municipal utilities"]
    for cx, cap in zip(cols, captions):
        _add_text(slide, cx, Inches(5.42), col_w, Inches(0.32), cap,
                  size=13, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)

    _add_rounded_filled_box(slide, Inches(1.15), Inches(5.68),
                            Inches(1.70), Inches(0.44), "TODAY",
                            fill=GOLD, text_color=NAVY, size=16,
                            corner_pct=0.30)
    _add_takeaway_bar(
        slide, "Today: firms with no market power at all",
        top=Inches(6.45), width=Inches(7.2), height=Inches(0.55),
        left=(SLIDE_W - Inches(7.2)) // 2, fill=GOLD, text_color=NAVY,
        size=19, bold=True, rounded=True, shadow=True)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[7])
    return slide


# ==========================================================================
#  2 · PERFECT COMPETITION — slides 8 – 11
# ==========================================================================

def slide_09_price_taker(prs, page_num):
    """Characteristics of a price-taker market.  Rebuilt as a native
    two-column table (adopted from MW slide 9) instead of two facing text
    columns."""
    rows = [
        ["", "Perfect competition — price takers"],
        ["Number of actual or\npotential competitors", "Many small sellers"],
        ["Product differentiation", "None"],
        ["Entry conditions", "No barriers"],
        ["Profit potential",
         "Short run: can be positive, negative or zero\n"
         "Long run: zero economic profit"],
        ["Examples",
         "Agricultural commodities; currency markets; unskilled labor"],
    ]
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_PC)
    _draw_action_title(slide, "Perfect Competition: The Price Taker")
    tbl_w = Inches(11.8)
    _add_styled_table(
        slide, (SLIDE_W - tbl_w) // 2, Inches(1.85), tbl_w, Inches(4.60),
        rows, col_widths=[Inches(4.3), Inches(7.5)],
        row_heights=[Inches(0.62), Inches(0.90), Inches(0.72),
                     Inches(0.72), Inches(1.00), Inches(0.80)],
        font_size=20, header_size=22)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[8])
    return slide


def slide_10_market_and_firm(prs, page_num):
    """Two panels: the market sets P*, and the individual firm faces that
    price as a horizontal demand curve d = MR."""

    def draw(slide):
        mk = SimpleFig(1.40, 6.35, 3.50, 3.55, xmax=10.0, ymax=10.0)
        fm = SimpleFig(7.90, 6.35, 3.80, 3.55, xmax=10.0, ymax=10.0)

        # panel titles sit above the y-axis titles, centred on their panel
        _add_text(slide, Inches(1.40), Inches(1.50), Inches(3.5),
                  Inches(0.40), "Market", size=22, bold=True, color=NAVY,
                  font="Calibri", align=PP_ALIGN.CENTER)
        _add_text(slide, Inches(7.90), Inches(1.50), Inches(3.8),
                  Inches(0.40), "Firm (price taker)", size=22, bold=True,
                  color=NAVY, font="Calibri", align=PP_ALIGN.CENTER)

        for fig in (mk, fm):
            _fig_axes(slide, fig, x_title="Quantity", y_title="Price ($)",
                      label_size=17)

        # market: S up, D down, equilibrium at (5, 5)
        _fig_line(slide, mk, (0.6, 1.0), (9.2, 9.0), color=NAVY,
                  weight_pt=2.5)
        _fig_line(slide, mk, (0.6, 9.0), (9.2, 1.0), color=NAVY,
                  weight_pt=2.5)
        _fig_curve_label(slide, mk, 9.35, 9.0, "S", size=20)
        _fig_curve_label(slide, mk, 9.35, 1.1, "D", size=20)
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, int(mk.x(4.9) - Inches(0.07)),
            int(mk.y(5.0) - Inches(0.07)), int(Inches(0.14)),
            int(Inches(0.14)))
        dot.fill.solid()
        dot.fill.fore_color.rgb = GOLD
        dot.line.color.rgb = NAVY
        dot.line.width = Pt(1.0)
        dot.shadow.inherit = False
        _add_arrow(slide, (mk.x(4.9), mk.y(5.0)), (fm.x(0.0), fm.y(5.0)),
                   color=GOLD, weight_pt=2.0, head=True, dash="dash")
        _fig_line(slide, mk, (0.0, 5.0), (4.9, 5.0), color=GRAY,
                  weight_pt=1.25, dash="dash")
        _fig_ylab(slide, mk, 5.0, "P*", size=18, bold=True)
        _add_text(slide, Inches(2.05), Inches(2.72), Inches(1.9),
                  Inches(0.60), "Market\nequilibrium", size=15,
                  italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)

        # firm: horizontal d = MR at the market price
        _fig_line(slide, fm, (0.0, 5.0), (9.4, 5.0), color=NAVY,
                  weight_pt=3.0)
        _fig_ylab(slide, fm, 5.0, "P*", size=18, bold=True)
        _add_text(slide, Inches(8.20), Inches(3.85), Inches(4.3),
                  Inches(0.40), "Demand curve  d  =  Marginal Revenue (MR)",
                  size=17, bold=True, color=NAVY, font="Calibri")

    slide = make_diagram_slide(prs, page_num, TAG_PC,
                               "Perfect Competition: The Market and the Firm",
                               draw)
    _set_notes(slide, NOTES[10])
    return slide


def slide_11_farmers(prs, page_num):
    """Germany 2024: farmers protest subsidy cuts because they cannot pass
    costs on.  The minister's line is the quote callout."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_PC)
    _draw_action_title(slide, "Why Farmers Cannot Pass On Their Costs")

    _add_hierarchical_bullets(
        slide, MARGIN, Inches(1.60), Inches(10.4), Inches(1.30),
        [("Germany cut its farm subsidies in 2024", 0),
         ("Some farmers stood to lose up to €10,000 a year", 0)],
        size=26, line_spacing_pts=14)

    _add_media_image(slide, "NV_s11_4_e363d6f5.png",
                     left=Inches(0.50), top=Inches(3.05), width=Inches(5.60))
    _add_text(slide, Inches(0.50), Inches(6.48), Inches(5.60), Inches(0.32),
              "Farmer protests in Berlin, January 2024. "
              "Photo: The Guardian", size=12, italic=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.CENTER)

    _add_convention_box(
        slide, Inches(6.80), Inches(3.60), Inches(6.05), Inches(2.35),
        runs=[("“For farmers there is a structural problem: they cannot "
               "pass on their production costs, because the prices are not "
               "made by them.”", {'italic': True, 'size': 19}),
              ("\n", {}),
              ("— Germany's Minister of Economic Affairs, 2024",
               {'bold': True, 'size': 17})],
        corner_pct=0.10, size=19)
    _add_media_image(slide, "NV_s11_5_d816cef7.png",
                     left=Inches(11.10), top=Inches(1.55), width=Inches(1.75),
                     rounded=False, shadow=False)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


# ==========================================================================
#  Figure primitives added for Module 4
# ==========================================================================

def _fig_curve(slide, fig, fn, x0, x1, *, color=NAVY, weight_pt=2.5,
               segments=4, dash=None):
    """An editable freeform through fn over [x0, x1], built from a handful
    of cubic Bézier anchors (Catmull-Rom → Bézier), not a dense polyline,
    so "Edit Points" shows a few handles.  The shape gets its own TIGHT
    bounding box hugging just this curve, per the course chart rules.
    """
    n = segments
    pts = [(x0 + (x1 - x0) * i / n, None) for i in range(n + 1)]
    pts = [(xv, fn(xv)) for xv, _ in pts]
    dev = [(int(fig.x(xv)), int(fig.y(yv))) for xv, yv in pts]

    # Catmull-Rom control points
    ext = [dev[0]] + dev + [dev[-1]]
    beziers = []
    for i in range(1, len(ext) - 2):
        p0, p1, p2, p3 = ext[i - 1], ext[i], ext[i + 1], ext[i + 2]
        c1 = (p1[0] + (p2[0] - p0[0]) / 6.0, p1[1] + (p2[1] - p0[1]) / 6.0)
        c2 = (p2[0] - (p3[0] - p1[0]) / 6.0, p2[1] - (p3[1] - p1[1]) / 6.0)
        beziers.append((c1, c2, p2))

    xs = [p[0] for p in dev] + [c[0] for b in beziers for c in b[:2]]
    ys = [p[1] for p in dev] + [c[1] for b in beziers for c in b[:2]]
    left, top = int(min(xs)), int(min(ys))
    w, h = max(int(max(xs) - left), 1), max(int(max(ys) - top), 1)

    def rel(p):
        return int(p[0] - left), int(p[1] - top)

    parts = ["<a:moveTo><a:pt x=\"%d\" y=\"%d\"/></a:moveTo>" % rel(dev[0])]
    for c1, c2, p in beziers:
        parts.append(
            "<a:cubicBezTo>"
            "<a:pt x=\"%d\" y=\"%d\"/><a:pt x=\"%d\" y=\"%d\"/>"
            "<a:pt x=\"%d\" y=\"%d\"/></a:cubicBezTo>"
            % (rel(c1) + rel(c2) + rel(p)))

    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, w, h)
    spPr = shp._element.spPr
    old = spPr.find(qn('a:prstGeom'))
    dash_xml = ('<a:prstDash val="%s"/>' % dash) if dash else ''
    geom = _H.ET.fromstring(
        '<a:custGeom xmlns:a="%s"><a:avLst/><a:gdLst/><a:ahLst/>'
        '<a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
        '<a:pathLst><a:path w="%d" h="%d">%s</a:path></a:pathLst>'
        '</a:custGeom>' % (_H.A_NS, w, h, "".join(parts)))
    old.addnext(geom)
    spPr.remove(old)
    shp.fill.background()
    shp.shadow.inherit = False
    # Go through python-pptx's line API, NOT a bare SubElement: inside
    # <a:spPr> the order is xfrm -> geometry -> fill -> ln -> effectLst,
    # and an <a:ln> appended after the effect list is ignored outright
    # (the curve then renders in the theme's thin blue default).
    shp.line.color.rgb = color
    shp.line.width = Pt(weight_pt)
    ln = spPr.find(qn('a:ln'))
    # inside <a:ln> the order is fill -> dash -> join
    if dash_xml:
        ln.append(_H.ET.fromstring(
            dash_xml.replace('<a:prstDash',
                             '<a:prstDash xmlns:a="%s"' % _H.A_NS)))
    ln.append(_H.ET.fromstring('<a:round xmlns:a="%s"/>' % _H.A_NS))
    return shp


def _xlab_n(slide, fig, xv, label, *, size=17, bold=True, w=0.66,
            color=NAVY):
    """A NARROW x-axis label.  The stock _fig_xlab centres a 1.0" box on
    the tick, so two quantities less than an inch apart overlap; these
    charts mark two or three quantities side by side."""
    return _add_text(slide, fig.x(xv) - Inches(w / 2),
                     Inches(fig.b + 0.06), Inches(w), Inches(0.30), label,
                     size=size, bold=bold, italic=True, color=color,
                     font="Calibri", align=PP_ALIGN.CENTER)


def _fig_dot(slide, fig, xv, yv, *, d=Inches(0.15), fill=GOLD, line=NAVY):
    dot = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, int(fig.x(xv) - d // 2), int(fig.y(yv) - d // 2),
        int(d), int(d))
    dot.fill.solid()
    dot.fill.fore_color.rgb = fill
    dot.line.color.rgb = line
    dot.line.width = Pt(1.0)
    dot.shadow.inherit = False
    return dot


def _fig_region(slide, fig, x_lo, x_hi, y_lo, y_hi, *, fill=CREAM,
                alpha=None, line=None):
    """A shaded rectangle in logical coordinates (profit / loss areas)."""
    left, top = int(fig.x(x_lo)), int(fig.y(y_hi))
    w = int(fig.x(x_hi)) - left
    h = int(fig.y(y_lo)) - top
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1.25)
    shp.shadow.inherit = False
    if alpha is not None:
        sf = shp._element.spPr.find(qn('a:solidFill'))
        clr = sf.find(qn('a:srgbClr'))
        clr.append(_H.ET.fromstring(
            '<a:alpha xmlns:a="%s" val="%d"/>' % (_H.A_NS, int(alpha))))
    return shp


def _yi_badge(slide, label="Yi Family Example"):
    """The small gold-outlined corner tab that marks a slide as part of the
    running Yi-family example (my original deck carries it as loose text)."""
    return _add_outlined_box(
        slide, Inches(9.95), Inches(6.52), Inches(3.10), Inches(0.50),
        label, line=GOLD, text_color=NAVY, fill=WHITE, size=16,
        line_w=1.5, rounded=True, shadow=True, corner_pct=0.28)


# ==========================================================================
#  2a · PROFIT MAXIMIZATION IN THE SHORT RUN — slides 12 – 41
# ==========================================================================

def slide_13_yi_family(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "Example: The Yi Family")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.85), Inches(7.0),
        Inches(3.60),
        [("A South Korean family moves from California to rural Arkansas "
          "to start a farm", 0),
         ("Their goal: grow Korean produce and sell it to vendors in "
          "Dallas, at the market price", 0),
         ("We follow their cabbage crop — the raw material for kimchi", 0)],
        size=26, line_spacing_pts=20)
    _add_media_image(slide, "NV_s12_3_534a7244.png",
                     left=Inches(8.05), top=Inches(1.80), width=Inches(4.60))
    _add_text(slide, Inches(8.05), Inches(4.50), Inches(4.60), Inches(0.30),
              "Minari (2020)", size=12, italic=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _add_media_image(slide, "NV_s12_4_53752c12.jpg",
                     left=Inches(8.05), top=Inches(4.95), width=Inches(2.20))
    _add_media_image(slide, "NV_s12_5_042a9a61.png",
                     left=Inches(10.45), top=Inches(4.95), width=Inches(2.20))
    _yi_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[12])
    return slide


def slide_14_cost_table(prs, page_num):
    """The cabbage cost schedule as a NATIVE table (my original is a
    spreadsheet screenshot), condensed to six rows — the row selection is
    adopted from MW slide 15, the numbers are mine."""
    qs = [50, 200, 400, 600, 800, 1000]
    rows = [["Q  (tons)", "TC", "TFC", "TVC"]]
    for q in qs:
        rows.append(["{:,}".format(q), "{:,.0f}".format(tc(q)),
                     "{:,.0f}".format(TFC),
                     "{:,.0f}".format(tc(q) - TFC)])
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "Cabbage Production Costs")
    _add_styled_table(
        slide, Inches(0.75), Inches(1.95), Inches(6.30), Inches(4.20), rows,
        col_widths=[Inches(1.65)] * 4,
        row_heights=[Inches(0.66)] + [Inches(0.59)] * 6,
        font_size=19, header_size=19, first_col_align_left=False)

    _add_convention_box(
        slide, Inches(7.55), Inches(1.95), Inches(5.30), Inches(1.15),
        prefix="Fixed costs: ",
        body="land rent, equipment, and the salary the family gives up",
        corner_pct=0.12, size=17)
    _add_convention_box(
        slide, Inches(7.55), Inches(3.25), Inches(5.30), Inches(1.00),
        prefix="Variable costs: ",
        body="seeds, water, fertilizer, pesticides",
        corner_pct=0.12, size=17)
    for fn, x in (("NV_s13_5_cd8d8b01.jpg", 7.55),
                  ("NV_s13_4_3f29b09c.jpg", 9.35),
                  ("NV_s13_1_dd9ab3ce.jpg", 11.15)):
        _add_media_image(slide, fn, left=Inches(x), top=Inches(4.50),
                         width=Inches(1.70), height=Inches(1.28))
    _add_text(slide, Inches(7.55), Inches(5.85), Inches(5.30), Inches(0.30),
              "Seeds  ·  irrigation  ·  fertilizer", size=12, italic=True,
              color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
    _yi_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[13])
    return slide


def slide_15_tc_chart(prs, page_num):
    """Total cost against output, drawn natively.  The cost function is
    written on the plot (adopted from MW slide 16)."""

    def draw(slide):
        # the figure stops at y 6.00 so its x-axis title clears the
        # Yi-family badge in the bottom-right corner
        fig = SimpleFig(1.55, 6.00, 8.05, 3.55, xmax=1100.0,
                        ymax=tc(1050) * 1.10)
        _fig_axes(slide, fig, x_title="Quantity (tons)",
                  y_title="Total cost ($)", label_size=18)
        _fig_curve(slide, fig, tc, 0, 1050, color=NAVY, weight_pt=3.5,
                   segments=4)
        _fig_line(slide, fig, (0, TFC), (1050, TFC), color=GRAY,
                  weight_pt=1.75, dash="dash")
        _fig_ylab(slide, fig, TFC, _num(TFC), size=16)
        _add_text(slide, Inches(5.60), Inches(5.18), Inches(3.9),
                  Inches(0.34), "TFC — paid whatever the crop",
                  size=16, italic=True, color=GRAY, font="Calibri")
        for xv in (250, 500, 750, 1000):
            _fig_xlab(slide, fig, xv, "{:,}".format(xv), size=16)
        _add_mixed_textbox(
            slide, Inches(4.10), Inches(1.72), Inches(5.4), Inches(0.60),
            [("omml", _omml_tc(), {'size': 26})],
            align=PP_ALIGN.CENTER)
        _fig_curve_label(slide, fig, 1010, tc(1050) * 1.06, "TC", size=20)

    slide = make_diagram_slide(prs, page_num, TAG_SR,
                               "Cabbage Total Cost Function", draw)
    _yi_badge(slide)
    _set_notes(slide, NOTES[14])
    return slide


def slide_16_business_relevance(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "Business Relevance")
    _add_text(slide, MARGIN + Inches(0.15), Inches(1.75), Inches(8.0),
              Inches(0.45), "Key questions:", size=28, bold=True,
              color=NAVY, font="Calibri")
    for i, (num, q) in enumerate([
            ("1", "Should the Yi family operate at all?"),
            ("2", "How much should they produce?")]):
        y = Inches(2.55) + i * Inches(1.05)
        circ = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, int(Inches(1.05)), int(y), int(Inches(0.55)),
            int(Inches(0.55)))
        circ.fill.solid()
        circ.fill.fore_color.rgb = GOLD
        circ.line.fill.background()
        circ.shadow.inherit = False
        p = circ.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = num
        r.font.size = Pt(24)
        r.font.bold = True
        r.font.color.rgb = NAVY
        r.font.name = "Calibri"
        _add_text(slide, Inches(1.95), y + Inches(0.03), Inches(10.5),
                  Inches(0.50), q, size=28, color=NAVY, font="Calibri")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(4.85), Inches(12.4),
        Inches(1.30),
        [("We start with the second question: profit maximization", 0),
         ([("Key concept: ", {}),
           ("marginal analysis", {'bold': True, 'color': BLUE_PED})], 0, {})],
        size=28, line_spacing_pts=18)
    _yi_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[15])
    return slide


def slide_17_profit_max_rule(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "Price Taker Profit Maximization")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.80), Inches(12.4),
        Inches(3.10),
        [("Set MR = MC — the rule for every profit maximizer", 0),
         ("But MR is far easier to compute for a price taker", 0),
         ("For a price taker, MR = P, the market price", 1),
         ("P is given: the firm cannot set it", 1),
         ("So there is only one decision left: how much to produce", 0)],
        size=26, sub_size=24, line_spacing_pts=16)
    _add_math_equation(
        slide, Inches(3.85), Inches(5.20), Inches(5.6), Inches(1.05),
        _omml_text('Find ') + _omml_sup(_omml_run('Q'), _omml_text('*'))
        + _omml_text(' such that  ') + _omml_run('P')
        + _omml_text(' = ') + _omml_text('MC'),
        size_pt=30, color=NAVY, fill=CREAM, line=NAVY, rounded=True,
        shadow=True)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[16])
    return slide


def slide_18_revenue_conditions(prs, page_num):
    """The price-taker revenue table beside the two things it implies: TR
    rises along a straight ray, and d = P = MR is horizontal."""

    def draw(slide):
        rows = [["P", "Q", "TR", "MR"],
                ["$3", "10", "$30", "$3"],
                ["3", "11", "33", "3"],
                ["3", "12", "36", "3"],
                ["3", "13", "39", "3"]]
        _add_styled_table(
            slide, Inches(0.75), Inches(2.30), Inches(4.40), Inches(2.85),
            rows, col_widths=[Inches(1.10)] * 4,
            row_heights=[Inches(0.62)] + [Inches(0.56)] * 4,
            font_size=19, header_size=19, first_col_align_left=False)

        tr = SimpleFig(6.30, 5.95, 2.60, 3.30, xmax=15.0, ymax=48.0)
        _fig_axes(slide, tr, x_title="Quantity", y_title="TR ($)",
                  label_size=16)
        _fig_line(slide, tr, (0, 0), (14.0, 42.0), color=NAVY,
                  weight_pt=2.75)
        _add_text(slide, Inches(6.55), Inches(2.05), Inches(2.6),
                  Inches(0.34), "TR  =  P · Q", size=18, bold=True,
                  italic=True, color=NAVY, font="Calibri")
        _fig_guide(slide, tr, (10, 30))
        _fig_xlab(slide, tr, 10, "10", size=15)
        _fig_ylab(slide, tr, 30, "30", size=15)

        mr = SimpleFig(10.05, 5.95, 2.60, 3.30, xmax=15.0, ymax=6.0)
        _fig_axes(slide, mr, x_title="Quantity", y_title="P ($)",
                  label_size=16)
        _fig_line(slide, mr, (0, 3), (14.0, 3), color=NAVY, weight_pt=3.0)
        _fig_ylab(slide, mr, 3, "3", size=15)
        _add_text(slide, Inches(10.05), Inches(3.72), Inches(2.85),
                  Inches(0.34), "d  =  P  =  MR", size=18, bold=True,
                  italic=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)

        _add_convention_box(
            slide, Inches(0.75), Inches(5.55), Inches(4.40), Inches(0.95),
            body="Selling one more unit always brings in the same $3, "
                 "so MR = P at every quantity",
            corner_pct=0.12, size=16)

    slide = make_diagram_slide(prs, page_num, TAG_SR,
                               "Price Taker Revenue Conditions", draw)
    return slide


def slide_19_tr_tc_visual(prs, page_num):
    """TR and TC together: profit is the vertical gap, and it is widest
    where the two slopes match."""

    def draw(slide):
        fig = SimpleFig(2.40, 6.00, 7.15, 4.15, xmax=10.0, ymax=10.0)
        _fig_axes(slide, fig, x_title="Quantity (Q)",
                  y_title="Firm revenue, cost and profit ($)",
                  label_size=17)

        def tr_fn(q):
            return 1.05 * q

        def tc_fn(q):
            return 1.6 + 0.30 * q + 0.070 * q * q

        _fig_line(slide, fig, (0, 0), (9.0, tr_fn(9.0)), color=NAVY,
                  weight_pt=2.75)
        _fig_curve(slide, fig, tc_fn, 0, 9.0, color=GOLD, weight_pt=3.5)
        _fig_line(slide, fig, (0, 1.6), (9.0, 1.6), color=GRAY,
                  weight_pt=1.5, dash="dash")
        _fig_ylab(slide, fig, 1.6, "TFC", size=16)
        # the two labels are pulled apart along the curves so they do not
        # collide where TR and TC converge at the right edge
        _fig_curve_label(slide, fig, 8.60, tr_fn(8.60) - 0.60, "TR",
                         size=19)
        _fig_curve_label(slide, fig, 8.30, tc_fn(8.30) + 0.85, "TC",
                         size=19, color=GOLD)

        # profit is maximal where the slopes match: 1.05 = 0.30 + 0.14 q
        qs = (1.05 - 0.30) / 0.14
        _fig_line(slide, fig, (qs, tc_fn(qs)), (qs, tr_fn(qs)),
                  color=DARKRED, weight_pt=3.0)
        _fig_dot(slide, fig, qs, tr_fn(qs), d=Inches(0.13))
        _fig_dot(slide, fig, qs, tc_fn(qs), d=Inches(0.13))
        _fig_line(slide, fig, (qs, 0), (qs, tc_fn(qs)), color=GRAY,
                  weight_pt=1.25, dash="dash")
        _fig_xlab(slide, fig, qs, "Q*", size=18, bold=True)

        _add_text(slide, Inches(9.85), Inches(2.20), Inches(3.30),
                  Inches(0.70), "TR slope = MR = P", size=18, bold=True,
                  color=NAVY, font="Calibri")
        _add_text(slide, Inches(9.85), Inches(3.05), Inches(3.30),
                  Inches(0.70), "TC slope = MC", size=18, bold=True,
                  color=GOLD, font="Calibri")
        _add_convention_box(
            slide, Inches(9.85), Inches(3.95), Inches(3.30), Inches(1.55),
            runs=[("Profit = TR − TC", {'bold': True, 'size': 19}),
                  ("\n", {}),
                  ("It is widest where the two slopes are equal — "
                   "MR = MC", {'size': 17})],
            corner_pct=0.12, size=18)

    slide = make_diagram_slide(
        prs, page_num, TAG_SR,
        "Profit Maximization of a Price Taker: Visual Representation", draw)
    return slide


def slide_20_max_profit_setup(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "Maximizing Profit: Cabbage Production")
    _add_text(slide, MARGIN + Inches(0.15), Inches(1.75), Inches(12.4),
              Inches(0.45), "The Yi family has estimated its total cost:",
              size=26, color=NAVY, font="Calibri")
    _add_math_equation(
        slide, Inches(3.55), Inches(2.45), Inches(6.2), Inches(1.05),
        _omml_tc(),
        size_pt=30, color=NAVY, fill=CREAM, line=NAVY, rounded=True,
        shadow=True)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(3.85), Inches(12.4),
        Inches(2.20),
        [([("Q", {'italic': True}),
           (" is the quantity produced, in tons", {})], 0, {}),
         ("The market price of cabbage is $%s per ton" % _num(P_HIGH), 0),
         ([("How many tons should the Yi family produce to maximize "
            "profits?", {'bold': True})], 0, {})],
        size=26, line_spacing_pts=20)
    _yi_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[19])
    return slide


def slide_22_qstar_solution(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "Profit-Maximizing Quantity: Solution")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.70), Inches(12.4),
        Inches(0.85),
        [([("We need to set ", {}), ("MR = MC", {'bold': True}),
           (".  MC is the derivative of TC with respect to ", {}),
           ("Q", {'italic': True}), (":", {})], 0, {})],
        size=24, line_spacing_pts=0)
    _add_math_equation(
        slide, Inches(2.30), Inches(2.45), Inches(8.7), Inches(0.85),
        _omml_text('MC') + _omml_text(' = ')
        + _omml_frac(_omml_text('d') + _omml_text('TC'),
                     _omml_text('d') + _omml_run('Q'))
        + _omml_text(' = ') + _omml_mc_expr(),
        size_pt=28, color=NAVY)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(3.55), Inches(12.4),
        Inches(0.85),
        [("The cabbage market is competitive, so MR is just the market "
          "price:", 0)],
        size=24, line_spacing_pts=0)
    _add_math_equation(
        slide, Inches(2.30), Inches(4.15), Inches(8.7), Inches(0.70),
        _omml_text('MR') + _omml_text(' = ') + _omml_run('P')
        + _omml_text(' = %s' % _num(P_HIGH)),
        size_pt=28, color=NAVY)
    _add_math_equation(
        slide, Inches(2.30), Inches(5.10), Inches(8.7), Inches(0.70),
        _omml_mc_expr() + _omml_text(' = %s' % _num(P_HIGH)),
        size_pt=28, color=NAVY)
    _add_math_equation(
        slide, Inches(2.30), Inches(5.85), Inches(8.7), Inches(0.80),
        _omml_sup(_omml_run('Q'), _omml_text('*'))
        + _omml_text(' = %s tons' % _num(Q_HIGH)),
        size_pt=32, color=DARKRED)
    _yi_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_23_two_ways_atc(prs, page_num):
    """Adopted from MW slide 20: the profit identity written the two ways
    the deck goes on to use, stated before its first use."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "Two Ways to Write Profit")
    _add_rounded_filled_box(slide, Inches(1.35), Inches(1.92),
                            Inches(0.62), Inches(0.62), "1",
                            fill=GOLD, text_color=NAVY, size=25,
                            corner_pct=0.50)
    for i, eq in enumerate([
            _omml_text('Profit') + _omml_text(' = ') + _omml_text('TR')
            + _omml_text(' − ') + _omml_text('TC'),
            _omml_text('= (') + _omml_run('P') + _omml_text(' · ')
            + _omml_run('Q') + _omml_text(') − (') + _omml_text('ATC')
            + _omml_text(' · ') + _omml_run('Q') + _omml_text(')'),
            _omml_text('= (') + _omml_run('P') + _omml_text(' − ')
            + _omml_text('ATC') + _omml_text(') · ') + _omml_run('Q')]):
        _add_math_equation(
            slide, Inches(2.45), Inches(1.85) + i * Inches(0.92),
            Inches(8.4), Inches(0.80), eq, size_pt=30, color=NAVY)
    # the identity is only useful once ATC is on the table, so it is
    # defined here rather than assumed
    _add_math_equation(
        slide, Inches(2.45), Inches(4.72), Inches(8.4), Inches(1.00),
        _omml_text('where  ') + _omml_text('ATC') + _omml_text(' = ')
        + _omml_frac(_omml_text('TC'), _omml_run('Q'))
        + _omml_text('   (average total cost per ton)'),
        size_pt=24, color=GRAY)
    _add_math_equation(
        slide, Inches(2.45), Inches(5.90), Inches(8.4), Inches(0.85),
        _omml_text('Profit is positive if  ') + _omml_run('P')
        + _omml_text(' ≥ ') + _omml_text('ATC'),
        size_pt=28, color=NAVY, fill=CREAM, line=NAVY, rounded=True,
        shadow=True)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_24_profit_solution(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "Profit at Q*: Solution")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.68), Inches(12.4),
        Inches(0.80),
        [([("What is the economic profit at the profit-maximizing output "
            "level (", {}),
           ("Q* = %s" % _num(Q_HIGH), {'bold': True}),
           (") ?", {})], 0, {})],
        size=24, line_spacing_pts=0)
    eqs = [
        (_omml_text('Profit') + _omml_text(' = ') + _omml_text('TR')
         + _omml_text(' − ') + _omml_text('TC'), NAVY, 28),
        (_omml_text('TR') + _omml_text(' = ') + _omml_run('P')
         + _omml_text(' · ') + _omml_run('Q')
         + _omml_text(' = %s · %s = %s'
                      % (_num(P_HIGH), _num(Q_HIGH),
                         _num(P_HIGH * Q_HIGH))), NAVY, 26),
        (_omml_text('TC') + _omml_text(' = %s + %s · %s + %s · '
                                       % (_num(TFC), _num(B_LIN),
                                          _num(Q_HIGH), _num(B_QUAD)))
         + _omml_sup(_omml_text(_num(Q_HIGH)), _omml_text('2'))
         + _omml_text(' = %s' % _num(tc(Q_HIGH))), NAVY, 24),
        (_omml_text('Profit') + _omml_text(' = %s − %s = %s'
                                           % (_num(P_HIGH * Q_HIGH),
                                              _num(tc(Q_HIGH)),
                                              _num(P_HIGH * Q_HIGH
                                                   - tc(Q_HIGH)))),
         DARKRED, 30),
    ]
    for i, (eq, col, sz) in enumerate(eqs):
        _add_math_equation(
            slide, Inches(1.70), Inches(2.55) + i * Inches(0.88),
            Inches(10.0), Inches(0.80), eq, size_pt=sz, color=col)
    _add_takeaway_bar(
        slide, "The Yi family makes a positive economic profit at Q*",
        top=Inches(6.22), width=Inches(8.6), height=Inches(0.55),
        left=Inches(0.75), fill=GOLD, text_color=NAVY, size=19, bold=True,
        rounded=True, shadow=True)
    _yi_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def atc(q):
    return tc(q) / q


Q_ATC_MIN = (TFC / B_QUAD) ** 0.5          # 387.3 — where MC cuts ATC


def slide_25_profit_rectangle(prs, page_num):
    """The Yi family's profit as the rectangle between P and ATC.  Every
    marked point lies mathematically on the curves: MC and ATC cross at the
    minimum of ATC, and Q* is the true solution of MC = P."""

    def draw(slide):
        fig = SimpleFig(1.60, 6.00, 8.00, 3.85, xmax=950.0, ymax=500.0)
        _fig_axes(slide, fig, x_title="Quantity (tons)",
                  y_title="Price and cost ($/ton)", label_size=17)
        # The profit rectangle is drawn FIRST so the curves stay on top of
        # it — its lower edge runs along ATC, and ATC has to remain visible
        # through it.  Height P − ATC(Q*), width Q*.
        a_star = atc(Q_HIGH)
        _fig_region(slide, fig, 0, Q_HIGH, a_star, P_HIGH,
                    fill=CREAM, line=NAVY)

        _fig_line(slide, fig, (0, mc(0)), (800, mc(800)), color=NAVY,
                  weight_pt=2.75)
        _fig_curve_label(slide, fig, 778, mc(800) + 36, "MC", size=19)
        # ATC starts low enough on its falling branch for the U to read;
        # any earlier and it would run off the top of the box
        _fig_curve(slide, fig, atc, 210, 800, color=GOLD, weight_pt=3.25,
                   segments=5)
        _fig_curve_label(slide, fig, 778, atc(800) - 44, "ATC", size=19,
                         color=GOLD)
        # the margin is genuinely thin (400 against 358), so the label sits
        # inside the band rather than the band being blown up to fit it
        _add_text(slide, Inches(2.30), fig.y(P_HIGH) + Inches(0.03),
                  Inches(2.4), Inches(0.32), "Profit", size=16, bold=True,
                  color=NAVY, font="Calibri", align=PP_ALIGN.CENTER)

        _fig_line(slide, fig, (0, P_HIGH), (860, P_HIGH), color=NAVY,
                  weight_pt=2.5)
        _fig_curve_label(slide, fig, 590, P_HIGH + 42, "P = MR", size=18)
        _fig_ylab(slide, fig, P_HIGH, _num(P_HIGH), size=17, bold=True)
        _fig_ylab(slide, fig, a_star, _num(round(a_star)), size=17)
        _fig_line(slide, fig, (Q_HIGH, 0), (Q_HIGH, P_HIGH), color=GRAY,
                  weight_pt=1.5, dash="dash")
        _fig_xlab(slide, fig, Q_HIGH, "Q* = %s" % _num(Q_HIGH),
                  size=17, bold=True)
        _fig_dot(slide, fig, Q_HIGH, P_HIGH)

    slide = make_diagram_slide(
        prs, page_num, TAG_SR,
        "Illustrating Positive Profits of the Yi Family", draw)
    _yi_badge(slide)
    return slide


def slide_26_ross_stores(prs, page_num):
    """Where the cost concepts show up in a real income statement."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide,
                       "Relationship to Accounting: Ross Stores Annual Report")
    # the three crops are placed at my original relative geometry, scaled
    # from the 10" canvas — width AND height are given explicitly because
    # the strips are displayed at an aspect of my choosing, not their own
    for fn, l, t, w, h in (
            ("NV_s25_g2.1_6ae9af97.png", 0.35, 2.55, 3.19, 2.05),
            ("NV_s25_g2.2_79ddddc9.png", 3.98, 1.98, 2.34, 2.64),
            ("NV_s25_g2.3_1fcdd894.png", 6.41, 1.89, 1.91, 2.74)):
        _add_media_image(slide, fn, left=Inches(l), top=Inches(t),
                         width=Inches(w), height=Inches(h),
                         rounded=False, shadow=False)
    _add_text(slide, Inches(0.35), Inches(4.85), Inches(8.00), Inches(0.30),
              "Ross Stores, Form 10-K, fiscal 2021", size=12, italic=True,
              color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
    _add_convention_box(
        slide, Inches(8.85), Inches(1.98), Inches(4.20), Inches(3.10),
        runs=[("Cost of goods sold", {'bold': True, 'size': 17}),
              ("  —  variable\n", {'size': 17}),
              ("Selling, general and administrative",
               {'bold': True, 'size': 17}),
              ("  —  a mix of fixed and variable\n", {'size': 17}),
              ("Interest expense", {'bold': True, 'size': 17}),
              ("  —  fixed", {'size': 17})],
        corner_pct=0.10, size=17)
    _add_takeaway_bar(
        slide, "The accounting statement does not hand you TFC and TVC — "
               "you have to read them out of it",
        top=Inches(6.15), width=Inches(11.0), height=Inches(0.55),
        left=(SLIDE_W - Inches(11.0)) // 2, fill=GOLD, text_color=NAVY,
        size=18, bold=True, rounded=True, shadow=True)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[25])
    return slide


# --------------------------------------------------------------------------
#  The general-case cost curves, shared by slides 27 and 35 – 37.
#
#  AVC is U-shaped, ATC sits above it, and MC cuts each at its minimum —
#  which the algebra below guarantees rather than approximates:
#      AVC(q) = a(q − q0)² + m      TC(q) = q·AVC(q) + F
#      MC(q)  = AVC(q) + q·AVC'(q) = a(3q² − 4q0·q + q0²) + m
#  so MC = AVC exactly at q = q0, and MC cuts ATC at ATC's minimum.
# --------------------------------------------------------------------------

#
#  The level constant _GC_M is the MINIMUM of AVC, and the shut-down case
#  needs the price line to sit visibly below it, so it is set well clear of
#  the axis rather than at a token value.
_GC_A, _GC_Q0, _GC_M, _GC_F = 0.18, 4.0, 3.0, 12.0

PINK = RGBColor(0xF3, 0xD8, 0xD8)        # the loss-region fill


def gc_avc(q):
    return _GC_A * (q - _GC_Q0) ** 2 + _GC_M


def gc_atc(q):
    return gc_avc(q) + _GC_F / q


def gc_mc(q):
    return _GC_A * (3 * q * q - 4 * _GC_Q0 * q + _GC_Q0 ** 2) + _GC_M


def gc_q_star(p):
    """The output where MC = p, on MC's rising branch."""
    a, b = 3 * _GC_A, -4 * _GC_Q0 * _GC_A
    c = _GC_A * _GC_Q0 ** 2 + _GC_M - p
    return (-b + (b * b - 4 * a * c) ** 0.5) / (2 * a)


GC_XLO, GC_XHI = 3.0, 7.0
GC_MC_HI = 6.38                  # MC is clipped so it stays inside the box
GC_AVC_HI = 6.30                 # AVC stops short of the low price line
GC_ATC_MIN_Q = 5.222             # where MC cuts ATC, solved numerically
GC_XMAX, GC_YMAX = 8.6, 9.5
GC_XMAX_SMALL = 9.4              # wider box on the three small panels, so
                                 # the Q* label clears the axis title
P_CASE_HIGH, P_CASE_LOW, P_CASE_VLOW = 7.5, 4.0, 2.4
P_CASE_NEG = 4.4                 # the loss panel of the general-case slide


def _draw_cost_panel(slide, fig, price, *, label, show_avc=True,
                     region_fill=None, x_title="Q"):
    """MC / ATC (/ AVC) with a horizontal price line and the resulting Q*.

    The shaded profit / loss region is drawn FIRST, so the curves stay on
    top of it instead of being buried under the fill.
    """
    _fig_axes(slide, fig, x_title=x_title, y_title="$/Q", label_size=17)
    qs = gc_q_star(price)
    if region_fill is not None:
        lo, hi = sorted((price, gc_atc(qs)))
        _fig_region(slide, fig, 0, qs, lo, hi, fill=region_fill, line=NAVY)
    _fig_curve(slide, fig, gc_mc, GC_XLO, GC_MC_HI, color=NAVY,
               weight_pt=3.0, segments=4)
    _fig_curve(slide, fig, gc_atc, GC_XLO, GC_XHI, color=GOLD,
               weight_pt=3.0, segments=4)
    if show_avc:
        _fig_curve(slide, fig, gc_avc, GC_XLO, GC_AVC_HI, color=GRAY,
                   weight_pt=2.5, segments=4)
    _fig_line(slide, fig, (0, price), (GC_XHI + 0.15, price), color=NAVY,
              weight_pt=2.5)
    _fig_line(slide, fig, (qs, 0), (qs, price), color=GRAY,
              weight_pt=1.5, dash="dash")
    _fig_xlab(slide, fig, qs, "Q*", size=17, bold=True)
    _fig_ylab(slide, fig, price, label, size=17, bold=True)
    return qs


def slide_27_general_case(prs, page_num):
    """Positive, zero and negative profit in one row of three panels."""

    def draw(slide):
        panels = [
            ("Positive Profit", P_CASE_HIGH, "Profit, π > 0", CREAM),
            ("Zero Profit", gc_atc(GC_ATC_MIN_Q), None, None),
            ("Negative Profit (Loss)", P_CASE_NEG, "Loss, π < 0", RGBColor(
                0xF3, 0xD8, 0xD8)),
        ]
        for i, (heading, price, note, fill) in enumerate(panels):
            x = 0.55 + i * 4.32
            _add_rounded_filled_box(
                slide, Inches(x), Inches(1.62), Inches(3.90), Inches(0.52),
                heading, fill=NAVY, text_color=WHITE, size=18,
                corner_pct=0.14)
            fig = SimpleFig(x + 0.75, 6.15, 2.85, 3.30,
                            xmax=GC_XMAX_SMALL, ymax=GC_YMAX)
            _draw_cost_panel(slide, fig, price, label="P",
                             show_avc=False, region_fill=fill)
            _fig_curve_label(slide, fig, GC_MC_HI - 0.55,
                             gc_mc(GC_MC_HI) + 0.75, "MC", size=16)
            _fig_curve_label(slide, fig, GC_XHI - 0.10,
                             gc_atc(GC_XHI) + 0.75, "ATC", size=16,
                             color=GOLD)
            if note is not None:
                _add_text(slide, Inches(x + 0.10), Inches(5.55),
                          Inches(3.80), Inches(0.34), note, size=16,
                          bold=True, color=NAVY, font="Calibri",
                          align=PP_ALIGN.CENTER)
            else:
                _add_text(slide, Inches(x + 0.10), Inches(5.55),
                          Inches(3.80), Inches(0.34), "P = ATC at Q*",
                          size=16, bold=True, color=NAVY, font="Calibri",
                          align=PP_ALIGN.CENTER)

    slide = make_diagram_slide(
        prs, page_num, TAG_SR,
        "Maximizing Profits in the Short Run: The General Case", draw)
    _add_takeaway_bar(
        slide, "Profit = (P − ATC) · Q*, so the sign of the profit is the "
               "sign of P − ATC",
        top=Inches(6.45), width=Inches(10.4), height=Inches(0.52),
        left=(SLIDE_W - Inches(10.4)) // 2, fill=GOLD, text_color=NAVY,
        size=18, bold=True, rounded=True, shadow=True)
    _set_notes(slide, NOTES[26])
    return slide


def slide_28_stop_producing(prs, page_num):
    bullets = [
        ("Is a firm better off operating at a loss, or shutting down and "
         "producing nothing?", 0),
        ("Stopping production means something different in each horizon", 0),
        ("Long run: exit the industry altogether", 1),
        ("Short run: the fixed costs still have to be paid", 1),
        ("e.g. a land rental contract that runs to the end of the year", 1),
        ("We take the short-run decision first", 0),
    ]
    return content_slide(prs, page_num, TAG_SR,
                         "If Profit Is Negative, Should a Firm Stop "
                         "Producing?", bullets,
                         size=26, sub_size=24, line_spacing_pts=16)


def slide_29_two_ways_avc(prs, page_num):
    """Adopted from MW slide 25: both identities side by side, the second
    one being what the shut-down rule actually rests on."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "Two Ways to Write Profit")
    blocks = [
        ("1", Inches(1.72), [
            _omml_text('Profit') + _omml_text(' = ') + _omml_text('TR')
            + _omml_text(' − ') + _omml_text('TC'),
            _omml_text('= (') + _omml_run('P') + _omml_text(' − ')
            + _omml_text('ATC') + _omml_text(') · ') + _omml_run('Q')]),
        ("2", Inches(3.70), [
            _omml_text('Profit') + _omml_text(' = ') + _omml_run('P')
            + _omml_text(' · ') + _omml_run('Q') + _omml_text(' − ')
            + _omml_text('TVC') + _omml_text(' − ') + _omml_text('TFC'),
            _omml_text('= (') + _omml_run('P') + _omml_text(' − ')
            + _omml_text('AVC') + _omml_text(') · ') + _omml_run('Q')
            + _omml_text(' − ') + _omml_text('TFC')]),
    ]
    for num, y0, eqs in blocks:
        _add_rounded_filled_box(slide, Inches(1.30), y0 + Inches(0.30),
                                Inches(0.62), Inches(0.62), num,
                                fill=GOLD, text_color=NAVY, size=25,
                                corner_pct=0.50)
        for i, eq in enumerate(eqs):
            _add_math_equation(
                slide, Inches(2.40), y0 + i * Inches(0.85), Inches(9.4),
                Inches(0.78), eq, size_pt=28, color=NAVY)
    _add_math_equation(
        slide, Inches(2.40), Inches(5.70), Inches(9.4), Inches(0.95),
        _omml_text('The second form is what the decision to operate turns on'),
        size_pt=24, color=NAVY, fill=CREAM, line=NAVY, rounded=True,
        shadow=True)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_30_shutdown_rule(prs, page_num):
    """My slide 28, with the operate / shut-down rule set as a gold bar
    (the framing is adopted from MW slide 26)."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(
        slide, "If Profit Is Negative, Should a Firm Shut Down in the "
               "Short Run?")
    _add_text(slide, MARGIN + Inches(0.15), Inches(1.62), Inches(12.4),
              Inches(0.40),
              "Compare what the firm earns under the two options:",
              size=24, color=NAVY, font="Calibri")
    _add_rounded_filled_box(
        slide, Inches(0.70), Inches(2.25), Inches(5.85), Inches(0.55),
        "Option 1: shut down", fill=NAVY, text_color=WHITE, size=20,
        corner_pct=0.12)
    _add_rounded_filled_box(
        slide, Inches(6.90), Inches(2.25), Inches(5.85), Inches(0.55),
        "Option 2: operate", fill=NAVY, text_color=WHITE, size=20,
        corner_pct=0.12)
    _add_math_equation(
        slide, Inches(0.70), Inches(3.05), Inches(5.85), Inches(0.85),
        _omml_text('Profit') + _omml_text(' = − ') + _omml_text('TFC'),
        size_pt=26, color=NAVY)
    _add_math_equation(
        slide, Inches(6.90), Inches(3.05), Inches(5.85), Inches(0.85),
        _omml_text('Profit') + _omml_text(' = (') + _omml_run('P')
        + _omml_text(' − ') + _omml_text('AVC') + _omml_text(') · ')
        + _omml_run('Q') + _omml_text(' − ') + _omml_text('TFC'),
        size_pt=24, color=NAVY)
    _add_text(slide, Inches(0.70), Inches(3.95), Inches(5.85), Inches(0.42),
              "The fixed costs are lost either way", size=19, italic=True,
              color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
    _add_text(slide, Inches(6.90), Inches(3.95), Inches(5.85), Inches(0.42),
              "Operating adds (P − AVC) · Q on top", size=19, italic=True,
              color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
    _add_math_equation(
        slide, Inches(2.55), Inches(4.72), Inches(8.2), Inches(0.85),
        _omml_text('Difference') + _omml_text(' = ') + _omml_run('Q')
        + _omml_text(' · (') + _omml_run('P') + _omml_text(' − ')
        + _omml_text('AVC') + _omml_text(')'),
        size_pt=26, color=NAVY)
    _add_takeaway_bar(
        slide, "Operate if  P ≥ AVC        ·        Shut down if  P < AVC",
        top=Inches(5.85), width=Inches(10.2), height=Inches(0.62),
        left=(SLIDE_W - Inches(10.2)) // 2, fill=GOLD, text_color=NAVY,
        size=22, bold=True, rounded=True, shadow=True)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_31_new_price(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide,
                       "Optimizing in the Short Run: A New Market Price")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.65), Inches(12.4),
        Inches(1.45),
        [("The US starts importing large quantities of cabbage from China", 0),
         ("The US market price drops to $%s per ton" % _num(P_LOW), 0)],
        size=25, line_spacing_pts=14)
    _add_math_equation(
        slide, Inches(3.55), Inches(3.05), Inches(6.2), Inches(0.90),
        _omml_tc(),
        size_pt=26, color=NAVY, fill=CREAM, line=NAVY, rounded=True,
        shadow=True)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(4.15), Inches(11.6),
        Inches(2.00),
        [("1.  What quantity should the Yi family now produce?", 0),
         ("2.  What profit do they make at that quantity?", 0),
         ("3.  Should they continue to operate in the short run?", 0)],
        size=24, line_spacing_pts=14)
    _yi_badge(slide, "Yi Family Example  ·  Group Work")
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[29])
    return slide


def slide_33_new_qstar(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "New Q* and Profit: Solution")
    eqs = [
        (_omml_text('MC') + _omml_text(' = ') + _omml_text('MR')
         + _omml_text(' = ') + _omml_run('P'), NAVY, 26),
        (_omml_mc_expr() + _omml_text(' = %s' % _num(P_LOW))
         + _omml_text('     →     ')
         + _omml_sup(_omml_run('Q'), _omml_text('*'))
         + _omml_text(' = %s' % _num(Q_LOW)), NAVY, 26),
        (_omml_text('TR') + _omml_text(' = ') + _omml_run('P')
         + _omml_text(' · ') + _omml_run('Q')
         + _omml_text(' = %s · %s = %s'
                      % (_num(P_LOW), _num(Q_LOW),
                         _num(P_LOW * Q_LOW))), NAVY, 26),
        (_omml_text('TC') + _omml_text(' = %s + %s · %s + %s · '
                                       % (_num(TFC), _num(B_LIN),
                                          _num(Q_LOW), _num(B_QUAD)))
         + _omml_sup(_omml_text(_num(Q_LOW)), _omml_text('2'))
         + _omml_text(' = %s' % _num(tc(Q_LOW))), NAVY, 24),
        (_omml_text('Profit') + _omml_text(' = %s − %s = −%s'
                                           % (_num(P_LOW * Q_LOW),
                                              _num(tc(Q_LOW)),
                                              _num(tc(Q_LOW)
                                                   - P_LOW * Q_LOW))),
         DARKRED, 28),
    ]
    for i, (eq, col, sz) in enumerate(eqs):
        _add_math_equation(
            slide, Inches(1.55), Inches(1.85) + i * Inches(0.86),
            Inches(10.3), Inches(0.80), eq, size_pt=sz, color=col)
    _add_takeaway_bar(
        slide, "At the new price the Yi family makes a LOSS at its best "
               "output level",
        top=Inches(6.22), width=Inches(9.2), height=Inches(0.55),
        left=Inches(0.75), fill=GOLD, text_color=NAVY, size=19, bold=True,
        rounded=True, shadow=True)
    _yi_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_34_operate_solution(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "Decision to Operate: Solution")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.62), Inches(12.4),
        Inches(1.30),
        [("Compare the two options directly", 0),
         ("Operating: loss of %s" % _num(tc(Q_LOW) - P_LOW * Q_LOW), 1),
         ("Shutting down: loss of %s, the whole of TFC" % _num(TFC), 1)],
        size=24, sub_size=22, line_spacing_pts=10)
    _add_text(slide, MARGIN + Inches(0.15), Inches(3.05), Inches(12.4),
              Inches(0.40), "Or check the rule directly:", size=24,
              color=NAVY, font="Calibri")
    _add_math_equation(
        slide, Inches(1.55), Inches(3.55), Inches(10.3), Inches(1.00),
        _omml_text('AVC') + _omml_text(' = ')
        + _omml_frac(_omml_text(_num(B_LIN)) + _omml_run('Q')
                     + _omml_text(' + %s' % _num(B_QUAD))
                     + _omml_sup(_omml_run('Q'), _omml_text('2')),
                     _omml_run('Q'))
        + _omml_text(' = ') + _omml_avc_expr(),
        size_pt=26, color=NAVY)
    _add_math_equation(
        slide, Inches(1.55), Inches(4.70), Inches(10.3), Inches(0.80),
        _omml_text('AVC') + _omml_text(' at ')
        + _omml_sup(_omml_run('Q'), _omml_text('*'))
        + _omml_text(' = %s  is  %s + %s = %s'
                     % (_num(Q_LOW), _num(B_LIN), _num(B_QUAD * Q_LOW),
                        _num(avc(Q_LOW)))),
        size_pt=26, color=NAVY)
    _add_math_equation(
        slide, Inches(1.55), Inches(5.45), Inches(10.3), Inches(0.75),
        _omml_run('P') + _omml_text(' = %s  >  ' % _num(P_LOW))
        + _omml_text('AVC') + _omml_text(' = %s' % _num(avc(Q_LOW))),
        size_pt=28, color=DARKRED)
    _add_takeaway_bar(
        slide, "Keep operating in the short run",
        top=Inches(6.25), width=Inches(6.4), height=Inches(0.55),
        left=Inches(1.55), fill=GOLD, text_color=NAVY, size=20, bold=True,
        rounded=True, shadow=True)
    _yi_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def _price_case_slide(prs, page_num, title, price, price_label, *,
                      region_kind, footnote, avc_dy=-0.40):
    """One of the three price cases on the general cost curves."""

    def draw(slide):
        heading, fill = {
            "profit": ("Profit", CREAM),
            "loss": ("Loss if operating", PINK),
            "shutdown": ("Loss if operating —\nlarger than TFC", PINK),
        }[region_kind]
        fig = SimpleFig(1.55, 5.85, 7.30, 3.75,
                        xmax=GC_XMAX, ymax=GC_YMAX)
        _draw_cost_panel(slide, fig, price, label=price_label,
                         region_fill=fill)
        _fig_curve_label(slide, fig, GC_MC_HI - 0.45,
                         gc_mc(GC_MC_HI) + 0.65, "MC", size=18)
        _fig_curve_label(slide, fig, GC_XHI + 0.20,
                         gc_atc(GC_XHI) + 0.40, "ATC", size=18, color=GOLD)
        _fig_curve_label(slide, fig, GC_AVC_HI + 0.20,
                         gc_avc(GC_AVC_HI) + avc_dy, "AVC", size=18,
                         color=GRAY)
        _add_text(slide, Inches(9.30), Inches(2.35), Inches(3.60),
                  Inches(0.62), heading, size=21, bold=True, color=NAVY,
                  font="Calibri")
        _add_convention_box(
            slide, Inches(9.30), Inches(3.55), Inches(3.60), Inches(2.10),
            body=footnote, corner_pct=0.10, size=17)

    slide = make_diagram_slide(prs, page_num, TAG_SR, title, draw)
    return slide


def slide_35_high_price(prs, page_num):
    slide = _price_case_slide(
        prs, page_num, "High Price: Positive Profits in the Short Run",
        P_CASE_HIGH, "P high", region_kind="profit",
        footnote="Optimal output is where MC = MR = P.  At that output the "
                 "firm earns positive economic profits, because P > ATC.")
    _set_notes(slide, NOTES[33])
    return slide


def slide_36_low_price(prs, page_num):
    return _price_case_slide(
        prs, page_num, "Low Price: Operating at a Loss in the Short Run",
        P_CASE_LOW, "P low", region_kind="loss", avc_dy=-1.05,
        footnote="Economic profits are negative, because P < ATC.  But "
                 "P ≥ AVC, so the firm should keep operating in the short "
                 "run — it still covers part of its fixed costs.")


def slide_37_very_low_price(prs, page_num):
    return _price_case_slide(
        prs, page_num, "Very Low Price: Shutting Down in the Short Run",
        P_CASE_VLOW, "P very low", region_kind="shutdown",
        footnote="P < AVC now, so every ton produced loses money before "
                 "any fixed cost is covered.  The firm should shut down; "
                 "its loss is then TFC.")


def slide_38_coffee(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "Coffee Bean Producer")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.75), Inches(7.6),
        Inches(1.30),
        [("Your firm is a price taker in the coffee bean market", 0),
         ("You have just optimized your output level for the short run", 0)],
        size=25, line_spacing_pts=16)
    rows = [["", "Your costs at that output"],
            ["AVC", "$12"],
            ["ATC", "$15"],
            ["MC", "$10"]]
    _add_styled_table(
        slide, Inches(0.75), Inches(3.45), Inches(4.60), Inches(2.30), rows,
        col_widths=[Inches(1.75), Inches(2.85)],
        row_heights=[Inches(0.62)] + [Inches(0.56)] * 3,
        font_size=20, header_size=19)
    _add_text(slide, Inches(6.05), Inches(4.05), Inches(5.4), Inches(0.90),
              "Should you continue to operate\nin the short run?", size=26,
              bold=True, color=NAVY, font="Calibri")
    _add_media_image(slide, "NV_s36_4_7cb0fc52.jpg",
                     left=Inches(8.85), top=Inches(1.55), width=Inches(4.00))
    _add_pollbreak_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_40_coffee_solution(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide, "Coffee Bean Producer: Solution")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.70), Inches(12.0),
        Inches(1.10),
        [("The output level is already profit-maximizing, so for a price "
          "taker P = MC", 0)],
        size=24, line_spacing_pts=0)
    eqs = [
        (_omml_text('MC') + _omml_text(' = 10')
         + _omml_text('     →     ') + _omml_run('P')
         + _omml_text(' = 10'), NAVY, 28),
        (_omml_text('AVC') + _omml_text(' = 12')
         + _omml_text('     →     ') + _omml_run('P')
         + _omml_text(' < ') + _omml_text('AVC'), NAVY, 28),
    ]
    for i, (eq, col, sz) in enumerate(eqs):
        _add_math_equation(
            slide, Inches(2.30), Inches(2.80) + i * Inches(0.95),
            Inches(8.7), Inches(0.80), eq, size_pt=sz, color=col)
    _add_math_equation(
        slide, Inches(2.30), Inches(4.75), Inches(8.7), Inches(0.85),
        _omml_text('You lose money on every unit — shut down'),
        size_pt=30, color=DARKRED)
    _add_takeaway_bar(
        slide, "ATC does not enter the short-run shut-down decision at all",
        top=Inches(5.90), width=Inches(9.4), height=Inches(0.55),
        left=(SLIDE_W - Inches(9.4)) // 2, fill=GOLD, text_color=NAVY,
        size=19, bold=True, rounded=True, shadow=True)
    _add_pollbreak_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_41_sr_summary(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_SR)
    _draw_action_title(slide,
                       "Summary: Short-Run Decisions for a Price Taker")
    _add_math_equation(
        slide, Inches(2.55), Inches(1.62), Inches(8.2), Inches(0.80),
        _omml_text('Optimal output: ')
        + _omml_sup(_omml_run('Q'), _omml_text('*'))
        + _omml_text(' where  ') + _omml_text('MC') + _omml_text(' = ')
        + _omml_text('MR') + _omml_text(' = ') + _omml_run('P'),
        size_pt=26, color=NAVY)
    for i, (num, head, subs) in enumerate([
            ("1", "If P ≥ AVC at Q*:  operate",
             ["Profits can still be positive or negative",
              "Positive if P ≥ ATC, negative if P < ATC"]),
            ("2", "If P < AVC at Q*:  shut down",
             ["The loss is then TFC",
              "Producing anyway would make the loss larger than TFC"])]):
        y = Inches(2.65) + i * Inches(1.85)
        _add_rounded_filled_box(slide, Inches(0.90), y, Inches(0.58),
                                Inches(0.58), num, fill=GOLD,
                                text_color=NAVY, size=24, corner_pct=0.50)
        _add_text(slide, Inches(1.75), y + Inches(0.02), Inches(10.5),
                  Inches(0.50), head, size=26, bold=True, color=NAVY,
                  font="Calibri")
        _add_hierarchical_bullets(
            slide, Inches(1.95), y + Inches(0.62), Inches(10.3),
            Inches(1.00), [(s, 1) for s in subs],
            size=24, sub_size=22, line_spacing_pts=4)
    _add_ps_pointer(slide, label="Problem Set 3")
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


# ==========================================================================
#  2b · FIRM-LEVEL AND MARKET SUPPLY — slides 42 – 45
# ==========================================================================

def slide_43_supply_curve(prs, page_num):
    """Trace the firm's supply curve out of its MC curve: the two prices
    and the two quantities are the ones the worked example already
    delivered, read straight off the constants."""

    def draw(slide):
        left = SimpleFig(1.30, 6.05, 4.55, 3.70, xmax=820.0, ymax=500.0)
        right = SimpleFig(7.55, 6.05, 4.55, 3.70, xmax=820.0, ymax=500.0)
        for fig, curve_label in ((left, "MC"), (right, "S")):
            _fig_axes(slide, fig, x_title="Q", y_title="Price ($)",
                      label_size=17)
            _fig_line(slide, fig, (0, mc(0)), (760, mc(760)), color=NAVY,
                      weight_pt=3.0)
            _fig_curve_label(slide, fig, 740, mc(760) + 26, curve_label,
                             size=19)
            for p, q in ((P_LOW, Q_LOW), (P_HIGH, Q_HIGH)):
                _fig_line(slide, fig, (0, p), (q, p), color=GRAY,
                          weight_pt=1.5, dash="dash")
                _fig_line(slide, fig, (q, 0), (q, p), color=GRAY,
                          weight_pt=1.5, dash="dash")
                _fig_ylab(slide, fig, p, _num(p), size=16)
                _xlab_n(slide, fig, q, _num(q), size=16, bold=False, w=0.86)
                _fig_dot(slide, fig, q, p, d=Inches(0.13))
        # AVC only on the left panel: the supply curve is MC above AVC
        _fig_line(slide, left, (0, avc(0)), (760, avc(760)), color=GRAY,
                  weight_pt=2.25)
        _fig_curve_label(slide, left, 740, avc(760) + 26, "AVC", size=17,
                         color=GRAY)
        _add_text(slide, Inches(1.30), Inches(1.58), Inches(4.55),
                  Inches(0.38), "The firm's cost curves", size=20,
                  bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)
        _add_text(slide, Inches(7.55), Inches(1.58), Inches(4.55),
                  Inches(0.38), "The firm's supply curve", size=20,
                  bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)

    slide = make_diagram_slide(
        prs, page_num, TAG_SUPPLY,
        "The MC Curve Is the Price Taker's Short-Run Supply Curve", draw)
    _add_takeaway_bar(
        slide, "Read the quantity off MC at each price — that IS the "
               "supply curve",
        top=Inches(6.42), width=Inches(8.9), height=Inches(0.55),
        left=Inches(0.65), fill=GOLD, text_color=NAVY, size=19, bold=True,
        rounded=True, shadow=True)
    _yi_badge(slide)
    _set_notes(slide, NOTES[41])
    return slide


MC1_SLOPE = 0.30         # high-yield seeds: MC pivots down from 0.40


def mc1(q):
    return B_LIN + MC1_SLOPE * q


def slide_44_changing_mc(prs, page_num):
    """High-yield seeds lower marginal cost, so the firm supplies more at
    the same price.  Labels use the MC0 / MC1, q0 / q1 convention."""

    def draw(slide):
        fig = SimpleFig(2.10, 6.05, 7.40, 3.90, xmax=1150.0, ymax=530.0)
        _fig_axes(slide, fig, x_title="q", y_title="$", label_size=18)
        _fig_line(slide, fig, (0, mc(0)), (900, mc(900)), color=NAVY,
                  weight_pt=3.0)
        _fig_curve_label(slide, fig, 880, mc(900) + 30, "MC0", size=19)
        _fig_line(slide, fig, (0, mc1(0)), (1080, mc1(1080)), color=GOLD,
                  weight_pt=3.0)
        _fig_curve_label(slide, fig, 1055, mc1(1080) + 30, "MC1", size=19,
                         color=GOLD)
        _fig_line(slide, fig, (0, P_HIGH), (1100, P_HIGH), color=NAVY,
                  weight_pt=2.5)
        _fig_curve_label(slide, fig, 980, P_HIGH + 44, "P = MR", size=18)
        q0, q1 = Q_HIGH, (P_HIGH - B_LIN) / MC1_SLOPE
        for q, lab in ((q0, "q0"), (q1, "q1")):
            _fig_line(slide, fig, (q, 0), (q, P_HIGH), color=GRAY,
                      weight_pt=1.5, dash="dash")
            _fig_xlab(slide, fig, q, lab, size=18, bold=True)
            _fig_dot(slide, fig, q, P_HIGH, d=Inches(0.13))
        _add_arrow(slide, (fig.x(q0), fig.y(145)), (fig.x(q1), fig.y(145)),
                   color=GOLD, weight_pt=3.0, head=True)
        _add_text(slide, Inches(2.55), Inches(2.30), Inches(3.4),
                  Inches(0.75), "High-yield seeds\nlower marginal cost",
                  size=19, bold=True, color=GOLD, font="Calibri")

    slide = make_diagram_slide(
        prs, page_num, TAG_SUPPLY,
        "The Effect of Changing Marginal Costs on a Price Taker's Supply",
        draw)
    _set_notes(slide, NOTES[42])
    return slide


def slide_45_market_dynamics(prs, page_num):
    """When EVERY farmer gets the new seeds, market supply shifts, the
    price falls, and the firm ends up between its old output and what it
    would have produced as the only one with the new seeds."""

    def draw(slide):
        mk = SimpleFig(1.05, 6.02, 4.35, 3.35, xmax=11.0, ymax=11.0)
        fm = SimpleFig(7.35, 6.02, 4.45, 3.35, xmax=11.0, ymax=11.0)
        _add_text(slide, Inches(1.05), Inches(1.98), Inches(4.35),
                  Inches(0.38), "Market", size=21, bold=True, color=NAVY,
                  font="Calibri", align=PP_ALIGN.CENTER)
        _add_text(slide, Inches(7.35), Inches(1.98), Inches(4.45),
                  Inches(0.38), "Firm (farmer)", size=21, bold=True,
                  color=NAVY, font="Calibri", align=PP_ALIGN.CENTER)
        for fig in (mk, fm):
            _fig_axes(slide, fig, x_title="Q", y_title="P",
                      label_size=17)

        # market: S0 and S1 = S0 shifted right; D falls
        _fig_line(slide, mk, (1.0, 1.0), (9.4, 9.4), color=NAVY,
                  weight_pt=2.75)
        _fig_curve_label(slide, mk, 9.35, 9.7, "S0", size=18)
        _fig_line(slide, mk, (3.4, 1.0), (10.4, 8.0), color=GOLD,
                  weight_pt=2.75)
        _fig_curve_label(slide, mk, 10.3, 8.4, "S1", size=18, color=GOLD)
        _fig_line(slide, mk, (1.0, 9.0), (9.6, 0.4), color=NAVY,
                  weight_pt=2.75)
        _fig_curve_label(slide, mk, 9.5, 0.9, "D", size=18)
        for q, p, lab in ((5.0, 5.0, "Q0"), (6.2, 3.8, "Q1")):
            _fig_line(slide, mk, (0, p), (q, p), color=GRAY,
                      weight_pt=1.5, dash="dash")
            _fig_line(slide, mk, (q, 0), (q, p), color=GRAY,
                      weight_pt=1.5, dash="dash")
            _xlab_n(slide, mk, q, lab, w=0.50)
            _fig_dot(slide, mk, q, p, d=Inches(0.13))
        _fig_ylab(slide, mk, 5.0, "P0", size=17, bold=True)
        _fig_ylab(slide, mk, 3.8, "P1", size=17, bold=True)

        # firm: MC0 and MC1, the two price lines, three quantities
        _fig_line(slide, fm, (0, 0), (8.2, 9.84), color=NAVY,
                  weight_pt=2.75)
        _fig_curve_label(slide, fm, 8.05, 10.2, "MC0", size=18)
        _fig_line(slide, fm, (0, 0), (10.4, 7.8), color=GOLD,
                  weight_pt=2.75)
        _fig_curve_label(slide, fm, 10.2, 8.2, "MC1", size=18, color=GOLD)
        _fig_line(slide, fm, (0, 5.0), (10.6, 5.0), color=NAVY,
                  weight_pt=2.25)
        _fig_line(slide, fm, (0, 3.8), (10.6, 3.8), color=GOLD,
                  weight_pt=2.25)
        _add_text(slide, Inches(11.80), fm.y(5.0) - Inches(0.17),
                  Inches(1.50), Inches(0.34), "MR0 = P0", size=15,
                  bold=True, color=NAVY, font="Calibri")
        _add_text(slide, Inches(11.80), fm.y(3.8) - Inches(0.17),
                  Inches(1.50), Inches(0.34), "MR1 = P1", size=15,
                  bold=True, color=GOLD, font="Calibri")
        for q, p, lab in ((5.0 / 1.2, 5.0, "q0"),
                          (3.8 / 0.75, 3.8, "q2"),
                          (5.0 / 0.75, 5.0, "q1")):
            _fig_line(slide, fm, (q, 0), (q, p), color=GRAY,
                      weight_pt=1.25, dash="dash")
            _xlab_n(slide, fm, q, lab)
            _fig_dot(slide, fm, q, p, d=Inches(0.12))

    slide = make_diagram_slide(prs, page_num, TAG_SUPPLY,
                               "Firm and Market Dynamics When MC Falls", draw)
    _add_takeaway_bar(
        slide, "Everyone's costs fall, so the price falls too — the farmer "
               "gains less than if the seeds were his alone",
        top=Inches(6.45), width=Inches(11.4), height=Inches(0.55),
        left=(SLIDE_W - Inches(11.4)) // 2, fill=GOLD, text_color=NAVY,
        size=18, bold=True, rounded=True, shadow=True)
    _set_notes(slide, NOTES[43])
    return slide


# ==========================================================================
#  2c · LONG-RUN COMPETITIVE EQUILIBRIUM — slides 46 – 54
# ==========================================================================

def slide_47_long_run(prs, page_num):
    bullets = [
        ("What changes in the long run?", 0),
        ("Capital can be adjusted too, so every cost is variable", 1),
        ("The relevant curve is long-run average cost (LAC)", 1),
        ("Firms can enter and leave the market", 0),
        ("Entry and exit move the market price", 1),
    ]
    slide = content_slide(prs, page_num, TAG_LR,
                          "Perfect Competition in the Long Run", bullets,
                          size=28, sub_size=25, line_spacing_pts=18)
    _set_notes(slide, NOTES[45])
    return slide


# Long-run curves.  LAC is U-shaped; LMC is derived from it, so LMC cuts
# LAC exactly at LAC's minimum rather than near it:
#     LAC(q) = a(q − q0)² + m      LMC(q) = LAC + q·LAC'(q)
_LR_A, _LR_Q0, _LR_M = 0.25, 5.0, 3.0


def lr_lac(q):
    return _LR_A * (q - _LR_Q0) ** 2 + _LR_M


def lr_lmc(q):
    return lr_lac(q) + q * 2 * _LR_A * (q - _LR_Q0)


def lr_q_star(p):
    a = 3 * _LR_A
    b = -4 * _LR_A * _LR_Q0
    c = _LR_A * _LR_Q0 ** 2 + _LR_M - p
    return (-b + (b * b - 4 * a * c) ** 0.5) / (2 * a)


def slide_48_lr_equilibrium(prs, page_num):
    """Entry pushes the price down from P1, exit pushes it up from P2, and
    both stop at P_LR = min LAC, where LMC = LAC = P."""

    def draw(slide):
        fig = SimpleFig(2.55, 6.10, 7.10, 4.05, xmax=9.6, ymax=9.0)
        _fig_axes(slide, fig, x_title="Q", y_title="$/Q", label_size=18)
        _fig_curve(slide, fig, lr_lac, 1.5, 8.2, color=GOLD, weight_pt=3.25,
                   segments=5)
        _fig_curve_label(slide, fig, 8.20, lr_lac(8.2) + 0.35, "LAC",
                         size=19, color=GOLD)
        _fig_curve(slide, fig, lr_lmc, 1.5, 6.5, color=NAVY, weight_pt=3.0,
                   segments=4)
        _fig_curve_label(slide, fig, 6.35, lr_lmc(6.5) + 0.45, "LMC",
                         size=19)

        p_lr = lr_lac(_LR_Q0)
        for p, lab, col in ((6.0, "P1", NAVY), (p_lr, "PLR", DARKRED),
                            (1.5, "P2", NAVY)):
            q = lr_q_star(p)
            _fig_line(slide, fig, (0, p), (9.2, p), color=col,
                      weight_pt=2.5)
            _add_text(slide, Inches(9.80), fig.y(p) - Inches(0.17),
                      Inches(2.0), Inches(0.34), "MR = %s" % lab, size=17,
                      bold=True, color=col, font="Calibri")
            _fig_ylab(slide, fig, p, lab, size=17, bold=True)
            _fig_line(slide, fig, (q, 0), (q, p), color=GRAY,
                      weight_pt=1.4, dash="dash")
            _xlab_n(slide, fig, q, lab.replace("P", "Q"))
            _fig_dot(slide, fig, q, p, d=Inches(0.13))

        _add_arrow(slide, (fig.x(7.4), fig.y(5.85)),
                   (fig.x(7.4), fig.y(3.15)), color=NAVY, weight_pt=2.5,
                   head=True)
        _add_text(slide, fig.x(7.5), fig.y(4.6) - Inches(0.17),
                  Inches(1.3), Inches(0.34), "Entry", size=17, bold=True,
                  color=NAVY, font="Calibri")
        _add_arrow(slide, (fig.x(2.6), fig.y(1.65)),
                   (fig.x(2.6), fig.y(2.85)), color=NAVY, weight_pt=2.5,
                   head=True)
        _add_text(slide, fig.x(2.7), fig.y(2.3) - Inches(0.17),
                  Inches(1.3), Inches(0.34), "Exit", size=17, bold=True,
                  color=NAVY, font="Calibri")

    slide = make_diagram_slide(
        prs, page_num, TAG_LR,
        "Long-Run Competitive Equilibrium With Many Identical Producers",
        draw)
    _add_takeaway_bar(
        slide, "At QLR:  P = LAC = LMC, and economic profit is zero",
        top=Inches(6.45), width=Inches(8.2), height=Inches(0.55),
        left=(SLIDE_W - Inches(8.2)) // 2, fill=GOLD, text_color=NAVY,
        size=20, bold=True, rounded=True, shadow=True)
    _set_notes(slide, NOTES[46])
    return slide


def slide_49_chickens(prs, page_num):
    """Cecile Steele: an accidental order of 500 chicks, a large profit, and
    then the entry that competed it away."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_LR)
    _draw_action_title(slide,
                       "Entry Follows Profit: The Broiler Chicken Industry")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.70), Inches(6.9),
        Inches(2.90),
        [("In the 1920s the farmer Cecile Steele of Ocean View, Delaware "
          "kept a small flock of chickens for eggs", 0),
         ("In 1923 she ordered 50 chicks and, by accident, received 500. "
          "She kept them and made a sizable profit", 0),
         ("Word of her success spread, and by 1928 hundreds of farmers in "
          "the area were raising chickens for meat", 0)],
        size=23, line_spacing_pts=16)
    _add_media_image(slide, "NV_s47_2_f42d0c16.png",
                     left=Inches(7.55), top=Inches(1.85), width=Inches(4.30))
    _add_text(slide, Inches(7.55), Inches(5.30), Inches(4.30), Inches(0.30),
              "An early Delaware broiler farm. Photo: Vox", size=12,
              italic=True, color=GRAY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_takeaway_bar(
        slide, "Profit in a market with low entry barriers → new firms enter",
        top=Inches(6.10), width=Inches(9.0), height=Inches(0.58),
        left=Inches(0.60), fill=GOLD, text_color=NAVY, size=20, bold=True,
        rounded=True, shadow=True)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[47])
    return slide


def slide_50_lr_summary(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_LR)
    _draw_action_title(slide,
                       "Summary: Long-Run Decisions for a Price Taker")
    _add_math_equation(
        slide, Inches(2.55), Inches(1.60), Inches(8.2), Inches(0.80),
        _omml_text('Optimal output: ')
        + _omml_sup(_omml_run('Q'), _omml_text('*'))
        + _omml_text(' where  ') + _omml_text('LMC') + _omml_text(' = ')
        + _omml_text('MR') + _omml_text(' = ') + _omml_run('P'),
        size_pt=26, color=NAVY)
    for i, (num, head, subs) in enumerate([
            ("1", "If P ≥ LAC at Q*:  operate",
             ["Positive profits attract entry in the long run"]),
            ("2", "If P < LAC at Q*:  exit, or never enter",
             ["Losses cannot be sustained once capital is adjustable"])]):
        y = Inches(2.55) + i * Inches(1.55)
        _add_rounded_filled_box(slide, Inches(0.90), y, Inches(0.58),
                                Inches(0.58), num, fill=GOLD,
                                text_color=NAVY, size=24, corner_pct=0.50)
        _add_text(slide, Inches(1.75), y + Inches(0.02), Inches(10.5),
                  Inches(0.50), head, size=26, bold=True, color=NAVY,
                  font="Calibri")
        _add_hierarchical_bullets(
            slide, Inches(1.95), y + Inches(0.62), Inches(10.3),
            Inches(0.60), [(s, 1) for s in subs],
            size=24, sub_size=22, line_spacing_pts=4)
    _add_convention_box(
        slide, Inches(1.75), Inches(5.70), Inches(9.8), Inches(1.10),
        runs=[("Long-run competitive equilibrium:  LMC = P = LAC",
               {'bold': True, 'size': 20}),
              ("\n", {}),
              ("“Normal profits” — economic profit is zero in the long run",
               {'size': 18})],
        corner_pct=0.10, size=20, align=PP_ALIGN.CENTER)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_51_drug_market(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_LR)
    _draw_action_title(slide, "Long-Run Equilibrium: The Drug Market")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.65), Inches(12.4),
        Inches(1.20),
        [("Can we use the toolbox on the market for illegal drugs?", 0),
         ("Assume individual dealers are price takers", 1)],
        size=25, sub_size=23, line_spacing_pts=10)
    _add_text(slide, MARGIN + Inches(0.15), Inches(3.05), Inches(12.4),
              Inches(0.42),
              "The government has two one-time options:", size=25,
              color=NAVY, font="Calibri")
    for i, opt in enumerate(["Arrest the drug dealers",
                             "Arrest or punish the drug users"]):
        y = Inches(3.70) + i * Inches(0.85)
        _add_rounded_filled_box(slide, Inches(1.05), y, Inches(0.55),
                                Inches(0.55), str(i + 1), fill=GOLD,
                                text_color=NAVY, size=24, corner_pct=0.50)
        _add_text(slide, Inches(1.95), y + Inches(0.03), Inches(10.5),
                  Inches(0.50), opt, size=26, color=NAVY, font="Calibri")
    _add_takeaway_bar(
        slide, "Which policy does more to reduce the use of illegal drugs?",
        top=Inches(5.75), width=Inches(9.6), height=Inches(0.58),
        left=(SLIDE_W - Inches(9.6)) // 2, fill=GOLD, text_color=NAVY,
        size=21, bold=True, rounded=True, shadow=True)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[49])
    return slide


def slide_53_arrest_dealers(prs, page_num):
    """Deliberately blank below the title — I draw the diagram live in
    class (2026-08-28, Nico).  Chrome and speaker notes only."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_LR)
    _draw_action_title(slide, "Arrest Drug Dealers")
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[51])
    return slide


def slide_54_arrest_users(prs, page_num):
    """Deliberately blank below the title — see slide_53_arrest_dealers."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_LR)
    _draw_action_title(slide, "Arrest / Punish Drug Users")
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[52])
    return slide


# ==========================================================================
#  3 · MARKET DISTORTIONS AND REGULATIONS — slides 55 – 75
# ==========================================================================

def _fig_poly(slide, fig, pts, *, fill=CREAM, line=NAVY, alpha=None):
    """A filled polygon in logical coordinates (surplus triangles, tax
    areas).  Built as a custGeom freeform so it stays editable."""
    dev = [(int(fig.x(x)), int(fig.y(y))) for x, y in pts]
    left = min(p[0] for p in dev)
    top = min(p[1] for p in dev)
    w = max(max(p[0] for p in dev) - left, 1)
    h = max(max(p[1] for p in dev) - top, 1)
    path = ['<a:moveTo><a:pt x="%d" y="%d"/></a:moveTo>'
            % (dev[0][0] - left, dev[0][1] - top)]
    for x, y in dev[1:]:
        path.append('<a:lnTo><a:pt x="%d" y="%d"/></a:lnTo>'
                    % (x - left, y - top))
    path.append('<a:close/>')
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, w, h)
    spPr = shp._element.spPr
    old = spPr.find(qn('a:prstGeom'))
    geom = _H.ET.fromstring(
        '<a:custGeom xmlns:a="%s"><a:avLst/><a:gdLst/><a:ahLst/>'
        '<a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
        '<a:pathLst><a:path w="%d" h="%d">%s</a:path></a:pathLst>'
        '</a:custGeom>' % (_H.A_NS, w, h, "".join(path)))
    old.addnext(geom)
    spPr.remove(old)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1.25)
    shp.shadow.inherit = False
    if alpha is not None:
        sf = spPr.find(qn('a:solidFill'))
        sf.find(qn('a:srgbClr')).append(_H.ET.fromstring(
            '<a:alpha xmlns:a="%s" val="%d"/>' % (_H.A_NS, int(alpha))))
    return shp


def slide_56_distortions_intro(prs, page_num):
    bullets = [
        ("In a competitive market, equilibrium is where supply meets "
         "demand", 0),
        ("Outside forces can push the market away from it", 0),
        ("Price controls — rent control, minimum wages", 1),
        ("Taxes and subsidies", 1),
        ("How do we tell who gains and who loses?", 0),
        ("Consumer and producer surplus", 1),
        ("Deadweight loss", 1),
    ]
    return content_slide(prs, page_num, TAG_DIST,
                         "Market Distortions: Introduction", bullets,
                         size=26, sub_size=24, line_spacing_pts=14)


# a shared supply / demand frame for the surplus and price-control slides
def _sd_fig(left=1.90, bottom=6.35, w=4.30, h=4.20):
    return SimpleFig(left, bottom, w, h, xmax=11.0, ymax=11.0)


def _sd_curves(slide, fig, *, d_label="D", s_label="S"):
    """D falls, S rises, and they cross at (Q*, P*) = (3.75, 5.5)."""
    _fig_line(slide, fig, (0.0, 10.0), (9.6, 0.4), color=NAVY,
              weight_pt=2.75)
    _fig_curve_label(slide, fig, 9.5, 0.95, d_label, size=19)
    _fig_line(slide, fig, (0.0, 1.0), (8.0, 10.6), color=NAVY,
              weight_pt=2.75)
    _fig_curve_label(slide, fig, 7.9, 10.9, s_label, size=19)


SD_QSTAR, SD_PSTAR = 3.75, 5.5          # the exact crossing of the two lines


def slide_57_consumer_surplus(prs, page_num):
    def draw(slide):
        fig = _sd_fig(bottom=6.40, h=3.95)
        _fig_axes(slide, fig, x_title="Q", y_title="Price", label_size=18)
        _fig_poly(slide, fig,
                  [(0, SD_PSTAR), (0, 10.0), (SD_QSTAR, SD_PSTAR)],
                  fill=CREAM, line=NAVY)
        _sd_curves(slide, fig)
        _fig_line(slide, fig, (0, SD_PSTAR), (SD_QSTAR, SD_PSTAR),
                  color=GRAY, weight_pt=1.5, dash="dash")
        _fig_line(slide, fig, (SD_QSTAR, 0), (SD_QSTAR, SD_PSTAR),
                  color=GRAY, weight_pt=1.5, dash="dash")
        _fig_ylab(slide, fig, SD_PSTAR, "P*", size=18, bold=True)
        _fig_xlab(slide, fig, SD_QSTAR, "Q*", size=18, bold=True)
        _fig_dot(slide, fig, SD_QSTAR, SD_PSTAR)
        _add_text(slide, Inches(2.25), fig.y(7.4) - Inches(0.20),
                  Inches(1.2), Inches(0.40), "CS", size=22, bold=True,
                  color=NAVY, font="Calibri", align=PP_ALIGN.CENTER)
        _add_convention_box(
            slide, Inches(7.30), Inches(2.70), Inches(5.55), Inches(1.90),
            body="Consumer surplus is the area below the demand curve and "
                 "above the price line — what buyers were willing to pay, "
                 "less what they actually paid.",
            corner_pct=0.10, size=19)

    slide = make_diagram_slide(prs, page_num, TAG_DIST, "Consumer Surplus",
                               draw)
    _add_text(slide, MARGIN + Inches(0.15), Inches(1.36), Inches(12.4),
              Inches(0.42),
              "CS: the gap between what consumers would pay (the D curve) "
              "and what they do pay (P*)", size=21, color=NAVY,
              font="Calibri")
    _set_notes(slide, NOTES[55])
    return slide


def slide_58_producer_surplus(prs, page_num):
    def draw(slide):
        fig = _sd_fig(bottom=6.40, h=3.95)
        _fig_axes(slide, fig, x_title="Q", y_title="Price", label_size=18)
        _fig_poly(slide, fig,
                  [(0, SD_PSTAR), (0, 1.0), (SD_QSTAR, SD_PSTAR)],
                  fill=CREAM, line=NAVY)
        _sd_curves(slide, fig)
        _fig_line(slide, fig, (0, SD_PSTAR), (SD_QSTAR, SD_PSTAR),
                  color=GRAY, weight_pt=1.5, dash="dash")
        _fig_line(slide, fig, (SD_QSTAR, 0), (SD_QSTAR, SD_PSTAR),
                  color=GRAY, weight_pt=1.5, dash="dash")
        _fig_ylab(slide, fig, SD_PSTAR, "P*", size=18, bold=True)
        _fig_xlab(slide, fig, SD_QSTAR, "Q*", size=18, bold=True)
        _fig_dot(slide, fig, SD_QSTAR, SD_PSTAR)
        _add_text(slide, Inches(2.25), fig.y(3.4) - Inches(0.20),
                  Inches(1.2), Inches(0.40), "PS", size=22, bold=True,
                  color=NAVY, font="Calibri", align=PP_ALIGN.CENTER)
        _add_convention_box(
            slide, Inches(7.30), Inches(2.70), Inches(5.55), Inches(1.90),
            body="Producer surplus is the area above the supply curve and "
                 "below the price line — what sellers received, less the "
                 "least they would have accepted.",
            corner_pct=0.10, size=19)

    slide = make_diagram_slide(prs, page_num, TAG_DIST, "Producer Surplus",
                               draw)
    _add_text(slide, MARGIN + Inches(0.15), Inches(1.36), Inches(12.4),
              Inches(0.42),
              "PS: the gap between the price at which producers would sell "
              "(the S curve) and what they do receive (P*)", size=21,
              color=NAVY, font="Calibri")
    _set_notes(slide, NOTES[56])
    return slide


def slide_59_deadweight_loss(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_DIST)
    _draw_action_title(slide, "Deadweight Loss")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.75), Inches(12.4),
        Inches(2.30),
        [("The inefficiency that a market intervention creates", 0),
         ("Mutually beneficial transactions that no longer happen", 0),
         ("Workers willing to work at the market wage and firms willing "
          "to hire them — but a minimum wage stands in the way", 1)],
        size=26, sub_size=23, line_spacing_pts=18)
    _add_math_equation(
        slide, Inches(1.20), Inches(4.55), Inches(11.0), Inches(1.30),
        _omml_text('Deadweight loss') + _omml_text(' = ')
        + _omml_sub(_omml_text('Welfare'), _omml_text('free market'))
        + _omml_text(' − ')
        + _omml_sub(_omml_text('Welfare'), _omml_text('regulation')),
        size_pt=26, color=NAVY, fill=CREAM, line=NAVY, rounded=True,
        shadow=True)
    _add_text(slide, Inches(1.20), Inches(5.98), Inches(11.0), Inches(0.40),
              "with welfare measured as CS + PS", size=19, italic=True,
              color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_60_sales_tax(prs, page_num):
    """A 10 % sales tax modelled as an upward shift of supply.  Every
    marked price and quantity is the exact intersection of the lines drawn.
    """

    def draw(slide):
        fig = SimpleFig(1.45, 6.05, 5.60, 4.30, xmax=11.0, ymax=11.0)
        _fig_axes(slide, fig, x_title="Q", y_title="Price ($)",
                  label_size=18)
        # a wide wedge: the A / B / C / D areas have to be big enough to
        # carry a readable label, and the axes are schematic anyway
        t = 2.0

        def s0(q):
            return 2.0 + 0.8 * q

        def s1(q):
            return s0(q) + t

        def dd(q):
            return 10.0 - q

        q0, p0 = 4.4444, 5.5556          # 10 − q = 2 + 0.8q
        q1, pb = 3.3333, 6.6667          # 10 − q = 4 + 0.8q
        ps = pb - t

        _fig_poly(slide, fig, [(0, p0), (0, pb), (q1, pb), (q1, p0)],
                  fill=CREAM, line=NAVY)
        _fig_poly(slide, fig, [(0, ps), (0, p0), (q1, p0), (q1, ps)],
                  fill=CREAM, line=NAVY)
        _fig_poly(slide, fig, [(q1, p0), (q1, pb), (q0, p0)],
                  fill=PINK, line=NAVY)
        _fig_poly(slide, fig, [(q1, ps), (q1, p0), (q0, p0)],
                  fill=PINK, line=NAVY)

        _fig_line(slide, fig, (0, s0(0)), (8.6, s0(8.6)), color=NAVY,
                  weight_pt=2.75)
        _fig_curve_label(slide, fig, 8.5, s0(8.6) + 0.35, "S", size=19)
        _fig_line(slide, fig, (0, s1(0)), (6.6, s1(6.6)), color=GOLD,
                  weight_pt=2.75)
        _fig_curve_label(slide, fig, 6.5, s1(6.6) + 0.35, "S’", size=19,
                         color=GOLD)
        _fig_line(slide, fig, (0, dd(0)), (9.6, dd(9.6)), color=NAVY,
                  weight_pt=2.75)
        _fig_curve_label(slide, fig, 9.5, dd(9.6) + 0.5, "D", size=19)

        for p, lab in ((pb, "PB"), (p0, "P0"), (ps, "PS")):
            _fig_line(slide, fig, (0, p), (q0 if p == p0 else q1, p),
                      color=GRAY, weight_pt=1.4, dash="dash")
            _fig_ylab(slide, fig, p, lab, size=17, bold=True)
        for q, lab in ((q1, "Q1"), (q0, "Q0")):
            _fig_line(slide, fig, (q, 0), (q, dd(q)), color=GRAY,
                      weight_pt=1.4, dash="dash")
            _xlab_n(slide, fig, q, lab)
        for lx, ly, lab in ((1.5, (p0 + pb) / 2, "A"), (1.5, (ps + p0) / 2, "C"),
                            (q1 + 0.34, p0 + 0.30, "B"),
                            (q1 + 0.34, p0 - 0.42, "D")):
            _add_text(slide, fig.x(lx) - Inches(0.25),
                      fig.y(ly) - Inches(0.17), Inches(0.5), Inches(0.34),
                      lab, size=17, bold=True, color=NAVY, font="Calibri",
                      align=PP_ALIGN.CENTER)

        _add_convention_box(
            slide, Inches(7.70), Inches(1.60), Inches(5.15), Inches(1.55),
            body="The price the seller keeps (PS) falls and the price the "
                 "buyer pays (PB) rises — but by less than the tax.",
            corner_pct=0.10, size=18)
        _add_hierarchical_bullets(
            slide, Inches(7.90), Inches(3.45), Inches(4.95), Inches(1.90),
            [("Buyers lose A + B", 0), ("Sellers lose C + D", 0),
             ("The government collects A + C", 0),
             ("Deadweight loss: B + D", 0)],
            size=21, line_spacing_pts=10)
        _add_text(slide, Inches(7.90), Inches(5.55), Inches(4.95),
                  Inches(0.70),
                  "A tax on gasoline can still make sense — because of "
                  "the externality, which is next", size=17, italic=True,
                  color=GRAY, font="Calibri")

    slide = make_diagram_slide(
        prs, page_num, TAG_DIST,
        "Sales Tax: a 10 % Tax on Top of the Listed Price", draw)
    _add_ps_pointer(slide, label="Problem Set 3")
    _set_notes(slide, NOTES[58])
    return slide


def slide_61_tax_incidence(prs, page_num):
    """Adopted from MW slide 65.  In my deck this argument lived only in
    the speaker notes of the sales-tax slide."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_DIST)
    _draw_action_title(slide, "Tax Incidence: Who Really Pays?")
    _add_convention_box(
        slide, Inches(1.20), Inches(1.62), Inches(11.0), Inches(1.05),
        prefix="Tax incidence:  ",
        body="which side of the market — buyers or sellers — ends up "
             "bearing the burden of the tax",
        corner_pct=0.10, size=21)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.35), Inches(3.05), Inches(12.2),
        Inches(2.85),
        [("Who hands the money to the government does not matter", 0),
         ("The incidence is the same whether the tax is collected from "
          "the buyer or from the seller", 1),
         ("What matters is the elasticity of supply and demand", 0),
         ("The side that responds less to price bears more of the tax", 1)],
        size=26, sub_size=23, line_spacing_pts=16)
    _add_takeaway_bar(
        slide, "The less price-responsive side of the market pays the "
               "larger share",
        top=Inches(6.20), width=Inches(9.8), height=Inches(0.58),
        left=(SLIDE_W - Inches(9.8)) // 2, fill=GOLD, text_color=NAVY,
        size=20, bold=True, rounded=True, shadow=True)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_62_minwage_map(prs, page_num):
    """Adopted from MW slide 53; the map is re-fetched from Wikimedia
    Commons rather than lifted, and is the July 2026 edition."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_DIST)
    _draw_action_title(slide, "The Minimum Wage: A Popular Policy Tool")
    _add_media_image(slide, "web_minwage_map_commons.png",
                     left=Inches(2.85), top=Inches(1.55),
                     height=Inches(4.95), width=Inches(7.43))
    _add_text(slide, Inches(2.85), Inches(6.60), Inches(7.43), Inches(0.28),
              "Map: Wikimedia Commons, from US Department of Labor data",
              size=11, italic=True, color=GRAY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_63_min_wage(prs, page_num):
    """The labour-market price floor, with the circumvention card adopted
    from MW slide 52."""

    def draw(slide):
        fig = SimpleFig(1.30, 6.15, 4.55, 4.15, xmax=11.0, ymax=11.0)
        _fig_axes(slide, fig, x_title="Labor", y_title="Wage",
                  label_size=18)
        w_star, l_star = 5.5, 3.75
        w_min = 7.0
        l_d, l_s = (10.0 - w_min) / 1.2, (w_min - 1.0) / 1.2

        _fig_poly(slide, fig, [(0, w_star), (0, w_min), (l_d, w_min),
                               (l_d, w_star)], fill=CREAM, line=NAVY)
        _fig_poly(slide, fig, [(l_d, w_star), (l_d, w_min),
                               (l_star, w_star)], fill=PINK, line=NAVY)
        _fig_poly(slide, fig, [(l_d, w_star), (l_star, w_star),
                               (l_d, (1.0 + 1.2 * l_d))], fill=PINK,
                  line=NAVY)

        _fig_line(slide, fig, (0.0, 10.0), (8.0, 0.4), color=NAVY,
                  weight_pt=2.75)
        _fig_curve_label(slide, fig, 7.7, 1.0, "DLabor", size=17)
        _fig_line(slide, fig, (0.0, 1.0), (8.0, 10.6), color=NAVY,
                  weight_pt=2.75)
        _fig_curve_label(slide, fig, 7.5, 10.9, "SLabor", size=17)
        _fig_line(slide, fig, (0, w_star), (l_star, w_star), color=GRAY,
                  weight_pt=1.4, dash="dash")
        _fig_line(slide, fig, (l_star, 0), (l_star, w_star), color=GRAY,
                  weight_pt=1.4, dash="dash")
        _fig_ylab(slide, fig, w_star, "w*", size=17, bold=True)
        _xlab_n(slide, fig, l_star, "L*")
        _fig_line(slide, fig, (0, w_min), (l_s, w_min), color=DARKRED,
                  weight_pt=2.75)
        _fig_ylab(slide, fig, w_min, "wmin", size=17, bold=True)
        for l, lab in ((l_d, "Ld"), (l_s, "Ls")):
            _fig_line(slide, fig, (l, 0), (l, w_min), color=GRAY,
                      weight_pt=1.4, dash="dash")
            _xlab_n(slide, fig, l, lab)
        # the brace sits ABOVE the price floor, clear of the x labels
        _add_arrow(slide, (fig.x(l_d), fig.y(w_min + 0.75)),
                   (fig.x(l_s), fig.y(w_min + 0.75)),
                   color=DARKRED, weight_pt=2.0, head=True)
        _add_text(slide, fig.x(l_d) - Inches(0.55),
                  fig.y(w_min + 0.80) - Inches(0.34), Inches(2.6),
                  Inches(0.32), "Ls − Ld = unemployment", size=14,
                  bold=True, color=DARKRED, font="Calibri",
                  align=PP_ALIGN.CENTER)
        for lx, ly, lab in ((1.2, 6.25, "A"), (l_d + 0.30, 6.15, "B"),
                            (l_d + 0.30, 4.95, "C")):
            _add_text(slide, fig.x(lx) - Inches(0.25),
                      fig.y(ly) - Inches(0.17), Inches(0.5), Inches(0.34),
                      lab, size=17, bold=True, color=NAVY, font="Calibri",
                      align=PP_ALIGN.CENTER)

        _add_hierarchical_bullets(
            slide, Inches(6.55), Inches(1.62), Inches(6.35), Inches(2.20),
            [("Welfare effects", 0, {'bold': True}),
             ("Some workers win A — they are paid more", 1),
             ("Some workers lose C — they are not hired", 1),
             ("Firms lose A + B", 1),
             ("Deadweight loss: B + C", 1)],
            size=21, sub_size=19, line_spacing_pts=8)
        _add_convention_box(
            slide, Inches(6.55), Inches(4.15), Inches(6.35), Inches(2.05),
            runs=[("Circumvention?", {'bold': True, 'size': 20}),
                  ("\nUnderground labor market\n"
                   "Cutting non-wage benefits\n"
                   "Unpaid internships", {'size': 18})],
            corner_pct=0.10, size=19)

    slide = make_diagram_slide(prs, page_num, TAG_DIST,
                               "Price Floor: The Minimum Wage", draw)
    return slide


def slide_64_minwage_evidence(prs, page_num):
    """My evidence bullets, with the two headline clippings and the
    three-perspective discussion prompt adopted from MW slide 54."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_DIST)
    _draw_action_title(slide,
                       "What Does the Evidence Say on the Minimum Wage?")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.55), Inches(6.30),
        Inches(3.95),
        [("Theory predicts a negative effect on employment", 0),
         ("The evidence is mixed", 0),
         ("Employment itself does not seem to fall much", 1),
         ("But hours worked decline (the Seattle study)", 1),
         ("Experienced workers: hours fall moderately, so income rises", 1),
         ("Low-skilled workers: hours and income both fall", 1),
         ("New entrants find the first job harder to get", 1),
         ("What are the long-run effects of a high minimum wage?", 0)],
        size=21, sub_size=19, line_spacing_pts=9)
    _add_convention_box(
        slide, MARGIN + Inches(0.15), Inches(5.70), Inches(6.30),
        Inches(1.10),
        runs=[("How much is too much — from whose seat?",
               {'bold': True, 'size': 19}),
              ("\nBusiness owners  ·  workers  ·  consumers", {'size': 18})],
        corner_pct=0.10, size=19)
    _add_media_image(slide, "mw_headline_krueger_nyt_2015.png",
                     left=Inches(6.95), top=Inches(1.62), width=Inches(5.90),
                     rounded=False, shadow=True)
    _add_media_image(slide, "mw_headline_vox_2017.png",
                     left=Inches(6.95), top=Inches(3.30), width=Inches(5.90),
                     rounded=False, shadow=True)
    _add_media_image(slide, "NV_s59_3_3e55694d.png",
                     left=Inches(6.95), top=Inches(4.85),
                     height=Inches(1.85), width=Inches(2.73))
    _add_text(slide, Inches(9.95), Inches(5.35), Inches(2.90), Inches(0.90),
              "One way firms respond to a higher wage floor",
              size=17, italic=True, color=GRAY, font="Calibri")
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[59])
    return slide


def slide_65_guess_rent(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_DIST)
    _draw_action_title(slide, "Guess the Rent!")
    _add_media_image(slide, "NV_s61_6_fb7ae0ee.png",
                     left=Inches(0.55), top=Inches(1.65),
                     height=Inches(4.75), width=Inches(5.60))
    _add_hierarchical_bullets(
        slide, Inches(7.90), Inches(2.10), Inches(5.00), Inches(3.60),
        [("1 bed, 1 bath", 0), ("900 sqft", 0),
         ("401 San Vicente Blvd., Santa Monica", 0),
         ("5-minute walk to the beach", 0),
         ("Best public school district in West LA", 0)],
        size=24, line_spacing_pts=16)
    _add_pollbreak_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_67_the_rent_is(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_DIST)
    _draw_action_title(slide, "The Rent Is…")
    # my original composition, centred on the wider canvas: the browser
    # chrome, the rent-board page under it, and the current figures zoomed
    # in over its right-hand side
    _add_media_image(slide, "NV_s63_g3.1_6544d890.png",
                     left=Inches(2.16), top=Inches(1.45),
                     width=Inches(9.02), height=Inches(1.55),
                     rounded=False, shadow=False)
    _add_media_image(slide, "NV_s63_g3.2_9862f8e4.png",
                     left=Inches(2.14), top=Inches(3.00),
                     width=Inches(9.06), height=Inches(3.33),
                     rounded=False, shadow=False)
    _add_media_image(slide, "NV_s63_5_d33bea6f.png",
                     left=Inches(6.25), top=Inches(3.85),
                     width=Inches(5.23), height=Inches(2.37),
                     rounded=False, shadow=True)
    _add_text(slide, Inches(0.45), Inches(6.55), Inches(7.20), Inches(0.30),
              "Source: City of Santa Monica Rent Control Board  —  "
              "rent checked on 10/22/2025", size=12, italic=True,
              color=GRAY, font="Calibri")
    _add_pollbreak_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[63])
    return slide


def slide_68_rent_control(prs, page_num):
    def draw(slide):
        fig = SimpleFig(1.30, 6.15, 4.55, 4.15, xmax=11.0, ymax=11.0)
        _fig_axes(slide, fig, x_title="Q", y_title="P", label_size=18)
        p_max = 3.5
        q_s, q_d = (p_max - 1.0) / 1.2, (10.0 - p_max) / 1.2

        _fig_poly(slide, fig, [(0, p_max), (0, SD_PSTAR),
                               (q_s, SD_PSTAR), (q_s, p_max)],
                  fill=CREAM, line=NAVY)
        _fig_poly(slide, fig, [(q_s, SD_PSTAR), (SD_QSTAR, SD_PSTAR),
                               (q_s, 10.0 - q_s)], fill=PINK, line=NAVY)
        _fig_poly(slide, fig, [(q_s, SD_PSTAR), (SD_QSTAR, SD_PSTAR),
                               (q_s, p_max)], fill=PINK, line=NAVY)

        _sd_curves(slide, fig)
        _fig_line(slide, fig, (0, SD_PSTAR), (SD_QSTAR, SD_PSTAR),
                  color=GRAY, weight_pt=1.4, dash="dash")
        _fig_line(slide, fig, (SD_QSTAR, 0), (SD_QSTAR, SD_PSTAR),
                  color=GRAY, weight_pt=1.4, dash="dash")
        _fig_ylab(slide, fig, SD_PSTAR, "P*", size=17, bold=True)
        _xlab_n(slide, fig, SD_QSTAR, "Q*")
        _fig_line(slide, fig, (0, p_max), (q_d, p_max), color=DARKRED,
                  weight_pt=2.75)
        _fig_ylab(slide, fig, p_max, "Pmax", size=17, bold=True)
        for q, lab in ((q_s, "Qs"), (q_d, "Qd")):
            _fig_line(slide, fig, (q, 0), (q, p_max), color=GRAY,
                      weight_pt=1.4, dash="dash")
            _xlab_n(slide, fig, q, lab)
        _add_arrow(slide, (fig.x(q_s), fig.y(1.55)),
                   (fig.x(q_d), fig.y(1.55)), color=DARKRED,
                   weight_pt=2.0, head=True)
        _add_text(slide, fig.x(q_s) - Inches(0.5),
                  fig.y(1.55) + Inches(0.08), Inches(2.6), Inches(0.32),
                  "Qd − Qs = shortage", size=14, bold=True, color=DARKRED,
                  font="Calibri", align=PP_ALIGN.CENTER)
        for lx, ly, lab in ((1.0, 4.45, "A"), (q_s + 0.32, 6.15, "B"),
                            (q_s + 0.32, 4.95, "C")):
            _add_text(slide, fig.x(lx) - Inches(0.25),
                      fig.y(ly) - Inches(0.17), Inches(0.5), Inches(0.34),
                      lab, size=17, bold=True, color=NAVY, font="Calibri",
                      align=PP_ALIGN.CENTER)

        _add_hierarchical_bullets(
            slide, Inches(6.55), Inches(1.75), Inches(6.35), Inches(3.60),
            [("Welfare effects", 0, {'bold': True}),
             ("Some renters win A — those lucky enough to hold a place", 1),
             ("Some renters lose B — they cannot find one, because "
              "fewer are listed", 1),
             ("Landlords lose A + C", 1),
             ("A: the rent they no longer collect", 2),
             ("C: the flats they stop listing at the lower price", 2),
             ("Deadweight loss: B + C", 1)],
            size=21, sub_size=19, line_spacing_pts=8)

    slide = make_diagram_slide(prs, page_num, TAG_DIST,
                               "Price Ceiling: Rent Control", draw)
    return slide


def slide_69_prop33(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_DIST)
    _draw_action_title(slide, "Vote on Prop 33 in California, 2024")
    _add_media_image(slide, "NV_s65_5_2f36cabb.png",
                     left=Inches(3.55), top=Inches(1.55),
                     height=Inches(2.75), width=Inches(6.20))
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.35), Inches(4.55), Inches(11.4),
        Inches(1.90),
        [("California limits how far cities may go with rent control; "
          "Prop 33 would have removed those limits", 0),
         ("Cities could then have controlled rents on any housing — "
          "single-family homes, new apartments, and new tenants", 0),
         ("It did not pass", 0, {'bold': True})],
        size=21, line_spacing_pts=10)
    _add_text(slide, Inches(0.45), Inches(6.62), Inches(8.4), Inches(0.28),
              "Source: calmatters.org, California voter guide 2024",
              size=11, italic=True, color=GRAY, font="Calibri")
    _add_pollbreak_badge(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[65])
    return slide


def slide_71_landlord_reaction(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_DIST)
    _draw_action_title(slide, "How Do Landlords React to Rent Control?")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.75), Inches(5.90),
        Inches(2.60),
        [("Non-price adjustments — let maintenance slide", 0),
         ("Buy the tenants out (legal)", 0),
         ("Harass the tenants (illegal)", 0)],
        size=25, line_spacing_pts=20)
    _add_media_image(slide, "NV_s67_5_9b9eb200.jpg",
                     left=Inches(0.55), top=Inches(3.95),
                     height=Inches(2.55), width=Inches(3.40))
    _add_media_image(slide, "NV_s67_4_868e46cb.jpg",
                     left=Inches(7.20), top=Inches(1.62),
                     height=Inches(4.80), width=Inches(3.59))
    _add_text(slide, Inches(0.55), Inches(6.58), Inches(3.40), Inches(0.30),
              "Deferred maintenance", size=12, italic=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _add_text(slide, Inches(7.20), Inches(6.50), Inches(3.59), Inches(0.30),
              "A window left taped over", size=12, italic=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_72_cambridge_map(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_DIST)
    _draw_action_title(slide, "Evidence on Rent Control: Cambridge, MA")
    _add_text(slide, MARGIN + Inches(0.15), Inches(1.52), Inches(12.4),
              Inches(0.42),
              "A natural experiment: rent control in Cambridge ended "
              "suddenly in 1995", size=23, color=NAVY, font="Calibri")
    _add_media_image(slide, "NV_s68_5_5777e352.png",
                     left=Inches(3.35), top=Inches(2.10), width=Inches(6.65))
    _add_text(slide, Inches(3.35), Inches(6.35), Inches(6.65), Inches(0.30),
              "Source: Autor, Palmer and Pathak (2014)", size=12,
              italic=True, color=GRAY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    _set_notes(slide, NOTES[68])
    return slide


def slide_73_rent_questions(prs, page_num):
    bullets = [
        ("Rent control in Cambridge ended suddenly in 1995. What happened "
         "to…", 0),
        ("the rent of the units that had been controlled?", 1),
        ("investment in those units?", 1),
        ("the price of and investment in nearby units that were never "
         "controlled?", 1),
    ]
    slide = content_slide(prs, page_num, TAG_DIST,
                          "What Happened When Rent Control Ended?", bullets,
                          size=27, sub_size=25, line_spacing_pts=20)
    _set_notes(slide, NOTES[69])
    return slide


def slide_74_rent_results(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_DIST)
    _draw_action_title(slide, "Evidence on Rent Control: Results")
    panels = [
        ("NV_s70_2_95e7216e.png", 0.45, 1.95, 3.90, 2.43,
         "Price of decontrolled units"),
        ("NV_s70_4_5ecaba44.png", 4.75, 1.95, 3.30, 2.43,
         "Price of never-controlled units"),
        ("NV_s70_6_08feed5a.png", 8.90, 1.95, 3.20, 3.72,
         "Investment at decontrolled units ($1,000s)"),
    ]
    for fn, x, y, w, h, cap in panels:
        _add_text(slide, Inches(x - 0.35), Inches(y - 0.50),
                  Inches(w + 0.70), Inches(0.44), cap, size=15, bold=True,
                  italic=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)
        _add_media_image(slide, fn, left=Inches(x), top=Inches(y),
                         width=Inches(w), height=Inches(h),
                         rounded=False, shadow=True)
    _add_text(slide, Inches(0.45), Inches(5.85), Inches(7.60), Inches(0.30),
              "Source: Autor, Palmer and Pathak (2014)", size=12,
              italic=True, color=GRAY, font="Calibri")
    _add_takeaway_bar(
        slide, "Rents rose, and so did investment — including at the units "
               "next door",
        top=Inches(6.45), width=Inches(9.8), height=Inches(0.55),
        left=(SLIDE_W - Inches(9.8)) // 2, fill=GOLD, text_color=NAVY,
        size=19, bold=True, rounded=True, shadow=True)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_75_argentina(prs, page_num):
    """Adopted from MW slide 63: a second, very recent natural experiment
    beside Cambridge 1995.  The magnitudes are press-reported, and the
    slide says so."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_DIST)
    _draw_action_title(slide, "Argentina Scrapped Its Rent Controls")
    _add_media_image(slide, "mw_headline_wsj_argentina.png",
                     left=Inches(1.55), top=Inches(1.65), width=Inches(10.20),
                     rounded=False, shadow=True)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.55), Inches(3.90), Inches(11.6),
        Inches(1.90),
        [("Argentina repealed its rent-control law at the end of 2023", 0),
         ("Reported since: many more flats listed, and rents lower "
          "in real terms", 0),
         ("Not everyone gains — sitting tenants lost their protection", 0)],
        size=24, line_spacing_pts=16)
    _add_text(slide, Inches(0.45), Inches(6.35), Inches(9.5), Inches(0.28),
              "Magnitudes as reported in the press, not from a study. "
              "Source: The Wall Street Journal", size=11, italic=True,
              color=GRAY, font="Calibri")
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


# ==========================================================================
#  4 · EXTERNALITIES — slides 76 – 83
# ==========================================================================

def slide_77_externalities_def(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_EXT)
    _draw_action_title(slide, "Externalities: Definitions")
    _add_convention_box(
        slide, Inches(1.00), Inches(1.60), Inches(11.3), Inches(1.00),
        prefix="Externality:  ",
        body="a cost or a benefit that falls on someone who is not part "
             "of the transaction",
        corner_pct=0.10, size=22)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.35), Inches(3.00), Inches(12.2),
        Inches(3.30),
        [([("Negative externality", {'bold': True, 'color': BLUE_PED}),
           (" — a cost imposed on an outsider", {})], 0, {}),
         ("Air pollution from coal-fired power plants", 1),
         ("The external cost of gasoline is about $1.50 per gallon", 1),
         ([("Positive externality", {'bold': True, 'color': BLUE_PED}),
           (" — a benefit conferred on an outsider", {})], 0, {}),
         ("Your neighbour's well-kept front yard is a pleasure for you "
          "to look at too", 1)],
        size=24, sub_size=22, line_spacing_pts=16)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_78_correcting(prs, page_num):
    """Adopted from MW slide 67, with one worked example for each kind of
    quota (2026-08-28, Nico)."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_EXT)
    _draw_action_title(slide, "Correcting for Externalities")
    _add_text(slide, MARGIN + Inches(0.15), Inches(1.50), Inches(8.30),
              Inches(0.75),
              "The aim of the intervention: move the market toward the "
              "quantity that would be chosen if the outside cost or "
              "benefit were counted", size=21, color=NAVY, font="Calibri")
    _add_rounded_filled_box(
        slide, Inches(0.55), Inches(2.55), Inches(8.30), Inches(0.55),
        "Price mechanisms:  taxes and subsidies", fill=NAVY,
        text_color=WHITE, size=20, corner_pct=0.12)
    _add_hierarchical_bullets(
        slide, Inches(0.85), Inches(3.20), Inches(8.00), Inches(0.90),
        [("Tax the activity that imposes the cost, subsidize the one "
          "that confers the benefit", 0)],
        size=20, line_spacing_pts=0)
    _add_rounded_filled_box(
        slide, Inches(0.55), Inches(4.15), Inches(8.30), Inches(0.55),
        "Quantity mechanisms:  quotas and mandates", fill=NAVY,
        text_color=WHITE, size=20, corner_pct=0.12)
    _add_hierarchical_bullets(
        slide, Inches(0.85), Inches(4.85), Inches(8.00), Inches(1.60),
        [([("Negative externality — cap it:", {'bold': True}),
           (" a limit on the tonnes of sulphur dioxide a power plant may "
            "emit in a year", {})], 0, {}),
         ([("Positive externality — require it:", {'bold': True}),
           (" childhood vaccination as a condition of school entry", {})],
          0, {})],
        size=19, line_spacing_pts=12)
    _add_media_image(slide, "NV_s73_4_e69d8d25.png",
                     left=Inches(9.35), top=Inches(1.55), width=Inches(3.20))
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


# The gasoline externality figure, shared by slides 79 and 80.
EMC = 1.5


def _gas_d(q):
    return 8.0 - 0.6 * q


def _gas_mci(q):
    return 1.0 + 0.55 * q


def _gas_smc(q):
    return _gas_mci(q) + EMC


GAS_Q_MARKET = 7.0 / 1.15          # 8 − 0.6q = 1 + 0.55q
GAS_Q_EXT = 5.5 / 1.15             # 8 − 0.6q = 2.5 + 0.55q


def _draw_gas_fig(slide, *, with_tax):
    fig = SimpleFig(2.30, 6.35, 6.15, 4.35, xmax=9.4, ymax=9.0)
    _fig_axes(slide, fig, x_title="Gallons of gasoline per year",
              y_title="$ / gallon", label_size=17)
    _fig_line(slide, fig, (0, _gas_d(0)), (8.8, _gas_d(8.8)), color=NAVY,
              weight_pt=2.75)
    _fig_curve_label(slide, fig, 8.30, _gas_d(8.8) + 1.15, "Demand",
                     size=18)
    _fig_line(slide, fig, (0, _gas_mci(0)), (8.8, _gas_mci(8.8)),
              color=NAVY, weight_pt=2.75)
    _fig_line(slide, fig, (0, _gas_smc(0)), (8.0, _gas_smc(8.0)),
              color=GOLD, weight_pt=2.75)
    for q, p, lab, col in (
            (GAS_Q_MARKET, _gas_d(GAS_Q_MARKET), "QMarket", NAVY),
            (GAS_Q_EXT, _gas_d(GAS_Q_EXT), "QExt", GOLD)):
        _fig_line(slide, fig, (0, p), (q, p), color=GRAY, weight_pt=1.4,
                  dash="dash")
        _fig_line(slide, fig, (q, 0), (q, p), color=GRAY, weight_pt=1.4,
                  dash="dash")
        _xlab_n(slide, fig, q, lab, size=16, w=0.95)
        _fig_dot(slide, fig, q, p, d=Inches(0.13), fill=col)
    _add_text(slide, Inches(8.60), fig.y(_gas_smc(8.0)) - Inches(0.46),
              Inches(4.35), Inches(0.62),
              "MCI + EMC = SMC\n(social marginal cost)", size=17, bold=True,
              color=GOLD, font="Calibri")
    _add_text(slide, Inches(8.60), fig.y(_gas_mci(8.8)) + Inches(0.02),
              Inches(4.35), Inches(0.62),
              "Supply = internal\nmarginal cost (MCI)", size=17, bold=True,
              color=NAVY, font="Calibri")
    # the EMC wedge, drawn where the two supply curves are furthest apart
    qw = 2.2
    _add_arrow(slide, (fig.x(qw), fig.y(_gas_mci(qw))),
               (fig.x(qw), fig.y(_gas_smc(qw))), color=DARKRED,
               weight_pt=2.5, head=True)
    _add_text(slide, fig.x(qw) + Inches(0.12),
              fig.y((_gas_mci(qw) + _gas_smc(qw)) / 2) - Inches(0.17),
              Inches(3.0), Inches(0.34),
              ("Tax = EMC = $1.50" if with_tax
               else "EMC = $1.50 per gallon"),
              size=17, bold=True, color=DARKRED, font="Calibri")
    return fig


def slide_79_gas_wedge(prs, page_num):
    """Adopted from MW slide 68: the diagnosis on its own slide, before the
    cure.  My original carried both on one slide."""

    def draw(slide):
        _draw_gas_fig(slide, with_tax=False)
        _add_convention_box(
            slide, Inches(8.75), Inches(4.35), Inches(4.15), Inches(1.85),
            body="Drivers weigh only their own cost, so the market settles "
                 "at QMarket. Counting the cost borne by everyone else, "
                 "the efficient quantity is the smaller QExt.",
            corner_pct=0.10, size=18)

    slide = make_diagram_slide(prs, page_num, TAG_EXT,
                               "The Externality Wedge: Gasoline", draw)
    return slide


def slide_80_carbon_tax(prs, page_num):
    def draw(slide):
        _draw_gas_fig(slide, with_tax=True)
        _add_media_image(slide, "NV_s74_g13.g2.1_862b02c9.jpg",
                         left=Inches(11.25), top=Inches(1.60),
                         width=Inches(1.55))
        _add_text(slide, Inches(10.95), Inches(4.05), Inches(2.15),
                  Inches(0.30), "Arthur Pigou", size=13, italic=True,
                  color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
        _add_convention_box(
            slide, Inches(8.75), Inches(4.65), Inches(4.15), Inches(1.55),
            body="A tax equal to the external cost makes each driver face "
                 "the social marginal cost, and the market lands on QExt "
                 "on its own.",
            corner_pct=0.10, size=18)

    slide = make_diagram_slide(
        prs, page_num, TAG_EXT,
        "A Carbon Tax to Correct the Negative Externality", draw)
    _set_notes(slide, NOTES[74])
    return slide


def slide_81_airplane_noise(prs, page_num):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_EXT)
    _draw_action_title(slide, "Externality: Airplane Noise")
    _add_media_image(slide, "NV_s75_3_4d362b83.wmf",
                     left=Inches(0.60), top=Inches(1.60), width=Inches(7.60),
                     rounded=False, shadow=False)
    # 2026-08-28: Nico replaced the screenshot that was here with this
    # Santa Monica Lookout article, at exactly this position and size
    # (hand-edit ported from the canonical deck).
    _add_media_image(slide, "nv_smo_lookout_article.png",
                     left=Inches(8.21), top=Inches(4.10),
                     width=Inches(5.13), height=Inches(3.06),
                     rounded=False, shadow=True)
    # the caption sits right under the map so the grouping pass pairs the
    # two and they reveal on one click (course rule: a picture and its
    # caption are ONE object)
    _add_text(slide, Inches(0.60), Inches(5.66), Inches(7.60), Inches(0.30),
              "Noise contours around Santa Monica Airport", size=12,
              italic=True, color=GRAY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def slide_82_smo_corner(prs, page_num):
    """The corner solution: once the noise cost is added, social marginal
    cost lies above willingness to pay at every quantity, so Q* = 0."""

    def draw(slide):
        fig = SimpleFig(1.85, 6.15, 6.30, 4.10, xmax=9.4, ymax=9.0)
        _fig_axes(slide, fig, x_title="Flights per day", y_title="Price",
                  label_size=17)
        mci, smc = 2.0, 6.0
        _fig_line(slide, fig, (0, 5.0), (8.6, 0.5), color=NAVY,
                  weight_pt=2.75)
        _add_text(slide, Inches(8.45), fig.y(0.9) - Inches(0.30),
                  Inches(3.6), Inches(0.62), "D  (MB of flying)", size=17,
                  bold=True, color=NAVY, font="Calibri")
        _fig_line(slide, fig, (0, mci), (8.6, mci), color=NAVY,
                  weight_pt=2.5)
        _add_text(slide, Inches(8.45), fig.y(mci) - Inches(0.30),
                  Inches(4.30), Inches(0.62),
                  "Pilots' internal\nmarginal cost (MCI)", size=17,
                  bold=True, color=NAVY, font="Calibri")
        _fig_line(slide, fig, (0, smc), (8.6, smc), color=GOLD,
                  weight_pt=2.75)
        _add_text(slide, Inches(8.45), fig.y(smc) - Inches(0.30),
                  Inches(4.30), Inches(0.62),
                  "MCI + tax = SMC\n(social marginal cost)", size=17,
                  bold=True, color=GOLD, font="Calibri")
        _add_arrow(slide, (fig.x(2.4), fig.y(mci)), (fig.x(2.4), fig.y(smc)),
                   color=DARKRED, weight_pt=2.5, head=True)
        _add_text(slide, fig.x(2.55), fig.y((mci + smc) / 2) - Inches(0.17),
                  Inches(3.0), Inches(0.34), "Noise tax = EMC", size=17,
                  bold=True, color=DARKRED, font="Calibri")
        _fig_xlab(slide, fig, 0.0, "Q* = 0", size=17, bold=True)
        _fig_dot(slide, fig, 0.0, 5.0, d=Inches(0.15))

    slide = make_diagram_slide(
        prs, page_num, TAG_EXT,
        "A Pigouvian Tax for Noise? The Case of Santa Monica Airport", draw)
    _add_takeaway_bar(
        slide, "SMC lies above willingness to pay at every quantity — a "
               "corner solution at Q* = 0, so close the airport",
        top=Inches(6.42), width=Inches(11.4), height=Inches(0.58),
        left=(SLIDE_W - Inches(11.4)) // 2, fill=GOLD, text_color=NAVY,
        size=19, bold=True, rounded=True, shadow=True)
    return slide


def slide_83_summary(prs, page_num):
    bullets = [
        ("Perfect competition: the firm is a price taker, so MR = P", 0),
        ("Profit maximization: MC = MR = P gives Q*", 1),
        ("The decision to operate in the short run", 0),
        ("Operate if P ≥ AVC; shut down if P < AVC", 1),
        ("The decision to operate in the long run", 0),
        ("Operate if P ≥ LAC; exit or never enter if P < LAC", 1),
        ("Long-run equilibrium: P = LMC = LAC, so profit is zero", 1),
        ("Market distortions: price controls and taxes", 0),
        ("Winners, losers, and deadweight loss", 1),
        ("Externalities, and the Pigouvian tax that corrects them", 0),
    ]
    return content_slide(prs, page_num, TAG_SUMMARY, "Module 4: Summary",
                         bullets, size=23, sub_size=21,
                         line_spacing_pts=11)


# ==========================================================================
#  ASSEMBLY
# ==========================================================================

def _poll(prs, page_num, tag, title):
    """A marked placeholder for a PollEverywhere slide.

    The title carries POLL_MARK so the pending slides stand out when
    paging through the deck; the body says what has to be done.
    """
    return make_stub(prs, page_num, tag,
                     "%s  %s" % (title, POLL_MARK), STUB_POLL)


def _strip_unused_layouts(prs):
    """Drop the python-pptx template's unused slide layouts.

    Course rule: ONE slide master for the deck, with the template defaults
    stripped.  Every slide here is built on the blank layout, so the other
    ten are dead weight that invite drift the moment someone inserts a
    slide by hand.
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


def build(out_path=None):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    n = [0]

    def nxt():
        n[0] += 1
        return n[0]

    # ---- front matter --------------------------------------------------
    slide_01_title(prs); nxt()
    slide_02_logistics(prs, nxt())
    make_m4_roadmap(prs, nxt())
    make_m4_outline(prs, nxt(), descriptions=True)

    # ---- 1 · introduction to market structures -------------------------
    make_m4_outline(prs, nxt(), highlight_idx=0)
    slide_06_market_structures(prs, nxt())
    slide_07_market_power(prs, nxt())

    # ---- 2 · perfect competition ---------------------------------------
    make_m4_outline(prs, nxt(), highlight_idx=1)
    slide_09_price_taker(prs, nxt())
    slide_10_market_and_firm(prs, nxt())
    slide_11_farmers(prs, nxt())

    # ---- 2a · profit maximization in the short run ---------------------
    make_m4_outline(prs, nxt(), highlight_idx=2)
    slide_13_yi_family(prs, nxt())
    slide_14_cost_table(prs, nxt())
    slide_15_tc_chart(prs, nxt())
    slide_16_business_relevance(prs, nxt())
    slide_17_profit_max_rule(prs, nxt())
    slide_18_revenue_conditions(prs, nxt())
    slide_19_tr_tc_visual(prs, nxt())
    slide_20_max_profit_setup(prs, nxt())
    _poll(prs, nxt(), TAG_SR, "Poll: The Optimal Quantity")
    slide_22_qstar_solution(prs, nxt())
    slide_23_two_ways_atc(prs, nxt())
    slide_24_profit_solution(prs, nxt())
    slide_25_profit_rectangle(prs, nxt())
    slide_26_ross_stores(prs, nxt())
    slide_27_general_case(prs, nxt())
    slide_28_stop_producing(prs, nxt())
    slide_29_two_ways_avc(prs, nxt())
    slide_30_shutdown_rule(prs, nxt())
    slide_31_new_price(prs, nxt())
    _poll(prs, nxt(), TAG_SR, "Poll: The Optimal Quantity")
    slide_33_new_qstar(prs, nxt())
    slide_34_operate_solution(prs, nxt())
    slide_35_high_price(prs, nxt())
    slide_36_low_price(prs, nxt())
    slide_37_very_low_price(prs, nxt())
    slide_38_coffee(prs, nxt())
    _poll(prs, nxt(), TAG_SR, "Poll: The Coffee Bean Producer")
    slide_40_coffee_solution(prs, nxt())
    slide_41_sr_summary(prs, nxt())

    # ---- 2b · firm-level and market supply -----------------------------
    make_m4_outline(prs, nxt(), highlight_idx=3)
    slide_43_supply_curve(prs, nxt())
    slide_44_changing_mc(prs, nxt())
    slide_45_market_dynamics(prs, nxt())

    # ---- 2c · long-run competitive equilibrium -------------------------
    make_m4_outline(prs, nxt(), highlight_idx=4)
    slide_47_long_run(prs, nxt())
    slide_48_lr_equilibrium(prs, nxt())
    slide_49_chickens(prs, nxt())
    slide_50_lr_summary(prs, nxt())
    slide_51_drug_market(prs, nxt())
    _poll(prs, nxt(), TAG_LR, "Poll: Which Policy?")
    slide_53_arrest_dealers(prs, nxt())
    slide_54_arrest_users(prs, nxt())

    # ---- 3 · market distortions and regulations ------------------------
    make_m4_outline(prs, nxt(), highlight_idx=5)
    slide_56_distortions_intro(prs, nxt())
    slide_57_consumer_surplus(prs, nxt())
    slide_58_producer_surplus(prs, nxt())
    slide_59_deadweight_loss(prs, nxt())
    slide_60_sales_tax(prs, nxt())
    slide_61_tax_incidence(prs, nxt())
    slide_62_minwage_map(prs, nxt())
    slide_63_min_wage(prs, nxt())
    slide_64_minwage_evidence(prs, nxt())
    slide_65_guess_rent(prs, nxt())
    _poll(prs, nxt(), TAG_DIST, "Poll: Guess the Rent")
    slide_67_the_rent_is(prs, nxt())
    slide_68_rent_control(prs, nxt())
    slide_69_prop33(prs, nxt())
    _poll(prs, nxt(), TAG_DIST, "Poll: How Would You Vote?")
    slide_71_landlord_reaction(prs, nxt())
    slide_72_cambridge_map(prs, nxt())
    slide_73_rent_questions(prs, nxt())
    slide_74_rent_results(prs, nxt())
    slide_75_argentina(prs, nxt())

    # ---- 4 · externalities ---------------------------------------------
    make_m4_outline(prs, nxt(), highlight_idx=6)
    slide_77_externalities_def(prs, nxt())
    slide_78_correcting(prs, nxt())
    slide_79_gas_wedge(prs, nxt())
    slide_80_carbon_tax(prs, nxt())
    slide_81_airplane_noise(prs, nxt())
    slide_82_smo_corner(prs, nxt())
    slide_83_summary(prs, nxt())

    # deck-wide passes
    apply_symbol_subscripts(prs)
    _strip_unused_layouts(prs)

    out = Path(out_path) if out_path else OUT
    prs.save(str(out))
    print("%d slides -> %s" % (len(prs.slides._sldIdLst), out))
    return out


if __name__ == "__main__":
    import sys
    build(sys.argv[1] if len(sys.argv) > 1 else None)
