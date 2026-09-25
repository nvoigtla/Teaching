# -*- coding: utf-8 -*-
# ==========================================================================
#  _build_M6_practice.py — "Module 6 - Practice Video - Optimal Pricing in
#  Two Markets.pptx"
#
#  The full PG BMW example (PG1 31 – 38, 40) as its own 10-slide deck, for
#  retaping the practice video the course calendar lists as
#  "Practice Video: Optimal Pricing in two Markets" (m6p1, 19 min).
#
#  Requested by Nico 2026-09-09: "create a separate slide deck where you
#  adopt the full PG BMW example, including the illustrative slides that
#  you put in the main deck.  I may use this to tape a new version of the
#  practice video."  So slide 9 here is the same three-panel figure that
#  appears as slide 29 of the main deck — the deck has to stand alone.
#
#  Every number is DERIVED, not transcribed: the BMW constants and the
#  `_bmw_opt` solver live in _build_Module6.py, which also carries an
#  assert that segment profit beats uniform profit.  Importing that module
#  is safe — its work sits behind `if __name__ == "__main__"`.
#
#      python _build_M6_practice.py
# ==========================================================================

from pathlib import Path

from pptx import Presentation

import _m6_helpers as _H
import _build_Module6 as M6
from _m6_helpers import (
    CREAM, DARKRED, GOLD, GOLD_W, GRAY, MARGIN, NAVY, RULE, SLIDE_H, SLIDE_W,
    WHITE, Inches, Pt, PP_ALIGN, MSO_ANCHOR, MSO_SHAPE,
    _add_convention_box, _add_hierarchical_bullets, _add_math_equation,
    _add_media_image,
    _add_ps_pointer, _add_rect, _add_rounded_filled_box, _add_text,
    _blank_slide, _draw_action_title, _draw_footer, _draw_top_bar_tc,
    _omml_frac, _omml_sub, _omml_sup, apply_symbol_subscripts,
    content_slide, make_diagram_slide,
)

# the escaped OMML wrappers from the main build script (a bare "<" or ">"
# in a comparison would otherwise produce an invalid slide part)
_omml_run = M6._omml_run
_omml_text = M6._omml_text

OUT = (Path(__file__).resolve().parents[1]
       / "Module 6 - Practice Video - Optimal Pricing in Two Markets.pptx")

_H.FOOTER_TEXT = ("Management 405  ·  Module 6  ·  Practice Video: "
                  "Optimal Pricing in Two Markets")
FOOTER_TEXT = _H.FOOTER_TEXT

TAG = "Module 6 · Practice Video"
DECK_TITLE = "Optimal Pricing in Two Markets"
DECK_SUB = "Module 6  ·  Practice Video"

# --- the worked numbers, all derived in _build_Module6 -------------------
MC = M6.BMW_MC                       # 40
A_US, B_US = M6.BMW_US               # 160, 4
A_DE, B_DE = M6.BMW_DE               # 140, 5
Q_US, P_US = M6.BMW_QUS, M6.BMW_PUS  # 15, 100
Q_DE, P_DE = M6.BMW_QDE, M6.BMW_PDE  # 10, 90


# 2026-09-15 (Nico): "use US for US and D for Germany whenever you use
# abbreviations".  Every subscript in the deck comes from these two names
# rather than from the call sites, so the convention cannot drift.
K_US, K_DE = "US", "D"


def _elasticity(a, b, q, p):
    """E_D at the profit-maximising point for P = a − bQ.

    Q = a/b − (1/b)P, so dQ/dP = −1/b and E_D = (dQ/dP)(P/Q).
    """
    return (-1.0 / b) * (p / q)


def _lerner(p, mc=MC):
    return (p - mc) / p


E_US, E_DE = _elasticity(A_US, B_US, Q_US, P_US), _elasticity(A_DE, B_DE,
                                                              Q_DE, P_DE)
L_US, L_DE = _lerner(P_US), _lerner(P_DE)
# the Lerner identity must hold in both markets, or the slides disagree
assert abs(L_US - 1.0 / abs(E_US)) < 1e-9
assert abs(L_DE - 1.0 / abs(E_DE)) < 1e-9


def _eq(slide, y_in, omml, *, size=23, left=1.30, width=10.6, color=None):
    """`color` puts an ANSWER line in dark red -- one per solution slide,
    so it reads as the answer (2026-09-14, Nico)."""
    kw = {} if color is None else {"color": color}
    return _add_math_equation(slide, Inches(left), Inches(y_in),
                              Inches(width), Inches(0.55), omml,
                              size_pt=size, **kw)


# 2026-09-15 (Nico): "for all Your turn boxes, let's find a common
# format.  Let's use dark green line with dark green transparent fill."
DARKGREEN = _H.RGBColor(0x1B, 0x5E, 0x20)


def _yourturn(slide, y_in, height_in, body, *, left=0.85, width=11.6,
              size=21):
    """The deck's question prompt, in one place so the two instances
    cannot drift apart.  A light wash rather than a solid fill, so the
    dark green text stays legible on it."""
    return _add_convention_box(
        slide, Inches(left), Inches(y_in), Inches(width),
        Inches(height_in), prefix="Your turn:  ", body=body,
        # _set_fill_alpha takes OPACITY AS A PERCENT and multiplies by
        # 1000 itself -- not the 1/1000-percent value _fig_poly's
        # `alpha` takes.  Passing 14000 here wrote alpha=14000000,
        # which is well outside the 0-100000 range, and PowerPoint
        # refused to open the file while python-pptx read it happily.
        fill_rgb=DARKGREEN, fill_alpha=15, border=DARKGREEN,
        text_color=DARKGREEN, line_w=1.5, size=size,
        anchor=MSO_ANCHOR.MIDDLE)


def _section_header(slide, y_in, text, *, left=0.85, size=25):
    """A header INSIDE the body of a slide (2026-09-15, Nico: "on both
    slides, distinguish the 'what we know' more clearly so it's a
    header-type").  Set like the deck's action title one step down --
    navy bold with a short gold strip under it -- so it reads as a
    heading rather than as the first line of the list it introduces."""
    _add_text(slide, Inches(left), Inches(y_in), Inches(6.0),
              Inches(0.42), text, size=size, bold=True, color=NAVY,
              font="Calibri", align=PP_ALIGN.LEFT)
    _add_rect(slide, int(Inches(left)), int(Inches(y_in + 0.46)),
              int(Inches(1.05)), int(Inches(0.045)), GOLD)


def _lead(slide, y_in, text, *, size=22, left=0.80):
    return _add_text(slide, Inches(left), Inches(y_in), Inches(11.9),
                     Inches(0.42), text, size=size, bold=True, color=NAVY,
                     font="Calibri", align=PP_ALIGN.LEFT)


def _coef(c):
    """"1 · Q" is not how anyone writes Q.  Both demands now have
    b = 0.5, so 2b is exactly 1 and the MR lines printed a unit
    coefficient that reads as a slip rather than as algebra."""
    return "" if abs(c - 1.0) < 1e-9 else "%g · " % c


def _sub(v):
    return _omml_sub(_omml_run("Q"), _omml_text(v))


def _psub(v):
    return _omml_sub(_omml_run("P"), _omml_text(v))


# ==========================================================================
#  Slides
# ==========================================================================

def s01_card(prs):
    """The video title card: title-slide layout, no top bar, no footer text
    and no page number.  Carries no speaker notes, per the taping rule."""
    slide = _blank_slide(prs)
    # 2026-09-15 (Nico): "new image; all text boxes moved. Adopt this."
    # The whole text block rises 1.32" to free the bottom third for two
    # much larger cars, and the byline loses "· UCLA Anderson" -- his
    # edit, kept verbatim even though the other title cards carry the
    # full line.
    _add_text(slide, 0, Inches(0.78), SLIDE_W, Inches(1.1), DECK_TITLE,
              size=54, bold=True, color=NAVY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_text(slide, 0, Inches(1.98), SLIDE_W, Inches(0.75), DECK_SUB,
              size=36, bold=True, color=GOLD, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_rect(slide, int((SLIDE_W - Inches(4.0)) / 2), Inches(2.96),
              Inches(4.0), 54864, GOLD)
    _add_text(slide, 0, Inches(3.30), SLIDE_W, Inches(0.55),
              "Management 405", size=26, bold=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _add_text(slide, 0, Inches(4.00), SLIDE_W, Inches(0.37),
              "Prof. Nico Voigtländer", size=22,
              color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
    # his own two images, placed where he put them (2026-09-14).  The iX
    # on the right is a studio shot already standing on plain white, so it
    # is FLAT -- rounding and a drop shadow around its white margin would
    # read as a floating card rather than as a car.  The M2 on the left is
    # still a road photograph and keeps the deck's picture treatment until
    # its background comes off.
    _add_media_image(slide, "PV_s01_car.png", left=Inches(0.28),
                     top=Inches(4.19), width=Inches(4.77))
    # his replacement for the right-hand car, 2026-09-15, at his frame.
    # The image it replaced was a studio shot standing on plain white and
    # was therefore flat; this one is a road photograph like the M2 beside
    # it, so it takes the deck's ordinary picture treatment.
    _add_media_image(slide, "PV_s01_right2.png", left=Inches(8.21),
                     top=Inches(4.29), width=Inches(4.84))
    _add_rect(slide, 0, Inches(7.15), SLIDE_W, Inches(0.02), RULE)
    _add_rect(slide, MARGIN, Inches(7.135), GOLD_W, Inches(0.05), GOLD)
    return slide


def s02_setup(prs, n):
    """[NV 29] + [NEW – PG1 31] — the rule, before any numbers."""
    def draw(slide):
        _add_hierarchical_bullets(
            slide, Inches(0.80), Inches(1.75), Inches(11.9), Inches(1.5),
            [([("A firm produces the same product for two markets", {})], 0,
              {}),
             ([("Assume marginal cost is the same in both", {})], 0, {})],
            size=24, line_spacing_pts=13)
        _lead(slide, 3.45, "Profit maximization:")
        # 2026-09-14 (Nico): "red box with red background around the
        # formula" -- the same device the main deck now gives its own rule
        # equation on slide 39, so the two decks read as one.
        _add_math_equation(
            slide, Inches(4.16), Inches(3.86), Inches(5.00), Inches(0.66),
            _omml_sub(_omml_run("MR"), _omml_text("a")) + _omml_text(" = ")
            + _omml_sub(_omml_run("MR"), _omml_text("b"))
            + _omml_text("  ( = ") + _omml_run("MC") + _omml_text(" )"),
            size_pt=24, color=WHITE, fill=DARKRED, line=DARKRED,
            rounded=True, shadow=True, corner_pct=14000)
        _add_hierarchical_bullets(
            slide, Inches(0.80), Inches(4.75), Inches(11.9), Inches(0.8),
            [([("If the two marginal revenues differ, one price for both "
                "markets cannot be optimal", {})], 0, {})],
            size=21, line_spacing_pts=0)
        # his own size for it (2026-09-14): sized to the sentence rather
        # than spanning the slide
        _add_rounded_filled_box(
            slide, Inches(2.80), Inches(5.89), Inches(7.74), Inches(0.63),
            "RULE:  charge the higher price in the less elastic market",
            # 2026-09-15 (Nico): "make the box in the bottom dark blue
            # rather than yellow"
            fill=NAVY, text_color=WHITE, size=20, bold=True,
            corner_pct=0.14)
    return make_diagram_slide(prs, n, TAG, "Segment Pricing Analytically",
                              draw)


def s03_problem(prs, n):
    """[NEW – PG1 32] — the set-up, as a group exercise."""
    def draw(slide):
        # 2026-09-15 (Nico): "move the 'Your Turn' box to the top.  Then
        # say 'What we know', just like on slide 6."  The question now
        # frames the givens instead of arriving after them.
        _yourturn(slide, 1.55, 0.64,
                  "1. What is the profit-maximizing price in the US and "
                  "in Germany?")
        _section_header(slide, 2.52, "What we know:")
        _add_hierarchical_bullets(
            slide, Inches(0.85), Inches(3.14), Inches(11.6), Inches(0.6),
            [([("Q : number of cars (in 1,000)", {})], 0, {})],
            size=23, line_spacing_pts=0)
        _lead(slide, 3.78, "(Inverse) demand in the US:", size=23)
        _eq(slide, 4.26, _omml_run("P") + _omml_text(" = %g − %g" % (A_US, B_US))
            + _omml_run("Q"), size=26, left=5.10, width=6.0)
        _lead(slide, 4.96, "(Inverse) demand in Germany:", size=23)
        _eq(slide, 5.44, _omml_run("P") + _omml_text(" = %g − %g" % (A_DE, B_DE))
            + _omml_run("Q"), size=26, left=5.10, width=6.0)
        _lead(slide, 6.14, "MC = %g in both markets   [in $1,000]" % MC, size=23)
    return make_diagram_slide(
        prs, n, TAG, "BMW and Segmentation: Setup and Question 1", draw)


def _market_solution(prs, n, *, title, key, a, b, qstar, pstar):
    """One market's derivation.  The ANSWER line is dark red — one red line
    per solution slide, so it reads as the answer."""
    def draw(slide):
        _lead(slide, 1.70, "Compute MR:")
        _eq(slide, 2.20, _psub(key) + _omml_text(" = %g − %g" % (a, b))
            + _sub(key))
        _eq(slide, 2.78, _omml_sub(_omml_run("TR"), _omml_text(key))
            + _omml_text(" = ") + _psub(key) + _omml_text("·") + _sub(key)
            + _omml_text(" = %g·" % a) + _sub(key)
            + _omml_text(" − %g" % b) + _omml_sup(_sub(key),
                                                  _omml_text("2")))
        # his notation (2026-09-14): a real interpunct between the
        # coefficient and Q, and MR carries its market subscript in the
        # "set equal to MC" line as well
        _eq(slide, 3.36, _omml_sub(_omml_run("MR"), _omml_text(key))
            + _omml_text(" = %g − %s" % (a, _coef(2 * b)))
            + _sub(key))
        _eq(slide, 4.10, _omml_text("Set   ")
            + _omml_sub(_omml_run("MR"), _omml_text(key))
            + _omml_text(" = ") + _omml_run("MC")
            + _omml_text("     (recall: ") + _omml_run("MC")
            + _omml_text(" = %g)" % MC), size=22, left=0.80)
        _eq(slide, 4.60, _omml_text("%g − %s" % (a, _coef(2 * b))) + _sub(key)
            + _omml_text(" = %g" % MC))
        _eq(slide, 5.18, _omml_sup(_sub(key), _omml_text("*"))
            + _omml_text(" = %d" % int(qstar)))
        # the answer
        _add_math_equation(
            slide, Inches(1.30), Inches(5.80), Inches(10.6), Inches(0.62),
            _omml_sup(_psub(key), _omml_text("*"))
            + _omml_text(" = %g − %g·" % (a, b))
            + _omml_sup(_sub(key), _omml_text("*"))
            + _omml_text("   ⇒   ") + _omml_sup(_psub(key), _omml_text("*"))
            + _omml_text(" = %d" % int(pstar)),
            size_pt=27, color=DARKRED)
    return make_diagram_slide(prs, n, TAG, title, draw)


def s04_us(prs, n):
    """[NEW – PG1 33]."""
    return _market_solution(prs, n, title="Solution:  Profit Maximization "
                            "in the US Market", key=K_US, a=A_US, b=B_US,
                            qstar=Q_US, pstar=P_US)


def s05_de(prs, n):
    """[NEW – PG1 34]."""
    return _market_solution(prs, n, title="Solution:  Profit Maximization in Germany",
                            key=K_DE, a=A_DE, b=B_DE, qstar=Q_DE,
                            pstar=P_DE)


def s06_recap(prs, n):
    """[NEW – PG1 35] — both answers, then the next question."""
    def draw(slide):
        # 2026-09-14 (Nico): "add a textbox on top that says
        # 'What we know:'" -- promoted to a real header on 2026-09-15,
        # with the question moved above it.
        _yourturn(slide, 1.52, 0.90,
                  "2. What is the demand elasticity at the "
                  "profit-maximizing price, and the Lerner Index, in the "
                  "two countries?")
        _section_header(slide, 2.62, "What we know:")
        _lead(slide, 3.22, "MC = %g" % MC, size=23)
        _lead(slide, 3.64, "US market", size=23)
        # the two optima were HARD-CODED at 15 / 100 and 10 / 90 and so
        # went stale the moment the numbers changed; they are derived
        # from the same constants as the rest of the deck now.
        _eq(slide, 4.06, _psub(K_US) + _omml_text(" = %g − %g" % (A_US, B_US)) + _sub(K_US),
            size=23, left=1.65, width=5.6)
        _eq(slide, 4.52, _omml_sup(_sub(K_US), _omml_text("*"))
            + _omml_text(" = %g,   " % Q_US) + _omml_sup(_psub(K_US),
                                                  _omml_text("*"))
            + _omml_text(" = %g" % P_US), size=23, left=1.65, width=5.6)
        _lead(slide, 5.10, "German market", size=23)
        _eq(slide, 5.52, _psub(K_DE) + _omml_text(" = %g − %g" % (A_DE, B_DE)) + _sub(K_DE),
            size=23, left=1.65, width=5.6)
        _eq(slide, 5.98, _omml_sup(_sub(K_DE), _omml_text("*"))
            + _omml_text(" = %g,   " % Q_DE) + _omml_sup(_psub(K_DE),
                                                  _omml_text("*"))
            + _omml_text(" = %g" % P_DE), size=23, left=1.65, width=5.6)
    return make_diagram_slide(
        prs, n, TAG, "BMW and Segmentation: Question 2", draw)


def _elasticity_slide(prs, n, *, title, key, a, b, qstar, pstar, elas,
                      lerner, recall):
    def draw(slide):
        _lead(slide, 1.65, "Elasticity:")
        _eq(slide, 2.12, _psub(key) + _omml_text(" = %g − %g" % (a, b))
            + _sub(key) + _omml_text("   ⇒   ") + _sub(key)
            + _omml_text(" = %g − %g" % (a / float(b), 1.0 / b))
            + _psub(key), size=22)
        _eq(slide, 2.70, _omml_frac(_omml_text("d") + _sub(key),
                                    _omml_text("d") + _psub(key))
            + _omml_text(" = −%g   (the slope of the demand curve)"
                         % (1.0 / b)), size=22)
        # the ANSWER line in dark red -- one red line per solution
        # slide, so it reads as the answer (2026-09-14, Nico)
        _eq(slide, 3.34, _omml_sub(_omml_run("E"), _omml_text(key))
            + _omml_text(" = −%g · (%g/%g) = %.2f"
                         % (1.0 / b, int(pstar), int(qstar), elas)),
            size=24, color=M6.DARKRED)
        _lead(slide, 4.10, "Lerner Index (mark-up):")
        _eq(slide, 4.58, _omml_frac(_omml_run("P") + _omml_text(" − ")
                                    + _omml_run("MC"), _omml_run("P"))
            + _omml_text(" = (%g − %g)/%g = %.2f"
                         % (pstar, MC, pstar, lerner)), size=24,
            color=M6.DARKRED)
        # 2026-09-14 (Nico): "recall box: show the formula that you
        # mention."  The sentence said what the rule is without ever
        # writing it, which is the one thing a recall box is for.
        # ONE line of text, so the box stays 0.52" and the formula
        # underneath keeps its air.  The clause that used to follow --
        # "so the less elastic market carries the higher mark-up" --
        # ran the box to two lines and squeezed the formula onto the
        # footer rule; the formula says the same thing, and slide 2's
        # RULE box and slide 10 both carry the sentence in words.
        _add_convention_box(
            slide, Inches(0.95), Inches(5.50), Inches(11.4), Inches(0.52),
            prefix=recall + "  ",
            body="the mark-up is one over the absolute elasticity",
            size=18, anchor=MSO_ANCHOR.MIDDLE)
        _eq(slide, 6.15, _omml_frac(_omml_run("P") + _omml_text(" − ")
                                    + _omml_run("MC"), _omml_run("P"))
            + _omml_text("  =  ")
            + _omml_frac(_omml_text("1"),
                         _omml_text("|") + _omml_sub(_omml_run("E"),
                                                     _omml_text("D"))
                         + _omml_text("|")), size=22, left=0.95,
            width=11.4)
    return make_diagram_slide(prs, n, TAG, title, draw)


def s07_us_elas(prs, n):
    """[NEW – PG1 36]."""
    return _elasticity_slide(
        prs, n, title="Solution:  Demand Elasticity and Mark-up in the US",
        key=K_US,
        a=A_US, b=B_US, qstar=Q_US, pstar=P_US, elas=E_US,
        lerner=L_US, recall="Recall (Module 5):")


def s08_de_elas(prs, n):
    """[NEW – PG1 37]."""
    return _elasticity_slide(
        prs, n, title="Solution:  Demand Elasticity and Mark-up in Germany",
        key=K_DE,
        a=A_DE, b=B_DE, qstar=Q_DE, pstar=P_DE, elas=E_DE,
        lerner=L_DE, recall="Lerner Index:")


def _panel_numbers(slide, fig, *, a, b, qstar, pstar, qmax, pmax,
                   yint=None):
    """Module 5's markings, carrying this deck's numbers (2026-09-15,
    Nico: "add a P* for the price in each graph.  Also, add numbers to
    the axes.  The two graphs should be the same as in Module 5, but
    with numbers").

    Module 5's price-searcher panel marks P* and Q* on the axes with
    dashed guides and nothing else; here each of those carries its value,
    and the two curve intercepts and the MC level are numbered too, so a
    viewer can read the profit rectangle straight off the axes instead of
    taking the arithmetic underneath on trust.

    Prices are in $1,000 and quantities in 1,000 cars -- the deck's own
    units, the ones the profit line multiplies to millions.
    """
    # y axis: the demand intercept, the optimal price, the MC level.
    # `yint` exists for the KINKED joint demand, whose curve starts at
    # Germany's choke price -- `a` there is only where the lower segment
    # would reach the axis if it were extended, which is also where MR
    # starts, so labelling it in demand red would name the wrong line.
    _yi = a if yint is None else yint
    _H._fig_ylab(slide, fig, _yi, "$%g" % _yi, color=M6.DARKRED, size=13)
    _H._fig_ylab(slide, fig, pstar, "P* = $%g" % pstar, color=M6.NAVY,
                 size=14, bold=True)
    _H._fig_ylab(slide, fig, MC, "$%g" % MC, color=M6.NAVY, size=13)
    # x axis: the optimal quantity and where demand runs out
    _H._fig_xlab(slide, fig, qstar, "Q* = %g" % qstar, color=M6.NAVY,
                 size=14, bold=True)
    _H._fig_xlab(slide, fig, a / b, "%g" % (a / b), color=M6.DARKRED,
                 size=13)


def _markup_arrow(slide, fig, *, pstar, qmax, pmax):
    """The mark-up, measured against the y AXIS (2026-09-15, Nico: "move
    the markup arrow and text to the left side, next to the y axis").

    It used to stand out in the empty space to the right of Q*, where it
    measured the same distance but pointed at nothing.  Against the axis
    it is the left edge of the profit rectangle, so the arrow and the
    rectangle say the same thing in the same place.

    The LABEL goes immediately below the arrow's foot, under the MC line.
    Beside the arrow it was struck through by MR: the profit rectangle is
    only 15 price units tall and 30 quantity units wide in the US panel,
    and MR cuts across it, so there is no room for the word inside it.
    Below MC the band is empty until MR arrives, which at this height is
    most of the panel away.
    """
    xa = qmax * 0.030
    M6._add_arrow(slide, (fig.x(xa), fig.y(MC)), (fig.x(xa), fig.y(pstar)),
                  color=M6.DARKRED, weight_pt=1.25, head_both=True,
                  head_size="sm")
    M6._fig_curve_label(slide, fig, qmax * 0.012, MC - pmax * 0.085,
                        "Mark-up", color=M6.DARKRED, size=12)


def s09_graphical(prs, n):
    """[NEW – PG1 40] — the two market panels.

    2026-09-14 (Nico): "keep only the left two graphs.  Below each,
    write out the profit for that country, and also illustrate the
    mark-up."  The joint-market panel now has slides 10 and 11 of its
    own, so this slide is the segmented case alone.

    2026-09-15: the shared quantity scale drops from 360 to 210.  360 was
    set by the JOINT demand, which reaches Q = 340 and no longer appears
    on this slide; at that scale each market's curves used only the left
    40 % of their panel, which reads as a mistake once the axes carry
    numbers.  210 still covers both markets (Germany runs out at 190) and
    keeps ONE scale across the two panels, so the rectangles stay
    comparable by eye.
    """
    QM, PM = 210.0, 105.0

    def draw(slide):
        for left_in, (a, b), q, pr, key, letter, head in (
                (1.10, M6.BMW_US, Q_US, P_US, ":us", None, "US market"),
                (7.20, M6.BMW_DE, Q_DE, P_DE, ":de", None,
                 "German market")):
            fig = M6._bmw_panel(slide, left_in, 4.70, a=a, b=b, qstar=q,
                                pstar=pr, dlabel="D", mrlabel="MR",
                                key=key, region_letter=letter,
                                heading=head, qmax=QM, pmax=PM,
                                top_in=2.15, height_in=3.55)
            _markup_arrow(slide, fig, pstar=pr, qmax=QM, pmax=PM)
            _panel_numbers(slide, fig, a=a, b=b, qstar=q, pstar=pr,
                           qmax=QM, pmax=PM)
            profit = (pr - MC) * q
            _add_text(slide, Inches(left_in - 0.30), Inches(6.05),
                      Inches(5.30), Inches(0.36),
                      "Profit = ($%g − $%g) · %g = $%s m"
                      % (pr, MC, q, format(int(profit), ",")),
                      size=17, bold=True, color=M6.NAVY,
                      font="Calibri", align=PP_ALIGN.CENTER)
            _add_text(slide, Inches(left_in - 0.30), Inches(6.44),
                      Inches(5.30), Inches(0.36),
                      "Mark-up = ($%g − $%g)/$%g = %.0f %%"
                      % (pr, MC, pr, (pr - MC) / pr * 100),
                      size=17, bold=True, color=M6.DARKRED,
                      font="Calibri", align=PP_ALIGN.CENTER)
    return make_diagram_slide(prs, n, TAG,
                              "Solution:  Illustrating Segment Pricing",
                              draw)


# --- the uniform-price counterfactual (2026-09-15, Nico: "add slides
#     after slide 9: 'What if they charged the same price in both
#     markets?'  Then list the calculation on the first slide (including
#     the aggregate demand) and then the graph on the second slide.")
#
# Every number below is DERIVED from the same BMW constants as the rest
# of the deck -- M6._JA / M6._JB are the aggregate inverse demand and
# M6.BMW_QJ / M6.BMW_PJ its optimum -- so the counterfactual cannot drift
# away from the two market panels it is being compared with.
JA, JB = M6._JA, M6._JB                       # P = 85 − 0.25Q
Q_UNI, P_UNI = M6.BMW_QJ, M6.BMW_PJ           # 80, 65
KINK_Q, KINK_P = M6.BMW_KINK_Q, M6.BMW_KINK_P  # 40, 75
PROFIT_SEG, PROFIT_UNI = M6.BMW_PROFIT_SEG, M6.BMW_PROFIT_UNI


def s10_uniform_calc(prs, n):
    """The algebra behind one price for both markets.

    The horizontal sum has to be done on the DEMAND functions, not the
    inverse demands -- quantities add at a given price, prices do not add
    at a given quantity -- so the slide inverts both lines first, adds
    them, and only then inverts back.  That is the step the class gets
    wrong, so it gets its own line rather than a result box.
    """
    _qint_us, _qint_de = A_US / B_US, A_DE / B_DE        # 150, 190
    _slope = 1.0 / B_US                                   # 2 per $1,000

    def draw(slide):
        _lead(slide, 1.62, "Step 1 — invert each demand (quantity at a "
              "given price):", size=21)
        _eq(slide, 2.06, _sub(K_US) + _omml_text(" = %g − %g" % (_qint_us, _slope))
            + _psub(K_US) + _omml_text("          ") + _sub(K_DE)
            + _omml_text(" = %g − %g" % (_qint_de, _slope)) + _psub(K_DE),
            size=22, left=1.05, width=11.0)
        _lead(slide, 2.66, "Step 2 — add the QUANTITIES at each price "
              "(horizontal summation):", size=21)
        _eq(slide, 3.10, _omml_run("Q") + _omml_text(" = ") + _sub(K_US)
            + _omml_text(" + ") + _sub(K_DE)
            + _omml_text(" = %g − %g" % (_qint_us + _qint_de, 2 * _slope))
            + _omml_run("P") + _omml_text("     ⇒     ") + _omml_run("P")
            + _omml_text(" = %g − %g" % (JA, JB)) + _omml_run("Q"),
            size=22, left=1.05, width=11.0)
        _lead(slide, 3.72, "Step 3 — set MR = MC on the aggregate demand:",
              size=21)
        _eq(slide, 4.16, _omml_run("MR") + _omml_text(" = %g − %g" % (JA, 2 * JB))
            + _omml_run("Q") + _omml_text("     ⇒     %g − %g" % (JA, 2 * JB))
            + _omml_run("Q") + _omml_text(" = %g" % MC), size=22, left=1.05,
            width=11.0)
        _eq(slide, 4.72, _omml_sup(_omml_run("Q"), _omml_text("*"))
            + _omml_text(" = %g" % Q_UNI), size=22, left=1.05, width=11.0)
        # the answer
        _add_math_equation(
            slide, Inches(1.05), Inches(5.30), Inches(11.0), Inches(0.62),
            _omml_sup(_omml_run("P"), _omml_text("*"))
            + _omml_text(" = %g − %g·%g" % (JA, JB, Q_UNI))
            + _omml_text("   ⇒   ") + _omml_sup(_omml_run("P"),
                                                _omml_text("*"))
            + _omml_text(" = %g     (profit = $%s m)"
                         % (P_UNI, format(int(PROFIT_UNI), ","))),
            size_pt=27, color=DARKRED)
        # the consistency check.  Adding the two demand functions is only
        # valid while BOTH markets are buying, i.e. below the lower choke
        # price; if the optimum came out above it the aggregate line used
        # here would be the wrong branch, so the slide checks rather than
        # assumes.
        _add_convention_box(
            slide, Inches(1.05), Inches(6.16), Inches(11.0), Inches(0.62),
            prefix="Check:  ",
            body="P* = $%g is below the US choke price of $%g, so both "
                 "markets really do buy at this price" % (P_UNI, KINK_P),
            size=19, anchor=MSO_ANCHOR.MIDDLE)
    return make_diagram_slide(
        prs, n, TAG,
        "What if They Charged the Same Price in Both Markets?", draw)


def s11_uniform_graph(prs, n):
    """The same picture Module 5 draws for a price searcher, on the
    AGGREGATE demand curve and carrying this deck's numbers.

    One panel, so it can be scaled to its own curves: aggregate demand
    runs out at Q = 340, which is why this slide needs a wider quantity
    axis than slide 9's two markets do.
    """
    QM, PM = 370.0, 105.0

    def draw(slide):
        fig = M6._bmw_panel(slide, 3.55, 6.30, a=JA, b=JB, qstar=Q_UNI,
                            pstar=P_UNI, dlabel="D", mrlabel="MR",
                            key=":uni", region_letter=None,
                            heading="Joint market (US + Germany)",
                            qmax=QM, pmax=PM, top_in=1.90, height_in=4.05,
                            kink=(KINK_Q, KINK_P),
                            kink_a=M6.BMW_TOP[0], kink_b=M6.BMW_TOP[1],
                            # demand reaches the axis at 340 of 370, so
                            # there is room to put D past the intercept
                            dend=None)
        # no mark-up arrow here: Module 5's panel does not carry one, and
        # this slide's question is the total profit rather than the margin
        _panel_numbers(slide, fig, a=JA, b=JB, qstar=Q_UNI, pstar=P_UNI,
                       qmax=QM, pmax=PM, yint=M6.BMW_TOP[0])
        # the kink is the one feature of this curve the two market panels
        # do not have, so it is named -- the same aside the main deck's
        # slide 40 carries.
        _kb = (Inches(4.55), Inches(2.05))
        _kw = (Inches(3.10), Inches(0.62))
        _add_text(slide, _kb[0], _kb[1], _kw[0], _kw[1],
                  "Kink results from demand aggregation "
                  "(beyond the scope of this class)",
                  size=14, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)
        M6._arrow_from_box(slide, _kb, _kw,
                           (fig.x(KINK_Q), fig.y(KINK_P)),
                           color=GRAY, weight_pt=1.5, head=True)
        _add_text(slide, Inches(3.25), Inches(6.30), Inches(6.90),
                  Inches(0.40),
                  "Profit = ($%g − $%g) · %g = $%s m"
                  % (P_UNI, MC, Q_UNI, format(int(PROFIT_UNI), ",")),
                  size=19, bold=True, color=M6.NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)
    return make_diagram_slide(
        prs, n, TAG, "Solution:  Illustrating Uniform Pricing", draw)


def s12_takeaway(prs, n):
    """[NEW – PG1 38] — retitled 2026-09-15 (Nico: "change the title to
    'Comparing Segment Pricing to Uniform Pricing'.  Then add the
    comparison of profits"), now that slides 10 and 11 have worked the
    uniform price out.  The take-away bullets stay as they were; what is
    new is the row that puts the two profits side by side."""
    def draw(slide):
        _add_hierarchical_bullets(
            slide, Inches(0.85), Inches(1.72), Inches(11.6), Inches(3.0),
            [([("Optimization when selling in multiple markets:", {})], 0,
              {}),
             ([("Set MR = MC in each market separately", {})], 1, {}),
             ([("Elasticity:", {})], 0, {}),
             ([("Charge the higher price in the LESS elastic market", {})],
              1, {}),
             # derived, so the sentence cannot contradict the figures
             # above it again (it still said −1.67 / −1.80 and $100 / $90)
             ([("BMW: Germany is less elastic (E = %s against %s), "
                "so the German price is higher – $%g against $%g"
                % (("%.1f" % E_DE).replace("-", "−"),
                   ("%.1f" % E_US).replace("-", "−"),
                   P_DE, P_US), {})], 1, {})],
            size=24, sub_size=22, line_spacing_pts=14)
        # the comparison he asked for: each market's profit, their sum,
        # and the best ONE price can do.  Derived, so it can never
        # contradict slides 9 and 11.
        _gap = 0.28
        _bw = (11.4 - 2 * _gap) / 3.0
        for _i, (_head, _body, _col) in enumerate((
                ("Segment pricing",
                 "$%s m  +  $%s m  =  $%s m"
                 % (format(int((P_US - MC) * Q_US), ","),
                    format(int((P_DE - MC) * Q_DE), ","),
                    format(int(PROFIT_SEG), ",")), NAVY),
                ("Uniform pricing",
                 "$%s m   (one price of $%g)"
                 % (format(int(PROFIT_UNI), ","), P_UNI), GRAY),
                ("Gained by segmenting",
                 "$%s m   (+%.0f %%)"
                 % (format(int(PROFIT_SEG - PROFIT_UNI), ","),
                    (PROFIT_SEG / PROFIT_UNI - 1.0) * 100), DARKRED))):
            _x = 0.95 + _i * (_bw + _gap)
            _add_text(slide, Inches(_x), Inches(4.62), Inches(_bw),
                      Inches(0.32), _head, size=17, bold=True, color=_col,
                      font="Calibri", align=PP_ALIGN.CENTER)
            _add_rounded_filled_box(
                slide, Inches(_x), Inches(4.98), Inches(_bw), Inches(0.56),
                _body, fill=WHITE, line=_col, text_color=_col, size=17,
                bold=True, corner_pct=0.14)
        # 2026-09-15 (Nico): "make the text box in the bottom dark blue
        # and adopt my text change" -- "than a uniform price", now that
        # the deck has actually computed that price.
        _add_rounded_filled_box(
            slide, Inches(0.95), Inches(5.78), Inches(11.4), Inches(0.68),
            "Two markets, two prices – and more profit than a uniform "
            "price can earn", fill=NAVY, text_color=WHITE, size=22,
            bold=True, corner_pct=0.14)
        # back at the convention position: the bar now stops at 6.46
        _add_ps_pointer(slide, top=Inches(6.53), label=M6.PS_BMW)
    return make_diagram_slide(
        prs, n, TAG, "Comparing Segment Pricing to Uniform Pricing", draw)


# ==========================================================================
#  Speaker notes
# ==========================================================================
# Keyed by display number with a TITLE GUARD, the same device the main
# deck uses: two slides were inserted into the middle of this deck on
# 2026-09-15, and a silently mis-attached note would be worse than none.
# The title card (slide 1) carries no notes, per the taping rule.
PV_NOTES = {
    2: ("Segment Pricing Analytically",
        "The set-up for the whole video. One firm, one product, two "
        "separate markets, and the same marginal cost of serving either. "
        "Profit maximization means setting marginal revenue equal in the "
        "two markets, and equal to marginal cost. Here is the intuition for "
        "why. If the two marginal revenues differ, you can move one unit "
        "from the low-marginal-revenue market to the high one and earn more "
        "without producing anything extra, so a single price for both "
        "cannot have been optimal. The rule in the box is what that works "
        "out to in practice, and the rest of the video is the arithmetic "
        "behind it."),
    3: ("BMW and Segmentation: Setup and Question 1",
        "Here are the numbers. Quantities are in thousands of cars and "
        "prices in thousands of dollars, so a price of %g means $%g,000. "
        "The two demand curves have the SAME slope; what differs is the "
        "intercept, %g in the US against %g in Germany. Marginal cost is "
        "%g in both markets. Pause here and let them work out the two "
        "profit-maximizing prices before moving on."
        % (A_US, A_US, A_US, A_DE, MC)),
    4: ("Solution:  Profit Maximization in the US Market",
        "The mechanics, done slowly, because the next three slides all "
        "reuse them. Write total revenue as price times quantity using the "
        "inverse demand, then differentiate to get marginal revenue, which "
        "for a linear demand keeps the same intercept and doubles the "
        "slope. Set that equal to marginal cost of %g and solve for "
        "quantity, then put the quantity back into demand to get the "
        "price. The US answer is %g thousand cars at $%g thousand, and it "
        "is set in dark red because that is the line delivering the "
        "answer."
        % (MC, Q_US, P_US)),
    5: ("Solution:  Profit Maximization in Germany",
        "Exactly the same three steps in Germany, with the higher intercept "
        "of %g. Marginal revenue is %g minus Q, setting it equal to %g "
        "gives %g thousand cars, and demand then gives a price of $%g "
        "thousand. Notice that the higher price came out in Germany, and we "
        "have not said a word about elasticity yet. That is the next "
        "question."
        % (A_DE, A_DE, MC, Q_DE, P_DE)),
    6: ("BMW and Segmentation: Question 2",
        "Everything we have so far, gathered in one place before the second "
        "question. Both optima are on the table: %g thousand cars at $%g "
        "thousand in the US, %g thousand at $%g thousand in Germany, with "
        "marginal cost %g in both. Now ask them for the demand elasticity "
        "at each of those prices, and the Lerner Index that goes with it. "
        "Pause here as well."
        % (Q_US, P_US, Q_DE, P_DE, MC)),
    7: ("Solution:  Demand Elasticity and Mark-up in the US",
        "Start by inverting demand so that quantity is a function of price, "
        "because elasticity is defined on that derivative. The slope is "
        "minus %g, and elasticity is that slope times price over quantity, "
        "which gives minus %.2f in the US. The Lerner Index is the mark-up "
        "as a "
        "share of price, %g over %g, or %.2f. The recall box is the "
        "identity that ties the two together: the mark-up equals one over "
        "the absolute elasticity, and one over %g is indeed %.2f."
        % (1.0 / B_US, abs(E_US), P_US - MC, P_US, L_US, abs(E_US), L_US)),
    8: ("Solution:  Demand Elasticity and Mark-up in Germany",
        "Germany the same way. The slope is again minus %g, but price is %g "
        "and quantity %g, so elasticity comes out at minus %.2f. That is LESS "
        "elastic than the US. The Lerner Index is %g over %g, or %.2f, and "
        "again it is one over the absolute elasticity. So the market with "
        "the less elastic demand carries the bigger mark-up, which is the "
        "rule from the second slide arriving as a result rather than as an "
        "assertion."
        % (1.0 / B_DE, P_DE, Q_DE, abs(E_DE), P_DE - MC, P_DE, L_DE)),
    9: ("Solution:  Illustrating Segment Pricing",
        "The same two answers, drawn. Each panel carries its own demand "
        "curve and marginal revenue, the flat marginal cost line at %g, and "
        "the shaded rectangle between price and marginal cost out to the "
        "optimal quantity, which is the profit. The two panels share one "
        "quantity scale, so the rectangles really are comparable by eye. "
        "The double-headed arrow at the left edge of each rectangle is the "
        "mark-up, and it is visibly taller in Germany. Profit is $%s "
        "million in the US against $%s million in Germany."
        % (MC, format(int((P_US - MC) * Q_US), ","),
            format(int((P_DE - MC) * Q_DE), ","))),
    10: ("What if They Charged the Same Price in Both Markets?",
         "Now the counterfactual. To find the best single price you need "
         "the two markets combined, and the step people get wrong is that "
         "you add QUANTITIES at a given price, not prices at a given "
         "quantity. So invert both demands first, add them, then invert "
         "back. That gives P equals %g minus %g Q. Set marginal revenue "
         "equal to marginal cost on that curve and the best single price is "
         "$%g thousand on %g thousand cars, worth $%s million. The check at "
         "the bottom matters: adding the two demands is only valid while "
         "BOTH markets are still buying, and $%g is below the US choke "
         "price of $%g, so it is."
         % (JA, JB, P_UNI, Q_UNI, format(int(PROFIT_UNI), ","),
             P_UNI, KINK_P)),
    11: ("Solution:  Illustrating Uniform Pricing",
         "The same picture for the single price. The kink in the demand "
         "curve is where the US drops out: above $%g thousand only Germany "
         "is still buying, so the upper segment IS Germany's demand curve, "
         "and below it the two add together. Name the kink and then leave "
         "it alone; how to aggregate demand formally is beyond this class. "
         "The profit rectangle here is $%s million, and the next slide puts "
         "it beside the two we found separately."
         % (KINK_P, format(int(PROFIT_UNI), ","))),
    12: ("Comparing Segment Pricing to Uniform Pricing",
         "The comparison, and the take-away. Two separate prices earn $%s "
         "million plus $%s million, or $%s million. The best single price "
         "earns $%s million. The $%s million gap, about %.0f percent, is "
         "what segmenting is worth here, and the direction never reverses: "
         "one price applied to two different demand curves can never beat "
         "giving each curve its own. And the reason the German price is the "
         "higher one is the elasticity result from earlier, not anything "
         "about German buyers being wealthier."
         % (format(int((P_US - MC) * Q_US), ","),
             format(int((P_DE - MC) * Q_DE), ","),
             format(int(PROFIT_SEG), ","), format(int(PROFIT_UNI), ","),
             format(int(PROFIT_SEG - PROFIT_UNI), ","),
             (PROFIT_SEG / PROFIT_UNI - 1.0) * 100)),
}


def _apply_pv_notes(prs):
    """Attach PV_NOTES, checking each slide's action title first."""
    slides = list(prs.slides)
    n = 0
    for disp, (want, note) in sorted(PV_NOTES.items()):
        slide = slides[disp - 1]
        title = ""
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            try:
                x, y = sh.left / 914400.0, sh.top / 914400.0
            except TypeError:
                continue
            if abs(x - 0.28) < 0.03 and abs(y - 0.55) < 0.04:
                title = sh.text_frame.text.strip()
        if not title.startswith(want):
            raise ValueError(
                "practice note %d expects a slide titled %r but found %r"
                % (disp, want, title[:60]))
        _H._set_notes(slide, note)
        n += 1
    return n


def main():
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    n = 1
    s01_card(prs); n += 1
    s02_setup(prs, n); n += 1
    s03_problem(prs, n); n += 1
    s04_us(prs, n); n += 1
    s05_de(prs, n); n += 1
    s06_recap(prs, n); n += 1
    s07_us_elas(prs, n); n += 1
    s08_de_elas(prs, n); n += 1
    s09_graphical(prs, n); n += 1
    s10_uniform_calc(prs, n); n += 1
    s11_uniform_graph(prs, n); n += 1
    s12_takeaway(prs, n); n += 1
    apply_symbol_subscripts(prs)
    n_notes = _apply_pv_notes(prs)
    prs.save(str(OUT))
    print("%s: %d slides, %d notes"
          % (OUT.name, len(prs.slides._sldIdLst), n_notes))
    print("  US  Q*=%.0f P*=%.0f  E_D=%.2f  Lerner=%.2f"
          % (Q_US, P_US, E_US, L_US))
    print("  DE  Q*=%.0f P*=%.0f  E_D=%.2f  Lerner=%.2f"
          % (Q_DE, P_DE, E_DE, L_DE))
    print("  UNI Q*=%.0f P*=%.0f   aggregate P = %g - %gQ  (kink at %g)"
          % (Q_UNI, P_UNI, JA, JB, KINK_Q))
    print("  segment profit %.0f  >  uniform profit %.0f  (+%.1f%%)"
          % (PROFIT_SEG, PROFIT_UNI,
             (PROFIT_SEG / PROFIT_UNI - 1.0) * 100))


if __name__ == "__main__":
    main()
