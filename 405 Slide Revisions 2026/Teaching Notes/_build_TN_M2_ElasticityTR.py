"""Build: Module 2 - Teaching Note Demand Elasticity and Total Revenue.

Nico's original, restyled to the course theme and checked against the
Module 2 decks.  The video deck's slide 2 carries the pointer "Teaching
Note - Demand Elasticity and Total Revenue"; the material itself is on
in-class slide 48 ("Elasticity Along the Demand Curve: Linear Case") and
video slides 4-6 (demand -> total revenue, and the elastic / inelastic
regions).  The note's three-part decomposition and its endpoint cases
match slide 48 exactly.

STATUS: reviewed and ACCEPTED in full on 2026-09-06, so this script now
emits the agreed text with no revision marks.  Any future proposal goes
back in as `ins` / `dele`.

WHAT CHANGED relative to the original (all accepted 2026-09-06)
  1. "The first term dQ/dP is the inverse slope of the demand curve" ->
     "the slope of the demand function (Q as a function of P)".  Deck
     slide 69 and slide 42 both call it the slope of the demand function,
     and the note's own footnote explains why the plotted line is the
     inverse.  Calling it the "inverse slope" while the slides call it the
     "slope" is the kind of clash that trips students up.
  2. Footnote 1 (Q = 1600 - 4P) gains the thousands separator, matching the
     decks and the other notes.

FORMATTING (untracked)
  - Ed -> ED throughout: the course convention is a capital-D subscript
    (Teaching CLAUDE.md, "Elasticity symbol"), and every deck slide that
    shows it uses ED.
  - Footnote 1 becomes a cream "Recall:" card.  A teaching note is read for
    the asides, and the deck's own device for a short aside is the cream
    callout; a 9 pt footnote at the foot of the page is easy to miss.
  - The two loose black line drawings are rebuilt as one native figure:
    the demand panel and the total-revenue panel stacked and sharing a Q
    axis, with a dashed guide joining the unit-elastic point to the peak of
    total revenue.  That is how the decks present the pair (video slides 6
    and 21).  Demand is dark red C00000 and total revenue gold, per the
    palette.

Run:  python _build_TN_M2_ElasticityTR.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tn_theme as T
from _tn_theme import (CREAM, DARKRED, GOLD, GRAY, NAVY, body, dele, equation,
                       heading, ins, masthead, mfrac, mrun, para, run,
                       text_width)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(
    HERE, "Module 2 - Teaching Note Demand Elasticity and Total Revenue.docx")

MINUS = "−"
INF = "∞"
DQ_DP = mfrac(mrun("∆Q"), mrun("∆P"))
P_Q = mfrac(mrun("P"), mrun("Q"))
E_D = T.msub(mrun("E"), mrun("D", italic=False))


def ed_runs(tail, size=10):
    """The ED label, capital-D subscript, as figure runs."""
    return [("E", dict(bold=True, color=NAVY, size=size)),
            ("D", dict(bold=True, color=NAVY, size=size, subscript=True)),
            (tail, dict(bold=True, color=NAVY, size=size))]


# ==========================================================================
# One figure, two stacked panels sharing the Q axis
# ==========================================================================
def build_figure():
    W, H = 6.30, 4.80
    f = T.Fig(W, H, name="Elasticity along linear demand, and total revenue")

    OX = 1.05                 # shared left edge
    QSPAN = 4.00              # inches for Q = 0 .. 1600
    XTIP = 5.78
    QMID = OX + QSPAN / 2.0   # Q = 800, where ED = -1 and TR peaks

    # ---- panel A: the demand curve ---------------------------------------
    AY = 2.15                 # x-axis of panel A
    PSPAN = 1.60              # inches for P = 0 .. 400
    ATOP = AY - PSPAN         # P = 400
    AYTIP = ATOP - 0.22

    f.line(OX, AY, XTIP, AY, color=NAVY, w_pt=1.25, arrow=True, name="A x-axis")
    f.line(OX, AY, OX, AYTIP, color=NAVY, w_pt=1.25, arrow=True,
           name="A y-axis")
    f.line(OX, ATOP, OX + QSPAN, AY, color=DARKRED, w_pt=1.75, name="D")
    f.dot(QMID, AY - PSPAN / 2.0, d=0.070, color=NAVY, name="unit elastic")

    pt = [("P", dict(bold=True, italic=True, color=NAVY, size=12))]
    w = text_width(pt) + 0.06
    f.label(OX - 0.08 - w, AYTIP - 0.11, pt, w=w, h=0.22, align="r",
            name="A P title")
    qt = [("Q", dict(bold=True, italic=True, color=NAVY, size=12))]
    w = text_width(qt) + 0.06
    f.label(XTIP, AY + 0.04, qt, w=w, h=0.22, align="c", name="A Q title")

    # Region names and the curve label all sit ABOVE the demand line, each at
    # its own x so none of the three collides.  dy(x) is the line's own y.
    def dy(x):
        return ATOP + PSPAN * (x - OX) / QSPAN

    # x chosen so each label sits in its own region and clear of the chips:
    # "Elastic" in the upper half, "Inelastic" past the midpoint at Q = 800
    # (x = 3.05), "D" at the far end of the line.
    for x, txt, size, gap in ((1.95, "Elastic", 11, 0.20),
                              (3.90, "Inelastic", 11, 0.22),
                              (4.70, "D", 12, 0.22)):
        f.label(x, dy(x) - gap - 0.12,
                [(txt, dict(bold=True, color=DARKRED, size=size))],
                align="c", h=0.24, name=txt + " label")

    # ---- panel B: total revenue ------------------------------------------
    BY = 4.42                 # x-axis of panel B
    TRSPAN = 1.55             # inches for TR = 0 .. TRmax
    BTOP = BY - TRSPAN
    BYTIP = BTOP - 0.22

    f.line(OX, BY, XTIP, BY, color=NAVY, w_pt=1.25, arrow=True, name="B x-axis")
    f.line(OX, BY, OX, BYTIP, color=NAVY, w_pt=1.25, arrow=True,
           name="B y-axis")
    # TR = 400Q - Q^2/4 : a parabola from (0,0) through the peak to (1600,0).
    # A quadratic Bezier whose control point sits at twice the peak height
    # passes exactly through the apex.
    ctrl_y = BY - 2 * TRSPAN
    f.curve([(OX, BY), ((QMID, ctrl_y), (OX + QSPAN, BY))],
            color=GOLD, w_pt=1.75, name="TR")
    f.dot(QMID, BTOP, d=0.070, color=NAVY, name="TR max")

    tt = [("TR", dict(bold=True, italic=True, color=NAVY, size=12))]
    w = text_width(tt) + 0.06
    f.label(OX - 0.08 - w, BYTIP - 0.11, tt, w=w, h=0.22, align="r",
            name="B TR title")
    qt2 = [("Q", dict(bold=True, italic=True, color=NAVY, size=12))]
    w = text_width(qt2) + 0.06
    f.label(XTIP, BY + 0.04, qt2, w=w, h=0.22, align="c", name="B Q title")
    # "TR" sits clear of its own curve: the Bezier's y at x = 4.20 is ~3.38,
    # so the label goes 0.23" above it and left of the ED = 0 chip.
    f.label(4.20, 3.03, [("TR", dict(bold=True, color=GOLD, size=12))],
            align="c", h=0.24, name="TR label")

    # ---- the guide that ties the two panels together ---------------------
    f.line(QMID, AY - PSPAN / 2.0, QMID, BTOP, color=NAVY, w_pt=1.0,
           dash="dash", name="Q=800 guide")

    # ---- elasticity callouts ---------------------------------------------
    def callout(x, y, tail, name):
        runs = ed_runs(tail)
        w = text_width(runs) + 0.30
        h = 0.30
        f.label(x, y, runs, w=w, h=h, align="c", fill=CREAM, border=NAVY,
                name=name, pad=0.04)
        return x - w / 2, y, w, h

    # perfectly elastic, at the P-intercept of demand
    bx, by, bw, bh = callout(1.90, 0.10, " = " + MINUS + INF, "A Ed=-inf")
    f.line(bx + 0.02, by + bh / 2, OX + 0.07, ATOP - 0.03,
           color=NAVY, w_pt=0.75, arrow=True, name="A ptr -inf")

    # unit elastic, at the midpoint of demand
    bx, by, bw, bh = callout(4.60, 0.52, " = " + MINUS + "1", "A Ed=-1")
    f.line(bx + 0.02, by + bh / 2, QMID + 0.09, AY - PSPAN / 2.0 - 0.05,
           color=NAVY, w_pt=0.75, arrow=True, name="A ptr -1")

    # perfectly inelastic, at the Q-intercept of demand
    bx, by, bw, bh = callout(5.30, AY - 0.95, " = 0", "A Ed=0")
    f.line(bx + 0.12, by + bh, OX + QSPAN + 0.04, AY - 0.10,
           color=NAVY, w_pt=0.75, arrow=True, name="A ptr 0")

    # and the same three points on the total-revenue panel
    bx, by, bw, bh = callout(2.00, BTOP - 0.34, " = " + MINUS + "1", "B Ed=-1")
    f.line(bx + bw - 0.02, by + bh / 2, QMID - 0.07, BTOP - 0.04,
           color=NAVY, w_pt=0.75, arrow=True, name="B ptr -1")

    bx, by, bw, bh = callout(5.05, BY - 1.00, " = 0", "B Ed=0")
    f.line(bx + 0.12, by + bh, OX + QSPAN + 0.02, BY - 0.10,
           color=NAVY, w_pt=0.75, arrow=True, name="B ptr 0")

    return f


# ==========================================================================
def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)
    masthead(doc,
             "Teaching Note – Module 2: Demand Elasticity and Total Revenue")

    body(doc, "This note discusses how the own-price elasticity of demand "
              "varies along a linear demand curve and then establishes the "
              "relationship between demand elasticity and total revenue.")

    # ---------------- elasticity along the curve ---------------------------
    heading(doc, "Own-price elasticity of demand along a linear demand curve",
            before=12)

    body(doc, "The own-price elasticity of demand is defined as follows:")
    equation(doc,
             E_D + mrun("=") + mfrac(mrun("%∆Q"), mrun("%∆P"))
             + mrun("=") + DQ_DP + mrun("∙") + P_Q)

    p = body(doc)
    run(p, "The first term ")
    T.equation_inline(p, DQ_DP)
    run(p, " is the slope of the demand function (")
    run(p, "Q", italic=True)
    run(p, " as a function of ")
    run(p, "P", italic=True)
    run(p, "). This slope term is constant in a ")
    run(p, "linear", italic=True)
    run(p, " demand function. Thus, the second term in the formula, ")
    T.equation_inline(p, P_Q)
    run(p, ", will determine how the elasticity varies along the curve. We "
           "can decompose the curve into three parts:")

    # the original's footnote, promoted to the house cream aside
    T.callout(doc, "Recall:",
              "we write the demand function as, e.g., Q = 1,600 − 4P, but we "
              "plot inverse demand, with P on the y-axis and Q on the x-axis.")

    build_figure().place(doc, before=8, after=10)

    # ---------------- the three parts --------------------------------------
    p = numbered(doc, "1.", "Elastic")
    run(p, ": In the upper left portion of the demand curve, price is very "
           "high relative to quantity, making ")
    T.equation_inline(p, P_Q)
    # "slope", not "inverse slope" -- kept consistent with the paragraph
    # above, which the decks' wording settled
    run(p, " large. In the elasticity formula, the slope ")
    T.equation_inline(p, DQ_DP)
    run(p, " (which is negative) is thus multiplied by a large number, "
           "yielding a large negative number. For this reason, the upper "
           "left portion of the demand curve is highly elastic.")

    p = numbered(doc, "2.", "Inelastic")
    run(p, ": In the lower right portion of the demand curve, price is low "
           "relative to quantity, making ")
    T.equation_inline(p, P_Q)
    run(p, " small. This gives a small negative number for the elasticity – "
           "demand is inelastic.")

    p = numbered(doc, "3.", "Unit elastic")
    run(p, ": Unit elasticity is the point on the demand curve at which the "
           "elasticity of demand is equal to -1. At this point, the "
           "percentage change in price is exactly offset by the percentage "
           "change in quantity. This can be illustrated by writing the "
           "elasticity formula in terms of percentage changes:")

    equation(doc,
             E_D + mrun("=")
             + mfrac(mrun("∆Q/Q"), mrun("∆P/P")) + mrun("=-1"))

    p = body(doc)
    run(p, "In addition to these three elasticities, we can also compute the "
           "elasticity where the demand curve intersects the x and y axes. "
           "At the intersection with the y-axis, the quantity is zero, so ")
    T.equation_inline(p, P_Q + mrun("→∞"))
    run(p, ", and the elasticity is negative infinity. The intuition behind "
           "this is that a small decrease in price will yield a change in "
           "quantity demanded from zero to a positive value. Even if this "
           "change in quantity is small, the ")
    run(p, "percentage", italic=True)
    run(p, " change is infinite because it is rising from zero.")

    p = body(doc)
    run(p, "At the intersection with the x-axis, the price is zero, so ")
    T.equation_inline(p, P_Q + mrun("=0"))
    run(p, ", and the elasticity of demand is also zero.")

    # ---------------- elasticity and total revenue -------------------------
    heading(doc, "Relationship between elasticity of demand and total revenue",
            before=14)

    p = body(doc)
    run(p, "We can now explore the relationship between the elasticity of "
           "demand and total revenue. Total revenue is defined as ")
    T.equation_inline(p, T.acr("TR") + mrun("=P∙Q"))
    run(p, ". Starting at the leftmost portion of the graph, when ")
    T.equation_inline(p, mrun("Q=0"))
    run(p, ", we know that ")
    T.equation_inline(p, T.acr("TR") + mrun("=0"))
    run(p, ". As the price falls, quantity rises. Note that these changes "
           "have opposing effects on total revenue. Since we are on the "
           "elastic portion of the demand curve, we know that a small "
           "decrease in price causes a large increase in quantity, and total "
           "revenue increases. This pattern will hold as price falls through "
           "the point of unit elasticity on the demand curve. At this point, "
           "total revenue reaches its maximum.")
    ins(p, " This is also the point where marginal revenue is zero, as shown "
           "in the Teaching Note on Marginal Revenue.")
    run(p, " After this point, we enter "
           "the inelastic portion of the demand curve. On the inelastic "
           "portion of the demand curve, we know that a small decrease in "
           "price causes an even smaller increase in quantity, and total "
           "revenue falls. Total revenue continues to fall until the demand "
           "curve intersects the x-axis, where ")
    T.equation_inline(p, mrun("P=0"))
    run(p, ", and ")
    T.equation_inline(p, T.acr("TR") + mrun("=0"))
    run(p, ".")

    doc.save(OUT)
    print("wrote", OUT)


def numbered(doc, marker, lead):
    """A numbered paragraph whose italic lead-in names the case."""
    p = para(doc, before=6, after=6, left=0.35, hang=0.35,
             align=T.WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, marker + "\t")
    run(p, lead, italic=True)
    return p


if __name__ == "__main__":
    main()
