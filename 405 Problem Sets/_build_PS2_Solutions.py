"""Build: Problem Set 2 - Solutions.docx  (MGMT 405, Fall 2026 format).

Source: `_originals/Problem Set 2 - Solutions (2025 original).docx` -- a
BUILD INPUT, never deleted.

The job is FORMATTING: the original's reasoning, numbers, sample answers
and point breakdowns are preserved.  Everything beyond the restyle is a
real Word revision (`w:ins` / `w:del`, author "Claude (proposed)").

ARITHMETIC RE-CHECKED, all of it correct as written:
  1(a)  (200/1200) / (-5/45) = (1/6)/(-1/9) = -1.5
  1(d)  (-140/1400) / (-1/5) = (-1/10)/(-1/5) = 0.5
  3(a)  806,584/246,555 = 3.2714 ; -41.81/86.50 = -0.4833 ; E = -6.77
        (and the table itself is consistent: 246,555 x $86.50 =
        $21,327,008, 1,053,139 x $44.69 = $47,064,782)
  4     Q = 250 - 140 = 110 ; MR = 250 - 2Q = 250 - 220 = 30
  5(a)  67.4 - 211.2 + 52.78 + 149.34 = 58.32
  5(b)  -16 x 13.2/58.32 = -3.62
  5(c)  269.52/16 - (2/16)(58.32) = 16.845 - 7.29 = 9.555

FORMATTING (untracked)
  - Course masthead, navy problem headings with a gold points chip, the
    bare centred page number.
  - Every derivation set as native OMML with true subscripts and stacked
    fractions; the line that delivers the answer is dark red, the course
    convention for a solution's result.
  - Problem 3's data block becomes the same native four-column table the
    problem statement now uses.
  - BOTH figures rebuilt as native, editable Word shapes.  The originals
    were matplotlib screenshots; the 1(c) one carried a leftover prompt
    artifact in its title -- "Revenue vs. Price when Demand is Elastic
    (Labels repositioned: left label left, right label right)".
  - 4(c)'s figure KEEPS the original's orientation, Q on the vertical
    axis and P on the horizontal.  That is not a slip: deck slide 101
    ("a = intercept with the y-axis (Q on y-axis)") sets exactly this
    convention for an estimated demand function, and slides 103-104 plot
    the airline scatter the same way.
  - 1(c)'s figure drops the original's two overlapping revenue
    rectangles.  They obscured each other, and the point the figure has
    to make is that the elasticity DIFFERS at the two observed points.
    The two revenue figures ($54,000 and $56,000) are in the caption
    instead, where they can be read exactly.

TYPOS FIXED DIRECTLY (unambiguous; reported rather than tracked)
  - 5(c): "Q=58,32" -> "Q = 58.32" (decimal comma; the value is computed
    as 58.32 in part (a) of the same solution).
  - 1(d): "1, 260 units" -> "1,260 units".
  - Problem 3's table: "Market revenue IN DURING price war" -> "during".

PROPOSED (tracked -- these go beyond formatting)
  1. 1(a): the original's note that the elasticity is measured at the
     initial point is kept, and a sentence is added giving the reason it
     matters here -- measured from the SECOND point the same two
     observations give -1.14, not -1.5.  Deck slide 38 carries the
     convention; the solution asserted it without showing the stakes.
  2. 3(a): the same point, and here it is not a fine distinction.  From
     the initial point the elasticity is -6.77; from the final point it
     is -0.82, which would flip "elastic" to "inelastic".  A sentence
     is added saying so.
  3. 1(c): a sentence naming the two total-revenue figures ($54,000 at
     $45, $56,000 at $40), so the "revenues will fall if you raise the
     price" claim is backed by the arithmetic rather than only by the
     sign of the elasticity.

Run:  python _build_PS2_Solutions.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _ps_theme as S
from _ps_theme import (DARKRED, GOLD, GRAY, NAVY, Panel, WD_ALIGN_PARAGRAPH,
                       body, caption, cross, dele, ins, para, part, problem,
                       run)
import _tn_theme as T
from _tn_theme import mfrac, mrun, msub

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Problem Set 2 - Solutions.docx")

# -- OMML symbols ----------------------------------------------------------
Q = mrun("Q")
P = mrun("P")
PX = msub(mrun("P"), mrun("X"))
QX = msub(mrun("Q"), mrun("X"))
PY = msub(mrun("P"), mrun("Y"))
Y = mrun("Y")
DQ = mrun("Δ", italic=False) + Q
DP = mrun("Δ", italic=False) + P
DQX = mrun("Δ", italic=False) + QX
DPY = mrun("Δ", italic=False) + PY
E = mrun("E", italic=False)
ED = msub(E, mrun("D"))
# The BASE of every percentage change: the START of the interval.  Written
# with its subscript so the reader can see which end it is, rather than
# reconstructing it from the arithmetic (2026-09-13, Nico).
Q0 = msub(mrun("Q"), mrun("0", italic=False))
P0 = msub(mrun("P"), mrun("0", italic=False))


def EIV(interval, sym=None):
    """An elasticity subscripted by the INTERVAL it describes.

    "E_D, $45 -> $40" -- the same comma-arrow form the change tables use
    for a marginal cost ("MC, 0 -> 1,000"), so a student meets one
    convention rather than two.  The NUMBER is unaffected: the base is
    still the start of the interval, per deck slide 38.
    """
    if sym is None:
        # ONE subscript, "D, $45 -> $40", not a subscript nested on E_D:
        # nesting renders as "E_D ,$45" with a gap before the comma.
        return msub(E, mrun("D, " + interval, italic=False))
    return msub(sym, mrun(", " + interval, italic=False))
EXY = msub(E, mrun("X,Y"))
TR = T.acr("TR")
MR = T.acr("MR")


def draws_on(doc, text):
    # after 8 -> 5 on 2026-09-12, part of returning ~30 pt of pure
    # spacing so a page break is never decided by a few points.
    p = para(doc, before=0, after=5, keep_next=True)
    run(p, "Draws on:  " + text, italic=True, color=GRAY, size=9.5)
    return p


def answer(doc, content, before=4, after=8):
    """A display equation in dark red -- the line that delivers the result."""
    return T.equation(doc, T.color_omml(content, DARKRED),
                      before=before, after=after)


def bullet(doc, text=None, level=0, before=2, after=2):
    left = 0.32 + 0.28 * level
    p = para(doc, before=before, after=after, left=left, hang=0.18,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, ("•  " if level == 0 else "–  "), color=NAVY, size=11)
    if text:
        run(p, text, size=11)
    return p


# ==========================================================================
# Figures
# ==========================================================================
def fig_coffee():
    """1(c): one demand curve, two observed points, two different
    elasticities.

    Drawn to the actual numbers.  The two observations (1200, $45) and
    (1400, $40) determine a linear demand P = 75 - 0.025 Q, and on that
    line the initial-point elasticity IS the point elasticity:
        dQ/dP = 200 / -5 = -40
        E(1200, 45) = -40 x 45/1200  = -1.50
        E(1400, 40) = -40 x 40/1400  = -1.14
    which is exactly what parts (b) and (c) have to show.

    Axis ranges are chosen so the two points separate: at P 0-80 and
    Q 0-1800 the price ticks sit 0.19" apart and the quantity ticks
    0.50" apart, so nothing has to share space.
    """
    f = S.fig(5.90, 4.05, name="1(c) - elasticity varies along the demand curve")
    pn = Panel(f, 0.72, 3.42, 4.50, 3.05, ylabel="P", xlabel="Q")

    def u(q):
        return q / 1800.0 * 100.0

    def v(p):
        return p / 80.0 * 100.0

    d0, d1 = (u(0), v(75)), (u(1710), v(32.25))     # P = 75 - 0.025 Q
    # "D" is lifted clear of the line end so it cannot run into the
    # elasticity label that sits below the right-hand point
    pn.demand(d0, d1, label="D", lbl_dx=0.10, lbl_dy=-0.20)

    # The INTERVAL, drawn as a thicker stretch of the demand curve between
    # the two observations (2026-09-13, Nico).  Emitted BEFORE the dots and
    # labels so it cannot paint over them -- regions first, then curves,
    # then labels.
    pn.span((u(1200), v(45)), (u(1400), v(40)), color=DARKRED, w_pt=4.5,
            name="interval $45 -> $40")

    # The two elasticity labels take opposite sides of the curve -- the
    # left one above it, the right one well below -- so that neither can
    # reach the other, nor the "D" label at the end of the line.  Each is
    # now labelled by the INTERVAL it describes and the end it is measured
    # from: the SAME two observations give -1.50 from $45 and -1.14 from
    # $40, which is exactly the point parts (b) and (c) make.
    for q, p_, e_lab, iv, base, dv in (
            (1200, 45, "−1.50", "$45 → $40", True, +12.0),
            (1400, 40, "−1.14", "$40 → $45", False, -16.0)):
        pn.equilibrium(u(q), v(p_), guides=True)
        if not base:
            # HOLLOW: this end is not the base the class convention uses,
            # so it is drawn open over the solid equilibrium dot.
            pn.dot_open(u(q), v(p_), d=0.095, color=DARKRED)
        pn.tick_y(v(p_), "$%d" % p_, h=0.17)
        pn.tick_x(u(q), "%d" % q, h=0.19)
        runs = [("E", dict(bold=True, color=DARKRED, size=10.5)),
                ("D", dict(bold=True, color=DARKRED, size=10.5,
                           subscript=True)),
                (", " + iv + " = " + e_lab,
                 dict(bold=True, color=DARKRED, size=10.5))]
        pn.text(u(q) + 2.0, v(p_) + dv, runs, align="l")
    return f


def fig_boeing():
    """4(c): the quantity Boeing sells at each price from 100 to 200.

    Q on the VERTICAL axis and P on the horizontal, which is this
    course's convention for an estimated demand function (deck slide
    101: "a = intercept with the y-axis (Q on y-axis)").  The asked
    segment is solid; the rest of Q = 250 - P is dashed, so it is
    visible that the segment is part of the whole demand curve -- which
    is the answer to "what is this curve?".
    """
    f = S.fig(5.90, 4.05, name="4(c) - Boeing demand, prices 100 to 200")
    pn = Panel(f, 0.72, 3.42, 4.50, 3.05, ylabel="Q", xlabel="P")

    def u(p_):
        return p_ / 250.0 * 100.0

    def v(q):
        return q / 250.0 * 100.0

    # the whole demand line, dashed outside the asked range
    pn.demand((u(10), v(240)), (u(100), v(150)), label=None, dash="dash",
              w_pt=1.25, name="D below 100")
    pn.demand((u(200), v(50)), (u(240), v(10)), label=None, dash="dash",
              w_pt=1.25, name="D above 200")
    # the asked segment, solid
    pn.demand((u(100), v(150)), (u(200), v(50)), label="D", lbl_dx=0.10,
              lbl_dy=-0.30, w_pt=2.25, name="D segment")

    # the point from part (a)
    pn.equilibrium(u(140), v(110), guides=True)
    pn.tick_x(u(140), "140")
    pn.tick_y(v(110), "110")
    # the ends of the asked range
    for p_, q in ((100, 150), (200, 50)):
        pn.f.dot(pn.x(u(p_)), pn.y(v(q)), d=0.055, color=NAVY,
                 name="segment end")
        pn.tick_x(u(p_), "%d" % p_)
        pn.tick_y(v(q), "%d" % q)
    return f


# ==========================================================================
def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)

    S.ps_masthead(doc, "Problem Set 2 – Solutions",
                  covers="Covers Module 2", due="100 points")

    # ======================================================================
    # Problem 1
    # ======================================================================
    problem(doc, 1, "Coffee Demand", 20, before=12)
    draws_on(doc, "Module 2 – Elasticities (own-price and cross-price)")

    p = part(doc, "a", 5)
    run(p, "Use the formula for own-price elasticity, dividing each "
           "change by the START of the interval:")
    T.equation(doc, ED + mrun("=") + mfrac(mfrac(DQ, Q0), mfrac(DP, P0)),
               before=3, after=5)

    body(doc, "Since", after=2)
    T.equation(doc,
               mfrac(DQ, Q0) + mrun("=") + mfrac(mrun("1400") + mrun("−")
                                                 + mrun("1200"), mrun("1200"))
               + mrun("=") + mfrac(mrun("1"), mrun("6")),
               before=2, after=3)
    T.equation(doc,
               mfrac(DP, P0) + mrun("=") + mfrac(mrun("40") + mrun("−")
                                                 + mrun("45"), mrun("45"))
               + mrun("=") + mrun("−") + mfrac(mrun("1"), mrun("9")),
               before=2, after=5)
    body(doc, "the elasticity over the interval $45 → $40, measured from "
              "its starting point, is:", after=2)
    # The INTERVAL is the subscript now, not the anchor point alone
    # (2026-09-13, Nico), in the same comma-arrow form the change tables
    # use for a marginal cost.
    answer(doc, EIV("$45 → $40") + mrun("=")
           + mfrac(mrun("1"), mrun("6")) + mrun("÷") + mrun("(") + mrun("−")
           + mfrac(mrun("1"), mrun("9")) + mrun(")") + mrun("=") + mrun("−1.5"))
    body(doc, "Thus, demand is elastic.")

    S.note(doc,
           "As we said in class, we always use the initial price and "
           "quantity as the reference point. The interval is what the "
           "elasticity describes; the initial point is only the base we "
           "divide by. Because the two observations lie on one straight "
           "demand curve, this number is also the elasticity at $45 "
           "itself.",
           prefix="Note:")

    # PROPOSED 1: say why the reference point matters.  This sits in an
    # Accepted 2026-09-09, and he put it in square brackets -- an
    # editorial aside to the grader rather than solution prose.
    p = body(doc, before=2)
    run(p, "[It matters which point you use: measured from the second "
           "observation instead, the same two data points give −1.14 rather "
           "than −1.5.]")

    body(doc,
         "In this case, demand is elastic probably because of (1) low "
         "product differentiation, since coffee is a relatively homogeneous "
         "product (think of cars, for instance, that can differ from each "
         "other in many more dimensions) and (2) high market fragmentation, "
         "since coffee is a global commodity and roasters face a highly "
         "competitive market with each producer holding a relatively small "
         "market share.", before=10)

    p = part(doc, "b", 5)
    run(p, "Since ")
    T.equation_inline(p, EIV("$45 → $40") + mrun("=") + mrun("−1.5"))
    run(p, ", the firm is operating in the elastic portion of the demand "
           "curve. Thus, revenues will fall if the price goes up. If you "
           "change your price a second time you cannot use the same "
           "elasticity as before. This is because the elasticity changes as "
           "we move along the demand curve: at a different price-quantity "
           "point, the elasticity will be different.")

    p = part(doc, "c", 5)
    run(p, "First, revenues will fall because ")
    T.equation_inline(p, EIV("$45 → $40") + mrun("=") + mrun("−1.5"))
    run(p, ", so the firm is operating on the elastic portion of the demand "
           "curve.")
    # Accepted 2026-09-09.
    run(p, " Concretely, total revenue is 1,200 × $45 = $54,000 at the "
           "higher price and 1,400 × $40 = $56,000 at the lower one, so "
           "cutting the price raised revenue — and raising it would lower "
           "revenue.")
    run(p, " Second, if you change your price a second time you cannot use "
           "the same elasticity because the firm will be operating on a "
           "different point of the demand curve, and the elasticity varies "
           "across the demand curve.")

    S.place(fig_coffee(), doc, before=8, after=2)
    caption(doc, "The two observations pin down the demand curve "
                 "P = 75 − 0.025Q. The thick stretch is the interval they "
                 "define; the solid dot is the base the class convention "
                 "divides by, the hollow one the other end. From $45 the "
                 "elasticity is −1.50, from $40 it is −1.14 — same "
                 "interval, different base, and on this straight line each "
                 "is also the elasticity at its own point.")

    p = part(doc, "d", 5)
    run(p, "Use the formula for cross-price elasticity, where ")
    run(p, "X", italic=True)
    run(p, " refers to your product and ")
    run(p, "Y", italic=True)
    run(p, " to the competitor’s product:")
    T.equation(doc, EXY + mrun("=")
               + mfrac(mfrac(DQX, msub(QX, mrun("0", italic=False))),
                       mfrac(DPY, msub(PY, mrun("0", italic=False)))),
               before=3, after=5)

    body(doc, "Since", after=2)
    T.equation(doc,
               mfrac(DQX, QX) + mrun("=") + mfrac(mrun("1260") + mrun("−")
                                                  + mrun("1400"), mrun("1400"))
               + mrun("=") + mrun("−") + mfrac(mrun("1"), mrun("10")),
               before=2, after=3)
    T.equation(doc,
               mfrac(DPY, PY) + mrun("=") + mrun("−20%") + mrun("=")
               + mrun("−") + mfrac(mrun("1"), mrun("5")),
               before=2, after=5)
    body(doc, "the cross-price elasticity over that interval, measured "
              "from its starting point, is:", after=2)
    # E_XY is already subscripted, so a second subscript renders as
    # "E_X,Y ,year 2 -> year 3" with a gap before the comma.  The interval
    # is named in the sentence above instead, which reads better here.
    answer(doc, EXY + mrun("=") + mrun("(")
           + mrun("−")
           + mfrac(mrun("1"), mrun("10")) + mrun(")") + mrun("÷") + mrun("(")
           + mrun("−") + mfrac(mrun("1"), mrun("5")) + mrun(")") + mrun("=")
           + mrun("0.5"))

    body(doc,
         "This cross-price elasticity seems relatively high (for example, "
         "compared to the 0.15 that we have seen for different types of "
         "cereals in class). This is also due to low product differentiation "
         "as we discussed in (a): coffee is a relatively homogeneous good, so "
         "customers view coffee from different producers as close "
         "substitutes. Although there is some differentiation in coffee "
         "quality/flavors, these differences are relatively subtle to most "
         "customers, and many roasters/producers can produce varieties of "
         "coffee that compete with each other at relatively similar quality "
         "levels.")

    # ======================================================================
    # Problem 2
    # ======================================================================
    problem(doc, 2, "Consulting for Dollar General and Apple", 20)
    draws_on(doc, "Module 2 – Elasticities;  Demand Estimation")

    p = part(doc, "a", 10, after=2)
    run(p, "Dollar General.")
    S.note(doc,
           "Below we provide a sample response. Full credit will be given to "
           "answers that make reasonable assumptions regarding the elasticity "
           "of demand for Dollar General and predict the change in total "
           "revenue consistent with the assumptions.",
           prefix="Grading:")

    p = part(doc, "i", None, level=1, before=6)
    run(p, "Holding everything else constant, a reduction in prices is "
           "likely to lead to an increase in the quantity of goods sold due "
           "to the law of demand. However, the effect of the decrease in "
           "prices on total revenue is ambiguous. It depends on the "
           "elasticity of demand for their products. If demand is elastic at "
           "the initial price, a reduction in prices will result in a bigger "
           "increase in quantity, thus total revenue will increase. However, "
           "if the demand for their products is inelastic, the increase in "
           "quantity is smaller, thus not enough to offset the decrease in "
           "prices. Total revenue in this case will fall. If the demand is "
           "unit-elastic, total revenue is likely not to change much.")

    body(doc,
         "Since Dollar General sells everyday products that customers could "
         "find at many other retailers (Walmart or any convenience store), we "
         "would expect its demand elasticity to be relatively high. In "
         "addition, the typical Dollar General customer is likely cost "
         "conscious and willing to shop around for low prices. So a price "
         "decrease would likely result in an increase in total revenue.",
         left=0.64)

    body(doc,
         "Finally, note that if Dollar General is the only store that does "
         "not raise prices, then the effects are essentially the same as when "
         "it’s the only store that lowers prices. So you could also make the "
         "above argument focusing on the relative prices of Dollar General "
         "vis-à-vis its competitors.", left=0.64)

    p = part(doc, "ii", None, level=1)
    run(p, "If Walmart also reduces their prices at the same time, the demand "
           "curve that Dollar General is facing is likely to shift inward "
           "(shift to the left), since these two stores sell many substitute "
           "products and serve a similar market segment (price-sensitive "
           "customers). Therefore, the quantity of goods sold by Dollar "
           "General could either increase or decrease, depending on which "
           "effect prevails: the effect of Dollar General’s own price "
           "reduction, explained in part (i), or the effect of the "
           "competitor’s price reduction.")

    p = part(doc, "b", 10, after=2)
    run(p, "Three methods to estimate the effect of a price increase:")

    p = bullet(doc, level=0, before=5)
    run(p, "Survey customers:", bold=True)
    run(p, " Apple can survey their customers to get a sense of how likely "
           "they would be to cancel or keep the subscription if the price "
           "increased by $X.")
    p = bullet(doc)
    run(p, "Transaction data:", bold=True)
    run(p, " Apple has data on its customers’ past purchases. Using these "
           "data, Apple can run a regression analysis. The independent "
           "variables can include not only prices, but also other factors "
           "that could affect the quantity sold such as characteristics of "
           "customers (age, gender, etc.), or whether users have other Apple "
           "products. Using the regression results, they can estimate the "
           "change in ")
    run(p, "Q", italic=True)
    run(p, " in response to the change in price of the Apple Music "
           "“Individual” Plan.")
    p = bullet(doc)
    run(p, "Market experiment (A/B testing):", bold=True)
    run(p, " Apple could conduct an experiment in which they change the "
           "price for a segment of the market. The response of that group "
           "compared to the other group can help estimate the change in the "
           "number of subscriptions sold.")

    S.note(doc,
           "Below we provide a sample response. Full credit will be given to "
           "answers that make reasonable assumptions regarding the "
           "feasibility, costs, and benefits of the method of estimating "
           "demand.", prefix="Grading:")

    body(doc,
         "Surveys could give Apple an instant, low-cost estimate of the "
         "effects of the price increase. Since Apple is a well-known company "
         "with a broad customer base, it should be straightforward to reach "
         "out to customers to conduct surveys. However, sometimes customers’ "
         "responses do not truthfully reflect what they would do in real "
         "life. Apple should also be concerned about sampling bias: the "
         "customers who respond to the survey may not represent their "
         "(potential) customers as a whole.")

    body(doc,
         "One advantage of using transaction data is that Apple already "
         "maintains such data in-house, which will save Apple money compared "
         "to running a survey. However, it is hard to capture every factor "
         "that could potentially affect the quantity of subscriptions. There "
         "are numerous non-price factors that could affect demand, including "
         "seasonality, changing consumer preferences, and changes in the "
         "prices of competitor music streaming services. Also, customers’ "
         "past purchases of other Apple products may not accurately predict "
         "their future purchases of the Apple music plan.")

    body(doc,
         "A market experiment is likely the most accurate way to estimate the "
         "change in subscriptions sold due to a price increase. This would be "
         "our recommendation. The online setting enhances the feasibility of "
         "randomizing prices. When signed into their Apple account, new "
         "customers could be randomly assigned the higher price or the "
         "current price. Apple could then observe the purchase decisions of "
         "new customers (while grandfathering all existing customers into the "
         "old price). Even though market experiments could potentially give "
         "the best estimate, this method could pose some challenges. It could "
         "be costly to run an experiment and to make sure the experiment is "
         "run correctly. If the process of assigning control and treatment "
         "group is not random and done correctly, the experimental results "
         "would not be informative. In addition, the experiment described "
         "above would only be informative for new customers, not existing "
         "customers. Finally, if Apple randomizes prices for its service and "
         "customers find out they paid different prices for the same service, "
         "it could lead to consumer backlash.")

    # ======================================================================
    # Problem 3
    # ======================================================================
    problem(doc, 3, "Air Travel Demand Elasticity", 15)
    draws_on(doc, "Module 2 – Elasticities (computing elasticity from two "
                  "observations)")

    T.table(doc, [
        ["", "Before the price war", "During the price war", "Change"],
        ["Passengers", "246,555", "1,053,139", "+806,584"],
        ["Average one-way fare", "$86.50", "$44.69", "−$41.81"],
        ["Market revenue", "$21,327,008", "$47,064,782", "+$25,737,774"],
    ], widths_in=[1.60, 1.65, 1.65, 1.60], size=10, align_right=(1, 2, 3))

    p = part(doc, "a", 10, before=10)
    run(p, "Use the definition of elasticity and plug in the values:")

    T.equation(doc,
               ED + mrun("=") + mfrac(mfrac(msub(Q, mrun("2", italic=False))
                                            + mrun("−")
                                            + msub(Q, mrun("1", italic=False)),
                                            msub(Q, mrun("1", italic=False))),
                                      mfrac(msub(P, mrun("2", italic=False))
                                            + mrun("−")
                                            + msub(P, mrun("1", italic=False)),
                                            msub(P, mrun("1", italic=False)))),
               before=4, after=4)
    T.equation(doc,
               mrun("=") + mfrac(mfrac(mrun("1,053,139") + mrun("−")
                                       + mrun("246,555"), mrun("246,555")),
                                 mfrac(mrun("$44.69") + mrun("−")
                                       + mrun("$86.50"), mrun("$86.50")))
               + mrun("=") + mfrac(mrun("3.2714"), mrun("−0.4833")),
               before=2, after=4)
    answer(doc, EIV("$86.50 → $44.69") + mrun("=") + mrun("−6.77"))

    body(doc,
         "Quantity demanded increased by 327%, while the price dropped by "
         "48.33%. This implies that demand is elastic – the response in "
         "quantity demanded is proportionately stronger than the change in "
         "price.")

    # Accepted 2026-09-09.  He struck "wrongly": the -0.82 figure is not
    # an error, it is the same data measured from the other end point.
    p = body(doc)
    run(p, "Both numbers describe the SAME interval; what differs is the "
           "base we divide by. Measured from the situation before the price "
           "war, as the class convention requires, the elasticity is −6.77. "
           "Measured from the situation during the price war it would be "
           "−0.82, which would suggest that demand is inelastic. That is "
           "why the interval alone is not enough: the elasticity has to "
           "name its starting point too.")

    p = part(doc, "b", 5)
    run(p, "Below we provide a sample response. Full credit will be given to "
           "answers that choose either airport and that make reasonable "
           "assumptions regarding the determinants of demand elasticity. The "
           "ranking is ambiguous and depends on which forces dominate:")

    p = bullet(doc, before=5)
    run(p, "Argument for LAX more elastic:", bold=True)
    run(p, " LAX is a much larger airport, with more carriers and higher "
           "frequency, which expands same-day and same-time substitutes and "
           "intensifies price competition. This can make Oakland–LAX more "
           "elastic.")
    p = bullet(doc)
    run(p, "Argument for Burbank more elastic:", bold=True)
    run(p, " Burbank may attract a more price-sensitive customer segment that "
           "travels mostly on low-cost airlines, while business and wealthier "
           "(less price-sensitive) travelers prefer LAX. Also, LAX offers "
           "international connections, so people traveling abroad will prefer "
           "to use LAX and will thus be less elastic, being reluctant to "
           "switch airports.")

    # ======================================================================
    # Problem 4
    # ======================================================================
    problem(doc, 4, "Aircraft Demand", 20)
    draws_on(doc, "Module 2 – Demand and Revenue (marginal revenue)")

    p = part(doc, "a", 5)
    run(p, "Plug in the price to find the corresponding quantity:")
    answer(doc, Q + mrun("=") + mrun("250") + mrun("−") + P + mrun("=")
           + mrun("250") + mrun("−") + mrun("140") + mrun("=") + mrun("110"))

    p = part(doc, "b", 10)
    run(p, "First, solve the demand function for ")
    run(p, "P", italic=True)
    run(p, ":")
    T.equation(doc, Q + mrun("=") + mrun("250") + mrun("−") + P
               + mrun("  ⇒  ") + P + mrun("=") + mrun("250") + mrun("−") + Q,
               before=3, after=5)
    body(doc, "Second, compute total revenue:", after=2)
    T.equation(doc, TR + mrun("=") + Q + mrun("×") + P + mrun("=") + Q
               + mrun("(") + mrun("250") + mrun("−") + Q + mrun(")")
               + mrun("=") + mrun("250") + Q + mrun("−")
               + T.msup(Q, mrun("2", italic=False)), before=2, after=5)
    body(doc, "Third, take the derivative of total revenue with respect to "
              "Q to get MR:", after=2)
    T.equation(doc, MR + mrun("=") + mfrac(mrun("d", italic=False) + TR,
                                           mrun("d", italic=False) + Q)
               + mrun("=") + mrun("250") + mrun("−") + mrun("2") + Q,
               before=2, after=5)
    body(doc, "Finally, plug in the quantity corresponding to a price of "
              "$140 million (from part (a)):", after=2)
    answer(doc, MR + mrun("=") + mrun("250") + mrun("−") + mrun("2")
           + mrun("×") + mrun("110") + mrun("=") + mrun("30"))

    p = part(doc, "c", 5)
    run(p, "It is the demand curve for Boeing’s new aircraft.")

    S.place(fig_boeing(), doc, before=8, after=2)
    caption(doc, "Quantity in units, price in millions of dollars. The solid "
                 "segment is the range the question asks for; the dashed "
                 "extensions are the rest of the same demand curve.")

    # ======================================================================
    # Problem 5
    # ======================================================================
    problem(doc, 5, "Lyft Demand", 25)
    draws_on(doc, "Module 2 – Demand Estimation;  Demand and Revenue")

    p = part(doc, "a", 5)
    run(p, "Substitute the assumed values into the estimated demand "
           "equation:")
    answer(doc, Q + mrun("=") + mrun("67.4") + mrun("−") + mrun("16")
           + mrun("×") + mrun("13.2") + mrun("+") + mrun("2.6") + mrun("×")
           + mrun("20.3") + mrun("+") + mrun("39.3") + mrun("×") + mrun("3.8")
           + mrun("=") + mrun("58.32"))

    p = part(doc, "b", 7)
    run(p, "From the elasticity formula:")
    T.equation(doc, E + mrun("=") + mfrac(DQ, DP) + mrun("×") + mfrac(P, Q),
               before=3, after=5)
    p = body(doc)
    run(p, "From the regression, the slope of the demand function is ")
    T.equation_inline(p, mfrac(DQ, DP) + mrun("=") + mrun("−16"))
    run(p, ", and we know that ")
    T.equation_inline(p, P + mrun("=") + mrun("13.20"))
    run(p, " and ")
    T.equation_inline(p, Q + mrun("=") + mrun("58.32"))
    run(p, ". Plugging into the elasticity formula:")
    answer(doc, ED + mrun("=") + mrun("−16") + mrun("×")
           + mfrac(mrun("13.2"), mrun("58.32")) + mrun("=") + mrun("−3.6"))
    body(doc, "Demand is elastic.")

    p = part(doc, "c", 13)
    run(p, "From the regression equation we can obtain an expression for ")
    run(p, "Q", italic=True)
    run(p, " that only depends on ")
    run(p, "P", italic=True)
    run(p, " under the above conditions:")
    T.equation(doc,
               Q + mrun("=") + mrun("67.4") + mrun("−") + mrun("16") + P
               + mrun("+") + mrun("2.6") + mrun("×") + mrun("20.3")
               + mrun("+") + mrun("39.3") + mrun("×") + mrun("3.8")
               + mrun("=") + mrun("269.52") + mrun("−") + mrun("16") + P,
               before=3, after=5)
    p = body(doc, after=2)
    run(p, "Solving for ")
    run(p, "P", italic=True)
    run(p, ":")
    T.equation(doc, P + mrun("=") + mfrac(mrun("269.52") + mrun("−") + Q,
                                          mrun("16")), before=2, after=5)
    body(doc, "Total revenue can then be written as:", after=2)
    T.equation(doc,
               TR + mrun("=") + P + mrun("×") + Q + mrun("=")
               + mfrac(mrun("269.52") + mrun("−") + Q, mrun("16")) + mrun("×")
               + Q + mrun("=") + mfrac(mrun("269.52") + Q, mrun("16"))
               + mrun("−") + mfrac(T.msup(Q, mrun("2", italic=False)),
                                   mrun("16")),
               before=2, after=5)
    body(doc, "Taking the derivative with respect to Q gives MR:", after=2)
    T.equation(doc,
               MR + mrun("=") + mfrac(mrun("269.52"), mrun("16")) + mrun("−")
               + mfrac(mrun("2"), mrun("16")) + Q,
               before=2, after=5)
    # TYPO fixed: the original wrote "Q=58,32" with a decimal comma
    p = body(doc)
    run(p, "From part (a), at ")
    T.equation_inline(p, P + mrun("=") + mrun("13.20"))
    run(p, " we have ")
    T.equation_inline(p, Q + mrun("=") + mrun("58.32"))
    run(p, ". Then:")
    answer(doc, MR + mrun("=") + mfrac(mrun("269.52"), mrun("16")) + mrun("−")
           + mfrac(mrun("2"), mrun("16")) + mrun("×") + mrun("58.32")
           + mrun("=") + mrun("9.55"))

    S.save(doc, OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
