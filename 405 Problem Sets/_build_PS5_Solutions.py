"""Build: Problem Set 5 -- Solutions.docx  (MGMT 405, Fall 2026 format).

Source: `_originals/Problem Set 5 -- Solutions (2025 original).docx` -- a
BUILD INPUT, never deleted.

ARITHMETIC RE-CHECKED, all of it correct as written:
  1     12 miles x $3 = $36 marginal cost for the trip.  WTP = $50 plus
        half an hour of the traveller's time: CEO 50 + 119 = 169,
        business 50 + 50 = 100, tourist 50 + 12.5 = 62.5, student
        50 + 9 = 59.
  2(a)  P = MC = 2, Q = 8, CS = 0.5 x 8 x 8 = 32 = F, profit 32.
  2(b)  academic Q = 6, CS = 0.5 x 6 x 6 = 18 < 32, so they do not buy.
  3(b)  MR1 = 110 - 2q1 - q2 = 15 -> q1 = 47.5 - 0.5 q2
        MR2 = 110 - q1 - 2q2 = 25 -> q2 = 42.5 - 0.5 q1
  3(c)  q1 = 35, q2 = 25, Q = 60, P = 50 ;
        profit1 = 1750 - 825 = 925 ; profit2 = 1250 - 825 = 425
  3(e)  AVC2 = 25 < 50, and -75 beats -700, so it keeps producing.
  3(f)  ATC2(25) = 700/25 + 25 = 53 > 50, so it exits in the long run.
  3(g)  110 - 2q1 = 15 -> q1 = 47.5, P = 62.5, profit = 1,956.25
  3(h)  CS = 0.5 x 60 x 60 = 1,800 with two firms ;
        CS = 0.5 x 47.5 x 47.5 = 1,128.1 with one.
  4     Low is dominant for both; Nash is (Low, Low) paying (15, 15).

FORMATTING (untracked)
  - Course masthead, navy problem headings with a gold points chip, the
    bare centred page number.
  - Derivations set as native OMML; the answer line is dark red.
  - PROFIT IS A LOWER-CASE pi (Teaching CLAUDE.md, 2026-08-30).
  - All figures rebuilt as native, editable Word shapes.  Demand is dark
    red, marginal cost navy, marginal revenue concept blue, consumer
    surplus the course's red wash at 26 %, each surplus label at its
    region's computed centroid.
  - Problem 4's matrix uses Module 7's payoff-matrix format, and marks
    each best response by underlining it -- the Word equivalent of the
    deck's circled number -- with the Nash cell also taking the
    pale-gold fill.  The original showed the same grid twice with the
    reasoning only in prose.

TYPOS FIXED DIRECTLY (unambiguous; reported rather than tracked)
  - Problem 3(h): the original's figure and its closing sentence refer to
    the one-firm scenario as "f.", but the one-firm case is part (g);
    (f) is the exit decision.  Corrected to (g) throughout.

PROPOSED (tracked -- these go beyond formatting)
  1. The Problem 2 heading read "(25 points)" while the problem set
     awards it 30, and its own parts sum to 30 (8+6+5+4+2+5).  Corrected
     to 30.
  2. Problem 3's parts sum to 40 against a 35-point header -- the same
     discrepancy flagged in the problem set itself.  Noted here so the
     two documents say the same thing.

Run:  python _build_PS5_Solutions.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _ps_theme as S
from _ps_theme import (DARKRED, GOLD, GRAY, NAVY, Panel, WD_ALIGN_PARAGRAPH,
                       body, caption, centroid, dele, ins, para, part,
                       problem, run)
import _tn_theme as T
from _tn_theme import CBLUE, acr, mfrac, mrun, msub, msup

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Problem Set 5 -- Solutions.docx")

CS_RED = "C0201B"

Q, P, F = mrun("Q"), mrun("P"), mrun("F")
PI = mrun("π")
q1 = msub(mrun("q"), mrun("1", italic=False))
q2 = msub(mrun("q"), mrun("2", italic=False))
TC1 = msub(acr("TC"), mrun("1", italic=False))
TC2 = msub(acr("TC"), mrun("2", italic=False))
MC1 = msub(acr("MC"), mrun("1", italic=False))
MC2 = msub(acr("MC"), mrun("2", italic=False))
MR1 = msub(acr("MR"), mrun("1", italic=False))
MR2 = msub(acr("MR"), mrun("2", italic=False))
TR1 = msub(acr("TR"), mrun("1", italic=False))
TR2 = msub(acr("TR"), mrun("2", italic=False))
PI1 = msub(PI, mrun("1", italic=False))
PI2 = msub(PI, mrun("2", italic=False))
ATC2 = msub(acr("ATC"), mrun("2", italic=False))

PAYOFFS = [[(20, 20), (10, 26)],
           [(26, 12), (15, 15)]]
# Firm A's best response to each of B's columns, and B's to each of A's
BEST_A = [(1, 0), (1, 1)]
BEST_B = [(0, 1), (1, 1)]

# Problem 1: the four passenger types.  WTP = $50 for the taxi ride plus
# half an hour of the traveller's own time.
WTP = [(400, 169.0), (700, 100.0), (900, 62.5), (1000, 59.0)]


def _fit(points):
    """Least-squares line through the four willingness-to-pay points.

    The question asks for "an approximately linear demand curve" through
    them, and the marked optimum has to sit on whatever line is drawn --
    so the line is fitted here and every annotation is computed from it.
    """
    n = len(points)
    mx = sum(x for x, _ in points) / n
    my = sum(y for _, y in points) / n
    sxy = sum((x - mx) * (y - my) for x, y in points)
    sxx = sum((x - mx) ** 2 for x, _ in points)
    slope = sxy / sxx
    return my - slope * mx, slope       # intercept, slope


A_INT, A_SLOPE = _fit(WTP)              # P = A_INT + A_SLOPE * Q
MC_TRIP = 36.0
Q_STAR = (MC_TRIP - A_INT) / (2 * A_SLOPE)      # MR = a + 2bQ = MC
P_STAR = A_INT + A_SLOPE * Q_STAR


def draws_on(doc, text):
    # after 8 -> 5 on 2026-09-12, part of returning ~30 pt of pure
    # spacing so a page break is never decided by a few points.
    p = para(doc, before=0, after=5, keep_next=True)
    run(p, "Draws on:  " + text, italic=True, color=GRAY, size=9.5)
    return p


def answer(doc, content, before=4, after=8):
    return T.equation(doc, T.color_omml(content, DARKRED),
                      before=before, after=after)


def bullet(doc, text=None, before=2, after=2):
    p = para(doc, before=before, after=after, left=0.32, hang=0.18,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, "•  ", color=NAVY, size=11)
    if text:
        run(p, text, size=11)
    return p


# ==========================================================================
# Figures
# ==========================================================================
def fig_airtaxi():
    """1(b): the four willingness-to-pay points and the linear demand
    curve through them, with MR, MC and the optimal price."""
    f = S.fig(5.90, 4.05, name="1(b) - air taxi demand")
    pn = Panel(f, 0.98, 3.42, 4.24, 3.00, ylabel="P", xlabel="Q")
    QMAX, PMAX = 1100.0, 260.0

    def u(q):
        return q / QMAX * 100.0

    def v(p):
        return p / PMAX * 100.0

    def dem(q):
        return A_INT + A_SLOPE * q

    mr_zero = -A_INT / (2 * A_SLOPE)
    # MC = $36 runs close to the axis on a $260 price scale, so the three
    # curve labels are lifted clear of it and of each other: D well above
    # the MC line, MR just above the axis but below MC, and the MC label
    # itself pulled left so it does not reach D's.
    pn.demand((u(0), v(dem(0))), (u(QMAX), v(dem(QMAX))), label="D",
              lbl_dx=0.07, lbl_dy=-0.40)
    pn.curve((u(0), v(A_INT)), (u(mr_zero), v(0)), CBLUE, label="MR",
             lbl_dx=0.05, lbl_dy=-0.26, w_pt=1.75, name="MR")
    pn.curve((u(0), v(MC_TRIP)), (u(QMAX), v(MC_TRIP)), NAVY, label=None,
             dash="dash", w_pt=1.5, name="MC")
    # far LEFT, just above the MC line: demand and MR are both up at
    # v ~ 90 there, so nothing else can reach it.  Set anywhere to the
    # right and the descending MR line runs through it.
    pn.text(3.0, v(MC_TRIP) + 9.0,
            [("MC = $36", dict(bold=True, color=NAVY, size=10))], align="l")

    for q, p_ in WTP:
        pn.f.dot(pn.x(u(q)), pn.y(v(p_)), d=0.075, color=NAVY, name="WTP")

    pn.equilibrium(u(Q_STAR), v(P_STAR), guides=True)
    pn.tick_y(v(P_STAR), "P*", h=0.19)
    pn.tick_x(u(Q_STAR), "Q*", h=0.19)
    for p_ in (169.0, 100.0):
        pn.tick_y(v(p_), "%g" % p_, h=0.19, italic=False)
    return f


def fig_twopart(business):
    """2(a) / 2(b): the consumer surplus a two-part tariff can extract."""
    a = 10.0 if business else 8.0
    qstar = a - 2.0
    cs = 0.5 * qstar * (a - 2.0)
    f = S.fig(2.95, 2.60, name=("2(a) - business customer" if business
                                else "2(b) - academic customer"))
    pn = Panel(f, 0.62, 2.10, 1.98, 1.62, ylabel="P", xlabel="Q")
    PMAX, QMAX = 12.0, 12.0

    def u(q):
        return q / QMAX * 100.0

    def v(p):
        return p / PMAX * 100.0

    tri = [(u(0), v(a)), (u(qstar), v(2)), (u(0), v(2))]
    pn.polygon(tri, fill=CS_RED, alpha=26000, line=CS_RED, name="CS")
    pn.demand((u(0), v(a)), (u(a), v(0)), label="D", lbl_dx=0.05,
              lbl_dy=-0.14)
    pn.curve((u(0), v(2)), (u(QMAX), v(2)), NAVY, label=None, w_pt=1.5,
             name="MC")
    pn.tick_y(v(2), "2", h=0.17, italic=False)
    pn.tick_y(v(a), "%g" % a, h=0.17, italic=False)
    pn.tick_x(u(qstar), "%g" % qstar, h=0.17, italic=False)
    cu, cv = centroid(tri)
    pn.text(cu, cv + 4.0,
            [("$%g" % cs, dict(bold=True, color=CS_RED, size=10))],
            align="c")
    pn.title("Business" if business else "Academic", size=10, dy=0.22)
    return f


def fig_reaction():
    """3(b) and 3(c): the two reaction functions and the Cournot point.

    Firm 1: q1 = 47.5 - 0.5 q2, i.e. q2 = 95 - 2 q1
    Firm 2: q2 = 42.5 - 0.5 q1
    They cross at (35, 25), which is computed here rather than read off.
    The monopoly output of part (g), q1 = 47.5, is where firm 1's own
    reaction function meets the q1 axis -- which is what part (g) asks
    the student to relate back to this graph.
    """
    f = S.fig(5.90, 4.05, name="3(b) - reaction functions")
    pn = Panel(f, 0.94, 3.42, 4.28, 3.00, ylabel="q₂", xlabel="q₁")
    M = 100.0

    def u(x):
        return x / M * 100.0

    r1 = ((u(0), u(95)), (u(47.5), u(0)))          # firm 1's reaction
    r2 = ((u(0), u(42.5)), (u(85), u(0)))          # firm 2's reaction
    # Both labels are anchored on their own curve's VERTICAL intercept
    # (`lbl_end=0`, i.e. p0) rather than its lower end: R1 has slope -2
    # in these units, so a label set beside the line is cut through by
    # it, while above the intercept there is clear space.  Anchoring on
    # the curve rather than on a typed coordinate is the standing rule.
    def rlab(n, color):
        return [("R", dict(bold=True, color=color, size=11)),
                (n, dict(bold=True, color=color, size=11, subscript=True)),
                (" (firm %s)" % n, dict(bold=True, color=color, size=10))]

    pn.curve(r1[0], r1[1], DARKRED, label=rlab("1", DARKRED), w_pt=2.0,
             name="R1", lbl_end=0, lbl_dx=0.171, lbl_dy=-0.120)
    pn.curve(r2[0], r2[1], NAVY, label=rlab("2", NAVY), w_pt=2.0,
             name="R2", lbl_end=0, lbl_dx=0.171, lbl_dy=-0.225)

    e = S.cross(r1[0], r1[1], r2[0], r2[1])
    pn.equilibrium(e[0], e[1], guides=True)
    pn.tick_x(u(35), "35", h=0.19, italic=False)
    pn.tick_y(u(25), "25", h=0.19, italic=False)

    # part (g): firm 1 alone produces 47.5, where R1 meets the q1 axis
    pn.f.dot(pn.x(u(47.5)), pn.y(u(0)), d=0.070, color=GOLD,
             name="monopoly output")
    pn.text(u(47.5) + 1.5, u(8.0),
            [("47.5", dict(bold=True, color=GOLD, size=10))], align="l")
    return f


def fig_cs(two_firms):
    """3(h): consumer surplus with two firms and with one."""
    qstar, pstar = (60.0, 50.0) if two_firms else (47.5, 62.5)
    cs = 0.5 * qstar * (110.0 - pstar)
    f = S.fig(5.90, 3.55, name=("3(h) - CS with two firms" if two_firms
                                else "3(h) - CS with firm 1 alone"))
    pn = Panel(f, 0.86, 2.92, 4.32, 2.50, ylabel="P", xlabel="Q")
    M = 120.0

    def w(x):
        return x / M * 100.0

    tri = [(w(0), w(110)), (w(qstar), w(pstar)), (w(0), w(pstar))]
    pn.polygon(tri, fill=CS_RED, alpha=26000, line=CS_RED, name="CS")
    pn.demand((w(0), w(110)), (w(110), w(0)), label="D", lbl_dx=0.07,
              lbl_dy=-0.14)
    pn.equilibrium(w(qstar), w(pstar), guides=True)
    pn.tick_y(w(pstar), "%g" % pstar, h=0.19, italic=False)
    pn.tick_x(w(qstar), "%g" % qstar, h=0.19, italic=False)
    pn.tick_y(w(110), "110", h=0.19, italic=False)
    cu, cv = centroid(tri)
    pn.text(cu, cv + 4.0,
            [("CS = %s" % "{:,.1f}".format(cs).rstrip("0").rstrip("."),
              dict(bold=True, color=CS_RED, size=10.5))], align="c")
    pn.title("Two firms (part c)" if two_firms else "Firm 1 alone (part g)",
             size=10.5, dy=0.24)
    return f


# ==========================================================================
def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)

    S.ps_masthead(doc, "Problem Set 5 – Solutions",
                  covers="Covers Modules 6 and 7")

    # ======================================================================
    # Problem 1
    # ======================================================================
    problem(doc, 1, "Pricing in the Market for Air Taxis", 20, before=12)
    draws_on(doc, "Module 6 – Complex Pricing (willingness to pay, "
                  "versioning)")

    p = part(doc, "a", 4)
    run(p, "The distance as the crow flies between LAX and downtown L.A. is "
           "12 miles (see e.g. freemaptools.com), so the marginal cost of "
           "the trip is:")
    answer(doc, mrun("3") + mrun("×") + mrun("12") + mrun("=") + mrun("$36"))

    p = part(doc, "b", 10)
    run(p, "To compute the average willingness to pay for each passenger "
           "type we consider each type’s opportunity cost of time. By "
           "switching from a regular taxi to an air taxi a passenger saves "
           "30 minutes of commuting time, so the willingness to pay is the "
           "$50 taxi fare plus the value of that half hour.")

    T.table(doc, [
        ["Passenger type", "Number", "Hourly rate", "Value of 30 min", "WTP"],
        ["CEOs", "400", "$238", "$119.00", "$169.00"],
        ["Other business travelers", "300", "$100", "$50.00", "$100.00"],
        ["Tourists", "200", "$25", "$12.50", "$62.50"],
        ["Students", "100", "$18", "$9.00", "$59.00"],
    ], widths_in=[2.10, 0.75, 1.00, 1.20, 1.00], size=10,
        align_right=(1, 2, 3, 4))

    S.note(doc, "Your assumptions may deviate from the ones above, which is "
                "fine as long as your reasoning is coherent.",
           prefix="Grading:")

    S.place(fig_airtaxi(), doc, before=8, after=2)
    caption(doc, "The four willingness-to-pay points and the linear demand "
                 "curve fitted through them, P ≈ {:.0f} − {:.3f}Q. Marginal "
                 "revenue meets MC = $36 at about Q* ≈ {:.0f}, giving "
                 "P* ≈ ${:.0f}.".format(A_INT, -A_SLOPE, Q_STAR, P_STAR))

    p = part(doc, "c", 6)
    run(p, "To extract more surplus from its customer base, Joby could offer "
           "a menu of differential services such that different passenger "
           "types select different ride experiences. For instance, a "
           "single-customer ride with luxurious amenities and additional "
           "services may be offered at a premium rate, alongside a basic "
           "shared multiple-customer ride at a cheaper rate. The set of "
           "prices should be incentive compatible, so that CEOs and other "
           "business travelers are better off choosing the premium "
           "experience while tourists and students prefer the standard "
           "shared ride. This is ")
    run(p, "versioning", bold=True)
    run(p, ", or second-degree price discrimination: the company "
           "discriminates between consumers by inducing those with "
           "different willingness to pay to choose different versions of the "
           "product. Ideally Joby would offer four versions, one for each "
           "passenger type.")

    # ======================================================================
    # Problem 2
    # ======================================================================
    # PROPOSED 1: the header read 25; the parts sum to 30, as does the
    # problem set's own header
    problem(doc, 2, "Two-Part Pricing", 30)
    draws_on(doc, "Module 6 – Complex Pricing (two-part tariffs, segment "
                  "pricing)")
    p = body(doc, before=0)
    dele(p, "(25 points)")
    ins(p, "(30 points — the parts sum to 8 + 6 + 5 + 4 + 2 + 5 = 30.)")

    p = part(doc, "a", 8)
    run(p, "You should charge a usage price equal to marginal cost, ")
    T.equation_inline(p, P + mrun("=") + mrun("2"))
    run(p, ". At that price a business customer uses ")
    T.equation_inline(p, Q + mrun("=") + mrun("10") + mrun("−") + mrun("2")
                      + mrun("=") + mrun("8"))
    run(p, " days of computing, and the consumer surplus is:")
    answer(doc, mfrac(mrun("1"), mrun("2")) + mrun("×") + mrun("8")
           + mrun("×") + mrun("(") + mrun("10") + mrun("−") + mrun("2")
           + mrun(")") + mrun("=") + mrun("$32"))
    body(doc,
         "So you would charge F = $32 per month and $2 per day in usage "
         "price. That leaves zero consumer surplus after the subscription "
         "fee, and each business customer yields a profit of $32 per month.")

    p = part(doc, "b", 6)
    run(p, "At a usage price of $2 an academic customer uses ")
    T.equation_inline(p, Q + mrun("=") + mrun("8") + mrun("−") + mrun("2")
                      + mrun("=") + mrun("6"))
    run(p, " days, so their consumer surplus is:")
    answer(doc, mfrac(mrun("1"), mrun("2")) + mrun("×") + mrun("6")
           + mrun("×") + mrun("(") + mrun("8") + mrun("−") + mrun("2")
           + mrun(")") + mrun("=") + mrun("$18"))

    two_up(doc, fig_twopart(True), fig_twopart(False))
    caption(doc, "The shaded triangle is each customer type’s consumer "
                 "surplus at a usage price of $2: $32 for a business "
                 "customer, $18 for an academic one.")

    body(doc,
         "Since $18 of consumer surplus is below the $32 subscription fee "
         "set for business customers, academics would not subscribe to your "
         "service — you would lose that market entirely.")

    p = part(doc, "c", 5)
    run(p, "You would charge P = $2 and F = $18 — the consumer surplus "
           "computed in (b) — to academics. Your profit from this group "
           "would then be $18 per customer.")

    p = part(doc, "d", 4)
    run(p, "The scheme combines a ")
    run(p, "two-part tariff", bold=True)
    run(p, " with ")
    run(p, "segment pricing", bold=True)
    run(p, ", distinguishing between the academic and business customer "
           "segments.")

    p = part(doc, "e", 2)
    run(p, "Business customers would also purchase the academic version, "
           "since it has the same usage price but a lower subscription fee.")

    p = part(doc, "f", 5)
    run(p, "Movie theaters offer senior discounts on tickets, charging lower "
           "prices to customers above a certain age. This creates the same "
           "issue as in (e), since younger customers would prefer to pay the "
           "cheaper senior rate. The problem is avoided by asking customers "
           "to show a government-issued ID confirming their age at entry, "
           "which prevents younger individuals from using the discounted "
           "ticket.")

    # ======================================================================
    # Problem 3
    # ======================================================================
    problem(doc, 3, "Duopoly", 35)
    draws_on(doc, "Module 7 – Oligopoly (Cournot, Bertrand);  Module 4 – "
                  "the produce / stop / exit decision")
    # PROPOSED 2: the same point-total discrepancy the problem set flags
    p = body(doc, before=0)
    ins(p, "[The parts below sum to 40 points, not the 35 in the header: "
           "1 + 7 + 7 + 2 + 4 + 3 + 6 + 6 + 4.]")

    p = part(doc, "a", 1)
    run(p, "Cournot competition.", bold=True)

    p = part(doc, "b", 7)
    run(p, "To derive the reaction functions we set MC = MR for each "
           "producer. Differentiating the cost functions:")
    T.equation(doc, MC1 + mrun("=") + mrun("15") + mrun("      ") + MC2
               + mrun("=") + mrun("25"), before=3, after=4)
    body(doc, "and the marginal revenues are:", after=2)
    T.equation(doc, TR1 + mrun("=") + P + mrun("⋅") + q1 + mrun("=")
               + mrun("110") + q1 + mrun("−") + q2 + q1 + mrun("−")
               + msup(q1, mrun("2", italic=False)) + mrun("  ⇒  ") + MR1
               + mrun("=") + mrun("110") + mrun("−") + q2 + mrun("−")
               + mrun("2") + q1, before=2, after=3)
    T.equation(doc, TR2 + mrun("=") + P + mrun("⋅") + q2 + mrun("=")
               + mrun("110") + q2 + mrun("−") + q1 + q2 + mrun("−")
               + msup(q2, mrun("2", italic=False)) + mrun("  ⇒  ") + MR2
               + mrun("=") + mrun("110") + mrun("−") + q1 + mrun("−")
               + mrun("2") + q2, before=2, after=5)
    body(doc, "Setting MC = MR and solving for each firm’s own output:",
         after=2)
    answer(doc, q1 + mrun("=") + mrun("47.5") + mrun("−") + mrun("0.5") + q2
           + mrun("            ") + q2 + mrun("=") + mrun("42.5")
           + mrun("−") + mrun("0.5") + q1)

    S.place(fig_reaction(), doc, before=8, after=2)
    caption(doc, "The two reaction functions. They cross at "
                 "(q₁, q₂) = (35, 25), the Cournot equilibrium of part (c). "
                 "Firm 1’s reaction function meets the horizontal axis at "
                 "47.5, which is its monopoly output in part (g).")

    body(doc,
         "These curves give the optimal, profit-maximizing quantity a firm "
         "should produce for any quantity produced by its competitor. For "
         "firm 2, the reaction function reads off the vertical axis the "
         "optimal q₂ for each q₁ that firm 1 produces on the horizontal "
         "axis.")

    p = part(doc, "c", 7)
    run(p, "Substituting firm 2’s reaction function into firm 1’s:")
    T.equation(doc, q1 + mrun("=") + mrun("47.5") + mrun("−") + mrun("0.5")
               + mrun("(") + mrun("42.5") + mrun("−") + mrun("0.5") + q1
               + mrun(")") + mrun("=") + mrun("26.25") + mrun("+")
               + mrun("0.25") + q1, before=3, after=4)
    answer(doc, q1 + mrun("=") + mrun("35") + mrun("            ") + q2
           + mrun("=") + mrun("42.5") + mrun("−") + mrun("0.5") + mrun("×")
           + mrun("35") + mrun("=") + mrun("25"))
    body(doc, "so total output and the price are:", after=2)
    answer(doc, Q + mrun("=") + mrun("60") + mrun("            ") + P
           + mrun("=") + mrun("110") + mrun("−") + mrun("60") + mrun("=")
           + mrun("50"), before=2)
    body(doc, "and the profits are:", after=2)
    answer(doc, PI1 + mrun("=") + mrun("50") + mrun("×") + mrun("35")
           + mrun("−") + mrun("(") + mrun("300") + mrun("+") + mrun("15")
           + mrun("×") + mrun("35") + mrun(")") + mrun("=") + mrun("925"),
           before=2, after=3)
    answer(doc, PI2 + mrun("=") + mrun("50") + mrun("×") + mrun("25")
           + mrun("−") + mrun("(") + mrun("200") + mrun("+") + mrun("25")
           + mrun("×") + mrun("25") + mrun(")") + mrun("=") + mrun("425"),
           before=2)

    p = part(doc, "d", 2)
    run(p, "The rent is a fixed cost:")
    answer(doc, TC2 + mrun("=") + mrun("500") + mrun("+") + mrun("200")
           + mrun("+") + mrun("25") + q2 + mrun("=") + mrun("700")
           + mrun("+") + mrun("25") + q2)

    p = part(doc, "e", 4)
    run(p, "Average variable cost is 25, which is below the price of 50, so "
           "by continuing to produce firm 2 makes a positive contribution to "
           "covering its fixed costs. Comparing the two situations directly: "
           "profits would be −700 (the full fixed cost) if it stopped "
           "producing now, against")
    answer(doc, PI2 + mrun("=") + mrun("50") + mrun("×") + mrun("25")
           + mrun("−") + mrun("(") + mrun("700") + mrun("+") + mrun("25")
           + mrun("×") + mrun("25") + mrun(")") + mrun("=") + mrun("−75"))
    body(doc,
         "if it keeps producing. The quantity does not change, because the "
         "rental cost is fixed and so does not affect the equilibrium price "
         "or quantity: firm 2 still produces 25 units.")

    p = part(doc, "f", 3)
    run(p, "Since profits after the rent and the other fixed costs are −75, "
           "firm 2 will exit the market next year, when it has the "
           "possibility to do so. Equivalently, compare the price with "
           "average total cost at ")
    T.equation_inline(p, q2 + mrun("=") + mrun("25"))
    run(p, ":")
    answer(doc, ATC2 + mrun("=") + mfrac(TC2, q2) + mrun("=")
           + mfrac(mrun("700"), mrun("25")) + mrun("+") + mrun("25")
           + mrun("=") + mrun("53") + mrun("  >  ") + P + mrun("=")
           + mrun("50"))
    body(doc, "so the firm makes negative profits and should exit in the "
              "long run.")

    p = part(doc, "g", 6)
    run(p, "Firm 1 is now a monopolist:")
    T.equation(doc, TR1 + mrun("=") + mrun("110") + q1 + mrun("−")
               + msup(q1, mrun("2", italic=False)) + mrun("  ⇒  ") + MR1
               + mrun("=") + mrun("110") + mrun("−") + mrun("2") + q1,
               before=3, after=4)
    body(doc, "Setting MR = MC = 15:", after=2)
    answer(doc, mrun("110") + mrun("−") + mrun("2") + q1 + mrun("=")
           + mrun("15") + mrun("  ⇒  ") + q1 + mrun("=") + mrun("47.5"))
    answer(doc, P + mrun("=") + mrun("110") + mrun("−") + mrun("47.5")
           + mrun("=") + mrun("62.5"), before=2)
    answer(doc, PI1 + mrun("=") + mrun("62.5") + mrun("×") + mrun("47.5")
           + mrun("−") + mrun("(") + mrun("300") + mrun("+") + mrun("15")
           + mrun("×") + mrun("47.5") + mrun(")") + mrun("=")
           + mrun("1,956.25"), before=2)
    body(doc,
         "On the graph in part (b) this is the point where firm 1’s reaction "
         "function meets the horizontal axis: with q₂ = 0, firm 1’s best "
         "response is 47.5.")

    p = part(doc, "h", 6)
    run(p, "Consumer surplus is the triangle between the demand curve and "
           "the market price:")
    answer(doc, msub(acr("CS"), mrun("c", italic=False)) + mrun("=")
           + mfrac(mrun("1"), mrun("2")) + mrun("×") + mrun("60") + mrun("×")
           + mrun("(") + mrun("110") + mrun("−") + mrun("50") + mrun(")")
           + mrun("=") + mrun("1,800"))
    answer(doc, msub(acr("CS"), mrun("g", italic=False)) + mrun("=")
           + mfrac(mrun("1"), mrun("2")) + mrun("×") + mrun("47.5")
           + mrun("×") + mrun("(") + mrun("110") + mrun("−") + mrun("62.5")
           + mrun(")") + mrun("=") + mrun("1,128.1"), before=2)

    S.place(fig_cs(True), doc, before=8, after=2)
    caption(doc, "Consumer surplus with two firms in the market (part c).")
    S.place(fig_cs(False), doc, before=8, after=2)
    caption(doc, "Consumer surplus with firm 1 alone (part g).")

    body(doc,
         "Consumers are worse off with only one producer: consumer surplus "
         "falls from 1,800 to 1,128.1. Don did consumers no favor by "
         "demanding high land rental payments and pushing firm 2 out of "
         "business.", before=8)

    p = part(doc, "i", 4)
    run(p, "This is ")
    run(p, "Bertrand competition", bold=True)
    run(p, ". Firm 1 will set P = 24.99, just below firm 2’s marginal cost "
           "of 25, which pushes firm 2 out of the market. Total quantity is "
           "then Q = 110 − 24.99 = 85.01.")
    body(doc,
         "Would firm 1 want to set an even lower price, since it is a "
         "monopolist at any price below 25? No. As we saw in part (g), a "
         "monopolist would produce 47.5 and charge 62.5, so if anything firm "
         "1 would want to charge a price above $25 — but that would bring "
         "firm 2 back into the market. So P = 24.99 is firm 1’s optimal "
         "price.")

    # ======================================================================
    # Problem 4
    # ======================================================================
    problem(doc, 4, "Price Wars", 15)
    draws_on(doc, "Module 7 – Game Theory (dominant strategies, Nash "
                  "equilibrium)")

    p = part(doc, "a", 10)
    run(p, "Each firm wants to maximize its profit. Firm A will always "
           "choose the Low Ad Rate, independent of firm B’s choice, and firm "
           "B will always choose the Low Ad Rate, independent of firm A’s "
           "choice. The Low Ad Rate is therefore a ")
    run(p, "dominant strategy", bold=True)
    run(p, " for both firms:")

    bullet(doc, "If firm B chooses the High Ad Rate, firm A chooses the Low "
                "Ad Rate, because it gives firm A 26 rather than 20. If firm "
                "B chooses the Low Ad Rate, firm A again chooses the Low Ad "
                "Rate, because it gives 15 rather than 10.")
    bullet(doc, "If firm A chooses the High Ad Rate, firm B chooses the Low "
                "Ad Rate, because it gives firm B 26 rather than 20. If firm "
                "A chooses the Low Ad Rate, firm B again chooses the Low Ad "
                "Rate, because it gives 15 rather than 12.")

    body(doc, "Each firm’s best response is underlined below. The one cell "
              "where both payoffs are underlined is the Nash equilibrium.",
         before=8, after=8)

    S.payoff_matrix(doc, "FIRM A", "FIRM B",
                    ["High Ad Rate", "Low Ad Rate"],
                    ["High Ad Rate", "Low Ad Rate"],
                    PAYOFFS, best_row=BEST_A, best_col=BEST_B)

    p = part(doc, "b", 5, before=10)
    run(p, "The Nash equilibrium of this game is ")
    run(p, "{firm A plays Low Ad Rate, firm B plays Low Ad Rate}", bold=True)
    run(p, ". The payoff for each firm is ")
    run(p, "15", bold=True, color=DARKRED)
    run(p, ".")

    S.save(doc, OUT)
    print("wrote", OUT)


def two_up(doc, fa, fb, gap=0.30):
    """Place two small figures side by side in one centred paragraph."""
    p = para(doc, before=8, after=2)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    r._r.append(T.parse_xml(fa.xml(doc_pr_id=811)))
    run(p, " " * 6)
    r2 = p.add_run()
    r2._r.append(T.parse_xml(fb.xml(doc_pr_id=812)))
    return p


if __name__ == "__main__":
    main()
