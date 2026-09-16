"""Build: Problem Set 4 -- Solutions.docx  (MGMT 405, Fall 2026 format).

Source: `_originals/Problem Set 4 -- Solutions (2025 original).docx` -- a
BUILD INPUT, never deleted.

ARITHMETIC RE-CHECKED, all of it correct as written:
  1(a)  Q = 2,500 ; TR = 5,000,000 ; TC = 1,150,000 ; profit = 3,850,000
  1(b)  E = -1 x 2000/2500 = -0.8  (inelastic)
  1(c)  MR = 4,500 - 2Q = 400 -> Q = 2,050, P = 2,450 ;
        TR = 5,022,500 ; TC = 970,000 ; profit = 4,052,500 ;
        E = -1 x 2450/2050 = -1.195  (elastic)
  1(d)  (2450 - 400)/2450 = 0.837 = 1/1.195
  2(a)  55 - 4Q = 2Q - 5 -> Q = 10, P = 35, profit = 350 - 150 = 200
  2(b)  2Q - 5 = 55 - 2Q -> Q = 15, P = 25, profit = 375 - 250 = 125
  2(c)  CS 225 / PS 225 competitive ; CS 100 / PS 300 monopoly ;
        DWL = 450 - 400 = 50, which is also 0.5 x 20 x 5.
  4(a)  (730 - 2Q)/175 = 0.17 -> Q = 350.13, P = 2.17
  4(b)  (855 - 2Q)/319 = 0.17 -> Q = 400.38, P = 1.43
  4(c)  -175 x 2.17/350.13 = -1.08 ; -319 x 1.43/400.38 = -1.14
  5     Lawyers 495/820 = 0.60 ; HR managers 304/823 = 0.37

FORMATTING (untracked)
  - Course masthead, navy problem headings with a gold points chip, the
    bare centred page number.
  - Derivations set as native OMML; the answer line is dark red.
  - PROFIT IS A LOWER-CASE pi (Teaching CLAUDE.md, 2026-08-30).  The
    original used the capital.
  - All figures rebuilt as native, editable Word shapes.
  - CURVE COLOURS follow Module 4's price-taker panel, which states the
    rule outright: the line a firm can sell at IS the demand curve it
    faces, so demand and any P = MR line are DARK RED, marginal cost is
    NAVY, average cost is GOLD.  Marginal revenue for a price setter is
    concept blue `0070C0`, the colour the course reserves for MR.
  - CONSUMER AND PRODUCER SURPLUS use the fixed washes (2026-08-30): CS
    a red `C0201B` wash at 26 %, PS a blue `4E79B5` wash at 34 %, and
    each label sits at its region's COMPUTED centroid, not eyeballed.
    Deadweight loss is a gray wash.  Each region is built from the
    curves' own points, so a sloped edge lies exactly on the curve.
  - Problem 5's two task tables become native tables; the AI-replacement
    index is recomputed from them rather than restated, so the totals
    and the index cannot drift apart.

PROPOSED (tracked -- these go beyond formatting)
  1. Problem 2's cost function is written "C = ..." here and "TC = ..."
     in the problem set.  Made TC in both.
  2. Problem 1: two sentences connecting (b) and (c).  At $2,000 demand
     is inelastic, and the profit-maximizing price turns out to be
     HIGHER, at $2,450 -- which is Module 2's lesson (slide 49) playing
     out in this firm's own numbers.  The original solution computes
     both elasticities and never connects them.
  3. Problem 3(b): the original solution said the market "is a
     monopolistically competitive market" and then listed the
     characteristics, but never said which of the three curves is the
     answer beyond pasting the picture.  Named explicitly as curve (C).

Run:  python _build_PS4_Solutions.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _ps_theme as S
from _ps_theme import (DARKRED, GOLD, GRAY, NAVY, Panel, WD_ALIGN_PARAGRAPH,
                       body, caption, centroid, cross, dele, ins, para, part,
                       problem, run)
import _tn_theme as T
from _tn_theme import CBLUE, acr, mfrac, mrun, msub, msup

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Problem Set 4 -- Solutions.docx")

CS_RED = "C0201B"          # consumer-surplus wash (MW palette, 2026-08-30)
PS_BLUE = "4E79B5"         # producer-surplus wash
PS_BLUE_LINE = "2E5AA8"

Q = mrun("Q")
P = mrun("P")
PI = mrun("π")
TR, TC, MC, MR = acr("TR"), acr("TC"), acr("MC"), acr("MR")
ED = msub(mrun("E", italic=False), mrun("D"))
QA, PA = msub(Q, mrun("A")), msub(P, mrun("A"))
QB, PB = msub(Q, mrun("B")), msub(P, mrun("B"))
MRA, MRB = msub(MR, mrun("A")), msub(MR, mrun("B"))
EDA, EDB = msub(ED, mrun("A")), msub(ED, mrun("B"))

LAWYERS = [
    (89, "Interpret laws, rulings, and regulations for individuals and "
         "businesses.", 1),
    (85, "Analyze the probable outcomes of cases, using knowledge of legal "
         "precedents.", 1),
    (85, "Gather evidence to formulate defense or initiate legal actions "
         "(e.g. interviewing clients and witnesses).", 1),
    (84, "Represent clients in court or before government agencies.", 0),
    (83, "Evaluate findings and develop strategies and arguments in "
         "preparation for cases.", 0),
    (82, "Advise clients concerning business transactions, liability, or "
         "legal rights and obligations.", 0),
    (80, "Examine legal data to determine advisability of defending or "
         "prosecuting lawsuits.", 1),
    (80, "Prepare, draft, and review legal documents (e.g. wills, deeds, "
         "patents, mortgages, leases, contracts).", 1),
    (76, "Study Constitution, statutes, decisions, and ordinances to "
         "determine case ramifications.", 1),
    (76, "Negotiate settlements of civil disputes.", 0),
]

HR = [
    (92, "Serve as a link between management and employees by handling "
         "questions and resolving problems.", 0),
    (89, "Plan, direct, supervise, and coordinate work activities related to "
         "employment and labor relations.", 0),
    (87, "Perform staffing duties (e.g. recruiting, interviewing, hiring, "
         "disciplinary procedures).", 0),
    (85, "Represent organization at personnel-related hearings and "
         "investigations.", 0),
    (83, "Negotiate bargaining agreements and interpret labor contracts.", 0),
    (83, "Advise managers on policy matters (e.g. EEO and sexual "
         "harassment).", 0),
    (77, "Plan and conduct new employee orientation programs.", 1),
    (76, "Analyze and modify compensation and benefits policies to ensure "
         "compliance and competitiveness.", 1),
    (76, "Identify staff vacancies and recruit, interview, and select "
         "applicants.", 1),
    (75, "Investigate and report on industrial accidents for insurance "
         "carriers.", 1),
]


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


def task_table(doc, rows, label):
    tot_imp = sum(i for i, _, _ in rows)
    tot_prod = sum(i * d for i, _, d in rows)
    body_rows = [["Importance (A)", "Task", "AI-replacement (B)", "A × B"]]
    for imp, task, dummy in rows:
        body_rows.append([str(imp), task, str(dummy), str(imp * dummy)])
    body_rows.append(["{:,}".format(tot_imp), "Total", "",
                      "{:,}".format(tot_prod)])
    n = len(body_rows) - 1
    T.table(doc, body_rows, widths_in=[0.85, 3.75, 1.00, 0.60], size=8.5,
            align_right=(0, 2, 3),
            highlight=[(n, c) for c in range(4)])
    caption(doc, label, before=4, after=8)
    return tot_imp, tot_prod


# ==========================================================================
# Figures
# ==========================================================================
def fig_welfare(monopoly):
    """2(c): consumer and producer surplus, competition vs monopoly.

    Exact functions, as the original plotted them:
        D  : P = 55 - 2Q      MC : P = 2Q - 5      MR : P = 55 - 4Q
        competitive (15, 25)  monopoly (10, 35)    MC(10) = 15
    MC is negative below Q = 2.5, so the frame runs down to P = -10 and
    the P = 0 line is drawn in as its own reference; putting the axis on
    P = 0 would cut the producer-surplus region off at the bottom.
    """
    f = S.fig(5.90, 4.15, name=("2(c) - monopoly" if monopoly
                                else "2(c) - perfect competition"))
    pn = Panel(f, 0.86, 3.50, 4.30, 3.05, ylabel="P", xlabel="Q")

    QMAX, YMIN, YMAX = 30.0, -10.0, 60.0

    def u(q):
        return q / QMAX * 100.0

    def v(p):
        return (p - YMIN) / (YMAX - YMIN) * 100.0

    def L(pts):
        return [(u(q), v(p)) for q, p in pts]

    qstar, pstar = (10.0, 35.0) if monopoly else (15.0, 25.0)
    mc_at = 2.0 * qstar - 5.0                      # 15 under monopoly, 25 else

    # --- regions first, so the curves are drawn over them --------------
    cs = [(0.0, 55.0), (qstar, pstar), (0.0, pstar)]
    ps = [(0.0, pstar), (qstar, pstar), (qstar, mc_at), (0.0, -5.0)]
    pn.polygon(L(cs), fill=CS_RED, alpha=26000, line=CS_RED, name="CS")
    pn.polygon(L(ps), fill=PS_BLUE, alpha=34000, line=PS_BLUE_LINE,
               name="PS")
    if monopoly:
        dwl = [(qstar, pstar), (15.0, 25.0), (qstar, mc_at)]
        pn.polygon(L(dwl), fill=GRAY, alpha=32000, line=GRAY, name="DWL")

    # --- curves ---------------------------------------------------------
    # D ends exactly ON the P = 0 reference line drawn below, so the
    # label is lifted a full label height clear of it -- at -0.05 the
    # dashed line ran straight through the letter.
    pn.demand((u(0), v(55)), (u(27.5), v(0)), label="D", lbl_dx=0.08,
              lbl_dy=-0.27)
    pn.supply((u(0), v(-5)), (u(QMAX), v(55)), label="MC", lbl_dx=0.07,
              lbl_dy=-0.20, name="MC")
    if monopoly:
        pn.curve((u(0), v(55)), (u(13.75), v(0)), CBLUE, label="MR",
                 lbl_dx=0.05, lbl_dy=0.02, w_pt=1.75, name="MR")

    # the P = 0 reference line, since the frame runs below it
    pn.f.line(pn.x(0), pn.y(v(0)), pn.x(100), pn.y(v(0)), color=S.LIGHT,
              w_pt=1.0, dash="dash", name="P = 0")

    # --- the optimum ----------------------------------------------------
    pn.equilibrium(u(qstar), v(pstar), guides=True)
    pn.tick_y(v(pstar), "%g" % pstar)
    pn.tick_x(u(qstar), "%g" % qstar)
    if monopoly:
        pn.f.dot(pn.x(u(15.0)), pn.y(v(25.0)), d=0.065, color=NAVY,
                 name="competitive point")

    # --- region labels, at each region's computed centroid ---------------
    # With a linear demand curve the MR line passes EXACTLY through the
    # consumer-surplus centroid: the centroid is at Q*/3, and MR there is
    # 55 - 4Q*/3, which is the centroid's own height.  So on the monopoly
    # panel the CS label is moved down and to the left, into the wide part
    # of the triangle on the far side of MR.  On the competitive panel
    # there is no MR line and the centroid is used as it is.
    for pts, lab, col in ((cs, "CS", CS_RED), (ps, "PS", PS_BLUE_LINE)):
        cu, cv = centroid(L(pts))
        if monopoly and lab == "CS":
            cu, cv = cu * 0.60, cv - 6.0
        pn.text(cu, cv + 3.5, [(lab, dict(bold=True, color=col, size=11))],
                align="c")
    if monopoly:
        # centred INSIDE the triangle, not offset to the side: the region
        # is 0.7" wide at its centroid and the label is 0.35", so it fits,
        # whereas offsetting it right put it on top of the demand curve
        cu, cv = centroid(L(dwl))
        pn.text(cu, cv + 3.5,
                [("DWL", dict(bold=True, color=NAVY, size=10))], align="c")
    return f


def fig_avocado(city):
    """4(a) / 4(b): the monopolist's optimum in each city.

    Both panels share one pair of axes so the two can be compared, which
    is what part (d) asks for.  Demand dark red, MR concept blue, MC
    navy -- the course's colours for a price setter.
    """
    if city == "Atlanta":
        a, b, qstar, pstar = 730.0, 175.0, 350.13, 2.17
    else:
        a, b, qstar, pstar = 855.0, 319.0, 400.38, 1.43
    mc = 0.17
    QMAX, PMAX = 900.0, 4.5
    f = S.fig(5.90, 3.95, name="4 - optimal price in " + city)
    pn = Panel(f, 0.86, 3.32, 4.30, 2.85, ylabel="P", xlabel="Q")

    def u(q):
        return q / QMAX * 100.0

    def v(p):
        return p / PMAX * 100.0

    # P = (a - Q)/b, so the price intercept is a/b and the quantity
    # intercept is a; MR has the same price intercept and half the reach
    pn.demand((u(0), v(a / b)), (u(a), v(0)), label=None)
    pn.curve((u(0), v(a / b)), (u(a / 2.0), v(0)), CBLUE, label=None,
             w_pt=1.75, name="MR")
    pn.curve((u(0), v(mc)), (u(QMAX), v(mc)), NAVY, label=None, w_pt=1.75,
             name="MC")

    # Marginal cost is $0.17 against prices of $1.43 to $2.17, so the MC
    # line necessarily runs close to the axis.  All three curve labels are
    # therefore placed by hand, in the clear band between two curves,
    # rather than at the curve ends where they would pile up on the axis
    # and on each other.
    def lab(uu, vv, text, color):
        pn.text(uu, vv, [(text, dict(bold=True, color=color, size=10.5))],
                align="l")

    lab(u(0.78 * a), v(a / b * 0.30) + 8.0, "D", DARKRED)
    lab(u(0.52 * a), v(mc) + 15.0, "MR", CBLUE)
    lab(u(0.06 * a), v(mc) + 8.0, "MC = $0.17", NAVY)

    pn.equilibrium(u(qstar), v(pstar), guides=True)
    pn.tick_y(v(pstar), "$%.2f" % pstar)
    pn.tick_x(u(qstar), "%.0f" % qstar)
    pn.title(city, size=11, dy=0.24)
    return f


# ==========================================================================
def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)

    S.ps_masthead(doc, "Problem Set 4 – Solutions",
                  covers="Covers Modules 5 and 6", due="100 points")

    # ======================================================================
    # Problem 1
    # ======================================================================
    problem(doc, 1, "Maguire Alarm Company", 25, before=12)
    draws_on(doc, "Module 5 – Monopoly;  Module 6 – Complex Pricing "
                  "(the Lerner Index)")

    p = part(doc, "a", 4)
    run(p, "Profits are total revenue minus total cost, so we need both. "
           "From the demand function at ")
    T.equation_inline(p, P + mrun("=") + mrun("2,000"))
    run(p, ":")
    T.equation(doc, Q + mrun("=") + mrun("4,500") + mrun("−") + mrun("2,000")
               + mrun("=") + mrun("2,500"), before=3, after=4)
    T.equation(doc, TR + mrun("=") + P + mrun("×") + Q + mrun("=")
               + mrun("2,000") + mrun("×") + mrun("2,500") + mrun("=")
               + mrun("5,000,000"), before=2, after=4)
    T.equation(doc, TC + mrun("=") + mrun("150,000") + mrun("+") + mrun("400")
               + mrun("×") + mrun("2,500") + mrun("=") + mrun("1,150,000"),
               before=2, after=4)
    answer(doc, PI + mrun("=") + mrun("5,000,000") + mrun("−")
           + mrun("1,150,000") + mrun("=") + mrun("3,850,000"))

    p = part(doc, "b", 4)
    run(p, "The own-price elasticity of demand is")
    T.equation(doc, ED + mrun("=") + mfrac(mrun("d", italic=False) + Q,
                                           mrun("d", italic=False) + P)
               + mrun("×") + mfrac(P, Q), before=3, after=4)
    p = body(doc)
    run(p, "From the demand equation the slope is ")
    T.equation_inline(p, mfrac(mrun("d", italic=False) + Q,
                               mrun("d", italic=False) + P) + mrun("=")
                      + mrun("−1"))
    run(p, ", and from part (a) we know that at ")
    T.equation_inline(p, P + mrun("=") + mrun("2,000"))
    run(p, " we have ")
    T.equation_inline(p, Q + mrun("=") + mrun("2,500"))
    run(p, ". Hence:")
    answer(doc, ED + mrun("=") + mrun("−1") + mrun("×")
           + mfrac(mrun("2,000"), mrun("2,500")) + mrun("=") + mrun("−0.8"))
    body(doc, "Demand is inelastic at the price of $2,000.")

    p = part(doc, "c", 12)
    run(p, "To find the optimum we set MC = MR. Marginal cost comes straight "
           "from the cost function:")
    T.equation(doc, MC + mrun("=") + mfrac(mrun("d", italic=False) + TC,
                                           mrun("d", italic=False) + Q)
               + mrun("=") + mrun("400"), before=3, after=4)
    body(doc, "and we use the three-step method for marginal revenue:",
         after=2)
    T.equation(doc, P + mrun("=") + mrun("4,500") + mrun("−") + Q,
               before=2, after=3)
    T.equation(doc, TR + mrun("=") + P + Q + mrun("=") + mrun("4,500") + Q
               + mrun("−") + msup(Q, mrun("2", italic=False)),
               before=2, after=3)
    T.equation(doc, MR + mrun("=") + mfrac(mrun("d", italic=False) + TR,
                                           mrun("d", italic=False) + Q)
               + mrun("=") + mrun("4,500") + mrun("−") + mrun("2") + Q,
               before=2, after=5)
    body(doc, "Setting MC = MR:", after=2)
    T.equation(doc, mrun("400") + mrun("=") + mrun("4,500") + mrun("−")
               + mrun("2") + Q + mrun("  ⇒  ") + Q + mrun("=")
               + mrun("2,050"), before=2, after=4)
    body(doc, "and substituting back into the demand equation:", after=2)
    answer(doc, P + mrun("=") + mrun("4,500") + mrun("−") + mrun("2,050")
           + mrun("=") + mrun("2,450"))

    body(doc, "Profits at the optimum:", after=2)
    T.equation(doc, TR + mrun("=") + mrun("2,450") + mrun("×") + mrun("2,050")
               + mrun("=") + mrun("5,022,500"), before=2, after=3)
    T.equation(doc, TC + mrun("=") + mrun("150,000") + mrun("+") + mrun("400")
               + mrun("×") + mrun("2,050") + mrun("=") + mrun("970,000"),
               before=2, after=4)
    answer(doc, PI + mrun("=") + mrun("5,022,500") + mrun("−")
           + mrun("970,000") + mrun("=") + mrun("4,052,500"))

    p = body(doc)
    run(p, "The elasticity at the optimum, with ")
    T.equation_inline(p, P + mrun("=") + mrun("2,450"))
    run(p, " and ")
    T.equation_inline(p, Q + mrun("=") + mrun("2,050"))
    run(p, ":")
    answer(doc, ED + mrun("=") + mrun("−1") + mrun("×")
           + mfrac(mrun("2,450"), mrun("2,050")) + mrun("=") + mrun("−1.195"))
    body(doc,
         "Demand is elastic at the optimal price. At the optimum, where "
         "MR = MC, the firm is on the elastic portion of the demand curve.")

    # PROPOSED 2: connect (b) and (c)
    p = body(doc)
    ins(p, "This is worth putting next to part (b). At $2,000 demand was "
           "inelastic, and the profit-maximizing price turns out to be "
           "higher, at $2,450. That is the general result from Module 2: a "
           "firm on the inelastic part of its demand curve can raise both "
           "revenue and profit by raising its price, so an inelastic point "
           "can never be the optimum.")

    p = part(doc, "d", 5)
    run(p, "The Lerner Index is the share of the markup in the price:")
    answer(doc, mrun("Lerner Index", italic=False) + mrun("=")
           + mfrac(P + mrun("−") + MC, P) + mrun("=")
           + mfrac(mrun("2,450") + mrun("−") + mrun("400"), mrun("2,450"))
           + mrun("=") + mrun("0.837"))
    body(doc,
         "where the optimal price is $2,450 and MC = 400, both computed in "
         "(c). The share of markup in the profit-maximizing price is "
         "therefore 83.7%. We have seen in class that at the "
         "profit-maximizing P and Q this also equals one over the absolute "
         "elasticity:")
    answer(doc, mfrac(P + mrun("−") + MC, P) + mrun("=")
           + mfrac(mrun("1"), mrun("|") + ED + mrun("|")) + mrun("=")
           + mfrac(mrun("1"), mrun("1.195")) + mrun("=") + mrun("0.837"))
    body(doc, "which confirms the calculation above.")

    # ======================================================================
    # Problem 2
    # ======================================================================
    problem(doc, 2, "Dayna’s Doorstops", 20)
    draws_on(doc, "Module 5 – Monopoly;  Module 4 – consumer surplus, "
                  "producer surplus and deadweight loss")

    # PROPOSED 1: TC, not C, matching the problem statement
    p = body(doc)
    run(p, "Dayna’s Doorstops, Inc. (DD) is a monopolist in the doorstop "
           "industry. Its total cost is ")
    dele(p, "C")
    ins(p, "TC")
    run(p, " = 100 − 5")
    run(p, "Q", italic=True)
    run(p, " + ")
    run(p, "Q", italic=True)
    run(p, "2", italic=True, superscript=True)
    run(p, ", and demand is ")
    run(p, "P", italic=True)
    run(p, " = 55 − 2")
    run(p, "Q", italic=True)
    run(p, ".")

    p = part(doc, "a", 5)
    run(p, "To find the optimal Q and P we set MR = MC. From total revenue:")
    T.equation(doc, TR + mrun("=") + P + mrun("×") + Q + mrun("=") + mrun("(")
               + mrun("55") + mrun("−") + mrun("2") + Q + mrun(")") + Q
               + mrun("=") + mrun("55") + Q + mrun("−") + mrun("2")
               + msup(Q, mrun("2", italic=False)), before=3, after=3)
    T.equation(doc, MR + mrun("=") + mrun("55") + mrun("−") + mrun("4") + Q,
               before=2, after=4)
    body(doc, "and from total cost:", after=2)
    T.equation(doc, MC + mrun("=") + mrun("2") + Q + mrun("−") + mrun("5"),
               before=2, after=4)
    body(doc, "Setting them equal:", after=2)
    T.equation(doc, mrun("55") + mrun("−") + mrun("4") + Q + mrun("=")
               + mrun("2") + Q + mrun("−") + mrun("5") + mrun("  ⇒  ")
               + msub(Q, mrun("monopoly", italic=False)) + mrun("=")
               + mrun("10"), before=2, after=4)
    answer(doc, msub(P, mrun("monopoly", italic=False)) + mrun("=")
           + mrun("55") + mrun("−") + mrun("2") + mrun("×") + mrun("10")
           + mrun("=") + mrun("35"))
    body(doc, "Profits for the monopolist are:", after=2)
    answer(doc, msub(PI, mrun("monopoly", italic=False)) + mrun("=")
           + mrun("(") + mrun("35") + mrun("×") + mrun("10") + mrun(")")
           + mrun("−") + mrun("(") + mrun("100") + mrun("−") + mrun("50")
           + mrun("+") + mrun("100") + mrun(")") + mrun("=") + mrun("200"))

    p = part(doc, "b", 7)
    run(p, "If DD acts as a perfect competitor it sets MC = P:")
    T.equation(doc, mrun("2") + Q + mrun("−") + mrun("5") + mrun("=")
               + mrun("55") + mrun("−") + mrun("2") + Q + mrun("  ⇒  ")
               + msub(Q, mrun("comp", italic=False)) + mrun("=") + mrun("15"),
               before=3, after=4)
    answer(doc, msub(P, mrun("comp", italic=False)) + mrun("=") + mrun("55")
           + mrun("−") + mrun("2") + mrun("×") + mrun("15") + mrun("=")
           + mrun("25"))
    body(doc, "and profits are:", after=2)
    answer(doc, msub(PI, mrun("comp", italic=False)) + mrun("=") + mrun("375")
           + mrun("−") + mrun("(") + mrun("100") + mrun("−") + mrun("75")
           + mrun("+") + mrun("225") + mrun(")") + mrun("=") + mrun("125"))

    p = part(doc, "c", 8)
    run(p, "Consumer surplus (CS) and producer surplus (PS) under perfect "
           "competition:")
    S.place(fig_welfare(False), doc, before=6, after=2)
    caption(doc, "Perfect competition: P = MC at Q = 15, P = 25. "
                 "CS = ½ × 15 × 30 = 225 and PS = 225, so total surplus "
                 "is 450.")

    body(doc, "Consumer and producer surplus under monopoly, with the "
              "deadweight loss:", before=8)
    S.place(fig_welfare(True), doc, before=6, after=2)
    caption(doc, "Monopoly: MR = MC at Q = 10, priced off demand at P = 35. "
                 "CS falls to 100 and PS rises to 300, so total surplus is "
                 "400 and the deadweight loss is 50.")

    body(doc,
         "Graphically, the deadweight loss of the monopoly scenario compared "
         "with perfect competition is the DWL triangle above. Under monopoly "
         "consumers are worse off (lower consumer surplus), producers are "
         "better off (higher producer surplus), and society overall is worse "
         "off because of the deadweight loss.")

    # ======================================================================
    # Problem 3
    # ======================================================================
    problem(doc, 3, "Reinventing Barnes & Noble", 20)
    draws_on(doc, "Module 5 – Monopolistic Competition")

    p = part(doc, "a", 5)
    run(p, "The market for books has changed since 2000, mostly because of "
           "competition from other companies (Amazon’s Kindle, for example) "
           "and the availability of substitutes. One of the main reasons for "
           "the decline in the fortunes of Barnes & Noble is that all its "
           "bookstores were identical up and down the country and did not "
           "meet local customers’ needs.")

    p = part(doc, "b", 5)
    # PROPOSED 3: name the answer
    run(p, "The best approximation for the demand Barnes & Noble faces is ")
    ins(p, "curve (C), the relatively flat downward-sloping one.")
    run(p, " The market in which Barnes & Noble operates is monopolistically "
           "competitive. Its characteristics are:")

    for t in ("There are many sellers.",
              "Differences can be perceived between products, but they are "
              "close substitutes.",
              "There are no or limited barriers to entry.",
              "Profit can be positive or negative in the short run, but "
              "economic profits are zero in the long run."):
        bullet(doc, t)

    body(doc,
         "It differs from perfect competition, since under perfect "
         "competition there is no product differentiation. It differs from "
         "monopoly, because in a monopoly there is one seller, restricted "
         "entry, and only one product. The market served by Barnes & Noble "
         "is not perfectly competitive because there is some product "
         "differentiation: their stores and the products sold within them "
         "are different. Some bookstores have a café and comfortable chairs; "
         "some specialize in certain genres.", before=6)

    p = part(doc, "c", 5)
    run(p, "The firm has some power to change the price of its product — "
           "more than in a perfectly competitive market, where a firm cannot "
           "change the price at all, but less than a monopolist. The firm "
           "could raise its price, but an increase would result in a sizable "
           "decrease in quantity demanded because of the relatively flat "
           "slope, making demand quite elastic. The magnitude depends on the "
           "slope of the demand curve: the flatter the curve, the larger the "
           "change in quantity demanded in response to a change in price.")

    p = part(doc, "d", 5)
    run(p, "Two ways of offering differentiated products to customers:")
    bullet(doc, "Bookstores could specialize in a few genres of books, "
                "targeted to the interests and tastes of local consumers.")
    bullet(doc, "Bookstores can offer more products and services, so that "
                "customers do not go to the store only to buy books. They "
                "can stock magazines, stationery and entertainment products, "
                "and serve coffee and pastries, so customers are more likely "
                "to spend time there — turning the bookstore into a place "
                "people come to on a day off, not only when they need a "
                "book.")
    body(doc, "Differentiation will make the demand curve less elastic.",
         before=6)

    # ======================================================================
    # Problem 4
    # ======================================================================
    problem(doc, 4, "Avocado Pricing Strategy", 20)
    draws_on(doc, "Module 6 – Complex Pricing (third-degree price "
                  "discrimination)")

    for letter, pts, city, a, b, qstar, pstar in (
            ("a", 5, "Atlanta", "730", "175", "350.13", "2.17"),
            ("b", 5, "San Diego", "855", "319", "400.38", "1.43")):
        p = part(doc, letter, pts)
        if letter == "a":
            run(p, "To find the optimal price we set marginal revenue equal "
                   "to marginal cost. First invert the demand function for ")
        else:
            run(p, "The same steps for ")
        run(p, city + ":")
        T.equation(doc,
                   (QA if letter == "a" else QB) + mrun("=") + mrun(a)
                   + mrun("−") + mrun(b) + (PA if letter == "a" else PB)
                   + mrun("  ⇒  ") + (PA if letter == "a" else PB)
                   + mrun("=") + mfrac(mrun(a) + mrun("−")
                                       + (QA if letter == "a" else QB),
                                       mrun(b)), before=3, after=4)
        T.equation(doc,
                   TR + mrun("=") + mfrac(mrun(a) + (QA if letter == "a"
                                                     else QB), mrun(b))
                   + mrun("−") + mfrac(msup(QA if letter == "a" else QB,
                                            mrun("2", italic=False)),
                                       mrun(b)), before=2, after=3)
        T.equation(doc,
                   (MRA if letter == "a" else MRB) + mrun("=")
                   + mfrac(mrun(a) + mrun("−") + mrun("2")
                           + (QA if letter == "a" else QB), mrun(b)),
                   before=2, after=4)
        body(doc, "Setting marginal revenue equal to the marginal cost "
                  "of $0.17:", after=2)
        T.equation(doc,
                   mfrac(mrun(a) + mrun("−") + mrun("2")
                         + (QA if letter == "a" else QB), mrun(b))
                   + mrun("=") + mrun("0.17") + mrun("  ⇒  ")
                   + (QA if letter == "a" else QB) + mrun("=") + mrun(qstar),
                   before=2, after=4)
        answer(doc, (PA if letter == "a" else PB) + mrun("=")
               + mfrac(mrun(a) + mrun("−") + mrun(qstar), mrun(b))
               + mrun("=") + mrun("$" + pstar))
        S.place(fig_avocado(city), doc, before=6, after=2)
        caption(doc, "Optimal price in {}: MR = MC at Q = {}, priced off "
                     "demand at ${}.".format(city, qstar, pstar))

    p = part(doc, "c", 3)
    run(p, "Using ")
    T.equation_inline(p, ED + mrun("=") + mfrac(mrun("d", italic=False) + Q,
                                                mrun("d", italic=False) + P)
                      + mrun("×") + mfrac(P, Q))
    run(p, ", with the slopes read straight off the demand functions:")
    answer(doc, EDA + mrun("=") + mrun("−175") + mrun("×")
           + mfrac(mrun("2.17"), mrun("350.13")) + mrun("=") + mrun("−1.08"))
    answer(doc, EDB + mrun("=") + mrun("−319") + mrun("×")
           + mfrac(mrun("1.43"), mrun("400.38")) + mrun("=") + mrun("−1.14"),
           before=2)

    p = part(doc, "d", 4)
    run(p, "Yes, they do. We learned that, everything else being equal, we "
           "want to charge higher prices in less elastic markets. Atlanta is "
           "the less elastic market here (|")
    T.equation_inline(p, ED)
    run(p, "| = 1.08 against 1.14 in San Diego), and it is also the market "
           "with the higher optimal price ($2.17 against $1.43).")

    p = part(doc, "e", 3)
    run(p, "This may lead to the problem of ")
    run(p, "parallel imports", bold=True)
    run(p, ". Selling the same item at different prices may result in "
           "traders purchasing at the lower price in San Diego and shipping "
           "the avocados to Atlanta, where the price is higher. That threat "
           "would undercut the firm’s pricing policy and may cause it to "
           "abandon the price discrimination scheme.")

    # ======================================================================
    # Problem 5
    # ======================================================================
    problem(doc, 5, "Working Interactively with AI: The Future of Work", 15)
    draws_on(doc, "Course-wide – marginal analysis applied to tasks; O*NET "
                  "occupational data")

    p = para(doc, before=10, after=5, keep_next=True)
    run(p, "Step 1 (2 points) — Choosing the Occupations", bold=True,
        color=NAVY, size=12)
    body(doc,
         "For this solution we focus on two occupations that are quite "
         "different from each other in terms of daily activities and "
         "required skills: Lawyers (SOC 23-1011.00) and Human Resources "
         "Managers (SOC 11-3121.00).")

    p = para(doc, before=12, after=5, keep_next=True)
    run(p, "Step 2 (5 points) — Listing and Classifying Tasks", bold=True,
        color=NAVY, size=12)
    li, lp = task_table(doc, LAWYERS, "Table 1 – Lawyers")
    hi, hp = task_table(doc, HR, "Table 2 – Human Resources Managers")

    p = para(doc, before=12, after=5, keep_next=True)
    run(p, "Step 3 (8 points) — Computing the AI-Replacement Index",
        bold=True, color=NAVY, size=12)

    T.table(doc, [
        ["Occupation", "Calculation", "Weighted AI-Replacement Index"],
        ["Lawyers", "{:,} ÷ {:,}".format(lp, li), "%.2f" % (lp / float(li))],
        ["Human Resources Managers", "{:,} ÷ {:,}".format(hp, hi),
         "%.2f" % (hp / float(hi))],
    ], widths_in=[2.30, 1.80, 2.40], size=10, align_right=(1, 2))

    body(doc,
         "This result suggests that a majority of core legal tasks — "
         "particularly those involving document drafting, legal research, "
         "precedent analysis and evidence review — could be automated or "
         "significantly accelerated by AI tools such as large-language-model "
         "legal assistants and automated document-search platforms. These "
         "systems can process vast texts and produce legal drafts much "
         "faster than humans.", before=10)

    body(doc,
         "In contrast, HR managers’ work remains more resistant to full "
         "automation. While AI can support recruiting (resume screening, "
         "scheduling interviews, analyzing turnover data), much of the role "
         "requires nuanced human interaction — resolving conflicts, "
         "negotiating agreements, motivating teams and building trust within "
         "organizations — tasks that rely on empathy, judgment and "
         "communication skills that current AI cannot replicate "
         "effectively.")

    body(doc,
         "At first glance it may seem surprising that lawyers score higher "
         "in exposure than HR managers, since legal work is often perceived "
         "as specialized and knowledge-intensive. However, this result "
         "aligns with the idea that AI excels at structured, text-based "
         "cognitive tasks. The legal profession generates enormous "
         "quantities of documents, precedents and standardized forms that "
         "can be processed algorithmically. By contrast, human resources "
         "managers perform tasks that are relational, situational and often "
         "emotional — attributes less easily codified into rules or "
         "patterns. AI may assist HR professionals by automating "
         "administrative work, but the human component of decision-making "
         "remains central.")

    S.save(doc, OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
