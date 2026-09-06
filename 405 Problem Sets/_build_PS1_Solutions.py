"""Build: Problem Set 1 -- Solutions.docx  (MGMT 405, Fall 2026 format).

Source: `_originals/Problem Set 1 -- Solutions (2025 original).docx` -- a
BUILD INPUT, never deleted.  The folder is untracked in git, so that copy
plus the rolling `_t-1` / `_t-2` files are the only way back.

The job is FORMATTING: the original's reasoning, numbers, grading notes
and point breakdowns are preserved.  Everything beyond the restyle is a
real Word revision (`w:ins` / `w:del`, author "Claude (proposed)").

FORMATTING (untracked)
  - Course masthead, navy problem headings with a gold points chip, the
    bare centred page number.
  - Problem 1's cost list becomes two native tables (navy header, cream /
    white banding), with the two totals in the house pale-gold highlight.
    Same components, same numbers, same wording.
  - All seven figures rebuilt as NATIVE, editable Word shapes.  The
    originals were matplotlib screenshots carrying chart junk (gridlines,
    a legend duplicating the curve labels, arrows) and, in the 3(c)
    stack, a stale third panel titled "Generic mobile plans" while the
    text discusses store-brand groceries.  Rebuilt per the course
    conventions: demand dark red `C00000`, supply navy, labels inside the
    plot in boxes measured to the text, no gridlines, no legend.
  - Every marked equilibrium is the COMPUTED intersection of its two
    lines (`_ps_theme.cross`), never eyeballed.
  - The "Note: In the figure above..." lines become figure captions.
  - The answer line of each derivation is set in dark red, the course
    convention for the line that delivers the result.
  - Elasticities written with a true subscript (E(I)), as in the Module 2
    deck.

PROPOSED (tracked -- these go beyond formatting)
  1. Problem 1's figures brought to 2026 (Nico, 2026-09-06), matching the
     problem statement.  Only the two components with a source move:
       * the forgone SALARY, $40,000 -> $46,000.  BLS OEWS May 2025 puts
         the median for Landscaping and Groundskeeping Workers (SOC
         37-3011) in California at $45,560, against $39,150 nationally.
       * GAS, $5.00 -> $5.80 a gallon, so $5,000 -> $5,800.  AAA had
         California at about $5.80 in early September 2026, the highest
         of any state.
     Everything else is unsourced and stays put, so the exercise does not
     acquire numbers that only look precise: the $40,000 truck is still
     right (a 2026 Ford F-150 XL starts at about $40,085), and tools,
     repairs, mileage and permits are assumptions the problem invites the
     student to make.  Totals follow: explicit $16,000 -> $16,800,
     implicit $40,000 -> $46,000, total cost $56,000 -> $62,800, revenue
     $70,000 -> $78,000, economic profit $14,000 -> $15,200.  The margin
     stays narrow, which is what makes the exercise worth doing.
  2. Problem 3(b) was graded out of 8 while the problem statement awards
     it 12, so the three parts summed to 31 rather than the stated 35.
     Re-broken as 2 points per industry plus 6 for the graph = 12.
  3. Problem 3(a), justification list: "Normal goods -- income elasticity
     is between 0 and 1" implies normal and luxury are separate
     categories.  Module 2 (deck slide 55) defines a LUXURY as a normal
     good with E(I) > 1, so luxuries are a subset of normal goods.
     Reworded, with the relationship stated.
  4. Problem 1: the solution stops at "revenues exceed total costs".
     Module 1 (deck slide 74, Flip a House) frames exactly this
     comparison as ECONOMIC PROFIT, so the closing line now names it and
     gives the figure ($15,200, on the 2026 numbers above).
  5. Problem 4: "(L1, w1)" -> "(L0, w0)", and the shifted supply curve
     becomes S' with the new equilibrium (L1, w1).  Module 1 numbers the
     initial equilibrium 0 and the post-shift one 1, and primes the
     shifted curve (deck slides 23, 26, 31-33); the solution used three
     different conventions across its four problems.

Run:  python _build_PS1_Solutions.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _ps_theme as S
from _ps_theme import (DARKRED, GOLD, GRAY, NAVY, Panel, WD_ALIGN_PARAGRAPH,
                       body, caption, cross, dele, ins, line_at, para, part,
                       problem, run)
import _tn_theme as T
from _tn_theme import mrun, msub

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Problem Set 1 -- Solutions.docx")

QD = msub(mrun("Q"), mrun("d", italic=False))
QS = msub(mrun("Q"), mrun("s", italic=False))
P = mrun("P")
Q = mrun("Q")
E_I = msub(mrun("E", italic=False), mrun("I"))


def draws_on(doc, text):
    p = para(doc, before=0, after=8, keep_next=True)
    run(p, "Draws on:  " + text, italic=True, color=GRAY, size=9.5)
    return p


def answer(doc, content, before=4, after=8):
    """A display equation in dark red -- the line that delivers the result."""
    return T.equation(doc, T.color_omml(content, DARKRED),
                      before=before, after=after)


def _ins_elasticity(p, tail):
    """E(I) plus a tail, as tracked insertions.

    Inside a `w:ins` the OMML helpers are awkward, so the symbol is built
    from ordinary runs: upright E, italic subscript I -- the same shape
    the Module 2 deck uses.
    """
    ins(p, "E", size=11)
    ins(p, "I", italic=True, subscript=True, size=11)
    ins(p, tail, size=11)


def _sub(p, base, idx, emit=run, **kw):
    """An indexed symbol -- italic letter, true subscript index."""
    emit(p, base, italic=True, size=11, **kw)
    emit(p, idx, italic=True, subscript=True, size=11, **kw)


def bullet(doc, text=None, level=0, before=2, after=2):
    """A hanging-indent bullet, matching the notes' list treatment."""
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
FIG_W, FIG_H = 5.30, 3.55          # one-panel figure, in inches
PLOT = dict(ox=0.62, oy=3.02, w=4.28, h=2.58)


def _panel(f, ylabel="P", xlabel="Q", **kw):
    d = dict(PLOT)
    d.update(kw)
    return Panel(f, d["ox"], d["oy"], d["w"], d["h"],
                 ylabel=ylabel, xlabel=xlabel)


def fig_p2c():
    """Problem 2(c): one supply curve, demand shifting out with income.

    Drawn to the actual numbers: Qd = 240 - P and Qd = 280 - P against
    Qs = P - 100, on a P axis of 0-300 and a Q axis of 0-250.
    """
    f = S.fig(FIG_W, FIG_H, name="Problem 2(c) - income shifts demand out")
    pn = _panel(f)

    def u(q):
        return q / 250.0 * 100.0

    def v(p):
        return p / 300.0 * 100.0

    # supply Qs = P - 100, from (Q=0, P=100) up to the top of the frame
    s0, s1 = (u(0), v(100)), (u(200), v(300))
    # demand at I = 10 (Qd = 240 - P) and I = 20 (Qd = 280 - P)
    da0, da1 = (u(0), v(240)), (u(240), v(0))
    db0, db1 = (u(30), v(250)), (u(250), v(30))

    pn.supply(s0, s1, label="S", lbl_dx=0.06, lbl_dy=-0.20)
    pn.demand(da0, da1, label=None)
    pn.demand(db0, db1, label=("D", "b"), lbl_dx=0.08, lbl_dy=-0.14)
    # Da is labelled ON its own line, in the clear band below Db, so the two
    # labels do not pile up at the bottom-right corner.
    pn.text(70, line_at(da0, da1, 70) + 6,
            S._label_runs(("D", "a"), dict(bold=True, color=DARKRED, size=11)),
            align="c")

    ea = cross(s0, s1, da0, da1)
    eb = cross(s0, s1, db0, db1)
    pn.equilibrium(ea[0], ea[1], ptick=("P", "a"), qtick=("Q", "a"))
    pn.equilibrium(eb[0], eb[1], ptick=("P", "b"), qtick=("Q", "b"))

    # the shift itself
    pn.arrow((u(150), v(120)), (u(180), v(120)), color=GOLD, w_pt=1.25,
             name="demand shifts out")
    return f


def fig_p3b():
    """Problem 3(b): supply fixed, three demand shifts of different size.

    The four demand curves are parallel, so they are labelled where they
    are furthest apart -- at their quantity intercepts along the bottom,
    12 to 14 logical units (about half an inch) from each other -- while
    the four equilibria sit together on S, up and to the left.  Nothing
    then has to share space with anything else.
    """
    f = S.fig(FIG_W, FIG_H, name="Problem 3(b) - three demand shifts")
    pn = _panel(f)

    s0, s1 = (0, 0), (72, 72)
    curves = [
        ((0, 56), (56, 0), ("D", "L"), "L"),        # luxury: large left shift
        ((0, 70), (70, 0), ("D", "N"), "N"),        # normal: moderate left
        ((0, 82), (82, 0), ("D", "0"), ("E", "0")),  # initial
        ((0, 96), (96, 0), ("D", "I"), "I"),        # inferior: right shift
    ]

    pn.supply(s0, s1, label="S", lbl_dx=0.06, lbl_dy=-0.20)
    for p0, p1, lbl, eq in curves:
        initial = lbl[1] == "0"
        pn.demand(p0, p1, label=None, w_pt=2.25 if initial else 1.5)
        pr = dict(bold=True, color=DARKRED, size=11)
        pn.text(p1[0] + 3.0, 14.0, S._label_runs(lbl, pr), align="c")
        e = cross(s0, s1, p0, p1)
        pn.equilibrium(e[0], e[1], guides=False)
        pr2 = dict(bold=True, italic=True, color=NAVY, size=11)
        pn.text(e[0] + 3.0, e[1] - 6.0, S._label_runs(eq, pr2), align="l")
    return f


D0_3C = ((0, 90), (90, 0))          # the common pre-shock demand curve


def fig_p3c(case):
    """One case of Problem 3(c), at full text width.

    Stacked full-width rather than a three-across row: at a third of the
    page the curve labels, the E0 / E1 pair and the axis titles all had to
    share the same corner, and each panel now sits directly under the
    bullet that describes it.

    case 1  smartphones        big leftward demand shift, supply unchanged
    case 2  restaurants        demand left AND supply left
    case 3  store-brand food   demand right, along an elastic supply curve
    """
    titles = {1: "Case 1 – high-end smartphones",
              2: "Case 2 – mid-price-range restaurants",
              3: "Case 3 – store-brand groceries"}
    f = S.fig(FIG_W, FIG_H - 0.30, name="Problem 3(c) - " + titles[case])
    pn = _panel(f, oy=2.72, h=2.28)
    pn.title(titles[case], size=10, dy=0.24)

    s0, s1 = (0, 0), (76, 76)
    s_shift = None
    if case == 1:
        d1 = ((0, 58), (58, 0))
    elif case == 2:
        d1 = ((0, 70), (70, 0))
        s_shift = ((0, 18), (68, 86))       # capacity cut: supply shifts in
    else:
        d1 = ((8, 100), (100, 8))
        s0, s1 = (0, 28), (96, 62)          # elastic (flat) supply, no shift

    pn.supply(s0, s1, label="S", lbl_dx=0.06, lbl_dy=-0.20)
    e1_supply = (s0, s1)
    if s_shift:
        pn.supply(s_shift[0], s_shift[1], label="S′", lbl_dx=0.06,
                  lbl_dy=-0.20, name="S prime")
        e1_supply = s_shift

    pn.demand(D0_3C[0], D0_3C[1], label=None, w_pt=2.25)
    pn.demand(d1[0], d1[1], label=None, name="D prime")
    pr = dict(bold=True, color=DARKRED, size=11)
    # each demand curve is named just past its own lower-right end
    pn.text(D0_3C[1][0] + 3.5, 14.0, S._label_runs("D", pr), align="c")
    if d1[1][1] > 0:                      # case 3: the end is off the axis
        pn.text(d1[1][0] + 1.0, d1[1][1] + 8.0, S._label_runs("D′", pr),
                align="c")
    else:
        pn.text(d1[1][0] + 3.5, 14.0, S._label_runs("D′", pr), align="c")

    e0 = cross(s0, s1, D0_3C[0], D0_3C[1])
    e1 = cross(e1_supply[0], e1_supply[1], d1[0], d1[1])
    # The two labels take opposite sides, chosen from the direction the
    # equilibrium actually moved, so they never crowd each other.
    right = e1[0] > e0[0]
    pn.equilibrium(e0[0], e0[1], label=("E", "0"),
                   dx=-0.36 if right else 0.07, dy=-0.30, guides=False)
    pn.equilibrium(e1[0], e1[1], label=("E", "1"),
                   dx=0.07 if right else -0.36, dy=-0.30, guides=False)
    return f


def fig_p4(shifted):
    """Problem 4: the market for low-skilled labor, before and after."""
    name = ("Problem 4 - labor market after benefits rise" if shifted
            else "Problem 4 - the market for low-skilled labor")
    f = S.fig(FIG_W, FIG_H, name=name)
    pn = _panel(f, ylabel="w", xlabel="L")

    s0, s1 = (4, 4), (92, 92)
    d0, d1 = (4, 96), (92, 8)
    pn.supply(s0, s1, label="S", lbl_dx=0.06, lbl_dy=-0.20)
    pn.demand(d0, d1, label="D", lbl_dx=0.08, lbl_dy=-0.18)

    e0 = cross(s0, s1, d0, d1)
    if not shifted:
        pn.equilibrium(e0[0], e0[1], ptick=("w", "0"), qtick=("L", "0"))
        return f

    # benefits raise the wage at which people are willing to work: S left
    t0, t1 = (0, 26), (74, 100)
    pn.supply(t0, t1, label="S′", lbl_dx=0.06, lbl_dy=-0.20, name="S prime")
    e1 = cross(t0, t1, d0, d1)
    pn.equilibrium(e0[0], e0[1], ptick=("w", "0"), qtick=("L", "0"))
    pn.equilibrium(e1[0], e1[1], ptick=("w", "1"), qtick=("L", "1"))
    # the shift itself, drawn where neither guide line runs: S sits at
    # v = u and S' at v = u + 26, so this spans the gap between them
    pn.arrow((15, 15), (15, 39), color=GOLD, name="supply shifts in")
    return f


# ==========================================================================
def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)

    S.ps_masthead(
        doc, "Problem Set 1 – Solutions",
        covers="Covers Modules 1 and 2",
        due="100 points")

    # ======================================================================
    # Problem 1
    # ======================================================================
    problem(doc, 1, "Business Choice", 25, before=12)
    draws_on(doc, "Module 1 – Economic Costs Include Opportunity Costs")

    S.note(
        doc,
        "The answer outlined below is one possible way. Your solution likely "
        "involved different numbers, and possibly even a different "
        "recommendation. For grading, we considered whether (i) your "
        "assumptions were broadly realistic and (ii) whether your arguments "
        "were internally consistent and referring to concepts from class, in "
        "particular, explicit and implicit costs.")

    p = para(doc, before=10, after=5, keep_next=True)
    run(p, "Explicit costs", bold=True, color=NAVY, size=11.5)

    # PROPOSED 1: the cost components brought to 2026.  Only the two with a
    # source move -- gas (AAA California, about $5.80 in September 2026) and
    # the forgone salary; the $40,000 truck still matches a 2026 F-150 XL.
    T.table(doc, [
        ["Component", "Assumption", "Per year"],
        ["Truck", "New price $40,000, depreciated over 10 years ($4,000), "
                  "plus $2,000 average repairs", "$6,000"],
        ["Tools", "Gardening tools, lawn mower, blower. Under high usage "
                  "these last about 1 year", "$4,000"],
        ["Gas", "20,000 miles a year at 20 MPG = 1,000 gallons at about "
                "$5.80 (California, 2026)", "$5,800"],
        ["Licenses and permits", "West L.A. may have specific regulations "
                                 "for landscaping businesses", "$1,000"],
        ["Rent (possible)", "Storage space for the equipment. Ignored here, "
                            "assuming he can store it at home", "–"],
        ["Total explicit costs", "", "$16,800"],
    ], widths_in=[1.35, 3.75, 0.85], size=9.5,
        highlight=[(6, 0), (6, 1), (6, 2)], align_right=(2,),
        replaced={
            (3, 1): ("20,000 miles a year at 20 MPG = 1,000 gallons at "
                     "about $5",
                     "20,000 miles a year at 20 MPG = 1,000 gallons at "
                     "about $5.80 (California, 2026)"),
            (3, 2): ("$5,000", "$5,800"),
            (6, 2): ("$16,000", "$16,800"),
        })

    p = para(doc, before=12, after=5, keep_next=True)
    run(p, "Implicit costs (opportunity costs)", bold=True, color=NAVY,
        size=11.5)

    T.table(doc, [
        ["Component", "Assumption", "Per year"],
        ["Owner’s wage", "By not working for the landscaping company, the "
                         "gardener forgoes a sure salary of $46,000",
         "$46,000"],
        ["Forgone benefits", "Health insurance, retirement benefits or paid "
                             "leave from the previous employer. Possible, "
                             "but not required for full credit", "–"],
        ["Total implicit costs", "", "$46,000"],
    ], widths_in=[1.35, 3.75, 0.85], size=9.5,
        highlight=[(3, 0), (3, 1), (3, 2)], align_right=(2,),
        replaced={
            (1, 1): ("By not working for the landscaping company, the "
                     "gardener forgoes a sure salary of $40,000",
                     "By not working for the landscaping company, the "
                     "gardener forgoes a sure salary of $46,000"),
            (1, 2): ("$40,000", "$46,000"),
            (3, 2): ("$40,000", "$46,000"),
        })

    p = body(doc, before=12)
    run(p, "Total costs = explicit + implicit costs = ")
    dele(p, "$16,000 + $40,000 = $56,000")
    ins(p, "$16,800 + $46,000 = $62,800")
    run(p, ". Expected revenues are ")
    dele(p, "$70,000")
    ins(p, "$78,000")
    run(p, ", which is above the total costs of ")
    dele(p, "$56,000")
    ins(p, "$62,800")
    run(p, ". ")
    # PROPOSED 3: name the comparison the way Module 1 does.
    ins(p, "In the language of Module 1, the economic profit of the business "
           "is $78,000 − $62,800 = $15,200, and it is positive. ")
    run(p, "Thus, under the above assumptions, you would suggest that your "
           "gardener opens his own business. Of course, this recommendation "
           "may vary depending on the assumptions that you have made. You "
           "will receive full credit if all assumptions were reasonable and "
           "you distinguished correctly between explicit and implicit costs.")

    # ======================================================================
    # Problem 2
    # ======================================================================
    problem(doc, 2, "Demand Curve", 15)
    draws_on(doc, "Module 1 – Demand and Supply; Equilibrium")

    p = part(doc, "a", 4)
    run(p, "We know that in equilibrium, ")
    T.equation_inline(p, QD + mrun("=") + QS)
    run(p, ". Substituting ")
    run(p, "I", italic=True)
    run(p, " = 10 gives ")
    T.equation_inline(p, QD + mrun("=") + mrun("240") + mrun("−") + P)
    run(p, " and ")
    T.equation_inline(p, QS + mrun("=") + P + mrun("−") + mrun("100"))
    run(p, ". Setting them equal yields 2")
    run(p, "P", italic=True)
    run(p, " = 340, and therefore:")

    answer(doc, P + mrun("=") + mrun("170") + mrun(",  ") + Q + mrun("=")
           + mrun("70"))

    S.note(
        doc,
        [("Once you obtain the equilibrium price, the equilibrium quantity "
          "can be found by substituting ", {}),
         ("P", dict(italic=True)),
         (" in either the demand or the supply equation. To be sure that you "
          "did not get the algebra wrong, substitute ", {}),
         ("P", dict(italic=True)),
         (" in both equations and check that the resulting quantity is in "
          "fact the same – that is, that quantity demanded equals quantity "
          "supplied.", {})],
        prefix="Suggestion to avoid algebraic mistakes:")

    p = part(doc, "b", 4)
    run(p, "Now ")
    run(p, "I", italic=True)
    run(p, " = 20 (higher income), so ")
    T.equation_inline(p, QD + mrun("=") + mrun("280") + mrun("−") + P)
    run(p, ". Using the same steps as in (a):")

    answer(doc, P + mrun("=") + mrun("190") + mrun(",  ") + Q + mrun("=")
           + mrun("90"))

    p = part(doc, "c", 7)
    run(p, "The graph should look like this, with subscripts referring to "
           "parts (a) and (b). Increasing income has shifted the demand "
           "curve outward, so both the equilibrium price and the "
           "equilibrium quantity rise.")

    fig_p2c().place(doc, before=8, after=2)
    caption(doc, "Higher income shifts the demand curve out, and the market "
                 "clears at both a higher price and a larger quantity.")

    # ======================================================================
    # Problem 3
    # ======================================================================
    problem(doc, 3,
            "Income Elasticity and Demand During Economic Uncertainty", 35)
    draws_on(doc, "Module 2 – Elasticities (income elasticity);  "
                  "Module 1 – Equilibrium")

    p = part(doc, "a", 12)
    run(p, "(6 points for the initial ranking, 2 points for each of the "
           "three shifts explained.)  Ranking by decline in quantity "
           "demanded, from largest to smallest:")

    bullet(doc, "Luxury goods and services (sports cars, jewelry, luxury "
                "travel, high-end smartphones, or other high-end "
                "non-essential goods). Below we focus on high-end "
                "smartphones as a concrete example.")
    bullet(doc, "Normal non-essential goods (regular restaurant visits, "
                "gyms, streaming subscriptions, mid-range clothing). Below "
                "we focus on mid-price-range restaurants.")
    bullet(doc, "Inferior goods (cheaper versions of essential goods: "
                "store-brand groceries, switching from a car to public "
                "transportation). Below we focus on store-brand groceries.")

    p = body(doc, before=8, after=4)
    run(p, "Justification:", bold=True, color=NAVY)

    p = bullet(doc)
    run(p, "Luxury goods", bold=True)
    run(p, " – income elasticity ")
    T.equation_inline(p, E_I + mrun(">") + mrun("1"))
    run(p, ". In crises, households cut or postpone expensive, non-essential "
           "purchases. Demand shifts left by a lot, so quantity falls the "
           "most.")

    # PROPOSED 2: luxuries are a SUBSET of normal goods (Module 2, slide 55)
    p = bullet(doc)
    dele(p, "Normal goods", bold=True)
    ins(p, "Normal, non-luxury goods", bold=True)
    run(p, " – income elasticity ")
    T.equation_inline(p, mrun("0") + mrun("<") + E_I + mrun("<") + mrun("1"))
    run(p, ". Non-essential goods that are part of routine consumption. "
           "People trim frequency and quality (cook at home, DIY grooming). "
           "Demand shifts left, but less than for luxury goods.")
    ins(p, " Note that a luxury good is itself a normal good: Module 2 "
           "defines any good with ")
    _ins_elasticity(p, " > 0")
    ins(p, " as normal, and a luxury as a normal good with ")
    _ins_elasticity(p, " > 1")
    ins(p, ".")

    p = bullet(doc)
    run(p, "Inferior goods", bold=True)
    run(p, " – income elasticity ")
    T.equation_inline(p, E_I + mrun("<") + mrun("0"))
    run(p, ". When budgets tighten, consumers purchase more affordable "
           "versions (these may be cheaper versions of essential goods, such "
           "as store-brand toilet paper). For inferior goods, demand shifts "
           "right during a crisis.")

    p = body(doc, before=8)
    run(p, "Link to the demand curves (supply unchanged): ")
    run(p, "the crisis moves demand, not supply. Luxury goods shift left the "
           "most; normal goods shift left moderately; inferior goods shift "
           "right.")

    # -- (b) -- PROPOSED 1: 8 points -> the 12 the problem awards -----------
    p = part(doc, "b", 12)
    dele(p, "(2 points for each correct price/quantity change; 2 points for "
            "a graph with the correct display of supply and demand curves.)")
    ins(p, "(2 points for each correct price/quantity change; 6 points for a "
           "graph with the correct display of supply and demand curves.)")
    run(p, "  With supply held constant, every new equilibrium lies on the "
           "unchanged supply curve. Luxury goods see the largest fall in "
           "both price and quantity, normal goods a moderate fall, and "
           "inferior goods a rise in both.")

    fig_p3b().place(doc, before=8, after=2)
    caption(doc, "E₀ is the pre-shock equilibrium; L, N and I are the "
                 "post-shock equilibria for luxury, normal and inferior "
                 "goods. Supply is unchanged, so all four lie on S.")

    # -- (c) ---------------------------------------------------------------
    p = part(doc, "c", 11)
    run(p, "(3 points for each correct reasoning and supply shift; 2 points "
           "for the correct new ranking.)")

    p = bullet(doc, before=6)
    run(p, "High-end smartphones", bold=True)
    run(p, " (luxury durable). Demand shock: big leftward shift. Supply "
           "(about 1 month): roughly unchanged – production takes time and "
           "decisions are set in advance, and inventory is already built. "
           "Net short-run price effect: ")
    run(p, "the largest price drop", bold=True)
    run(p, ", with quantity falling too.")

    fig_p3c(1).place(doc, before=6, after=2)
    caption(doc, "Demand shifts far to the left along an unchanged supply "
                 "curve, so the price falls a long way.")

    p = bullet(doc)
    run(p, "Mid-price-range restaurants", bold=True)
    run(p, " (normal services). Demand shock: leftward shift, as households "
           "trim dining out. Supply (about 1 month): leftward shift – "
           "restaurants cut hours, close some days, reduce freelance staff "
           "and buy less food, so capacity can be reduced quickly. Net "
           "short-run price effect: ")
    run(p, "a small price drop", bold=True)
    run(p, " (demand falls but supply falls too), with a larger quantity "
           "drop than in the first case.")

    fig_p3c(2).place(doc, before=6, after=2)
    caption(doc, "Both curves shift in. The two price effects work against "
                 "each other, so the price barely moves, while the two "
                 "quantity effects reinforce each other.")

    p = bullet(doc)
    run(p, "Store-brand groceries", bold=True)
    run(p, " (essential good). Demand shock: rightward shift, as households "
           "trade down toward cheaper staples. Supply (about 1 month): "
           "moderately elastic – inventories can be drawn down and extra "
           "shifts worked, so supply expands somewhat, though production and "
           "input constraints still bind. Net short-run price effect: ")
    run(p, "a small price increase", bold=True)
    run(p, ", with quantity rising.")

    fig_p3c(3).place(doc, before=6, after=2)
    caption(doc, "Supply does not shift, but it is flat enough (elastic) "
                 "that the extra demand is met mostly by more quantity and "
                 "only a little by a higher price.")

    # -- (d) ---------------------------------------------------------------
    p = part(doc, "d", 5, bonus=True)
    run(p, "(2 points for an explanation that mentions the ambiguity, "
           "1 point for each answer that is correct given the previous "
           "assumptions.)  In general we cannot make a clear prediction for "
           "prices. Once both demand and supply can shift, the price change "
           "is ambiguous: it depends on the direction and the size of each "
           "shift, and on the elasticities of the curves.")

    # ======================================================================
    # Problem 4
    # ======================================================================
    problem(doc, 4, "Consulting Project", 25)
    draws_on(doc, "Module 1 – Equilibrium;  Economic Costs Include "
                  "Opportunity Costs")

    p = part(doc, "a", 8)
    run(p, "Starting point. For simplicity we use “Labor” to refer to "
           "low-skilled labor. The equilibrium wage is where the supply of "
           "and the demand for labor intersect, at ")
    # PROPOSED 4: notation aligned with Module 1 (initial = 0, shifted = ')
    run(p, "the point ")
    dele(p, "(L1, w1)")
    ins(p, "(")
    _sub(p, "L", "0", emit=ins)
    ins(p, ", ")
    _sub(p, "w", "0", emit=ins)
    ins(p, ")")
    run(p, ".")

    fig_p4(False).place(doc, before=8, after=2)
    caption(doc, "The market for low-skilled labor. Wages on the vertical "
                 "axis, low-skill labor on the horizontal axis.")

    p = part(doc, "b", 12)
    run(p, "Effect of unemployment benefits. If unemployment benefits "
           "increase significantly, the (low-skilled) labor supply curve "
           "shifts up, to the left. The reason is that the wage at which "
           "people are willing to work is now higher, because they have a "
           "more attractive “outside option” (receiving benefits while "
           "staying at home). For any given wage, fewer individuals would be "
           "willing to work than in the original scenario.")

    p = body(doc)
    run(p, "Another way to look at this is that an increase in unemployment "
           "benefits raises the ")
    run(p, "opportunity cost of employment", bold=True)
    run(p, ": compared to the original scenario, the decision to work now "
           "implies giving up the additional benefits. Some individuals who "
           "were almost indifferent between working and not working, and "
           "leaned towards working, will change their decision and stay at "
           "home.")

    fig_p4(True).place(doc, before=8, after=2)
    cap = dict(italic=True, color=GRAY, size=9)
    p = para(doc, before=2, after=10)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, "Higher benefits shift labor supply in, from S to S′. The "
           "equilibrium wage rises and equilibrium employment falls.", **cap)

    p = part(doc, "c", 5)
    run(p, "The demand side. We do not have any indication that the labor "
           "demand curve will shift, at least not in the short run. You "
           "received full credit if you stated that labor demand remained "
           "unchanged. There are also other, more advanced possible answers: "
           "in the long run, for example, employers may increase automation, "
           "reducing their labor demand, which would shift the labor demand "
           "curve down and reduce the equilibrium wage. Answers like this "
           "also received full credit.")

    doc.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
