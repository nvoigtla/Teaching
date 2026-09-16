"""Build: Problem Set 3.docx  (MGMT 405, Fall 2026 format).

Source: `_originals/Problem Set 3 (2025 original).docx` -- a BUILD INPUT,
never deleted.

The job is FORMATTING: the original's problem titles, wording, point
allocations and examples are preserved.  Everything beyond the restyle is
a real Word revision (`w:ins` / `w:del`, author "Claude (proposed)").

CHECKED AGAINST `Module 3 - Revised.pptx` and `Module 4 - Revised.pptx`.
  - Problem 1 (marginal product, diminishing returns) is Module 3
    Video 2.  TWO separate deck conventions apply and they are easy to
    conflate:
      * HOW the change is computed -- slide 46, "Convention: use the
        next-higher input level (D goes forward)", i.e. a forward
        difference.  Unchanged.
      * WHERE it is PLOTTED -- at the MIDPOINT of its interval as of
        2026-09-13 (Nico).  It used to be the initial point; the midpoint
        makes the figure say what the staggered table says, that the value
        belongs to the STEP and not to either level on its own.
      * WHERE it is written in a table -- slide 19's MPL table leaves the
        DQ / DL / MPL columns empty and floats each value on the border
        between the two rows it comes from.  Adopted here (2026-09-09,
        Nico) as a thin row of its own between every pair of level rows;
        see `S.change_table`.
    The old "enter it at the initial point" instruction confused the two,
    so its table half is a tracked deletion and its plotting half is
    restated.
  - Problem 2 is the Bang-for-the-Buck rule, Module 3 slide 43.  That
    slide writes the price of capital as pK, so the solutions now do
    too (the original used r).
  - Problem 3 is Module 3's cost concepts plus Module 4's short-run
    produce / stop rule (slide 32: "Continue to produce if P >= AVC,
    Stop production if P < AVC").
  - Problem 4 is Module 4 Video 3, profit maximisation of a price taker.
  - Problem 5 is Module 4 in class: slide 69 is the sales-tax diagram and
    carries a "Problem Set 3" pointer; slide 70 is tax incidence; slide
    87 is the Pigouvian tax set equal to the external marginal cost.
  - Point totals: 15 + 15 + 25 + 30 + 25 = 110.  See PROPOSED 1.

FORMATTING (untracked)
  - Course masthead, navy problem headings with a gold points chip, the
    bare centred page number.
  - Problem 1's Excel screenshot of the production function becomes a
    native chart, plotted from the same nine observations the solutions
    tabulate (0, 500, 1100, 1700, 2200, 2500, 2600, 2500, 2000).  Light
    dashed gridlines are kept on purpose here: the student has to read
    values off this chart to answer part (a).
  - Problem 3's cost table becomes a native table with the same 17 rows
    and 11 columns, left blank for the student to fill.
  - Problem 4's two cost functions set as native OMML.
  - Problem 5's parts read "(10 points)" rather than "(10)", matching
    every other problem in the set.
  - The PBS article keeps its hyperlink.

PROPOSED (tracked -- this goes beyond formatting)
  1. The parts add up to 110, not the 100 the header claims:
     15 + 15 + 25 + 30 + 25.  Problem 4 is the outlier at 30 and is the
     only problem whose own parts (8 + 3 + 8 + 8 + 3) also sum to 30, so
     nothing inside it is mis-stated -- the set is simply over-weighted.
     Flagged in the header rather than silently rescaled, because which
     problem should give up the 10 points is Nico's call.

Run:  python _build_PS3.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _ps_theme as S
from _ps_theme import (GRAY, NAVY, Panel, WD_ALIGN_PARAGRAPH, body, dele,
                       ins, para, part, problem, run)
import _tn_theme as T
from _tn_theme import mrun, msub, msup

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Problem Set 3.docx")

PBS = ("https://www.pbs.org/newshour/economy/making-sense/"
       "how-taxing-sugary-drinks-affects-a-communitys-health-and-economy")

Q = mrun("Q")
TC_O = msub(T.acr("TC"), mrun("O"))
TC_N = msub(T.acr("TC"), mrun("N"))

# The tulip production function: tons of fertilizer -> tulips per month.
# Same nine observations the solutions tabulate.
TULIPS = [(0, 0), (1, 500), (2, 1100), (3, 1700), (4, 2200),
          (5, 2500), (6, 2600), (7, 2500), (8, 2000)]

# Problem 3's cost data.  Only Output and Total Cost are given.
BATS = [(0, 25600), (1000, 31360), (2000, 35920), (3000, 39520),
        (4000, 42400), (5000, 44800), (6000, 46960), (7000, 49120),
        (8000, 51520), (9000, 54400), (10000, 58000), (11000, 62560),
        (12000, 68320), (13000, 75520), (14000, 84400), (15000, 95200),
        (16000, 108160)]


def draws_on(doc, text):
    # after 8 -> 5 on 2026-09-12, part of returning ~30 pt of pure
    # spacing so a page break is never decided by a few points.
    p = para(doc, before=0, after=5, keep_next=True)
    run(p, "Draws on:  " + text, italic=True, color=GRAY, size=9.5)
    return p


# ==========================================================================
def fig_tulips():
    """The tulip production function, drawn from the nine observations.

    Gridlines are kept -- lightly, dashed, gray -- because part (a) asks
    the student to read the output level at each fertilizer quantity off
    this chart and difference it.  That is the one case the deck's
    "no gridlines" rule makes an exception for.
    """
    f = S.fig(5.90, 3.95, name="Problem 1 - tulip production function")
    pn = Panel(f, 0.86, 3.30, 4.35, 2.85, ylabel="Tulips",
               xlabel="Fertilizer")

    # Both axes run past the last labelled tick.  The axis titles are
    # anchored to the arrow tips, so a tick sitting AT the end of an axis
    # lands on top of its title -- the "8" ran into "Fertilizer".
    def u(tons):
        return tons / 9.0 * 100.0

    def v(n):
        return n / 3500.0 * 100.0

    # gridlines first, so every curve and marker sits on top of them
    for n in range(500, 3001, 500):
        pn.f.line(pn.x(0), pn.y(v(n)), pn.x(100), pn.y(v(n)),
                  color=S.LIGHT, w_pt=0.75, dash="dash", name="gridline")
        pn.tick_y(v(n), "%d" % n, size=9.5, italic=False)
    for tons in range(0, 9):
        pn.tick_x(u(tons), "%d" % tons, size=9.5, italic=False)

    pn.polyline([(u(t), v(n)) for t, n in TULIPS], color=NAVY, w_pt=2.0,
                name="total product")
    for t, n in TULIPS:
        pn.f.dot(pn.x(u(t)), pn.y(v(n)), d=0.070, color=NAVY, name="obs")
    return f


# The CHANGE columns sit at the far right, matching the solutions
# (2026-09-13, Nico).  The blank table a student fills in has to have the
# same shape as the key s/he checks it against.
BATS_HEAD = ["Output", "Total Cost", "TFC", "TVC", "AFC", "AVC", "ATC",
             "TR", "Profit", "MC", "MR"]

# Widths BY NAME, so reordering the header reorders the columns and
# nothing else has to change.
BATS_WIDTH = {"Output": 0.52, "Total Cost": 0.56, "TFC": 0.44, "TVC": 0.46,
              "AFC": 0.40, "AVC": 0.42, "ATC": 0.42, "TR": 0.52,
              "Profit": 0.50, "MC": 0.40, "MR": 0.40}


# Same widths as the solutions' filled tables, so the student's blank
# and the answer key line up column for column.
BATS_W = [BATS_WIDTH[h] for h in BATS_HEAD]


def bats_table(doc, size=8):
    """The blank table for the student to fill in.

    MC and MR are the change-based series, so their cells are MERGED
    across the boundary between two output levels -- the convention the
    slides use for DQ / DL / MPL.  Unlike the old half-row offset, the
    stagger is visible even when the cells are empty: the student can see
    that an MC box straddles two output rows, which is the whole point.
    The Convention note under the question says it in words as well.
    """
    # BY NAME, then emitted in BATS_HEAD order -- the order is carried by
    # the header list alone, so it cannot drift from the solutions again.
    given = {"Output": None, "Total Cost": None}
    rows = []
    for q, tc in BATS:
        given["Output"] = "{:,}".format(q)
        given["Total Cost"] = "{:,}".format(tc)
        rows.append([given.get(h, "") for h in BATS_HEAD])
    return S.stagger_table(doc, BATS_HEAD, rows,
                           (BATS_HEAD.index("MC"), BATS_HEAD.index("MR")),
                           widths_in=BATS_W, size=size, change_fade=True)


# ==========================================================================
def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)

    # PROPOSED 1: the parts sum to 110, not 100.
    S.ps_masthead(
        doc, "Problem Set 3",
        covers="Covers Modules 3 and 4",
        # The Class Website, not BruinLearn -- Nico's wording, made
        # permanent on 2026-09-08.  No points total on this line: the five
        # problems sum to 110, not 100, so stating a total here would need
        # his decision first.
        due="Due date: see the Course Calendar on the Class Website")
    p = para(doc, before=0, after=10)
    run(p, "100 points", italic=True, color=GRAY, size=10.5)
    ins(p, "     ·     [The parts below currently add up to 110 points: "
           "15 + 15 + 25 + 30 + 25.]", italic=True, color=GRAY, size=10.5)

    # -- AI policy, in Nico's wording (2026-09-08) -------------------------
    # -- Submission instruction, above the policy card (2026-09-12) -------
    # No hyperlink on "BruinLearn": the same PS goes to both sections.
    S.submit_card(doc)

    # The same cream card every problem set carries, from `S.policy_card`,
    # replacing the plain `S.note` this file used to build from the older
    # 2025 wording.  His rewrite drops the closing "ask specific questions
    # ... the answers are not always correct" advice and puts the screening
    # and the exams in its place.
    S.policy_card(doc)

    # ======================================================================
    # Problem 1
    # ======================================================================
    problem(doc, 1, "Tulip Production", 15, before=14)
    draws_on(doc, "Module 3 – The Production Function (marginal product, "
                  "diminishing returns)")

    body(doc,
         "A firm uses fertilizer, labor, and greenhouses as inputs to produce "
         "tulips. When the quantity of labor and greenhouses is held fixed, "
         "the relationship between the quantity of fertilizer and the number "
         "of tulips produced is given by the following graph:")

    S.place(fig_tulips(), doc, before=6, after=2)
    S.caption(doc, "Tulips produced per month against tons of fertilizer "
                   "per month, holding labor and greenhouses fixed.")

    p = part(doc, "a", 10)
    run(p, "Compute the marginal product of fertilizer for the different "
           "fertilizer quantities and plot the marginal product of "
           "fertilizer as a function of the quantity of fertilizer. ")
    # The old sentence made the initial point the rule for BOTH the table
    # and the plot.  The table now follows the between-rows convention (see
    # the note below), so only the plotting half survives.
    dele(p, "Remember to use the initial point as reference (e.g., the "
            "marginal product of moving from 3 to 4 tons should be entered "
            "at the initial point of 3 tons).")
    # MID-POINT, not the initial point (2026-09-13, Nico).  The solutions'
    # figure plots it that way, so a student following this sentence has to
    # land on the same picture.
    ins(p, "When you plot a marginal product, plot it at the mid-point of "
           "its interval (e.g. the marginal product of moving from 3 to 4 "
           "tons is plotted at 3.5 tons).")
    S.change_note(doc)

    p = part(doc, "b", 3)
    run(p, "Does the production function exhibit diminishing marginal "
           "returns to fertilizer? If so, over which range of fertilizer "
           "quantity do they occur?")

    p = part(doc, "c", 2)
    run(p, "Does total output always increase as more fertilizer is added?")

    # ======================================================================
    # Problem 2
    # ======================================================================
    problem(doc, 2, "Sludge Production", 15)
    draws_on(doc, "Module 3 – Long Run: The Optimal Input Mix "
                  "(bang-for-the-buck rule)")

    p = body(doc)
    run(p, "A bottling company uses two inputs to produce bottles of the "
           "soft drink Sludge: bottling machines (")
    run(p, "K", italic=True)
    run(p, ") and workers (")
    run(p, "L", italic=True)
    run(p, "). The machine costs $1,000 per day to operate and the workers "
           "earn $200 per day. At the current level of production, the "
           "marginal product of the machines is an additional 200 bottles "
           "per day and the marginal product of labor is 50 more bottles "
           "per day.")

    p = part(doc, "a", 8)
    run(p, "Is this firm producing its output in an optimized way? If so, "
           "say why. If not, explain in which direction the firm should "
           "change its quantity of machines and workers to optimize its "
           "production.")

    p = part(doc, "b", 7)
    run(p, "How does this relate to the idea of “bang for the buck” that we "
           "saw in class? Briefly illustrate in a short paragraph this "
           "concept with an example from your own experience/industry (no "
           "need to use actual numbers or real facts, you can use "
           "hypothetical scenarios if you wish).")

    # ======================================================================
    # Problem 3
    # ======================================================================
    problem(doc, 3, "Baseball Bats", 25)
    draws_on(doc, "Module 3 – Cost Concepts;  Module 4 – Profit "
                  "Maximization of a Price Taker in the Short Run")

    # keep_next so the heading, this paragraph and the table travel to the
    # next page together -- the table is 5.3" tall and cannot follow them
    # on a part-used page, which otherwise strands the heading alone.
    body(doc,
         "The cost data for a firm producing baseball bats appear below. Of "
         "the total cost, $25,600 is fixed at all levels of output. The bats "
         "sell for $6.40 each (we assume that the firm is a price taker).",
         keep_next=True)

    bats_table(doc)

    p = part(doc, "a", 10, before=10)
    run(p, "Fill out the above table, by entering the values of the "
           "following series: TFC, TVC, TC, AFC, AVC, ATC, MC, TR, MR, and "
           "profit. ")
    dele(p, "Remember to use the initial point as reference (e.g., the "
            "marginal cost of moving from 0 to 1,000 should be entered at "
            "the initial point of 0).")
    run(p, "You may copy the table to Excel or "
           "similar software and compute all values there.")
    S.change_note(doc)

    p = part(doc, "b", 5)
    run(p, "Assuming that production can only be adjusted in steps of 1,000, "
           "what is the optimal output and what is the level of profits at "
           "this output level? If you could instead adjust quantity in "
           "smaller steps, would you produce (slightly) more or less in "
           "order to further raise profits (based on the comparison of MR "
           "and MC at the optimal output level you just reported)?")

    p = part(doc, "c", 10)
    run(p, "Assume that the market price drops to $5.50. What is the new "
           "optimal level of output? Does the firm make a loss or a profit? "
           "Should the firm continue to produce in the short run at this new "
           "price? Continue to assume that production can only be adjusted "
           "in steps of 1,000.")

    # ======================================================================
    # Problem 4
    # ======================================================================
    problem(doc, 4, "Basketball Production", 30)
    draws_on(doc, "Module 4 – Profit Maximization of a Price Taker in the "
                  "Short Run")

    p = body(doc, after=2)
    run(p, "Balding is a basketball producing firm that has seen better "
           "days. Its technology is old and expensive. Total costs of the "
           "old basketball-producing technology are given by")
    T.equation(doc,
               TC_O + mrun("=") + mrun("50") + mrun("+") + mrun("2") + Q
               + mrun("+") + mrun("0.5") + msup(Q, mrun("2", italic=False)),
               before=2, after=4)
    p = body(doc)
    run(p, "where ")
    run(p, "Q", italic=True)
    run(p, " is balls produced per hour. The Balding managers are thinking "
           "about investment in a new technology. Total costs would then be "
           "given by:")
    T.equation(doc,
               TC_N + mrun("=") + mrun("100") + mrun("+") + mrun("0.1")
               + msup(Q, mrun("2", italic=False)), before=2, after=4)
    body(doc,
         "Balding is producing for the low-quality end of the market, and "
         "its basketballs are undifferentiated. This market is perfectly "
         "competitive and the typical market price for a ball is $10.")

    p = part(doc, "a", 8)
    run(p, "Under the old technology, what is the profit-maximizing output "
           "of Balding? What are profits at this output? Please explain how "
           "you obtained the profit-maximizing output and why following "
           "those steps makes sense. Here we want to know if you understand "
           "the concepts you apply, so avoid just solving a formula.")

    p = part(doc, "b", 3)
    run(p, "Illustrate your answer to (a) graphically. Show the point at "
           "which profit is maximized under the old technology, and the area "
           "representing the profits at the profit-maximizing output. You do "
           "not need to plot the exact functions; an illustration is fine. "
           "Just make sure to include all the relevant curves and to get "
           "their shape right.")
    run(p, "  You may use AI tools to produce graphs.", italic=True)

    p = part(doc, "c", 8)
    run(p, "If Balding cannot avoid paying the fixed cost of the old "
           "technology, should it continue to produce basketballs?")

    p = part(doc, "d", 8)
    run(p, "If Balding operated only the new technology, what would be the "
           "optimal output, and what would be its profits?")

    p = part(doc, "e", 3)
    run(p, "Illustrate your answer to (d) graphically. Show the point at "
           "which profit is maximized under the new technology, and the area "
           "representing the profits at the profit-maximizing output. You do "
           "not need to plot the exact functions; an illustration is fine. "
           "Just make sure to include all the relevant curves and to get "
           "their shape right.")
    run(p, "  You may use AI tools to produce graphs.", italic=True)

    # ======================================================================
    # Problem 5
    # ======================================================================
    problem(doc, 5, "Taxing Sugary Drinks", 25)
    draws_on(doc, "Module 4 – Market Distortions and Regulations "
                  "(taxes, tax incidence);  Externalities")

    body(doc,
         "You are working in the campaign team of the governor of "
         "California. The governor would like to implement a tax on sugary "
         "drinks. You are asked the following:")

    p = part(doc, "a", 10)
    run(p, "What should be the amount of the tax (per ounce)? Try your best "
           "to make a recommendation. There is no right or wrong answer. We "
           "want to evaluate your reasoning based on the economic concepts "
           "learned in class. If you want more background on taxing sugary "
           "drinks you can have a look at ")
    S.link(p, "this article", PBS)
    run(p, ". You are encouraged to use your own experience/knowledge or use "
           "other references to (online) sources.")

    p = part(doc, "b", 8)
    run(p, "Assume that the tax is levied on producers. How will the tax "
           "affect the equilibrium price of sugary drinks on the market? "
           "Illustrate your answer using a supply demand graph.")
    run(p, "  You may use AI tools to produce graphs.", italic=True)

    p = part(doc, "c", 7)
    run(p, "Who is more likely to bear the burden of the tax (and have the "
           "highest probability to vote for the governor’s opponent in the "
           "next elections)? How would your answer change if the tax is "
           "levied on consumers instead?")

    S.save(doc, OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
