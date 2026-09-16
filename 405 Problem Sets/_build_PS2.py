"""Build: Problem Set 2.docx  (MGMT 405, Fall 2026 format).

Source: `_originals/Problem Set 2 (2025 original).docx` -- a BUILD INPUT,
never deleted.

The job is FORMATTING: the original's problem titles, wording, point
allocations and examples are preserved, including its run-level emphasis
(the italic "exact" in 1(b), the italic AI-tools sentences, the italic
"holding everything else constant" and "estimate", the bold LAX).
Everything beyond the restyle is a real Word revision (`w:ins` / `w:del`,
author "Claude (proposed)").

CHECKED AGAINST `Module 2 - Revised.pptx`.  PS 2 is entirely Module 2.
  - The three estimation methods asked for in Problem 2 are exactly the
    three on deck slide 96: transaction data, surveys, market
    experimentation (A/B testing, slides 97-99).
  - Problem 5's regression mirrors deck slides 101-106, which close on
    "Q = 478.95 - 1.64 P ... What Is ED at P = 140? What Is MR at
    P = 140?" with a Problem Set 2 pointer.
  - Deck slide 101 says "a = intercept with the y-axis (Q on y-axis)",
    so in the ESTIMATION context this course puts Q on the vertical
    axis.  Problem 4(c)'s figure in the solutions therefore keeps the
    original's orientation (P horizontal, Q vertical); it is the class
    convention, not a slip.
  - Point totals check out: 20 + 20 + 15 + 20 + 25 = 100, and every
    problem's parts sum to its own header.

FORMATTING (untracked)
  - Course masthead, navy problem headings with a gold points chip, the
    bare centred page number (no year or term, so the file is reusable).
  - The AI-tools policy moves into the house cream callout.  The
    original set it in concept blue `0070C0`, which the course reserves
    for naming a concept, so it is navy here as in Problem Set 1.
  - The Wall Street Journal sentence on Dollar General becomes a cream
    quote callout with its attribution.
  - Problem 2's nested numbered list becomes (a) / (b) with (i) / (ii)
    sub-parts, so it is visible which two blocks carry the 10 points.
  - Problem 3's nine-row two-column data block becomes a native
    four-column table (measure / before / during / change).  Same
    numbers; they are internally consistent (246,555 x $86.50 =
    $21,327,008 and 1,053,139 x $44.69 = $47,064,782, and the two
    differences are as stated).
  - Problem 4's demand function and Problem 5's regression set as native
    OMML with true subscripts; the regression coefficients become a
    native table, matching how deck slide 105 presents its own
    regression results.

TYPOS FIXED DIRECTLY (unambiguous; reported rather than tracked)
  - Problem 4: "Boeing's new aircraft AS the following demand" -> "has".
    The solutions document already reads "has".
  - Problem 3's table: "Market revenue IN DURING price war" -> "during".

PROPOSED (tracked -- these go beyond formatting)
  1. Problems 1(a) and 3(a) do not say which point the percentage
     changes are measured from, and the answer depends on it.  In
     Problem 3 it depends enormously: from the initial point the
     elasticity is -6.77, from the final point -0.82.  Deck slide 38
     carries this as a CONVENTION box -- "always relative to the initial
     point (P0, Q0)" -- so a one-line reminder is added to each, in the
     deck's own words.
  2. Problem 5: "The analyst ran a price experiments" -> "ran a price
     experiment".  The rest of the paragraph describes a single
     experiment ("She then assembled data ...").

Run:  python _build_PS2.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _ps_theme as S
from _ps_theme import (GRAY, NAVY, WD_ALIGN_PARAGRAPH, body, dele, ins,
                       para, part, problem, run)
import _tn_theme as T
from _tn_theme import mrun, msub

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Problem Set 2.docx")

# -- OMML symbols ----------------------------------------------------------
Q = mrun("Q")
P = mrun("P")
PX = msub(mrun("P"), mrun("X"))
Y = mrun("Y")
B = [msub(mrun("b"), mrun(str(i), italic=False)) for i in range(4)]

AI_GRAPHS = "You may use AI tools to produce graphs."


def draws_on(doc, text):
    # after 8 -> 5 on 2026-09-12, part of returning ~30 pt of pure
    # spacing so a page break is never decided by a few points.
    p = para(doc, before=0, after=5, keep_next=True)
    run(p, "Draws on:  " + text, italic=True, color=GRAY, size=9.5)
    return p


def ai_note(p):
    """The original sets the AI-tools sentence in italics; keep that."""
    run(p, "  " + AI_GRAPHS, italic=True)


def no_right_answer(p, ai=False):
    """The recurring bracketed grading note, verbatim from the original."""
    txt = ("  [Note: There is no ‘right’ or ‘wrong’ answer. We want to "
           "evaluate your reasoning based on the economic concepts learned "
           "in class.")
    if ai:
        txt += (" You may use AI tools to produce graphs to explain your "
                "reasoning, but this is optional.")
    run(p, txt + "]")


# ==========================================================================
def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)

    S.ps_masthead(
        doc, "Problem Set 2",
        covers="Covers Module 2",
        # The Class Website, not BruinLearn -- Nico's wording, made
        # permanent on 2026-09-08 (it had been a tracked proposal).
        due="100 points     ·     Due date: see the Course Calendar on "
            "the Class Website")

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
    problem(doc, 1, "Coffee Demand", 20, before=14)
    draws_on(doc, "Module 2 – Elasticities (own-price and cross-price)")

    body(doc,
         "You run a small coffee roaster and would like to predict what will "
         "happen to the quantity demanded for your coffee if you raise your "
         "price. While you do not know the exact demand curve for your "
         "coffee, you do know that in the first year, you charged a price of "
         "$45 and sold 1,200 units, and in the second year, you charged a "
         "price of $40 and sold 1,400 units. The economic environment (e.g., "
         "competition, income, etc.) were similar in the two years.")

    # PROPOSED 1: name the reference point, per the deck's CONVENTION box.
    p = part(doc, "a", 5)
    run(p, "What is the implied own-price elasticity? Is demand elastic, "
           "inelastic or unit-elastic? Briefly discuss plausible reasons why "
           "demand is elastic or inelastic.")
    run(p, "  As in class, compute the percentage changes relative to the "
           "initial point, and name the interval your elasticity refers to.")
    S.elasticity_note(doc)

    p = part(doc, "b", 5)
    run(p, "Based on your answer in question (a): If you raise your price, "
           "will total revenues increase or decrease?  If you change the "
           "price a second time, can you use the same elasticity from (a) to "
           "predict the ")
    run(p, "exact", italic=True)
    run(p, " new quantity demanded?")

    p = part(doc, "c", 5)
    run(p, "Justify your answer in (b) and illustrate it graphically.")
    ai_note(p)

    p = part(doc, "d", 5)
    run(p, "Assume that you maintained a price of $40 in the third year. "
           "However, your main competitor decreased its own price by 20% "
           "relative to the second year. Following this change, your sales "
           "dropped from 1,400 to 1,260 units.  What is the magnitude of the "
           "cross-price elasticity between your coffee and your competitor "
           "one? Does this seem rather elastic or inelastic to you? What "
           "could be the reasons? Think in particular of product "
           "differentiation.")

    # ======================================================================
    # Problem 2
    # ======================================================================
    problem(doc, 2, "Consulting for Dollar General and Apple", 20)
    draws_on(doc, "Module 2 – Elasticities;  Demand Estimation")

    p = part(doc, "a", 10, after=2)
    run(p, "The Wall Street Journal recently reported on Dollar General, a "
           "discount retailer:")
    S.quote(doc,
            "has been trying to avoid passing along higher prices to its "
            "customers and has even reduced prices in some areas",
            "The Wall Street Journal, on Dollar General")

    p = part(doc, "i", None, level=1, before=4)
    run(p, "Using concepts from class, discuss the anticipated effects of "
           "these price reductions on Dollar General’s quantity of goods "
           "sold and total revenue, ")
    run(p, "holding everything else constant", italic=True)
    run(p, ".")
    no_right_answer(p, ai=True)

    p = part(doc, "ii", None, level=1)
    run(p, "Discuss whether a simultaneous price reduction at Walmart would "
           "alter the anticipated effects on quantity that you described "
           "in (i).")

    p = part(doc, "b", 10)
    run(p, "Apple is contemplating raising the price of the Apple Music "
           "“Individual” Plan. Before doing so, Apple would like to ")
    run(p, "estimate", italic=True)
    run(p, " the effects of a price increase on the number of subscriptions "
           "sold. Suppose that Apple has enlisted your help with this task.")

    p = part(doc, "i", None, level=1, before=4)
    run(p, "Using concepts from class, discuss three methods to ")
    run(p, "estimate", italic=True)
    run(p, " the effect of a price increase on the number of subscriptions "
           "sold.")

    p = part(doc, "ii", None, level=1)
    run(p, "Weighing the costs and benefits of the three approaches, which "
           "one would you recommend for Apple? Why?")
    no_right_answer(p)

    # ======================================================================
    # Problem 3
    # ======================================================================
    problem(doc, 3, "Air Travel Demand Elasticity", 15)
    draws_on(doc, "Module 2 – Elasticities (computing elasticity from two "
                  "observations)")

    body(doc,
         "The following data were collected to show the impact of a recent "
         "price war during which all the airlines serving the "
         "Oakland/Burbank market lowered their prices. You may assume that "
         "all other factors affecting air travel demand remained unchanged.")
    body(doc,
         "Your assignment is to calculate the implied (own-price) elasticity "
         "of demand for overall air travel between Oakland and Burbank. All "
         "numbers are measured on a per-month basis.")

    airline_table(doc)

    p = part(doc, "a", 10)
    run(p, "What is the own-price elasticity of demand? Is air travel demand "
           "between Oakland and Burbank elastic or inelastic?")
    # He cut the qualifier "-- here, the situation before the price war"
    # when he accepted this on 2026-09-09: the class convention is enough,
    # and the solutions make the reference point explicit anyway.
    run(p, "  As in class, compute the percentage changes relative to the "
           "initial point, and name the interval your elasticity refers to "
           "(see the Convention under Problem 1).")

    p = part(doc, "b", 5)
    run(p, "Do you expect the own-price elasticity of demand on the Oakland–")
    run(p, "LAX", bold=True)
    run(p, " route to be larger in magnitude, smaller, or about the same as "
           "the elasticity you estimated for Oakland–Burbank? Briefly "
           "justify your answer using economic concepts.")
    no_right_answer(p)

    # ======================================================================
    # Problem 4
    # ======================================================================
    problem(doc, 4, "Aircraft Demand", 20)
    draws_on(doc, "Module 2 – Demand and Revenue (marginal revenue)")

    # TYPO fixed directly: the original read "aircraft AS the following".
    p = body(doc, after=2)
    run(p, "Boeing’s new aircraft has the following demand:")
    T.equation(doc, Q + mrun("=") + mrun("250") + mrun("−") + P,
               before=2, after=4)
    p = body(doc)
    run(p, "where price ")
    run(p, "P", italic=True)
    run(p, " is in millions of dollars.")

    p = part(doc, "a", 5)
    run(p, "How many units will Boeing sell at a price of $140 million? Show "
           "your work.")

    p = part(doc, "b", 10)
    run(p, "Calculate the marginal revenue at the price of $140 million.")

    p = part(doc, "c", 5)
    run(p, "For prices between 100 and 200 million, plot the number of "
           "airplanes sold. What is this curve?")
    ai_note(p)

    # ======================================================================
    # Problem 5
    # ======================================================================
    problem(doc, 5, "Lyft Demand", 25)
    draws_on(doc, "Module 2 – Demand Estimation;  Demand and Revenue")

    # "ran a price experiments" -> "ran a price experiment", accepted
    # 2026-09-09.
    p = body(doc)
    run(p, "An analyst at Lyft initiated an empirical estimation of demand "
           "for Lyft trips in Boston. The analyst ran a price experiment "
           "in the Boston area, by randomizing the price charged for "
           "different routes. She then assembled data on the quantity of "
           "trips taken, depending on the different prices that Lyft "
           "charged. The analyst also collected data on the price charged by "
           "Uber, Lyft’s main competitor, at the time when the Lyft trips "
           "took place. Hence, for each route, she observes the quantity of "
           "trips taken, Lyft price, distance, and price of Uber.")

    body(doc, "Suppose that the analyst has run one multi-variate linear "
              "regression of the following form:", after=2)
    T.equation(doc,
               Q + mrun("=") + B[0] + mrun("+") + B[1] + P + mrun("+")
               + B[2] + PX + mrun("+") + B[3] + Y,
               before=2, after=6)

    p = para(doc, before=2, after=3)
    run(p, "Where:", italic=True, underline=True)

    for sym, sub, desc in ((("Q", None), None, "is the number of Lyft trips "
                                               "taken"),
                           (("P", None), None, "is Lyft’s price in dollars"),
                           (("P", "X"), None, "is Uber’s price in dollars"),
                           (("Y", None), None, "is the distance in miles")):
        p = para(doc, before=1, after=1, left=0.34, hang=0.18)
        run(p, "•  ", color=NAVY)
        run(p, sym[0], italic=True)
        if sym[1]:
            run(p, sym[1], italic=True, subscript=True)
        run(p, "  " + desc)

    body(doc, "As you will find out in your Stats class, the estimated "
              "regression coefficients are as follows:", before=8, after=4)

    tbl = T.table(doc, [
        ["Coefficient", "Variable", "Estimated value"],
        ["", "Constant (intercept)", "67.4"],
        ["", "", "−16.0"],
        ["", "", "2.6"],
        ["", "", "39.3"],
    ], widths_in=[1.15, 3.20, 1.40], size=10, align_right=(2,))
    # real subscripts, never Unicode subscript characters
    for r, idx in ((1, "0"), (2, "1"), (3, "2"), (4, "3")):
        S.cell_runs(tbl, r, 0, S.sub_runs("b", idx))
    S.cell_runs(tbl, 2, 1, [("Lyft price ", {})] + S.sub_runs("P", ""))
    S.cell_runs(tbl, 3, 1,
                [("Uber price ", {})] + S.sub_runs("P", "X"))
    S.cell_runs(tbl, 4, 1,
                [("Distance ", {}), ("Y", dict(italic=True)),
                 (" (miles)", {})])

    body(doc, "yielding the following demand equation:", before=8, after=2)
    T.equation(doc,
               Q + mrun("=") + mrun("67.4") + mrun("−") + mrun("16.0") + P
               + mrun("+") + mrun("2.6") + PX + mrun("+") + mrun("39.3") + Y,
               before=2, after=6)
    S.note(doc, "You do not need to perform the estimation for the Econ "
                "class.", prefix="Note:")

    p = part(doc, "a", 5)
    run(p, "Use the regression results to estimate the number of Lyft trips "
           "taken under the following assumptions: ")
    run(p, "P", italic=True)
    run(p, " = $13.2, ")
    run(p, "P", italic=True)
    run(p, "X", italic=True, subscript=True)
    run(p, " = $20.3, and ")
    run(p, "Y", italic=True)
    run(p, " = 3.8.")

    p = part(doc, "b", 7)
    run(p, "Calculate the own-price demand elasticity under the above "
           "conditions at a Lyft price of $13.2. Is demand elastic, "
           "inelastic or unit-elastic?")

    p = part(doc, "c", 13)
    run(p, "Calculate marginal revenue under the above conditions at a Lyft "
           "price of $13.2.")

    S.save(doc, OUT)
    print("wrote", OUT)


# --------------------------------------------------------------------------
def airline_table(doc):
    """The Oakland/Burbank price-war data.

    The original listed the same nine numbers as a two-column block of
    label/value pairs, with each measure's before, during and change on
    three separate rows.  Same numbers, laid out so a student can read a
    row across.
    """
    T.table(doc, [
        ["", "Before the price war", "During the price war", "Change"],
        ["Passengers", "246,555", "1,053,139", "+806,584"],
        ["Average one-way fare", "$86.50", "$44.69", "−$41.81"],
        ["Market revenue", "$21,327,008", "$47,064,782", "+$25,737,774"],
    ], widths_in=[1.60, 1.65, 1.65, 1.60], size=10,
        align_right=(1, 2, 3))


if __name__ == "__main__":
    main()
