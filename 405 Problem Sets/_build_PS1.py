"""Build: Problem Set 1.docx  (MGMT 405, Fall 2026 format).

Source: `_originals/Problem Set 1 (2025 original).docx` -- a BUILD INPUT,
never deleted.

STATUS: Nico reviewed the tracked round and ACCEPTED IT IN FULL on
2026-09-08, then edited on top of it.  This script now emits the agreed
text with NO revision marks; a future proposal goes back in as
`ins` / `dele`.  Accepted in that round: the 2026 salary and revenue
figures ($46,000 / $78,000), the truck in the bracketed note aligned with
the solution's own truck, "shifts of" rather than "movements of" the
demand curve, Problem 4 split into (a) / (b) / (c) on the solution's own
grading boundaries, and "its capacity" for "their capacity".

WHAT HE CHANGED ON TOP, adopted here (2026-09-08)
  1. "Due date: see the Course Calendar on BruinLearn"
     -> "... on the Class Website".
  2. The AI-tools policy is rewritten and gains a heading line,
     "Important information, please read carefully", set BOLD ITALIC DARK
     RED as its own centred first paragraph inside the card.  The policy
     drops "We encourage you to ask specific questions ... the answers are
     not always correct" and gains two things: a screening sentence ("We
     will screen for undisclosed AI-generated content using advanced
     AI-detection programs") and a study-partner framing that ends by
     pointing at the exams, where AI is prohibited.  See `S.policy_card`.
  3. Problem 1's bracketed note now says the grading is on the reasoning
     "not on the precision of your numerical assumptions", and caps the
     answer at one page.
  4. Problem 3(b): "(holding supply constant)" -> "(assuming the supply
     curve does not shift)", plus a note pointing the student at the
     SLOPE of the supply curve.

FORMATTING
  - Course masthead, now carrying "Prof. Nico Voigtlaender" at the right
    of the subtitle line, above the gold rule.
  - Navy problem headings with a gold points chip; bare centred page
    number, so the file is reusable next year.
  - Q(d) / Q(s) and the demand and supply functions as native OMML with
    true subscripts; parts as (a) / (b) / (c) run-in heads with their own
    points; a gray "Draws on:" line under each problem heading, verified
    against `Module 1 - Revised.pptx` and `Module 2 - Revised.pptx`.

The note added to Problem 3(b) used to end "...significantly higher?.]"
with the parenthesis opened at "(e.g. can this industry" never closed.  It
was reproduced verbatim rather than silently repaired; he closed it himself
on 2026-09-12 -- "...significantly higher?).]".

Run:  python _build_PS1.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _ps_theme as S
from _ps_theme import GRAY, NAVY, body, para, part, problem, run
import _tn_theme as T
from _tn_theme import mrun, msub

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Problem Set 1.docx")

# -- OMML symbols ----------------------------------------------------------
QD = msub(mrun("Q"), mrun("d", italic=False))
QS = msub(mrun("Q"), mrun("s", italic=False))
P = mrun("P")
I = mrun("I")


def draws_on(doc, text):
    """The gray provenance line under a problem heading (format layer)."""
    # after 8 -> 5 on 2026-09-12, part of returning ~30 pt of pure
    # spacing so a page break is never decided by a few points.
    p = para(doc, before=0, after=5, keep_next=True)
    run(p, "Draws on:  " + text, italic=True, color=GRAY, size=9.5)
    return p


# ==========================================================================
def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)

    S.ps_masthead(
        doc, "Problem Set 1",
        covers="Covers Modules 1 and 2",
        due="100 points     ·     Due date: see the Course Calendar on "
            "the Class Website")

    # -- Submission instruction, above the policy card (2026-09-12) -------
    # No hyperlink on "BruinLearn": the same PS goes to both sections.
    S.submit_card(doc)

    # -- AI policy, in Nico's wording (2026-09-08) -------------------------
    S.policy_card(doc)

    # ======================================================================
    # Problem 1
    # ======================================================================
    problem(doc, 1, "Business Choice", 25, before=14)
    draws_on(doc, "Module 1 – Economic Costs Include Opportunity Costs")

    body(doc,
         "Your gardener, who has thus far been employed by a large "
         "landscaping company (with an annual salary of $46,000), approaches "
         "you with a question: Should he open his own gardening business? He "
         "is located in West L.A. and heard from other gardeners that he can "
         "expect annual revenues of $78,000. Try your best to provide a "
         "recommendation, by making reasonable, well-informed assumptions "
         "about all the relevant cost components. For each cost component, "
         "state briefly what brought you to this assumption. You may use "
         "references to (online) sources. Distinguish between explicit and "
         "implicit costs of the gardening business.")

    # Nico's note, 2026-09-08: the reasoning is what is graded, and the
    # answer is capped at a page.
    body(doc,
         "[Note: There is no “right” or “wrong” answer. We want to evaluate "
         "your reasoning based on the economic concepts learned in class, "
         "not on the precision of your numerical assumptions. For "
         "simplicity, use average annual costs for all components – for "
         "example, if a new truck costs $40,000 and can be used for 10 "
         "years, the average annual cost would be $4,000. Your answer does "
         "not need to exceed 1 page.]")

    # ======================================================================
    # Problem 2
    # ======================================================================
    problem(doc, 2, "Demand Curve", 15)
    draws_on(doc, "Module 1 – Demand and Supply; Equilibrium")

    p = body(doc)
    run(p, "Assume that the demand curve for a product is given by ")
    T.equation_inline(p, QD + mrun("=") + mrun("200") + mrun("−") + P
                      + mrun("+") + mrun("4") + I)
    run(p, " where ")
    run(p, "I", italic=True)
    run(p, " is average income measured in thousands of dollars. The supply "
           "curve is ")
    T.equation_inline(p, QS + mrun("=") + P + mrun("−") + mrun("100"))
    run(p, ".")

    p = part(doc, "a", 4)
    run(p, "If ")
    run(p, "I", italic=True)
    run(p, " = 10, find the market-clearing price and quantity.")

    p = part(doc, "b", 4)
    run(p, "If ")
    run(p, "I", italic=True)
    run(p, " = 20, find the market-clearing price and quantity.")

    p = part(doc, "c", 7)
    run(p, "Draw a graph to illustrate your answers (draw the supply and "
           "demand curves and show how situations (a) and (b) differ). "
           "[Note: A sketch is sufficient – you do not need to plot the "
           "actual values]")

    # ======================================================================
    # Problem 3
    # ======================================================================
    problem(doc, 3,
            "Income Elasticity and Demand During Economic Uncertainty", 35)
    draws_on(doc, "Module 2 – Elasticities (income elasticity);  "
                  "Module 1 – Equilibrium")

    body(doc,
         "During major economic crises (such as the “Great Recession” "
         "in 2008), many households started experiencing a sharp increase in "
         "economic uncertainty and/or saw their assets losing value due to "
         "stock market devaluation. This can lead to a decrease in household "
         "expected and realized income.")

    # "shifts of", accepted 2026-09-08 (Module 1 slide 23 draws the
    # movement-along vs. shift distinction the question depends on)
    p = part(doc, "a", 12)
    run(p, "Choose three products or services and rank them according to how "
           "much you expect quantity demanded to fall after the onset of a "
           "crisis, from largest to smallest decline. Justify your ranking "
           "based on the concept of income elasticity of demand. Here you are "
           "encouraged to give examples from your own experience during times "
           "of economic crises or uncertainty, or from the sector you’re "
           "working in. For simplicity, you can assume that the supply curve "
           "remains unchanged. Tie your argument to shifts of the demand "
           "curve.")

    # Nico, 2026-09-08: "holding supply constant" -> "assuming the supply
    # curve does not shift", and a note pointing at the SLOPE of supply.
    p = part(doc, "b", 12)
    run(p, "Provide a visual representation of the changes in the three "
           "industries that you wrote about in point (a). What happens to "
           "quantities and prices in equilibrium (assuming the supply curve "
           "does not shift)? You may use AI tools to produce graphs. [Note: "
           "Although not required for full credit, here it will be useful "
           "for you to think about the slope of the supply curve (e.g. can "
           "this industry produce more units easily? Will the cost of the "
           "extra units be significantly higher?).]")

    p = part(doc, "c", 11)
    run(p, "Do you think the assumption of unchanged supply is reasonable for "
           "all three industries? Which industries/products do you think "
           "would see supply adjustments in the short term (say, 1 month)? "
           "Assuming the supply curve can shift, take into account this shift "
           "for the three cases you discussed under (a). Now, with both the "
           "demand and supply curves shifting, (re-)rank the three cases by "
           "their expected short-term price change. Illustrate your answer "
           "graphically again and explain. You can assume that the slopes of "
           "demand and supply curves in all industries/products stay "
           "constant.")

    p = part(doc, "d", 5, bonus=True)
    run(p, "Assuming your answers in part (a) are correct and allowing supply "
           "to adjust as in (c), can we be certain if the resulting prices "
           "will be above, below or the same as before the crisis?")

    # ======================================================================
    # Problem 4
    # ======================================================================
    problem(doc, 4, "Consulting Project", 25)
    draws_on(doc, "Module 1 – Equilibrium;  Economic Costs Include "
                  "Opportunity Costs")

    body(doc,
         "You are hired as a consultant by a large fast food chain. The "
         "company would like to understand whether an expansion of "
         "unemployment benefits in California could affect its capacity to "
         "hire low-skilled workers in the state. Use a supply-demand "
         "framework for labor to analyze this question.")

    # split into the three parts the solution grades, accepted 2026-09-08
    p = part(doc, "a", 8)
    run(p, "Draw and label a graph that depicts a demand curve and a supply "
           "curve in the market for low-skilled labor, and identify the "
           "equilibrium (low-skill) wage rate in the graph. Hint: Put "
           "low-skill labor on the x-axis and wages on the y-axis. Low-skill "
           "labor demand represents the amount of low-skill labor demanded by "
           "all firms in California; labor supply is the total quantity of "
           "low-skill labor supplied by all workers in the state.")

    p = part(doc, "b", 12)
    run(p, "Does an expansion of unemployment benefits shift the labor ")
    run(p, "supply", bold=True)
    run(p, " curve? Explain why or why not, and show the shift in your "
           "graph.")

    p = part(doc, "c", 5)
    run(p, "Does it shift the labor ")
    run(p, "demand", bold=True)
    run(p, " curve? Explain why or why not.")

    S.save(doc, OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
