"""Build: Problem Set 5.docx  (MGMT 405, Fall 2026 format).

Source: `_originals/Problem Set 5 (2025 original).docx` -- a BUILD INPUT,
never deleted.

CHECKED AGAINST `Module 7 - Revised.pptx` (oligopoly and game theory),
which IS revised.  Modules 5, 6 and 8 are not, so Problems 1 and 2 could
only be checked internally.
  - Problem 3 is Cournot (reaction functions, equilibrium), then the
    short-run / long-run produce-or-exit decision, then Bertrand -- all
    Module 7 material, with the produce / exit rule from Module 4.
  - Problem 4 is a 2x2 dominant-strategy game, and its matrix now uses
    Module 7's own payoff-matrix format.
  - Problem 2 is Module 6 (two-part pricing), recapped on Module 7 slide
    2: "Two-part pricing when MC>0 (e.g. Starbucks): per-unit fee = MC,
    flat fee = area under the demand curve above the per-unit fee" --
    which is exactly the method the solutions use.

FORMATTING (untracked)
  - Course masthead, navy problem headings with a gold points chip, the
    bare centred page number.
  - Problem 4's payoff matrix is rebuilt in Module 7's house style:
    the column player named on top in gold, the row player at the left in
    concept blue, white cells with navy borders, each cell reading
    "a , b" with the two payoffs in their own colours, and the caption
    "Payoffs to (Firm A, Firm B)".  The original was a plain grid whose
    header cells did not say which payoff belonged to whom.
  - Demand functions and the industry demand curve set as native OMML
    with true subscripts.
  - Parts renumbered onto the deck-standard (a) / (b) / (c) heads.

PROPOSED (tracked -- these go beyond formatting)
  1. Problem 3's parts add up to 40, not the 35 in its header:
     1 + 7 + 7 + 2 + 4 + 3 + 6 + 6 + 4, plus a further point offered in
     the closing note.  With Problems 1, 2 and 4 at 20 + 30 + 15, the set
     totals 105 (106 with the note) rather than 100.  Flagged rather than
     rescaled: which part gives up the points is Nico's call.
  2. Problem 3(d) says the rent is "500 per year" while the solutions say
     "per period"; and (e)-(f) then reason in years.  Made "per year"
     throughout.

Run:  python _build_PS5.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _ps_theme as S
from _ps_theme import GRAY, NAVY, body, ins, para, part, problem, run
import _tn_theme as T
from _tn_theme import mrun, msub

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Problem Set 5.docx")

Q, P, F = mrun("Q"), mrun("P"), mrun("F")
q1, q2 = msub(mrun("q"), mrun("1", italic=False)), \
    msub(mrun("q"), mrun("2", italic=False))
TC1 = msub(T.acr("TC"), mrun("1", italic=False))
TC2 = msub(T.acr("TC"), mrun("2", italic=False))

PAYOFFS = [[(20, 20), (10, 26)],
           [(26, 12), (15, 15)]]


def draws_on(doc, text):
    # after 8 -> 5 on 2026-09-12, part of returning ~30 pt of pure
    # spacing so a page break is never decided by a few points.
    p = para(doc, before=0, after=5, keep_next=True)
    run(p, "Draws on:  " + text, italic=True, color=GRAY, size=9.5)
    return p


# ==========================================================================
def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)

    # PROPOSED 1: the parts sum to 105, not 100
    S.ps_masthead(
        doc, "Problem Set 5",
        covers="Covers Modules 6 and 7",
        # The Class Website, not BruinLearn -- Nico's wording, made
        # permanent on 2026-09-08.  No points total on this line, for the
        # same reason as PS 3: the problem headers sum to 100, but the
        # tracked note below flags that Problem 3's parts sum to 40
        # against its header's 35, and resolving that changes the total.
        due="Due date: see the Course Calendar on the Class Website")
    p = para(doc, before=0, after=10)
    run(p, "100 points", italic=True, color=GRAY, size=10.5)
    ins(p, "     ·     [Problem 3’s parts add up to 40, not the 35 in its "
           "header, so the set currently totals 105.]",
        italic=True, color=GRAY, size=10.5)

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
    problem(doc, 1, "Pricing in the Market for Air Taxis", 20, before=14)
    draws_on(doc, "Module 6 – Complex Pricing (willingness to pay, "
                  "versioning)")

    body(doc,
         "The market for short-distance transportation may be facing a "
         "technological disruption with the imminent arrival of air taxis. "
         "Joby Aviation, Inc., a key player in this market, is consulting "
         "your advice on potential future pricing strategies. Joby "
         "explicitly states their preference for adopting a pricing model "
         "based on per-trip rates rather than a membership or flat fee.")
    body(doc,
         "Consider the commute from LAX airport to downtown L.A. Typically "
         "this commute takes an average of 40 minutes with a regular taxi "
         "and costs $50. In contrast, an air taxi would reach the same "
         "destination in just 10 minutes.")

    p = part(doc, "a", 4)
    run(p, "Joby targets marginal costs of $3 per mile. Provide an estimate "
           "of marginal costs (MC) for the trip. You should justify your "
           "assumptions and supply references when appropriate.")

    p = part(doc, "b", 10)
    run(p, "Let us approximate a linear demand curve for this particular "
           "trip, on a daily basis. Imagine a typical day with a potential "
           "customer base of 1,000 passengers, all of whom would otherwise "
           "opt for a regular taxi. These passengers can be categorized into "
           "four distinct types:")
    for t in ("400 CEOs with tight business schedules,",
              "300 other business travelers with greater flexibility in "
              "their time,",
              "200 tourists who plan to do city tours, and",
              "100 students who are very flexible with their time."):
        q = para(doc, before=2, after=2, left=0.64, hang=0.18)
        run(q, "•  ", color=NAVY)
        run(q, t)

    body(doc,
         "For each of these four passenger types, estimate the average "
         "willingness to pay for the trip and then plot these values in a "
         "graph. The graph should display dollars per ride on the y-axis and "
         "the number of passengers on the x-axis (ranging from 0 to 1,000). "
         "Connect the estimated data points into an approximately linear "
         "demand curve. Sketch the marginal revenue, the marginal costs and "
         "the optimal price (no computation required).", left=0.30,
         before=6)

    S.note(
        doc,
        "Relate the willingness to pay to the opportunity cost of time. For "
        "each type of customer, provide (and justify) assumptions regarding "
        "their willingness to pay for the time saved if they use an air "
        "instead of a regular taxi — that is, their willingness to pay for "
        "the air taxi should not be smaller than $50.", prefix="Hint:")

    p = part(doc, "c", 6)
    run(p, "Suppose Joby would like to extract more surplus from the "
           "different passenger types but cannot differentiate between them. "
           "Give an example of a pricing strategy that could incentivize the "
           "different passengers to pay different prices.")
    run(p, "  [No computation required. A simple description suffices. You "
           "may use AI tools to produce graphs if it helps you illustrate "
           "your argument.]", italic=True)

    # ======================================================================
    # Problem 2
    # ======================================================================
    problem(doc, 2, "Two-Part Pricing", 30)
    draws_on(doc, "Module 6 – Complex Pricing (two-part tariffs, segment "
                  "pricing)")

    body(doc,
         "You are an executive for a company that rents out server space on "
         "super computers. Your firm has two types of potential customers: "
         "businesses and academic researchers.")

    p = body(doc, after=2)
    run(p, "A typical business customer has the demand function ")
    T.equation_inline(p, Q + mrun("=") + mrun("10") + mrun("−") + P)
    run(p, ", where ")
    run(p, "Q", italic=True)
    run(p, " is in days of computing power used per month. The typical "
           "academic customer has the demand ")
    T.equation_inline(p, Q + mrun("=") + mrun("8") + mrun("−") + P)
    run(p, ". The marginal cost of additional computing is $2 per day, "
           "regardless of volume. Assume that fixed costs are zero. You are "
           "charging a two-part tariff with a monthly subscription fee (")
    run(p, "F", italic=True)
    run(p, ") and a usage price per day (")
    run(p, "P", italic=True)
    run(p, ").")

    for letter, pts, txt in (
            ("a", 8, "What subscription fee and usage price would you charge "
                     "for business customers? What would be your profits "
                     "from a typical business customer?"),
            ("b", 6, "Suppose you are initially only offering the business "
                     "option (that is, at the F and P computed in (a)). "
                     "Would a typical academic customer buy that option? Why "
                     "or why not? Hint: compute the consumer surplus of "
                     "academic users under the pricing scheme for business "
                     "customers."),
            ("c", 5, "After taking a class on complex pricing, you create a "
                     "special deal for people who can show proof of an "
                     "academic affiliation. What F and P would you charge "
                     "for academic customers? What would be your profits "
                     "from a typical academic customer?"),
            ("d", 4, "Which two complex pricing strategies does your new "
                     "pricing scheme combine?"),
            ("e", 2, "If you do not ask for proof of academic affiliation, "
                     "which problem can arise with your pricing scheme? "
                     "Hint: what would business customers do if they learned "
                     "about the academic deal?"),
            ("f", 5, "Give an example of a similar pricing scheme from your "
                     "work experience or an industry you know closely. Does "
                     "it face the same problem as the one you described in "
                     "(e)? If yes, how is it circumvented? If not, why is it "
                     "not circumvented?")):
        run(part(doc, letter, pts), txt)

    # ======================================================================
    # Problem 3
    # ======================================================================
    problem(doc, 3, "Duopoly", 35)
    draws_on(doc, "Module 7 – Oligopoly (Cournot, Bertrand);  Module 4 – "
                  "the produce / stop / exit decision")

    p = body(doc, after=2)
    run(p, "Two identical firms produce an identical product, and they are "
           "the only firms in this market. Their total annual costs are")
    T.equation(doc,
               TC1 + mrun("=") + mrun("300") + mrun("+") + mrun("15") + q1
               + mrun("      ") + TC2 + mrun("=") + mrun("200") + mrun("+")
               + mrun("25") + q2, before=3, after=4)
    p = body(doc, after=2)
    run(p, "where ")
    T.equation_inline(p, q1)
    run(p, " is the output of firm 1 and ")
    T.equation_inline(p, q2)
    run(p, " the output of firm 2. Price is determined by the industry "
           "demand curve")
    T.equation(doc, P + mrun("=") + mrun("110") + mrun("−") + Q,
               before=2, after=4)
    p = body(doc)
    run(p, "where ")
    T.equation_inline(p, Q + mrun("=") + q1 + mrun("+") + q2)
    run(p, ". Each firm decides how much to produce taking as given the "
           "output decision of the other one.")

    p = part(doc, "a", 1)
    run(p, "Which type of competition do the two firms engage in? A "
           "quantitative answer is not needed.")

    p = part(doc, "b", 7)
    run(p, "Derive the reaction function for each firm. Plot the two "
           "reaction functions in a single graph with ")
    S.sub_inline(p, "q", "1")
    run(p, " on the x-axis and ")
    S.sub_inline(p, "q", "2")
    run(p, " on the y-axis. Explain what these curves mean.")
    run(p, "  You may use AI tools to produce graphs.", italic=True)

    p = part(doc, "c", 7)
    run(p, "Find the equilibrium quantities and the price. Calculate the "
           "profit of each firm in equilibrium. Denote the equilibrium "
           "quantities in the graph from part (b).")

    # PROPOSED 2: "per year", matching (e) and (f)
    p = part(doc, "d", 2)
    run(p, "It turns out that the land on which firm 2 has been built is "
           "legally owned by Don, who offers to rent the land to firm 2 for "
           "500 per year. That is, this cost would occur in addition to all "
           "previously mentioned costs of firm 2. How would this be "
           "represented in firm 2’s total cost function?")

    p = part(doc, "e", 4)
    run(p, "A court rules that firm 2 has to pay the rent for this year "
           "regardless of whether or not it produces. Will firm 2 still "
           "produce this year? If so, what is the optimal quantity "
           "produced?")

    p = part(doc, "f", 3)
    run(p, "The court also rules that starting next year, the firm would not "
           "have to pay the rent if it declared bankruptcy. Thus “next year” "
           "can be considered the long run, when no fixed costs occur if the "
           "firm goes out of business (this includes all fixed costs, not "
           "only the rent). Will the firm produce next year, or instead shut "
           "down?")

    p = part(doc, "g", 6)
    run(p, "Suppose that we are in the second year now and firm 2 exits the "
           "market for good, so that firm 1 is the only producer. Calculate "
           "its optimal output ")
    S.sub_inline(p, "q", "1")
    run(p, "*, the new market price, and profits. Relate your findings to "
           "the graph from part (b).")

    p = part(doc, "h", 6)
    run(p, "Compute consumer surplus under the scenario with two firms "
           "(part (c)) and with only firm 1 (part (g)). Did Don do consumers "
           "a favor by demanding high land rental payments and pushing firm "
           "2 out of business? Explain your answer by drawing and labeling a "
           "graph showing the consumer surplus under both scenarios.")

    p = part(doc, "i", 4)
    run(p, "Let’s go back to the initial setup (the problem described before "
           "part (a)). Suppose that the two firms compete by setting prices. "
           "How is this competition called, and what is the most likely "
           "outcome (market price and total quantity produced)? Assume there "
           "is no collusion.")

    S.note(
        doc,
        "If one firm becomes the only producer for a certain price range, "
        "use your answer from (g) to check whether it may want to charge an "
        "even lower price to maximize its profits.",
        prefix="Note / hint (for 1 point):")

    # ======================================================================
    # Problem 4
    # ======================================================================
    problem(doc, 4, "Price Wars", 15)
    draws_on(doc, "Module 7 – Game Theory (dominant strategies, Nash "
                  "equilibrium)")

    body(doc,
         "The matrix below displays the advertising rates and profit results "
         "of the two rival newspapers in a major city.", after=8)

    S.payoff_matrix(doc, "FIRM A", "FIRM B",
                    ["High Ad Rate", "Low Ad Rate"],
                    ["High Ad Rate", "Low Ad Rate"],
                    PAYOFFS)

    p = part(doc, "a", 10, before=10)
    run(p, "The newspapers set their advertising rates independently. "
           "Determine the optimal strategy for each firm. Is there a "
           "dominant strategy for any of the two firms? If so, explain why.")

    p = part(doc, "b", 5)
    run(p, "What is the Nash equilibrium of this game, and the payoff for "
           "each firm?")

    S.save(doc, OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
