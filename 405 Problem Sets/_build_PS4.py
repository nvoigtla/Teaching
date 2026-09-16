"""Build: Problem Set 4.docx  (MGMT 405, Fall 2026 format).

Source: `_originals/Problem Set 4 (2025 original).docx` -- a BUILD INPUT,
never deleted.  So are the two O*NET screenshots it carries, which are
kept in `_source_images/` (see below).

The job is FORMATTING.  Everything beyond the restyle is a real Word
revision (`w:ins` / `w:del`, author "Claude (proposed)").

CHECKED AGAINST THE DECKS -- PARTIALLY.  PS 4 draws on Module 5
(monopoly) and Module 6 (complex pricing), NEITHER of which has been
revised yet, so those parts could not be checked against a deck.  What
could be checked:
  - Problem 2(c)'s welfare analysis uses Module 4's consumer surplus /
    producer surplus / deadweight loss apparatus, and the solutions'
    figures now follow that deck's colour convention.
  - Problem 4's price discrimination is recapped on Module 7 slide 2:
    "Third degree: segmentation -- Higher price to less elastic groups",
    which is exactly what part (d) asks the student to confirm.
  - Problem 1(b) finds demand INELASTIC at the current price and 1(c)
    finds the optimum at a HIGHER price, which is Module 2's own lesson
    (slide 49: "Firms that operate in the inelastic part of demand
    should raise their price").  See PROPOSED 2.
  - Point totals: 25 + 20 + 20 + 20 + 15 = 100.

FORMATTING (untracked)
  - Course masthead, navy problem headings with a gold points chip, the
    bare centred page number.
  - Problem 1's parts were run-in "a. / b. / c. / d." inside body
    paragraphs; they become the deck-standard (a) / (b) / (c) heads with
    their own point chips, as in every other problem.
  - Problem 3's three candidate demand curves are rebuilt as ONE native
    three-panel figure (they were two separate screenshots, with (A) and
    (B) in one and (C) in another).  Demand is dark red in all three, as
    it already was in the originals.
  - The two O*NET screenshots in Problem 5 are KEPT AS IMAGES: they are
    screenshots of a live website, which is exactly the case the
    standing rule leaves as an image rather than a re-creation.
  - Problem 4's demand functions and Problem 5's index formula set as
    native OMML.
  - The two Barnes & Noble podcast URLs become real hyperlinks on
    descriptive text rather than bare URLs printed in the body.

PROPOSED (tracked -- these go beyond formatting)
  1. Problem 2's cost function is stated as "TC = 100 - 5Q + Q2" in the
     problem set but as "C = 100 - 5Q + Q2" in the solutions.  Made TC in
     both, since the rest of the set uses TC.
  2. Problem 1: a closing sentence asking the student to connect (b) and
     (c) -- demand is inelastic at $2,000 and the optimum is at a HIGHER
     price of $2,450.  That is Module 2's result, and the numbers in this
     problem demonstrate it, but nothing currently asks the student to
     notice.  Worth 0 points as written; Nico may want to move points to
     it.

Run:  python _build_PS4.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _ps_theme as S
from _ps_theme import (DARKRED, GRAY, NAVY, Panel, body, dele, ins, para,
                       part, problem, run)
import _tn_theme as T
from _tn_theme import mfrac, mrun, msub, msup

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Problem Set 4.docx")
IMG = os.path.join(HERE, "_source_images")

VERGE = ("https://www.theverge.com/23642104/"
         "barnes-and-noble-amazon-bookshop-ecommerce-decoder-podcast")
NPR = ("https://www.npr.org/2023/03/07/1161295820/"
       "how-barnes-noble-turned-a-page-expanding-for-the-first-time-in-years")

Q = mrun("Q")
P = mrun("P")
QA, PA = msub(Q, mrun("A")), msub(P, mrun("A"))
QB, PB = msub(Q, mrun("B")), msub(P, mrun("B"))
TC = T.acr("TC")
SIGMA = mrun("Σ", italic=False)


def draws_on(doc, text):
    # after 8 -> 5 on 2026-09-12, part of returning ~30 pt of pure
    # spacing so a page break is never decided by a few points.
    p = para(doc, before=0, after=5, keep_next=True)
    run(p, "Draws on:  " + text, italic=True, color=GRAY, size=9.5)
    return p


def fig_three_demands():
    """Problem 3(b): the three candidate demand curves, side by side.

    (A) perfectly elastic, (B) steep, (C) flat.  One figure rather than
    the original's two screenshots, so the three can be compared at a
    glance -- which is what the question asks the student to do.
    """
    W, H = 6.30, 2.35
    f = S.fig(W, H, name="Problem 3 - three candidate demand curves")
    pw, ph, gap = 1.62, 1.52, 0.44

    specs = [("(A)", (0, 55), (88, 55)),      # perfectly elastic
             ("(B)", (10, 95), (34, 4)),      # steep
             ("(C)", (4, 46), (92, 6))]       # flat
    for k, (tag, p0, p1) in enumerate(specs):
        ox = 0.52 + k * (pw + gap)
        pn = Panel(f, ox, 1.94, pw, ph, ylabel="Price", xlabel="Output")
        pn.demand(p0, p1, label=None, w_pt=2.25)
        if tag == "(A)":
            pn.tick_y(p0[1], "P*", size=10)
        pn.f.label(ox + pw / 2.0 - 0.20, 2.06, [(tag, dict(
            bold=True, color=NAVY, size=11))], w=0.40, h=0.22, align="c",
            name="panel tag")
    return f


# ==========================================================================
def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)

    S.ps_masthead(
        doc, "Problem Set 4",
        covers="Covers Modules 5 and 6",
        # The Class Website, not BruinLearn -- Nico's wording, made
        # permanent on 2026-09-08.
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
    problem(doc, 1, "Maguire Alarm Company", 25, before=14)
    draws_on(doc, "Module 5 – Monopoly;  Module 6 – Complex Pricing "
                  "(the Lerner Index)")

    p = body(doc, after=2)
    run(p, "The Maguire Alarm Company is a medium-sized electronics company "
           "specializing in alarm devices. Maguire’s demand function is "
           "estimated as ")
    T.equation_inline(p, Q + mrun("=") + mrun("4,500") + mrun("−") + P)
    run(p, ", and the company has the following cost function:")
    T.equation(doc, TC + mrun("=") + mrun("150,000") + mrun("+")
               + mrun("400") + Q, before=3, after=4)
    body(doc, "Suppose that the current sales price is $2,000 per unit.")

    p = part(doc, "a", 4)
    run(p, "What is the current level of profits at this price?")

    p = part(doc, "b", 4)
    run(p, "What is the own-price elasticity at the current price of $2,000? "
           "Is demand elastic or inelastic at this price?")

    p = part(doc, "c", 12)
    run(p, "What is the optimal quantity and price (at which profits are "
           "maximized)? What is the level of profits at this optimal price? "
           "What is the own-price elasticity of demand at this optimal "
           "price? Is demand elastic or inelastic at this price?")
    # PROPOSED 2: ask the student to connect (b) and (c)
    ins(p, "  Comparing your answers to (b) and (c), what does the firm’s "
           "position on the demand curve at $2,000 tell you about whether "
           "that price could have been profit-maximizing?")

    p = part(doc, "d", 5)
    run(p, "What is the share of markup in the optimal price (Lerner Index)? "
           "Show the two ways of computing the Lerner Index by relating it "
           "to the demand elasticity estimated in (c).")

    # ======================================================================
    # Problem 2
    # ======================================================================
    problem(doc, 2, "Dayna’s Doorstops", 20)
    draws_on(doc, "Module 5 – Monopoly;  Module 4 – consumer surplus, "
                  "producer surplus and deadweight loss")

    p = body(doc, after=2)
    run(p, "(From Pindyck and Rubinfeld.)  Dayna’s Doorstops, Inc. (DD) is a "
           "monopolist in the doorstop industry. Its total cost is")
    T.equation(doc, TC + mrun("=") + mrun("100") + mrun("−") + mrun("5") + Q
               + mrun("+") + msup(Q, mrun("2", italic=False)),
               before=3, after=4)
    body(doc, "and (inverse) demand is", after=2)
    T.equation(doc, P + mrun("=") + mrun("55") + mrun("−") + mrun("2") + Q,
               before=2, after=4)

    p = part(doc, "a", 5)
    run(p, "What price should DD set to maximize profit? What output does "
           "the firm produce? How much profit does DD generate?")

    p = part(doc, "b", 7)
    run(p, "What would output be if DD acted like a perfect competitor and "
           "set MC = P? How much profit would DD generate?")

    p = part(doc, "c", 8)
    run(p, "Show graphically the consumer surplus and producer surplus under "
           "perfect competition (part (b)). Then draw a separate graph "
           "showing consumer and producer surplus under monopoly (part (a)). "
           "Illustrate on the second graph the deadweight loss due to "
           "monopoly. Explain in words the consequences of having a monopoly "
           "instead of perfect competition on consumer surplus, producer "
           "surplus, and on total welfare.")
    run(p, "  Note: You do not need to plot the exact functions from (a) and "
           "(b); an illustration is fine. You may use AI tools to produce "
           "graphs.", italic=True)

    # ======================================================================
    # Problem 3
    # ======================================================================
    problem(doc, 3, "Reinventing Barnes & Noble", 20)
    draws_on(doc, "Module 5 – Monopolistic Competition")

    body(doc,
         "James Daunt, the CEO of Barnes & Noble, was hired before the sales "
         "of Barnes & Noble reached the bottom. His goal is to change the "
         "business model of the company.")

    p = part(doc, "a", 5)
    run(p, "Describe how the market for books has changed since the early "
           "2000s. What do you think are the reasons for the decline in the "
           "fortunes of Barnes & Noble?")

    p = part(doc, "b", 5)
    run(p, "Consider the following three demand curves. Which one represents "
           "the best approximation for the demand Barnes & Noble faces, and "
           "why? Use the concepts learned in class to describe the "
           "characteristics of the market in which Barnes & Noble operates.")

    S.place(fig_three_demands(), doc, before=8, after=10)

    p = part(doc, "c", 5)
    run(p, "What does this type of demand curve imply about the ability of "
           "Barnes & Noble to increase the price of its products?")

    p = part(doc, "d", 5)
    run(p, "In a recent interview, James Daunt proposed to give more freedom "
           "to the individual bookstores:")
    S.quote(doc,
            "You should do what you want; there are no restrictions on what "
            "books you stock… Each bookshop is quite individual.",
            "James Daunt, CEO of Barnes & Noble")
    p = body(doc, left=0.30)
    run(p, "Propose two strategies for bookstores to differentiate "
           "themselves from competitors. What does differentiation imply for "
           "the price elasticity of demand?")

    p = body(doc, before=6)
    run(p, "Two podcasts you might consider listening to, to know more about "
           "the business strategy of Barnes & Noble: ")
    S.link(p, "The Verge — Decoder", VERGE)
    run(p, " and ")
    S.link(p, "NPR — How Barnes & Noble turned a page", NPR)
    run(p, ".")

    # ======================================================================
    # Problem 4
    # ======================================================================
    problem(doc, 4, "Avocado Pricing Strategy", 20)
    draws_on(doc, "Module 6 – Complex Pricing (third-degree price "
                  "discrimination)")

    p = body(doc, after=2)
    run(p, "From 2015 to 2018, the demand for avocados (both conventional "
           "and organic) in Atlanta was")
    T.equation(doc, QA + mrun("=") + mrun("730") + mrun("−") + mrun("175")
               + PA, before=3, after=4)
    p = body(doc, after=2)
    run(p, "where quantity ")
    run(p, "Q", italic=True)
    run(p, " is in 1,000 per week and ")
    S.sub_inline(p, "P", "A")
    run(p, " is the average weekly price per avocado across all retailers. "
           "For the same time period, the demand for avocados in San Diego "
           "was")
    T.equation(doc, QB + mrun("=") + mrun("855") + mrun("−") + mrun("319")
               + PB, before=3, after=4)
    body(doc,
         "Again, Q is in 1,000 per week and the price is per avocado. The "
         "marginal cost of production is the same in both markets and equal "
         "to $0.17 per avocado.")

    S.note(
        doc,
        "The fact that Q is given in 1,000 units should not affect any of "
        "your calculations. You can use the formulas and marginal cost "
        "exactly as given above, without multiplying or dividing by 1,000. "
        "Only at the end, when you interpret total quantity, total revenues "
        "or profits, note that all of those will be in 1,000. For example, "
        "if you compute total revenues of $5,000 those would actually "
        "correspond to $5,000,000. The price, however, is given per unit: if "
        "you compute an optimal price of $100 per unit, that is your final "
        "result. You may use AI tools to produce graphs; graphs do not need "
        "to be to scale, they can be used merely for illustration.",
        prefix="Note:")

    for letter, pts, txt in (
            ("a", 5, "What is the optimal price to charge in Atlanta? Show "
                     "your work both analytically and graphically."),
            ("b", 5, "What is the optimal price to charge in San Diego? Show "
                     "your work both analytically and graphically."),
            ("c", 3, "What is the own-price elasticity of demand at the "
                     "profit-maximizing prices in the two cities?"),
            ("d", 4, "Do your findings in (a) and (b) (price in each market) "
                     "and in (c) (own-price elasticity in each market) "
                     "confirm what we learned in class about elasticities "
                     "and price discrimination?")):
        run(part(doc, letter, pts), txt)

    p = part(doc, "e", 3)
    run(p, "Explain which problems may result from the pricing structure "
           "that you have calculated in the previous questions.")
    run(p, "  [A quantitative answer is not required. Hint: What if "
           "transportation costs are low relative to price differences?]",
        italic=True)

    # ======================================================================
    # Problem 5
    # ======================================================================
    problem(doc, 5, "Working Interactively with AI: The Future of Work", 15)
    draws_on(doc, "Course-wide – marginal analysis applied to tasks; O*NET "
                  "occupational data")

    body(doc,
         "Welcome to the AI & Labor Market Challenge. In this exercise you "
         "will collaborate with ChatGPT to explore how automation might "
         "reshape different occupations. Your task is to compute an index of "
         "exposure to AI for different jobs. You will use, interpret and may "
         "also challenge ChatGPT’s classification. Here are the step-by-step "
         "instructions.")

    step(doc, 1, 2, "Choosing the Occupations")
    body(doc,
         "Go to the O*NET website (see the screenshot below) and identify "
         "two occupations from your team that are as different from each "
         "other as possible. The goal is to compare jobs that rely on "
         "different skills and tasks.")
    S.note(doc,
           "If all team members have similar roles (for example, various "
           "types of managers), select one previous occupation or choose "
           "from this list: economist, lawyer, accountant, waiter/waitress, "
           "plumber, actor/actress, high-school teacher.", prefix="Note:")
    shot(doc, "onet_search.png", 5.5)

    step(doc, 2, 5, "Listing and Classifying Tasks")
    body(doc,
         "For each of the two chosen occupations, open the “Details → Tasks” "
         "tab in O*NET (see the screenshot below). Copy all core tasks for "
         "that occupation (supplemental tasks are not required), together "
         "with their importance scores (0–100), into an Excel spreadsheet.")
    shot(doc, "onet_tasks.png", 5.5)

    body(doc, "Paste this spreadsheet into ChatGPT (or any AI interface) and "
              "ask it to:", before=8)

    p = para(doc, before=4, after=3, left=0.32, hang=0.18)
    run(p, "•  ", color=NAVY)
    run(p, "Replacement dummy:", bold=True)
    run(p, " Classify, for each profession, which task ChatGPT believes will "
           "be replaced by AI-supported automation in the next five years. "
           "Hint: ask ChatGPT to add a new column “Replacement dummy” of 1s "
           "and 0s, where 1 indicates that the task will be replaced.")
    p = para(doc, before=2, after=3, left=0.50, hang=0.18)
    run(p, "–  ", color=NAVY)
    run(p, "If you want a more finely-grained index, you could instead ask "
           "ChatGPT to classify the extent to which each task will be "
           "replaced: 0 for “none”, 0.25 “some replacement”, 0.5 "
           "“substantial replacement”, 0.75 “almost completely replaced” and "
           "1 “completely replaced”. You will receive full credit for the "
           "simpler 0/1 dummy as well.", italic=True)

    p = para(doc, before=4, after=3, left=0.32, hang=0.18)
    run(p, "•  ", color=NAVY)
    run(p, "Review ChatGPT’s classification.", bold=True)
    run(p, " You may change its coding if you disagree. Report your final "
           "table with scores in your solutions. You do not need to document "
           "these steps: the idea is that you create an AI-supported coding "
           "of tasks, supported by your own judgment. You can of course ask "
           "ChatGPT for its reasoning behind each code.")

    step(doc, 3, 8, "Computing the AI-Replacement Index")
    body(doc,
         "For each of the two occupations, compute an AI-Replacement Index "
         "using the importance score of each task as a weight:", after=2)
    T.equation(doc,
               mrun("AI Replacement Index", italic=False) + mrun("=")
               + mfrac(SIGMA + mrun("(") + mrun("Importance", italic=False)
                       + mrun("×") + mrun("Replacement dummy", italic=False)
                       + mrun(")"),
                       SIGMA + mrun("(") + mrun("Importance", italic=False)
                       + mrun(")")),
               before=2, after=6)
    body(doc,
         "where Σ indicates “sum”. You may compute this in Excel, or ask "
         "ChatGPT to do the calculation. If the index equals 1, every single "
         "task of an occupation is replaceable by AI; a value of 0 implies "
         "that the occupation should be relatively safe from AI.")

    body(doc,
         "Compare the indexes for the two occupations. Which occupation "
         "appears more exposed to automation? Do you agree or disagree with "
         "this result? Discuss briefly.")

    S.save(doc, OUT)
    print("wrote", OUT)


def step(doc, n, pts, title):
    p = para(doc, before=12, after=5, keep_next=True)
    run(p, "Step {}  ".format(n), bold=True, color=NAVY, size=12)
    run(p, "({} points)  ".format(pts), bold=True, color=S.GOLD, size=10.5)
    run(p, "— " + title, bold=True, color=NAVY, size=12)
    return p


def shot(doc, fname, width_in):
    """An O*NET screenshot, kept as an image on purpose.

    These are screenshots of a live website: the standing rule rebuilds
    tables, charts and equations natively but leaves genuine screenshots
    alone, because a re-creation would be a different thing from what
    the student will actually see on the site.
    """
    path = os.path.join(IMG, fname)
    if not os.path.exists(path):
        raise SystemExit(
            "missing build input: {}\n(extract it from the 2025 original "
            "into _source_images/)".format(path))
    return T.picture(doc, path, width_in, rounded=False, shadow=True,
                     before=6, after=6)


if __name__ == "__main__":
    main()
