"""Build: Module 3 - Teaching Note Bang for Buck Rule.

Nico's original, restyled to the course theme and checked against the
Module 3 deck (`Module 3 - Revised.pptx`, slides 43-44 and 50; slide 43
carries the "Teaching Note: Bang-for-the-Buck Rule" pointer).

STATUS: reviewed and ACCEPTED in full on 2026-09-06, so this script now
emits the agreed text with no revision marks.  Any future proposal goes
back in as `ins` / `dele`.

WHAT CHANGED relative to the original (all accepted 2026-09-06)
  1. New opening sentence naming the regime: this is the LONG-run rule,
     both inputs flexible.  The note never said so, while deck slide 43
     leads with "Decision rule for the long run" and the podcast rules
     require every rule to state which regime it lives in.
  2. "the cost of capital (pK)" -> "the price of capital (pK)".  Deck
     slide 43's glossary reads "pk: Price of Capital", and "cost of
     capital" is a different thing in finance (WACC).
  3. "Recall from Module 3 that MRPL = MR . MPL" -> a cross-reference to
     the short-run hiring note.  The original pointed at Module 3 from
     inside Module 3; MRPL = MR x MPL is derived in the Hiring Decisions
     note and on deck slide 24.
  4. New closing paragraph with the three conditions from deck slide 43
     (both inputs flexible / a given output Q / constant input prices).
     The note stated only the third.

SWEEP, 2026-09-06.  The cross-reference is Title Case -- "the Teaching
Note on Hiring Decisions in the Short Run" -- following the convention
Nico set by hand on the MR=MC note.

FORMATTING (untracked)
  - Acronyms upright with italic subscripts (MRP, MP upright; L, K
    italic), per the OMML convention.
  - The bang-for-the-buck rule moves from bold body math into the cream
    hero card, matching how deck slide 43 presents it.

Run:  python _build_TN_M3_BangForBuck.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tn_theme as T
from _tn_theme import (NAVY, acr, body, equation, heading, masthead,
                       mfrac, mrun, para, run)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Module 3 - Teaching Note Bang for Buck Rule.docx")

MRP_L = acr("MRP", "L")
MRP_K = acr("MRP", "K")
MP_L = acr("MP", "L")
MP_K = acr("MP", "K")
W = mrun("w")
P_K = T.msub(mrun("p"), mrun("K"))
MR = acr("MR")


def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)
    masthead(doc, "Teaching Note – Module 3: Deriving the Bang-for-the-Buck Rule")

    # -- regime first: this is the long-run rule ---------------------------
    p = body(doc)
    run(p, "This note derives the rule for the optimal mix of capital and "
           "labor in the ")
    run(p, "long run", bold=True)
    run(p, ", when both inputs can be adjusted.")

    p = body(doc)
    run(p, "Marginal analysis requires that for the optimal use of a "
           "flexible input, its marginal benefit must equal its marginal "
           "cost (MB=MC). In the context of inputs in production, the "
           "marginal benefit is “by how much do our revenues increase "
           "if we use more of this input?” This is the marginal revenue "
           "product (")
    run(p, "MRP", italic=True)
    run(p, ") of this input. The marginal cost is: “By how much do our "
           "costs go up if we use more of this input?” In the case of "
           "workers, this is the wage (")
    run(p, "w", italic=True)
    run(p, "); in the case of capital, it is the price of capital (")
    run(p, "p", italic=True)
    run(p, "K", italic=True, subscript=True)
    run(p, ").")

    p = body(doc)
    run(p, "Note that we assume constant input prices, so we do not need to "
           "worry about ")
    run(p, "w", italic=True)
    run(p, " going up if we hire more workers, or ")
    run(p, "p", italic=True)
    run(p, "K", italic=True, subscript=True)
    run(p, " going up (or down) if we buy more machines. We thus have the "
           "following two formulas for optimization:")

    # -- the two first-order conditions (aligned on a shared tab stop) ------
    for label, eq in (("For labor:", MRP_L + mrun("=") + W),
                      ("For capital:", MRP_K + mrun("=") + P_K)):
        p = para(doc, before=2 if label.endswith("labor:") else 0,
                 after=4 if label.endswith("labor:") else 8, left=0.35)
        p.paragraph_format.tab_stops.add_tab_stop(T.Inches(1.10))
        run(p, label + "\t")
        T.equation_inline(p, eq)

    body(doc, "Consequently,")
    equation(doc, mfrac(MRP_L, MRP_K) + mrun("=") + mfrac(W, P_K))

    # -- decompose MRP into MR x MP -----------------------------------------
    p = body(doc)
    run(p, "Recall from the Teaching Note on Hiring Decisions in the "
           "Short Run that ")
    T.equation_inline(p, MRP_L + mrun("=") + MR + mrun("⋅") + MP_L)
    run(p, ", i.e., the product of marginal revenue and the marginal product "
           "of labor. Similarly, for capital we have: ")
    T.equation_inline(p, MRP_K + mrun("=") + MR + mrun("⋅") + MP_K)
    run(p, ". For example, if we use one more unit of capital (e.g., a "
           "robot), then output goes up by ")
    T.equation_inline(p, MP_K)
    run(p, " units, and each extra unit of output, in turn, raises our "
           "revenues by ")
    T.equation_inline(p, MR)
    run(p, ". Thus, ")
    T.equation_inline(p, MRP_K)
    run(p, " tells us by how much our revenues increase if we use one more "
           "unit of capital. Replacing these two in the equation above "
           "yields:")

    equation(doc,
             mfrac(MR + mrun("⋅") + MP_L, MR + mrun("⋅") + MP_K)
             + mrun("=") + mfrac(W, P_K))

    p = body(doc)
    run(p, "Now ")
    T.equation_inline(p, MR)
    run(p, " cancels out. This leaves:")

    equation(doc, mfrac(MP_L, MP_K) + mrun("=") + mfrac(W, P_K))

    p = body(doc, after=6)
    run(p, "Re-arranging then yields the ")
    run(p, "bang-for-the-buck rule", bold=True)
    run(p, ":")

    # -- the hero rule, in the cream card -----------------------------------
    def pop(cell):
        cp = cell.add_paragraph()
        cp.paragraph_format.space_before = T.Pt(2)
        cp.paragraph_format.space_after = T.Pt(2)
        cp.paragraph_format.alignment = T.WD_ALIGN_PARAGRAPH.CENTER
        T.equation_inline(cp, mfrac(MP_L, W) + mrun("=") + mfrac(MP_K, P_K),
                          size=14)

    T.cream_card(doc, pop, width_in=6.5)

    # -- the three conditions, from deck slide 43 ---------------------------
    p = body(doc, before=12)
    run(p, "Three conditions come with the rule. It holds in the long run, "
           "when both labor and capital are flexible. It refers to a given "
           "output quantity ")
    run(p, "Q", italic=True)
    run(p, ": it tells us the cheapest input mix for producing that "
           "quantity, not how much to produce. And it assumes the input "
           "prices ")
    run(p, "w", italic=True)
    run(p, " and ")
    run(p, "p", italic=True)
    run(p, "K", italic=True, subscript=True)
    run(p, " are constant.")

    doc.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
