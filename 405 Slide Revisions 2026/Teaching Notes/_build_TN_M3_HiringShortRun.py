"""Build: Module 3 - Teaching Note Hiring Decisions in the Short Run.

Nico's original, restyled to the course theme and checked against the
Module 3 deck (`Module 3 - Revised.pptx`): slide 24 (MRPL core concept),
slide 28 (the optimal hiring rule) and Video 4, slides 32-38 (wage
searchers).

STATUS: reviewed and ACCEPTED in full on 2026-09-06, so this script now
emits the agreed text with no revision marks.  Any future proposal goes
back in as `ins` / `dele`.

WHAT CHANGED relative to the original (all accepted 2026-09-06)
  1. ECONOMICS.  The wage-searcher rule read "An additional worker should
     be hired if MRPL >= w".  For a wage searcher the marginal labor cost
     EXCEEDS the wage -- hiring one more bids up the wage for everyone --
     which is the whole point of deck slides 32-38 (the 3rd designer costs
     $3M, not the $2M she is paid).  The rule has to be MRPL >= MCL, and
     the note's own preceding sentence already says MCL is not constant
     here.  Proposed as a tracked change so it is visible; flagged in chat.
  2. "offset by their contribution to costs" -> "her/his contribution".
     House style avoids gender-neutral "their".
  3. New sentence after the MRPL = w rule, naming the regime: this is the
     SHORT-run rule, capital fixed.  The note says "short run" in its
     title and first line but never ties it to the rule itself, and the
     companion bang-for-the-buck note is the long-run counterpart.

FORMATTING (untracked)
  - Acronyms upright with italic subscripts (MRP, MP, MC upright; L
    italic), per the OMML convention.
  - The two underlined bold headings become navy bold section headings.
  - The original's empty spacer paragraphs are dropped; spacing is set
    on the paragraphs themselves.
  - "wage takers" / "wage searchers" keep the original's italics; the
    terms match deck slides 32 and 34.

Run:  python _build_TN_M3_HiringShortRun.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tn_theme as T
from _tn_theme import (acr, body, equation, heading, ins, masthead,
                       mfrac, mrun, para, run)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(
    HERE, "Module 3 - Teaching Note Hiring Decisions in the Short Run.docx")

MRP_L = acr("MRP", "L")
MC_L = acr("MC", "L")
MP_L = acr("MP", "L")
MR = acr("MR")
MC = acr("MC")
TR = acr("TR")
W = mrun("w")


def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)
    masthead(doc, "Teaching Note – Module 3: Hiring Decisions in the Short Run")

    p = body(doc)
    run(p, "This note discusses how firms should optimally choose how much "
           "of an input to use in the ")
    run(p, "short run", bold=True)
    run(p, ".")

    # ---------------- MRP -------------------------------------------------
    heading(doc, "Marginal revenue product (MRP)", before=12)

    p = body(doc)
    run(p, "MRP is the contribution to revenues from adding one unit of a ")
    run(p, "factor of production", italic=True)
    run(p, ". For example, consider the standard short-run scenario where "
           "capital is fixed while labor (")
    run(p, "L", italic=True)
    run(p, ") is flexible. Then MRP")
    run(p, "L", subscript=True)
    run(p, " tells us how total revenues change if the firm uses more "
           "labor. This is the value to the company of one additional unit "
           "of labor. The marginal revenue product of labor is defined as:")

    equation(doc, MRP_L + mrun("=") + mfrac(mrun("d") + TR, mrun("d") + mrun("L")))

    p = body(doc)
    run(p, "We can rewrite this as (adding ")
    run(p, "dL", italic=True)
    run(p, " in the numerator and denominator):")

    # two-line derivation, second line aligned under the first's "="
    equation(doc,
             MRP_L + mrun("=")
             + mfrac(mrun("d") + mrun("Q"), mrun("d") + mrun("L"))
             + mrun("∙")
             + mfrac(mrun("d") + TR, mrun("d") + mrun("Q")),
             after=2)
    equation(doc, mrun("=") + MP_L + mrun("∙") + MR, before=0, after=10)

    p = body(doc)
    run(p, "Intuitively, this is how much more output the firm gets from "
           "one additional unit of labor (MP")
    run(p, "L", subscript=True)
    run(p, ") multiplied by how much revenue the firm gets from one more "
           "unit of output (MR).")

    # ---------------- MCL and the hiring rule -----------------------------
    # heading with a real subscript rather than a Unicode lookalike
    hp = para(doc, before=14, after=7, keep_next=True)
    run(hp, "Marginal labor cost (MC", bold=True, color=T.NAVY, size=13)
    run(hp, "L", bold=True, color=T.NAVY, size=13, subscript=True)
    run(hp, ") and optimal hiring decisions", bold=True, color=T.NAVY, size=13)

    p = body(doc)
    run(p, "The marginal labor cost tells us by how much total costs change "
           "when adding one unit of labor. Typically, we think of a unit of "
           "labor as a worker. Thus, MC")
    run(p, "L", subscript=True)
    run(p, " is the cost to the company of one additional worker. For the "
           "problems we consider in our class, we typically consider firms "
           "that are ")
    run(p, "wage takers", italic=True, underline=True)
    run(p, " (i.e., firms that can hire as much labor they would like at "
           "the constant market wage). In these cases, MC")
    run(p, "L", subscript=True)
    run(p, " is simply the wage (")
    run(p, "w", italic=True)
    run(p, "). As we saw in class, the optimal choice of labor then takes "
           "place where ")
    T.equation_inline(p, MRP_L + mrun("=") + W)
    run(p, ". This means that the firm should add more workers until the "
           "worker's contribution to revenue is offset by her/his "
           "contribution to costs. The intuition is similar to the ")
    T.equation_inline(p, MR + mrun("=") + MC)
    run(p, " rule for profit maximization")
    ins(p, " (see the Teaching Note on MR = MC)")
    run(p, ".")
    # regime, per deck slide 24 ("In the short run (capital K fixed)")
    run(p, " This is the short-run rule: capital is fixed, so labor is the "
           "only input being adjusted.")
    ins(p, " The long-run counterpart, where capital can be adjusted too, "
           "is in the Teaching Note on the Bang-for-the-Buck Rule.")

    p = body(doc)
    run(p, "If ")
    T.equation_inline(p, MRP_L + mrun("<") + W)
    run(p, ", adding an additional worker is a bad decision because "
           "revenues grow less than (labor) costs, thus lowering profits. "
           "If ")
    T.equation_inline(p, MRP_L + mrun(">") + W)
    run(p, ", adding an additional worker increases total profit.")

    p = body(doc)
    run(p, "In some instances, the firm's choice of labor may not be "
           "continuous. For example, the firm may be deciding between 5 and "
           "6 workers (a discrete choice), instead of 5 and 5.8 workers. In "
           "this case, the firm should hire the 6")
    run(p, "th", superscript=True)
    run(p, " worker if ")
    T.equation_inline(p, MRP_L + mrun("≥") + MC_L)
    run(p, ".")

    # ---------------- wage searchers --------------------------------------
    p = body(doc)
    run(p, "We also briefly discussed the case of ")
    run(p, "wage searchers", italic=True, underline=True)
    run(p, ", who may have to raise their wage offers in order to attract "
           "more workers (for example, highly specialized workers such as "
           "surgeons or top designers). In these cases, the marginal labor "
           "cost MC")
    run(p, "L", subscript=True)
    run(p, " is not constant. An additional worker should be hired if ")
    # For a wage searcher MCL > w, so the test is against MCL, not w
    # -- see deck slides 32-38.
    T.equation_inline(p, MRP_L + mrun("≥") + MC_L)
    run(p, ".")

    body(doc, "The reason is that hiring one more worker bids up the wage "
              "for everyone already employed, so the full marginal cost of "
              "that worker is larger than the wage she or he is paid. In the "
              "example we did in class, the third star designer is paid $2M, "
              "but the two existing designers each get a $0.5M raise, so the "
              "marginal labor cost of the hire is $3M rather than $2M.")

    doc.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
