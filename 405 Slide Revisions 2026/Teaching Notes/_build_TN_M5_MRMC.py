"""Build: Module 5 - Teaching Note MR=MC.

Nico's original, restyled to the course theme.

COHERENCE CAVEAT.  Module 5 has no rebuilt deck yet (only Modules 1, 2, 3,
4 and 7 exist under `405 Slide Revisions 2026\\`), so this note could not be
checked against Module 5's own in-class material.  What it WAS checked
against, and now matches:
  - the Module 2 teaching note on marginal revenue (same 3-step recipe,
    same demand function), and
  - `Module 2 - Video Part Revised.pptx` slide 18, whose speaker notes
    carry this exact worked example: MC = $200, Q = 400, P = $300.
Re-check this note once Module 5's deck is rebuilt.

STATUS: reviewed and ACCEPTED in full on 2026-09-06, so this script now
emits the agreed text with no revision marks.  Any future proposal goes
back in as `ins` / `dele`.

WHAT CHANGED relative to the original (all accepted 2026-09-06)
  1. Q = 1600 - 4P  ->  Q = 1,600 - 4P  (thousands separator, as elsewhere).
  2. The three step names aligned to the Marginal Revenue note and to deck
     slide 18: "Invert the demand function so that P is on the left-hand
     side" -> "Calculate inverse demand (solve for P)", and so on.  The
     original set them in italics; they are now bold, like the MR note's.
  3. Total revenue written as TR, not PQ:  "PQ = 400Q - Q^2/4" -> "TR =
     ...", and "MR = dPQ/dQ" -> "MR = dTR/dQ".  This is the notation the MR
     note now uses and the deck's own.
  4. A cross-reference to the Marginal Revenue teaching note where the
     original said only "You should already know how to calculate MR".
  5. "You obtain MC by deriving total cost (TC) with respect to quantity:
     Let's assume..." -> "by taking the derivative of total cost (TC) with
     respect to quantity. Let's assume..."  "Deriving with respect to"
     means differentiating, which is not what "derive" means in English,
     and the sentence carried two colons.

NICO'S HAND-EDITS, ported 2026-09-06
  - "(see the teaching note on marginal revenue)" -> "(see Module 2 and
    the Teaching Note on Marginal Revenue)".  The Title Case is his, and
    is swept into the Bang-for-the-Buck note's cross-reference too.
  - "Solution:" added to the takeaway card.  Rendered as the card's bold
    navy prefix -- the house callout pattern and the decks' own styling
    for a "Solution:" line (his typed run inherited the body's non-bold
    navy).  Flagged in chat; trivial to unbold.

FORMATTING (untracked)
  - The last equation of the derivation is set in dark red C00000 -- the
    house rule for the line that delivers the answer -- and the concluding
    Q / P values move into a cream takeaway card.
  - Figure rebuilt natively: demand dark red C00000 labelled D (the
    original's lowercase italic d; the decks label a firm's demand D),
    MR solid concept blue rather than dashed, MC and the solution guides
    in dark green as in the original.

Run:  python _build_TN_M5_MRMC.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tn_theme as T
from _tn_theme import (CBLUE, DARKRED, NAVY, acr, body, equation,
                       heading, masthead, mfrac, mrun, msup, para,
                       run, text_width)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Module 5 - Teaching Note MR=MC.docx")

DKGREEN = "1B5E20"          # deck-sanctioned accent, used for the solution
MINUS = "−"

MR = acr("MR")
MC = acr("MC")
TR = acr("TR")
TC = acr("TC")


# ==========================================================================
# Figure: D, MR and a constant MC, with the profit-maximizing P and Q
# ==========================================================================
def build_figure():
    W, H = 5.90, 3.30
    f = T.Fig(W, H, name="Profit maximization where MR = MC")

    OX, OY = 0.95, 2.75
    PX, PY = 2.10, 3.70         # inches per 400 of P, per 1600 of Q
    XTIP, YTIP = 5.45, 0.40

    def qx(q):
        return OX + PY * q / 1600.0

    def py(p):
        return OY - PX * p / 400.0

    f.line(OX, OY, XTIP, OY, color=NAVY, w_pt=1.25, arrow=True, name="x-axis")
    f.line(OX, OY, OX, YTIP, color=NAVY, w_pt=1.25, arrow=True, name="y-axis")

    # demand and MR share the P-intercept; MR is twice as steep
    f.line(qx(0), py(400), qx(1600), py(0), color=DARKRED, w_pt=1.75, name="D")
    q_end = 900.0
    f.line(qx(0), py(400), qx(q_end), py(400 - q_end / 2.0),
           color=CBLUE, w_pt=1.75, name="MR")

    # constant marginal cost at 200
    f.line(OX, py(200), qx(1550), py(200), color=DKGREEN, w_pt=1.75,
           name="MC = 200")

    # the optimum: MR = MC at Q = 400, priced off demand at P = 300
    f.line(qx(400), py(200), qx(400), OY, color=DKGREEN, w_pt=1.0,
           dash="dash", name="Q* guide")
    f.line(OX, py(300), qx(400), py(300), color=DKGREEN, w_pt=1.0,
           dash="dash", name="P* guide")
    f.line(qx(400), py(300), qx(400), py(200), color=DKGREEN, w_pt=1.0,
           dash="dash", name="P*-to-Q* guide")
    f.dot(qx(400), py(200), d=0.075, color=DKGREEN, name="MR = MC")
    f.dot(qx(400), py(300), d=0.075, color=DKGREEN, name="point on demand")

    # axis titles, sized to the label and anchored at the arrow tips
    pt = [("P", dict(bold=True, italic=True, color=NAVY, size=12))]
    w = text_width(pt) + 0.06
    f.label(OX - 0.08 - w, YTIP - 0.11, pt, w=w, h=0.22, align="r",
            name="P axis title")
    qt = [("Q", dict(bold=True, italic=True, color=NAVY, size=12))]
    w = text_width(qt) + 0.06
    f.label(XTIP, OY + 0.04, qt, w=w, h=0.22, align="c", name="Q axis title")

    # solution labels, in the same green as their guides
    sol = dict(bold=True, italic=True, color=DKGREEN, size=11)
    f.label(OX - 0.07, py(300) - 0.12, [("P = $300", sol)], align="r", h=0.24,
            name="P* label")
    f.label(OX - 0.07, py(200) - 0.12, [("MC = $200", sol)], align="r", h=0.24,
            name="MC label")
    f.label(qx(400), OY + 0.05, [("Q = 400", sol)], align="c", h=0.24,
            name="Q* label")

    # curve labels
    f.label(4.30, py(400 * (1 - (4.30 - OX) / PY)) - 0.30,
            [("D", dict(bold=True, color=DARKRED, size=12))],
            align="c", h=0.24, name="D label")
    f.label(qx(q_end) + 0.26, py(400 - q_end / 2.0) - 0.10,
            [("MR", dict(bold=True, color=CBLUE, size=12))],
            align="c", h=0.24, name="MR label")

    return f


# ==========================================================================
def step(doc, number, bold_name, tail_runs):
    """A numbered recipe step: bold name, then its parenthetical."""
    p = para(doc, before=6, after=6, left=0.35, hang=0.35)
    run(p, number + "\t")
    run(p, bold_name, bold=True)
    for t in tail_runs:
        run(p, t)
    return p


def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)
    masthead(doc, "Teaching Note – Module 5: MR = MC")

    # ---------------- Preliminaries ---------------------------------------
    heading(doc, "Preliminaries", before=8)

    p = body(doc)
    run(p, "This note explains how to maximize profits in the generic case "
           "of a firm facing a downward sloping, linear demand curve. Such a "
           "firm has a simple decision to make: choose a spot on the demand "
           "curve (i.e., a combination of price charged and quantity sold) "
           "such that its profits are maximized. This choice entails "
           "marginal decision making, picking ")
    run(p, "Q", italic=True)
    run(p, " to set marginal revenue equal to marginal cost: ")
    run(p, "MR=MC", bold=True, italic=True)
    run(p, ".")

    p = body(doc)
    run(p, "To set ")
    run(p, "MR=MC", italic=True)
    run(p, ", you need to know both ")
    run(p, "MR", italic=True)
    run(p, " and ")
    run(p, "MC", italic=True)
    run(p, ".")

    p = body(doc)
    run(p, "You should already know how to calculate ")
    run(p, "MR", italic=True)
    run(p, " (see Module 2 and the Teaching Note on Marginal Revenue). I "
           "will briefly repeat the steps starting from a specific demand "
           "function (from the lecture slides).")

    body(doc, "We start with the following demand:")
    equation(doc, mrun("Q=1,600-4P"))

    # ---------------- the three steps -------------------------------------
    step(doc, "1)", "Calculate inverse demand", [" (solve for P):"])
    equation(doc, mrun("P=400-") + mfrac(mrun("1"), mrun("4")) + mrun("Q"))

    step(doc, "2)", "Calculate total revenue", [" (multiply P(Q) by Q):"])
    equation(doc,
             TR + mrun("=400Q-") + mfrac(mrun("1"), mrun("4"))
             + msup(mrun("Q"), mrun("2")))

    step(doc, "3)", "Compute marginal revenue from total revenue",
         [" (MR = dTR/dQ):"])
    equation(doc,
             MR + mrun("=")
             + mfrac(mrun("d") + TR, mrun("d") + mrun("Q"))
             + mrun("=400-") + mfrac(mrun("1"), mrun("2")) + mrun("Q"))

    # ---------------- marginal cost ---------------------------------------
    p = body(doc)
    run(p, "Next, you need information on marginal cost. You obtain ")
    run(p, "MC", italic=True)
    run(p, " by taking the derivative of total cost (")
    run(p, "TC", italic=True)
    run(p, ") with respect to quantity. Let's assume the following total "
           "cost: ")
    T.equation_inline(p, TC + mrun("=10,000+200Q"))
    run(p, ". Then:")

    equation(doc, MC + mrun("=") + mfrac(mrun("d") + TC, mrun("d") + mrun("Q"))
             + mrun("=200"))

    body(doc, "Note that we're using the simplest case where marginal cost "
              "is constant.")

    # ---------------- The Final Step --------------------------------------
    heading(doc, "The Final Step")

    p = body(doc, after=4)
    run(p, "Now we set ")
    run(p, "MR", italic=True)
    run(p, " equal to ")
    run(p, "MC", italic=True)
    run(p, ":")

    equation(doc, MR + mrun("=") + MC, before=2, after=2)
    equation(doc,
             mrun("⇔") + mrun("400-") + mfrac(mrun("1"), mrun("2"))
             + mrun("Q=200"), before=0, after=2)
    equation(doc, mrun("⇔") + mrun("Q=400"), before=0, after=10)

    p = body(doc)
    run(p, "So now we know that maximizing profits entails setting ")
    run(p, "Q", italic=True)
    run(p, "=400. What about the price ")
    run(p, "P", italic=True)
    run(p, "? Well, that's easy to find: our price-quantity combination "
           "must be located on the (inverse) demand curve (i.e., the one "
           "that we plot, with ")
    run(p, "P", italic=True)
    run(p, " on the y-axis). So from the (inverse) demand we derived "
           "earlier, we find:")

    # the line that delivers the answer: dark red, per the house rule
    equation(doc, T.color_omml(
        mrun("P=400-") + mfrac(mrun("1"), mrun("4"))
        + mrun("Q=400-100=$300"), DARKRED))

    # ---------------- takeaway --------------------------------------------
    def pop(cell):
        cp = cell.add_paragraph()
        cp.paragraph_format.space_before = T.Pt(1)
        cp.paragraph_format.space_after = T.Pt(1)
        cp.paragraph_format.alignment = T.WD_ALIGN_PARAGRAPH.CENTER
        run(cp, "Solution:  ", bold=True, color=NAVY, size=11)
        run(cp, "Maximizing profits, here, entails setting ", color=NAVY,
            size=11)
        run(cp, "Q = 400 units", bold=True, color=NAVY, size=11)
        run(cp, " and ", color=NAVY, size=11)
        run(cp, "P = $300", bold=True, color=NAVY, size=11)
        run(cp, ".", color=NAVY, size=11)

    T.cream_card(doc, pop, width_in=6.5)

    body(doc, "Graphically:")
    build_figure().place(doc, before=4, after=8, scale=6.40 / 5.90)

    doc.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
