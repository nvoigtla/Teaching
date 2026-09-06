"""Build: Module 2 - Teaching Note Marginal Revenue.

Nico's 2021 teaching note, restyled to the course theme and aligned with
the Module 2 video deck (`Module 2 - Video Part Revised.pptx`, slides
15-21, whose slide 18 carries the "Teaching Note: Marginal Revenue"
pointer).

STATUS: the round of tracked changes proposed on 2026-09-06 was reviewed
and ACCEPTED in full, so this script now emits the agreed text with no
revision marks.  Any future proposal goes back in as `ins` / `dele`.

WHAT CHANGED relative to the 2021 original (all accepted 2026-09-06)
  1. Q = 1600 - 4P  ->  Q = 1,600 - 4P          (deck writes the separator)
  2. Step names aligned to slide 18:
       "Invert your demand function"   -> "Calculate inverse demand (solve for P)"
       "Write down total revenue"      -> "Calculate total revenue (multiply P(Q) by Q)"
       "Find Marginal Revenue"         -> "Compute marginal revenue from total
                                           revenue (MR = dTR/dQ)"
  3. MR notation: the fraction  DPQ/DQ  ->  dTR/dQ  (deck notation; the old
     numerator also said PQ where it meant TR)
  4. "the TA review videos on Canvas" -> "the TA Math Review videos"
  5. New closing paragraph after the figure, tying MR to the elasticity
     ranges and to MR = 0 at maximum total revenue (the point slide 21
     makes).

NICO'S HAND-EDITS, ported 2026-09-06
  - Step 1's aside gains a closing sentence naming P(Q), with P, Q and
    P(Q) set italic (see `step1_aside`).
  - "the TA's Math Review videos" -> "the TA Math Review videos".  His
    comment: a different TA recorded them, so the possessive is wrong.
    This wording holds THROUGHOUT, in every note and deck.
  - Footer reduced to a lone centred page number, no course line and no
    term, so the note is reusable in a later year unchanged.
  - A closing cross-reference to the Demand Elasticity and Total Revenue
    note, with that note's title bold, as he set it (see `main`).  It
    replaces nothing; the figure paragraph above it is unchanged.
  - NOT adopted: a pPrChange on step 1's aside that dropped `jc=both` and
    `spacing after=180` from that one paragraph.  It reads as collateral
    from typing at the paragraph end -- adopting it would leave a single
    unjustified body paragraph among justified ones -- so the paragraph
    keeps the deck-wide body formatting.  Flagged in chat 2026-09-06.

FIGURE.  Native, editable Word shapes.  Demand dark red C00000 and MR
concept blue 0070C0 per the 2026-08-30 palette rule; the axes carry the
example's numbers (400 / 800 / 1600) so the figure matches slide 21
instead of being schematic.

Run:  python _build_TN_M2_MarginalRevenue.py
"""

import os
import sys

from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tn_theme as T
from _tn_theme import (CBLUE, CREAM, DARKRED, NAVY, body, equation, heading,
                       ins, masthead, mfrac, mrun, msup, para, run,
                       text_width)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Module 2 - Teaching Note Marginal Revenue.docx")

MINUS = "−"    # true minus sign
INF = "∞"


# ==========================================================================
# The figure: linear demand, MR twice as steep, elasticity along demand
# ==========================================================================
def build_figure():
    W, H = 5.90, 3.15
    f = T.Fig(W, H, name="Marginal revenue and the elasticity of demand")

    # -- plot frame --------------------------------------------------------
    OX, OY = 0.85, 2.65          # origin
    PX, PY = 2.00, 3.80          # inches per 400 of P, per 1600 of Q
    XTIP, YTIP = 5.35, 0.42      # arrow tips

    def qx(q):
        return OX + PY * q / 1600.0

    def py(p):
        return OY - PX * p / 400.0

    f.line(OX, OY, XTIP, OY, color=NAVY, w_pt=1.25, arrow=True, name="x-axis")
    f.line(OX, OY, OX, YTIP, color=NAVY, w_pt=1.25, arrow=True, name="y-axis")

    # -- curves ------------------------------------------------------------
    # Demand: P = 400 - Q/4, from the P-intercept to Q = 1600
    f.line(qx(0), py(400), qx(1600), py(0),
           color=DARKRED, w_pt=1.75, name="D")
    # MR = 400 - Q/2: same intercept, twice as steep, carried past MR = 0
    # so the negative range is visible (as in the original note's figure)
    q_end = 940.0
    f.line(qx(0), py(400), qx(q_end), py(400 - q_end / 2.0),
           color=CBLUE, w_pt=1.75, name="MR")

    # -- unit-elastic point and its drop line -------------------------------
    f.line(qx(800), py(200), qx(800), OY,
           color=NAVY, w_pt=1.0, dash="dash", name="Q=800 guide")
    f.dot(qx(800), py(200), d=0.075, color=NAVY, name="unit-elastic point")

    # -- axis titles: box sized to the label, anchored to the arrow tip -----
    ptitle = [("P", dict(bold=True, italic=True, color=NAVY, size=12))]
    w = text_width(ptitle) + 0.06
    f.label(OX - 0.08 - w, YTIP - 0.11, ptitle, w=w, h=0.22, align="r",
            name="P axis title")

    qtitle = [("Q", dict(bold=True, italic=True, color=NAVY, size=12))]
    w = text_width(qtitle) + 0.06
    f.label(XTIP, OY + 0.04, qtitle, w=w, h=0.22, align="c",
            name="Q axis title")

    # -- tick labels --------------------------------------------------------
    tick = dict(color=NAVY, size=10)
    f.label(OX - 0.07, py(400) - 0.11, [("400", tick)], align="r", h=0.22,
            name="tick 400")
    f.label(qx(800), OY + 0.05, [("800", tick)], align="c", h=0.22,
            name="tick 800")
    f.label(qx(1600), OY + 0.05, [("1600", tick)], align="c", h=0.22,
            name="tick 1600")

    # -- curve labels -------------------------------------------------------
    # "D" sits 0.20" above the demand line, at the line's own y for that x
    dlx = 4.15
    f.label(dlx, (OY - PX * (1.0 - (dlx - OX) / PY)) - 0.32,
            [("D", dict(bold=True, color=DARKRED, size=12))],
            align="c", h=0.24, name="D label")
    f.label(3.34, py(400 - q_end / 2.0) - 0.07,
            [("MR", dict(bold=True, color=CBLUE, size=12))],
            align="c", h=0.24, name="MR label")

    # -- elasticity callouts ------------------------------------------------
    def ed(tail):
        return [("E", dict(bold=True, color=NAVY, size=10)),
                ("D", dict(bold=True, color=NAVY, size=10, subscript=True)),
                (tail, dict(bold=True, color=NAVY, size=10))]

    def callout(x, y, tail, name):
        runs = ed(tail)
        w = text_width(runs) + 0.30
        h = 0.30
        f.label(x, y, runs, w=w, h=h, align="c", fill=CREAM, border=NAVY,
                name=name, pad=0.04)
        return x - w / 2, y, w, h

    # perfectly elastic at the P-intercept
    bx, by, bw, bh = callout(1.62, 0.28, " = " + MINUS + INF, "callout Ed=-inf")
    f.line(bx + 0.02, by + bh / 2, qx(0) + 0.07, py(400) - 0.03,
           color=NAVY, w_pt=0.75, arrow=True, name="ptr Ed=-inf")

    # unit elastic at the midpoint of demand
    bx, by, bw, bh = callout(3.62, 1.00, " = " + MINUS + "1", "callout Ed=-1")
    f.line(bx + 0.02, by + bh / 2, qx(800) + 0.10, py(200) - 0.04,
           color=NAVY, w_pt=0.75, arrow=True, name="ptr Ed=-1")

    # perfectly inelastic at the Q-intercept
    bx, by, bw, bh = callout(5.10, 1.72, " = 0", "callout Ed=0")
    f.line(bx + 0.10, by + bh, qx(1600) + 0.05, OY - 0.10,
           color=NAVY, w_pt=0.75, arrow=True, name="ptr Ed=0")

    return f


# ==========================================================================
# Equations
# ==========================================================================
def eq_demand(doc):
    equation(doc, mrun("Q=1,600-4P"))


def eq_inverse(doc):
    equation(doc, mrun("P=400-") + mfrac(mrun("1"), mrun("4")) + mrun("Q"))


def eq_tr(doc):
    equation(doc,
             mrun("TR", italic=False) + mrun("=P⋅Q=400⋅Q-")
             + mfrac(mrun("1"), mrun("4")) + msup(mrun("Q"), mrun("2")))


def eq_mr(doc):
    equation(doc,
             mrun("MR", italic=False) + mrun("=")
             + mfrac(mrun("d") + mrun("TR", italic=False),
                     mrun("d") + mrun("Q"))
             + mrun("=400-") + mfrac(mrun("1"), mrun("2")) + mrun("Q"))


# ==========================================================================
# Document
# ==========================================================================
def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)                      # centred page number only
    masthead(doc, "Teaching Note – Module 2: Marginal Revenue")

    # ---------------- Preliminaries ---------------------------------------
    heading(doc, "Preliminaries", before=8)

    p = body(doc)
    run(p, "We are building our toolkit to maximize profits in the generic "
           "case of a firm facing a downward sloping, linear demand curve. "
           "Such a firm has a simple decision to make: choose a spot on the "
           "demand curve, i.e. a combination of price charged and quantity "
           "sold. We have already seen the marginal-decision rule MR=MC: the "
           "firm should produce and sell a quantity ")
    run(p, "Q", italic=True)
    run(p, " such that its marginal revenue equals its marginal cost. This "
           "note focuses on the marginal revenue side. Marginal cost will be "
           "covered in Module 3.")
    ins(p, " The two sides are brought together in the Teaching Note on "
           "MR = MC.")

    p = body(doc)
    run(p, "What is marginal revenue? In words, it is the change in revenues "
           "when you sell one more unit of the good. To get there, we start "
           "from total revenue. Let's write down a specific demand function "
           "(from the lecture slides")
    ins(p, "; in practice you would estimate it from data, as in the "
           "Teaching Note on Regressions")
    run(p, ") and work through an example using this demand function:")

    eq_demand(doc)

    # ---------------- The Recipe ------------------------------------------
    heading(doc, "The Recipe")

    body(doc, "We have to follow three steps to derive the equation for MR. "
              "You can apply this recipe whenever you face this type of "
              "problem (that will be often, throughout the MBA program!):")

    # -- Step 1 -------------------------------------------------------------
    p = step(doc, "1)")
    run(p, "Calculate inverse demand", bold=True)
    run(p, " (solve for P).")

    eq_inverse(doc)
    step1_aside(doc)

    # -- Step 2 -------------------------------------------------------------
    p = step(doc, "2)")
    run(p, "Calculate total revenue", bold=True)
    run(p, " (multiply P(Q) by Q). What's total revenue? It is ")
    run(p, "P", italic=True)
    run(p, " x ")
    run(p, "Q", italic=True)
    run(p, ". So:")

    eq_tr(doc)

    p = body(doc)
    run(p, "(Here I just multiplied the previous equation by ")
    run(p, "Q", italic=True)
    run(p, ": easy!)")

    # -- Step 3 -------------------------------------------------------------
    p = step(doc, "3)", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, "Compute marginal revenue from total revenue", bold=True)
    run(p, " (MR = dTR/dQ). That's the hardest step. It entails taking the "
           "derivative of TR with respect to ")
    run(p, "Q", italic=True)
    # "TA Math Review videos", never "the TA's" -- a different TA recorded
    # them (Nico, 2026-09-06).  Same wording in every note and deck.
    run(p, ". If you are unfamiliar with the underlying calculus, please "
           "watch the TA Math Review videos. All you need to know is that "
           "the derivative of ")
    run(p, "Q", italic=True)
    run(p, " with respect to ")
    run(p, "Q", italic=True)
    run(p, " is 1, and that the derivative of ")
    run(p, "Q", italic=True)
    run(p, "2", italic=True, superscript=True)
    run(p, " with respect to ")
    run(p, "Q", italic=True)
    run(p, " is 2")
    run(p, "Q", italic=True)
    run(p, ":")

    eq_mr(doc)

    # ---------------- Closing ---------------------------------------------
    p = body(doc)
    run(p, "So now we have MR. You will notice that MR has ")
    run(p, "twice the slope", italic=True)
    run(p, " of the inverse demand function. The slope of the inverse demand "
           "function is -1/4, and the slope of MR is -1/2. When demand is "
           "linear, MR is always twice as steep as (inverse) demand. "
           "Graphically, we obtain:")

    # 5.90" design width scaled to 6.40" -- the full text column
    build_figure().place(doc, before=8, after=10, scale=6.40 / 5.90)

    p = body(doc)
    run(p, "The figure also ties MR to the price elasticity of demand, ")
    run(p, "E")
    run(p, "D", subscript=True)
    run(p, ". On the upper half of the demand curve MR is positive and "
           "demand is elastic, so cutting the price raises total revenue. On "
           "the lower half MR is negative and demand is inelastic, so cutting "
           "the price lowers total revenue. The two halves meet where MR = 0. "
           "That is the unit-elastic point, ")
    run(p, "E")
    run(p, "D", subscript=True)
    run(p, " = " + MINUS + "1, and it is where total revenue reaches its "
           "maximum. In our example MR = 400 " + MINUS + " Q/2, so MR = 0 at "
           "Q = 800 – half of the 1,600 units that would be demanded at a "
           "price of zero.")

    # Nico's closing cross-reference, added by hand 2026-09-06.  Set plain:
    # cross-references between the notes are non-bold throughout
    # (2026-09-06, Nico) -- the same in the MR=MC and Bang-for-the-Buck notes.
    body(doc, "We will elaborate on the features displayed in the figure in "
              "the Teaching Note on Demand Elasticity and Total Revenue.")

    doc.save(OUT)
    print("wrote", OUT)


def step1_aside(doc):
    """Step 1's parenthetical, with Nico's closing sentence on P(Q).

    Hand-edit 2026-09-06: the sentence " This gives P as a function of Q,
    which we denote by P(Q)." was added, with P / Q / P(Q) italic.
    """
    p = body(doc)
    run(p, "(If you're wondering what I did here, I took our demand "
           "function, put P on the left side, Q on the right side and "
           "divided both sides by 4). This gives ")
    run(p, "P", italic=True)
    run(p, " as a function of ")
    run(p, "Q", italic=True)
    run(p, ", which we denote by ")
    run(p, "P(Q)", italic=True)
    run(p, ".")
    return p


def step(doc, number, align=None):
    """A numbered recipe item: hanging indent, the original's '1)' marker."""
    p = para(doc, before=6, after=6, left=0.35, hang=0.35, align=align)
    run(p, number + "\t")
    return p


if __name__ == "__main__":
    main()
