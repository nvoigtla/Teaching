"""Build: Module 2 - Teaching Note Regressions.

Nico's original, restyled to the course theme and checked against
`Module 2 - Video Part Revised.pptx` Video 3 (Demand Estimation): slide 29
(econometric estimates), 30 (OLS and the meaning of a and b), 33 (the
least-squares algorithm), 34 (the airline regression results), 36 (the
elasticity worked off that regression) and 39 (multivariate regression).

THE AIRLINE COEFFICIENTS NOW FOLLOW THE DECK (Nico, 2026-09-06)
The note used to read Q = 478.55 - 1.63 P (table: intercept 478.547, own
price -1.633) while deck slide 34 reads Q = 478.95 - 1.64 P, with slides
35/36 building the ED = -0.92 answer on it.  The deck is authoritative, so
the note now carries the deck's numbers.

What that means for the Excel table.  The original table was internally
consistent: 478.547 / 88.039 = 5.436 and 478.547 +- 2.1448 x 88.039 =
[289.722, 667.372] both reproduce the printed values exactly.  So the two
t-stats and the four confidence-interval bounds are RECOMPUTED from the new
coefficients and the unchanged standard errors, which is arithmetic, not
invention:
    t = coefficient / standard error
    CI = coefficient +- t(0.975, 14 df) x standard error,  t_crit = 2.144787
giving intercept t = 5.440, CI [290.125, 667.775]; own price t = -4.469,
CI [-2.427, -0.853].  P-values are unchanged at that precision.
Multiple R, R Square, Adjusted R Square, Standard Error, Observations and
the whole ANVOA block are LEFT ALONE -- they depend on the fit, and moving
the slope from -1.633 to -1.640 shifts them in the fourth decimal at most.
If the original Excel workbook is still around, refresh those from it
rather than from this note.  The multivariate table is untouched: deck
slide 39 gives only the model's form, no numbers, so nothing contradicts it.

STATUS: reviewed and ACCEPTED in full on 2026-09-06 (two rounds: the
restyle plus the cross-reference), so this script now
emits the agreed text with no revision marks.  Any future proposal goes
back in as `ins` / `dele`.

WHAT CHANGED relative to the original (all accepted 2026-09-06)
  0. The airline coefficients and everything derived from them, per the
     paragraph above: the displayed demand equation, the four numbers in
     the prose that quote it, and eight cells of the first table.
  1. "which I have circled in red" -> "which I have highlighted".  The red
     ovals were seven floating Word shapes sitting on top of the old table;
     the highlight is now the deck's own table-cell device (pale gold fill,
     bold navy figure), which cannot drift out of register with the cells.
  2. The multivariate demand curve's variable names aligned to deck slide
     39, which defines the model as Q = a + bP + cPc + dY:
     "+ 1.04 CompetPrice + 3.09 Income" -> "+ 1.04 Pc + 3.09 Y", with the
     names still spelled out in the surrounding prose.

FORMATTING (untracked)
  - The two Excel outputs become native course-styled tables (navy header,
    cream/white banding, thin light-gray borders) instead of the pasted
    grid; the seven red ovals are replaced by the cell highlight above.
  - The two figures stay as images: they are Nico's own charts and the
    underlying airline data is not in the note, so rebuilding them
    natively would mean inventing data points.  They now sit rounded with
    a soft shadow.  NOTE: `reg_leastsquares.png` has Excel's grey "Chart
    Area" tooltip baked into it, bottom left of the plot -- deck slide 33
    is a clean native rebuild of the same figure if that one should be
    ported instead.
  - Footnote 1 becomes a cream "Note:" card, as in the elasticity note.

NICO'S HAND-EDITS, ported 2026-09-06
  - A new opening note, which his comment asked to be set "within the
    yellow note box", so it is now the cream callout above "The Setting".
    His bold emphasis on the first claim is kept.
  - "regression out" set as "regression output" -- a dropped word, not a
    rewording; reported in chat.

Images are build inputs in `_source_images/` -- do not delete them.

Run:  python _build_TN_M2_Regressions.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tn_theme as T
from _tn_theme import (NAVY, body, equation, heading, masthead,
                       mrun, para, run)

HERE = os.path.dirname(os.path.abspath(__file__))
IMG = os.path.join(HERE, "_source_images")
OUT = os.path.join(HERE, "Module 2 - Teaching Note Regressions.docx")

# ---------------------------------------------------------------- table 1
# Excel's simple-regression output, verbatim from the original.
REG1 = [
    ["Regression Statistics", "", "", "", "", "", ""],
    ["Multiple R", "0.765", "", "", "", "", ""],
    ["R Square", "0.586", "", "", "", "", ""],
    ["Adjusted R Square", "0.556", "", "", "", "", ""],
    ["Standard Error", "18.607", "", "", "", "", ""],
    ["Observations", "16", "", "", "", "", ""],
    ["ANOVA", "", "", "", "", "", ""],
    ["", "df", "SS", "MS", "F", "Significance F", ""],
    ["Regression", "1", "6858.828", "6858.828", "19.810", "0.001", ""],
    ["Residual", "14", "4847.152", "346.225", "", "", ""],
    ["Total", "15", "11705.979", "", "", "", ""],
    ["", "Coefficients", "Standard Error", "t Stat", "P-value",
     "Lower 95%", "Upper 95%"],
    # coefficients per deck slide 34; t-stats and CI bounds recomputed from
    # the unchanged standard errors (see the docstring for the arithmetic)
    ["Intercept", "478.950", "88.039", "5.440", "0.000", "290.125", "667.775"],
    ["Own Price", "-1.640", "0.367", "-4.469", "0.001", "-2.427", "-0.853"],
]
# the cells the original circled in red: R Square, both coefficients,
# both t-stats, and the two confidence-interval columns
HL1 = ([(2, 1)]
       + [(12, 1), (13, 1)]
       + [(12, 3), (13, 3)]
       + [(12, 5), (13, 5), (12, 6), (13, 6)])
# Widths measured against the widest label in each column at 9 pt Calibri
# italic plus the 0.09" of cell margin: "Significance F" needs 0.91", so the
# column is 0.95".  Total 6.40", inside the 6.50" text width.
W1 = [1.30, 0.92, 1.02, 0.68, 0.70, 0.95, 0.83]

# ---------------------------------------------------------------- table 2
REG2 = [
    ["Regression Statistics", "", "", "", "", "", ""],
    ["Multiple R", "0.881", "", "", "", "", ""],
    ["R Square", "0.776", "", "", "", "", ""],
    ["Adjusted R Square", "0.721", "", "", "", "", ""],
    ["Standard Error", "14.766", "", "", "", "", ""],
    ["Observations", "16", "", "", "", "", ""],
    ["ANOVA", "", "", "", "", "", ""],
    ["", "df", "SS", "MS", "F", "Significance F", ""],
    ["Regression", "3", "9089.599", "3029.866", "13.896", "0.000", ""],
    ["Residual", "12", "2616.380", "218.032", "", "", ""],
    ["Total", "15", "11705.979", "", "", "", ""],
    ["", "Coefficients", "Standard Error", "t Stat", "P-value",
     "Lower 95%", "Upper 95%"],
    ["Intercept", "28.844", "174.665", "0.165", "0.872", "-351.719", "409.407"],
    ["Own Price", "-2.124", "0.340", "-6.238", "0.000", "-2.865", "-1.382"],
    ["Compet Price", "1.035", "0.467", "2.218", "0.047", "0.018", "2.051"],
    ["Income", "3.089", "0.999", "3.093", "0.009", "0.913", "5.266"],
]
HL2 = [(2, 1)] + [(r, 1) for r in (12, 13, 14, 15)] \
      + [(r, 3) for r in (12, 13, 14, 15)]

LABEL_ROWS = (0, 6, 7, 11)


def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)
    masthead(doc, "Teaching Note – Module 2: Regressions")

    # Nico's opening note, added by hand 2026-09-06, moved into the cream
    # card per his comment ("Make this a note within the yellow note box").
    # His bold emphasis on the first claim is kept; "regression out" is set
    # as "regression output" -- a dropped word, reported in chat.
    T.callout(doc, "First, a note:",
              [("You will not have to run regressions in the Managerial "
                "Economics class", dict(bold=True)),
               (". We will provide you with regression output that you need "
                "to interpret. This note is to familiarize yourself with "
                "regressions.", {})],
              before=2, after=10)

    # ---------------- The Setting -----------------------------------------
    heading(doc, "The Setting", before=8)

    p = body(doc)
    run(p, "Our goal is to estimate the demand function that a company "
           "faces. Recall that the general form of a linear demand function "
           "is:")
    equation(doc, mrun("Q=a+bP"))

    p = body(doc)
    run(p, "In particular, we want to estimate ")
    run(p, "a", italic=True)
    run(p, " (the intercept) and ")
    run(p, "b", italic=True)
    run(p, " (the slope).")

    body(doc, "Suppose you are given data coming from price "
              "experimentations. For instance, you are given the number of "
              "airline seats sold for a range of different prices that were "
              "varied randomly (independent of other factors that could "
              "affect demand). This is the example we have seen in class.")

    body(doc, "The first thing you might want to do is to plot these data, "
              "to eyeball them:")

    T.picture(doc, os.path.join(IMG, "reg_scatter.png"), width_in=6.0)

    body(doc, "The data seem to display a negative relationship between P "
              "and Q, much like a demand curve. But the points are "
              "“cloudy” – they do not align perfectly along a line. "
              "Nevertheless, we roughly see a linear relationship.")

    # ---------------- How to estimate -------------------------------------
    heading(doc, "How to estimate a linear demand curve")

    body(doc, "First, you might be tempted to just draw a line through the "
              "data. If you did that, you would probably try to pick a line "
              "that is “close” to the points, on average. If you did that "
              "manually, you could come up with several lines that “look” "
              "equally close to the points. How do you choose which one is "
              "best?")

    p = body(doc)
    run(p, "Once again, we want to estimate the parameters ")
    run(p, "a", italic=True)
    run(p, " (the intercept) and ")
    run(p, "b", italic=True)
    run(p, " (the slope) of the following equation:")
    equation(doc, mrun("Q=a+bP"))

    p = body(doc)
    run(p, "Regression analysis picks one criterion to choose the “best” "
           "estimates of ")
    run(p, "a", italic=True)
    run(p, " and ")
    run(p, "b", italic=True)
    run(p, ". The criterion is to pick the slope and position of the line to "
           "minimize the sum of squares of the vertical distances from the "
           "line to the data points. We call that way to estimate the "
           "position and slope of the line ")
    run(p, "least squares estimation", italic=True)
    run(p, ", or ")
    run(p, "ordinary least squares (OLS)", italic=True)
    run(p, ". The algorithm is described in the next Figure.")

    T.callout(doc, "Note:",
              "there are many other possible criteria. We won't study them "
              "in class. Least squares estimation has many good properties "
              "and is by far the most widespread method that is used in "
              "econometrics.")

    T.picture(doc, os.path.join(IMG, "reg_leastsquares.png"), width_in=6.0)

    p = body(doc)
    run(p, "It seems complicated, but fortunately, you won't have to do all "
           "that manually. The computer will do it for you! It will take a "
           "fraction of a second for the computer to give you its estimates "
           "of ")
    run(p, "a", italic=True)
    run(p, " and ")
    run(p, "b", italic=True)
    run(p, " from the data that you feed in.")

    body(doc, "If you implement regression analysis in Excel with the data "
              "on airline tickets given in class, you will obtain:")
    # deck slide 34's estimates
    equation(doc, mrun("Q=478.95-1.64P"))

    p = body(doc)
    run(p, "That is, above, ")
    run(p, "a", italic=True)
    run(p, "=478.95 and ")
    run(p, "b", italic=True)
    run(p, "=-1.64. The interpretation is that if you sell tickets at a "
           "price of zero, you will sell 478.95 tickets. And for every $1 "
           "increase in the price of tickets, you will sell 1.64 fewer "
           "tickets.")

    # ---------------- Diagnostics -----------------------------------------
    heading(doc, "Diagnostics")

    p = body(doc)
    run(p, "Excel will not just give you ")
    run(p, "a", italic=True)
    run(p, " and ")
    run(p, "b", italic=True)
    run(p, ". It will produce many more diagnostic statistics. The output "
           "you will get looks something like this:")

    T.table(doc, REG1, W1, header=False, size=9, highlight=HL1,
            align_right=(1, 2, 3, 4, 5, 6), italic_rows=(0, 7, 11),
            label_rows=LABEL_ROWS)

    p = body(doc, before=10)
    run(p, "We won't worry about all the statistics in this table. We will "
           "focus on the most important ones, which I have highlighted.")

    p = numbered(doc, "1)", "The coefficients")
    run(p, ". This is what we care about the most: they are the estimates "
           "of ")
    run(p, "a", italic=True)
    run(p, " (intercept) and ")
    run(p, "b", italic=True)
    run(p, " (slope) discussed above.")

    p = numbered(doc, "2)", "The t-stats and confidence interval")
    run(p, ". More precisely, the t-statistics. For ")
    run(p, "each coefficient", italic=True)
    run(p, ", they represent how accurately the coefficient is estimated. A "
           "t-stat with an absolute value above 2 indicates that you can be "
           "at least 95% certain that the coefficient is different from "
           "zero. The higher the absolute value of the t-stat, the better. "
           "Here, we are doing pretty well since the t-stat is -4.47, giving "
           "us 99.9% confidence that ")
    run(p, "b", italic=True)
    run(p, " is not zero. We are also 95% certain that ")
    run(p, "b", italic=True)
    run(p, " is between -2.43 and -0.85, as indicated by the confidence "
           "interval (“Lower 95%” and “Upper 95%”).")

    p = numbered(doc, "3)", "The R-square")
    run(p, ". The R-square gives us a notion of the ")
    run(p, "overall quality", italic=True)
    run(p, " of our model – that is, how good are we overall at explaining "
           "the variation in ")
    run(p, "Q", italic=True)
    run(p, " with variation in ")
    run(p, "P", italic=True)
    run(p, ". Here, the R-square is 0.586. That means the price of the "
           "tickets alone explains 58.6% of the variation in the quantity "
           "sold. In the figure above, if all the data points were arrayed "
           "exactly on a line, R-square would be one. Price would perfectly "
           "predict quantity. We can see graphically that this is not the "
           "case here. The more cloudy the data points around the line, the "
           "lower the R-square. The closer the data points are to the line "
           "on average, the higher the R-square. We want a high value of the "
           "R-square, but that is sometimes hard to obtain.")

    # ---------------- Multivariate ----------------------------------------
    heading(doc, "Multivariate Regression")

    p = body(doc)
    run(p, "What could we do to improve the explanatory power of our model? "
           "The answer lies in ")
    run(p, "multivariate regression", italic=True)
    run(p, ". That is, instead of only trying to predict quantity with "
           "price, we will add explanatory variables. In class, we have seen "
           "that there are many determinants of demand. Among them are the "
           "income of your customers (which influences demand positively "
           "for ")
    run(p, "normal goods", italic=True)
    run(p, ", and negatively for ")
    run(p, "inferior goods", italic=True)
    run(p, ") and the prices of complements and substitutes. So one thing we "
           "could do here, for instance, is to include in our model an index "
           "of income over time, and a measure of the ticket price charged "
           "by the airline's main competitor. Income is expected to come out "
           "with a positive sign because air travel is a normal good, and "
           "the competitor's price is also expected to influence demand "
           "positively: when our competitor raises prices, consumers will "
           "switch to our product. Let's add those two variables and "
           "estimate the model again:")

    T.table(doc, REG2, W1, header=False, size=9, highlight=HL2,
            align_right=(1, 2, 3, 4, 5, 6), italic_rows=(0, 7, 11),
            label_rows=LABEL_ROWS)

    body(doc, "A few additional points:", before=10)

    p = numbered(doc, "1)", "The estimated coefficients on a and b have changed")
    run(p, ". The intercept (")
    run(p, "a", italic=True)
    run(p, ") no longer has a natural interpretation here. It is the number "
           "of tickets sold if own price is zero, competitor's price is zero ")
    run(p, "and", italic=True)
    run(p, " the income index is zero. Of course, these conditions are "
           "unlikely to ever hold. The coefficient on own price (")
    run(p, "b", italic=True)
    run(p, ") is now -2.124, and its t-stat has risen to -6.24. So the "
           "demand curve is slightly steeper, and we can be even more "
           "confident of the negative relationship between price and "
           "quantity. That means there will be less uncertainty on our "
           "choice of the best ")
    run(p, "Q", italic=True)
    run(p, " and ")
    run(p, "P", italic=True)
    run(p, ", so it is good news!")

    p = numbered(doc, "2)",
                 "The estimates on competitor's price and income are "
                 "positive as expected.")
    run(p, " They also both have t-stats that are greater than 2, meaning "
           "that they are individually predictive of demand. So, we confirm "
           "that air travel is a normal good and that the competitor's "
           "product is a substitute for ours.")

    p = numbered(doc, "3)", "The R-square has increased to 0.776")
    run(p, ". That means we have improved the overall quality of our model, "
           "and can now explain 77.6% of the variation in quantity sold. Is "
           "that good? There is no absolute standard or cutoff for a “good” "
           "R-square, but it's pretty high.")

    body(doc, "We now have a more precisely estimated demand curve:")

    # variable names as on deck slide 39 (Q = a + bP + cPc + dY)
    equation(doc,
             mrun("Q=28.84-2.12P+1.04")
             + T.msub(mrun("P"), mrun("C", italic=False))
             + mrun("+3.09") + mrun("Y"))

    p = para(doc, after=9, align=T.WD_ALIGN_PARAGRAPH.JUSTIFY)
    # "(see the Teaching Note on ...)" is the house form for a parenthetical
    # cross-reference (Nico added "see the" to the first one by hand on
    # 2026-09-06; extended to the second so the pair matches, and so both
    # read like the ones in the MR=MC and Hiring notes).
    body(doc, "With an estimated demand curve in hand, you can compute the "
              "price elasticity at any point (see the Teaching Note on "
              "Demand Elasticity and Total Revenue). You can also derive "
              "marginal revenue with the 3-step recipe (see the Teaching "
              "Note on Marginal Revenue).")

    doc.save(OUT)
    print("wrote", OUT)


def numbered(doc, marker, lead):
    """A numbered paragraph with a bold lead-in, as in the original."""
    p = para(doc, before=6, after=6, left=0.40, hang=0.40,
             align=T.WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, marker + "\t")
    run(p, lead, bold=True)
    return p


if __name__ == "__main__":
    main()
