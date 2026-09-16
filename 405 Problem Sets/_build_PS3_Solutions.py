"""Build: Problem Set 3 -- Solutions.docx  (MGMT 405, Fall 2026 format).

Source: `_originals/Problem Set 3 -- Solutions (2025 original).docx` -- a
BUILD INPUT, never deleted.

ARITHMETIC RE-CHECKED, and every value in the original is correct except
the one noted under PROPOSED 1:
  P1  MP = 500, 600, 600, 500, 300, 100, -100, -500 -- differences of
      0, 500, 1100, 1700, 2200, 2500, 2600, 2500, 2000.
  P2  MPL/w = 50/200 = 0.25 ; MPK/pK = 200/1000 = 0.20.
  P3  TFC 25,600 throughout; TVC = TC - TFC; AFC, AVC, ATC and the
      forward-difference MC all reproduce (the MC values are unchanged --
      only their PLACEMENT moved, onto the strip between the two rows
      each one is derived from; see `S.change_table`).  Profit peaks at Q = 12,000 with
      8,480.  At $5.50 profit peaks at Q = 11,000 with -2,060, where
      ATC = 5.7 > 5.50 (a loss) but AVC = 3.4 < 5.50 (keep producing).
  P4  MCo = 2 + Q = 10 -> Q = 8 ; TCo(8) = 98 ; profit = 80 - 98 = -18.
      AVCo(8) = 2 + 0.5(8) = 6 < 10, so keep producing.
      MCn = 0.2Q = 10 -> Q = 50 ; TCn(50) = 350 ; profit = 500 - 350 = 150.
      ATCo(8) = 50/8 + 2 + 4 = 12.25 ; ATCn(50) = 2 + 5 = 7.
  P5  (0.02 x 13,240) / (12 x 365) = 264.8 / 4,380 = $0.0605 -> 6 cents.

FORMATTING (untracked)
  - Course masthead, navy problem headings with a gold points chip, the
    bare centred page number.
  - Derivations set as native OMML; the line that delivers the answer is
    dark red.
  - PROFIT IS A LOWER-CASE pi (Teaching CLAUDE.md, 2026-08-30).  The
    original used the capital .
  - All six figures rebuilt as native, editable Word shapes; the
    originals were Excel and matplotlib screenshots.
  - Profit and loss areas follow the course's fill convention
    (2026-08-30): a positive profit is dark red `C00000`, a loss the
    firm still produces through is GRAY `555B66`, both at about 22 %.
    The original shaded the Problem 4(b) LOSS in red and the 4(e) PROFIT
    in green, which is the convention backwards.
  - The 4(b) and 4(e) panels now start at Q = 0 so the profit / loss
    rectangle is drawn over its full width; the originals began the axis
    at Q = 2 and Q = 10, which cut the rectangle short.
  - Problem 5(b)'s figure adopts the deck's own tax notation (slide 69):
    S and S', P0, PB for the price paid by the buyer, PS for the price
    received by the seller, and the wedge PB - PS = t.  The original
    used P / P' / PC', which cannot express part (c)'s answer about who
    bears the burden.
  - Problem 3's two cost charts are plotted from the table itself, so
    figure and table cannot drift apart.

PROPOSED (tracked -- these go beyond formatting)
  1. Problem 3(c)'s second table has a STALE TR COLUMN.  It still shows
     the $6.40 revenues (6,400 / 12,800 / 19,200 ...) although the price
     is now $5.50, so TR and Profit contradict each other on every row:
     at Q = 1,000 it shows TR = 6,400 and profit = -25,860, which is
     5,500 - 31,360.  The Profit column is right; TR is wrong.  Corrected
     to 5.5 x Q throughout.
  2. Problem 3(c): "the firm should operate" -> "the firm should continue
     to produce".  Teaching CLAUDE.md and Module 4 slide 32 keep these
     apart: in the SHORT run a firm continues to produce or stops
     production; "shut down" is the long-run exit decision.  "Operate"
     blurs the distinction the question is testing.
  3. Problem 2: the price of capital is written pK, not r.  Module 3
     slide 43 glosses it "pK : Price of Capital", and the teaching note
     on the bang-for-the-buck rule uses pK throughout.
  4. Problem 2(a): the original answered the question twice, in two
     consecutive paragraphs saying the same thing (one of them with an
     unclosed parenthesis, "in order to reduce MPL."), and then repeated
     the point a third time under (b).  Consolidated to one answer under
     (a) and the "bang for the buck" framing under (b), which is what
     (b) actually asks for.
  5. Problem 2's restated question carried a duplicated sentence, "Is
     this firm producing its output at minimum cost? Is this firm
     producing its output in an optimized way?"  The problem set asks
     only the second.  Dropped.

Run:  python _build_PS3_Solutions.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _ps_theme as S
from _ps_theme import (DARKRED, GOLD, GRAY, NAVY, Panel, WD_ALIGN_PARAGRAPH,
                       body, caption, dele, ins, para, part, problem, run)
import _tn_theme as T
from _tn_theme import acr, mfrac, mrun, msub, msup

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Problem Set 3 -- Solutions.docx")

NCBI = "https://www.ncbi.nlm.nih.gov/pmc/articles/PMC6702611/"
PBS = ("https://www.pbs.org/newshour/economy/making-sense/"
       "how-taxing-sugary-drinks-affects-a-communitys-health-and-economy")
HSPH = ("https://www.hsph.harvard.edu/nutritionsource/healthy-drinks/"
        "sugary-drinks/")

Q = mrun("Q")
P = mrun("P")
PI = mrun("π")                      # lower-case pi, per the 2026-08-30 rule
MP_F = acr("MP", "F")
MP_L = acr("MP", "L")
MP_K = acr("MP", "K")
W = mrun("w")
P_K = msub(mrun("p"), mrun("K"))
MC = acr("MC")
MR = acr("MR")
AVC = acr("AVC")
ATC = acr("ATC")
TC = acr("TC")
MC_O, MC_N = msub(MC, mrun("O")), msub(MC, mrun("N"))
TC_O, TC_N = msub(TC, mrun("O")), msub(TC, mrun("N"))
ATC_O, ATC_N = msub(ATC, mrun("O")), msub(ATC, mrun("N"))
PI_O, PI_N = msub(PI, mrun("O")), msub(PI, mrun("N"))
Q_O, Q_N = msub(Q, mrun("O")), msub(Q, mrun("N"))

TULIPS = [(0, 0), (1, 500), (2, 1100), (3, 1700), (4, 2200),
          (5, 2500), (6, 2600), (7, 2500), (8, 2000)]
# Each marginal product is plotted at the MIDPOINT of its interval
# (2026-09-13, Nico), so the figure says the same thing the table's
# staggered column does: the value belongs to the STEP between two
# fertilizer levels, not to either level on its own.  It used to be
# plotted at TULIPS[i][0], the initial point.
MP = [((TULIPS[i][0] + TULIPS[i + 1][0]) / 2.0,
       TULIPS[i + 1][1] - TULIPS[i][1])
      for i in range(len(TULIPS) - 1)]

BATS = [(0, 25600), (1000, 31360), (2000, 35920), (3000, 39520),
        (4000, 42400), (5000, 44800), (6000, 46960), (7000, 49120),
        (8000, 51520), (9000, 54400), (10000, 58000), (11000, 62560),
        (12000, 68320), (13000, 75520), (14000, 84400), (15000, 95200),
        (16000, 108160)]
TFC = 25600


def _num(x):
    """Format with a true minus sign, never a hyphen."""
    return "{:,}".format(x) if x >= 0 else "−{:,}".format(-x)


# Column order.  MC and MR are the two change-based series, so they are
# the only ones written on the between-rows strips.
# The CHANGE columns sit at the far right (2026-09-13, Nico: "always move
# the change columns to the very right").  They used to sit inline -- MC
# after ATC, MR after TR -- which put two staggered columns in the middle
# of the run of level columns.
BATS_HEAD = ["Output", "Total Cost", "TFC", "TVC", "AFC", "AVC", "ATC",
             "TR", "Profit", "MC", "MR"]
I_MC = BATS_HEAD.index("MC")
I_MR = BATS_HEAD.index("MR")

# Widths BY NAME, so reordering the header reorders the columns and
# nothing else has to change.
BATS_WIDTH = {"Output": 0.52, "Total Cost": 0.56, "TFC": 0.44, "TVC": 0.46,
              "AFC": 0.40, "AVC": 0.42, "ATC": 0.42, "TR": 0.52,
              "Profit": 0.50, "MC": 0.40, "MR": 0.40}


def bats_rows(price):
    """The cost table at a given price.

    Everything is COMPUTED from (Output, Total Cost) and the price, so a
    column cannot go stale the way the original's TR column did at $5.50.
    MC and MR sit on the row the step STARTS from, as they always did --
    `S.change_table` then drops those two columns half a row, so each
    value lands on the border between the two rows it comes from.  The
    last row's MC and MR are blank: there is no next output level.
    """
    rows = []
    for i, (q, tc) in enumerate(BATS):
        tvc = tc - TFC
        nxt = BATS[i + 1] if i + 1 < len(BATS) else None
        tr = price * q
        # BY NAME, then emitted in BATS_HEAD order -- the old version
        # addressed cells as row[4] / row[8] / row[10], so moving a column
        # would have silently put values under the wrong headers.
        v = {"Output": "{:,}".format(q),
             "Total Cost": "{:,}".format(tc),
             "TFC": "{:,}".format(TFC),
             "TVC": "{:,}".format(tvc),
             "TR": "{:,.0f}".format(tr),
             "Profit": _num(int(round(tr - tc)))}
        if q:
            v["AFC"] = "{:.1f}".format(TFC / q)
            v["AVC"] = "{:.1f}".format(tvc / q)
            v["ATC"] = "{:.1f}".format(tc / q)
        if nxt:
            v["MC"] = "{:.1f}".format((nxt[1] - tc) / float(nxt[0] - q))
            v["MR"] = "{:.2f}".format(price)        # MR = dTR/dQ = price
        rows.append([v.get(h, "") for h in BATS_HEAD])
    return rows


# Column widths.  Every entry is a number, and the widest is "108,160"
# and "-25,600" at 8 pt (about 0.36"), so the columns are sized by their
# HEADERS, which wrap to two lines at these widths.  Narrower than the
# old 0.57" flat because the numbers are centred now rather than pushed
# against a right edge.
BATS_W = [BATS_WIDTH[h] for h in BATS_HEAD]


def bats_table(doc, price):
    best = max(range(len(BATS)), key=lambda i: price * BATS[i][0] - BATS[i][1])
    return S.stagger_table(doc, BATS_HEAD, bats_rows(price), (I_MC, I_MR),
                           widths_in=BATS_W, size=8, highlight=(best,),
                           change_fade=True)


def draws_on(doc, text):
    # after 8 -> 5 on 2026-09-12, part of returning ~30 pt of pure
    # spacing so a page break is never decided by a few points.
    p = para(doc, before=0, after=5, keep_next=True)
    run(p, "Draws on:  " + text, italic=True, color=GRAY, size=9.5)
    return p


def answer(doc, content, before=4, after=8):
    return T.equation(doc, T.color_omml(content, DARKRED),
                      before=before, after=after)


def bullet(doc, text=None, before=2, after=2):
    p = para(doc, before=before, after=after, left=0.32, hang=0.18,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, "•  ", color=NAVY, size=11)
    if text:
        run(p, text, size=11)
    return p


# ==========================================================================
# Figures
# ==========================================================================
def fig_mp():
    """1(a): the marginal product of fertilizer, plotted at the MIDPOINT
    of each interval (2026-09-13, Nico).

    The x tick marks run 0 - 8, the fertilizer LEVELS, so every plotted
    point visibly sits between two of them -- which is the whole reason
    for the midpoint.  The dots carry the reading; the polyline only
    joins them.
    """
    f = S.fig(5.90, 3.85, name="1(a) - marginal product of fertilizer")
    pn = Panel(f, 0.92, 3.20, 4.30, 2.70, ylabel="MP", xlabel="Fertilizer")

    # The frame's baseline is the BOTTOM of the value range, not MP = 0.
    # Putting the axis on the zero line would drop the x tick labels into
    # the part of the curve that goes negative, which is where the answer
    # to part (c) lives.  Zero is drawn as its own solid line instead.
    def u(t):
        # 8.8, not 8.0: the axis title is anchored to the arrow TIP at the
        # right end, and a tick drawn at u = 100 lands underneath it.  The
        # extra headroom keeps the "8" clear of "Fertilizer" (2026-09-13).
        return t / 8.8 * 100.0

    def v(mp):
        return (mp + 600.0) / 1300.0 * 100.0

    for n in (-500, -250, 250, 500):
        pn.f.line(pn.x(0), pn.y(v(n)), pn.x(100), pn.y(v(n)),
                  color=S.LIGHT, w_pt=0.75, dash="dash", name="gridline")
    for n in (-500, -250, 0, 250, 500):
        pn.tick_y(v(n), "%d" % n if n >= 0 else "−%d" % -n,
                  size=9.5, italic=False)
    # the zero line, drawn solid so the sign change is unmistakable
    pn.f.line(pn.x(0), pn.y(v(0)), pn.x(100), pn.y(v(0)), color=GRAY,
              w_pt=1.0, name="zero")
    # 0 - 8, not 0 - 7: the ticks are the fertilizer LEVELS, and each
    # plotted marginal product sits halfway between two of them.
    for t in range(0, 9):
        pn.tick_x(u(t), "%d" % t, size=9.5, italic=False)

    pn.polyline([(u(t), v(m)) for t, m in MP], color=NAVY, w_pt=2.0,
                name="MP")
    for t, m in MP:
        pn.f.dot(pn.x(u(t)), pn.y(v(m)), d=0.070, color=NAVY, name="obs")
    return f


def _bats_panel(name, series, ylab, ymin, ymax, ticks,
                legend_u, legend_v):
    """Shared frame for Problem 3's two cost charts.

    `series` is a list of (label, colour, [(Q, value), ...]).  Each is
    one polyline plus one label placed at its own right-hand end, so no
    two labels can land on the same spot.
    """
    f = S.fig(6.20, 3.95, name=name)
    pn = Panel(f, 1.02, 3.28, 4.55, 2.85, ylabel=ylab, xlabel="Q")

    def u(q):
        return q / 17000.0 * 100.0

    def v(y):
        return (y - ymin) / float(ymax - ymin) * 100.0

    for y in ticks:
        if y != ymin:
            pn.f.line(pn.x(0), pn.y(v(y)), pn.x(100), pn.y(v(y)),
                      color=S.LIGHT, w_pt=0.75, dash="dash", name="gridline")
        pn.tick_y(v(y), _num(y), size=9, italic=False)
    for q in (0, 4000, 8000, 12000, 16000):
        pn.tick_x(u(q), "{:,}".format(q), size=9, italic=False)

    for label, color, pts in series:
        pn.polyline([(u(q), v(y)) for q, y in pts], color=color, w_pt=1.75,
                    name=label)
    # A legend, not end-of-curve labels: in the average-cost panel ATC,
    # MR and AVC all arrive within one logical unit of each other at the
    # right-hand edge, so their labels cannot be separated there.
    pn.legend([(lab, col) for lab, col, _ in series], legend_u, legend_v)
    return f


def fig_bats_totals():
    rows = [(q, tc) for q, tc in BATS]
    tvc = [(q, tc - TFC) for q, tc in BATS]
    tfc = [(0, TFC), (16000, TFC)]
    tr = [(q, 6.4 * q) for q, _ in BATS]
    profit = [(q, 6.4 * q - tc) for q, tc in BATS]
    return _bats_panel(
        "3(a) - TC, TVC, TFC, TR and profit",
        [("TC", DARKRED, rows),
         ("TR", NAVY, tr),
         ("TVC", GOLD, tvc),
         ("TFC", "7A5C2E", tfc),
         ("π", "1B5E20", profit)],
        "$", -40000, 120000,
        [-40000, -20000, 0, 20000, 40000, 60000, 80000, 100000, 120000],
        legend_u=5, legend_v=99)


def fig_bats_averages():
    afc = [(q, TFC / q) for q, _ in BATS if q]
    avc = [(q, (tc - TFC) / q) for q, tc in BATS if q]
    atc = [(q, tc / q) for q, tc in BATS if q]
    # MC is plotted at the MIDPOINT of its output interval (2026-09-13,
    # Nico), the same convention as the tulip MP figure: a marginal cost
    # describes the STEP from one output level to the next, so it belongs
    # between them, not on either one.  It used to be plotted at `q`.
    mc = []
    for i, (q, tc) in enumerate(BATS[:-1]):
        q1 = BATS[i + 1][0]
        mc.append(((q + q1) / 2.0, (BATS[i + 1][1] - tc) / float(q1 - q)))
    # MR is marginal too, so its line spans the same midpoints rather than
    # the output levels -- it is flat at the price, but the two marginal
    # curves should cover the same domain.
    mr = [(mc[0][0], 6.4), (mc[-1][0], 6.4)]
    return _bats_panel(
        "3(a) - AFC, AVC, ATC, MC and MR",
        [("MC", "1B5E20", mc),
         ("ATC", DARKRED, atc),
         ("MR", NAVY, mr),
         ("AVC", GOLD, avc),
         ("AFC", "7A5C2E", afc)],
        "$", 0, 35, [0, 5, 10, 15, 20, 25, 30, 35],
        legend_u=40, legend_v=99)


def fig_balding(new):
    """4(b) and 4(e): the price taker's optimum under each technology.

    Both panels start at Q = 0 so the profit / loss rectangle is drawn
    over its whole width, and both shade it in the course's colour for
    the case: GRAY for a loss the firm still produces through (the old
    technology), dark red for a positive profit (the new one).
    """
    if new:
        qmax, ymax, qstar, atcstar = 100.0, 25.0, 50.0, 7.0
        def mc(q):
            return 0.2 * q

        def atc(q):
            return 100.0 / q + 0.1 * q
        name, fill = "4(e) - new technology", DARKRED
        word, pi_txt = "Profit", "π > 0"
    else:
        qmax, ymax, qstar, atcstar = 16.0, 30.0, 8.0, 12.25
        def mc(q):
            return 2.0 + q

        def atc(q):
            return 50.0 / q + 2.0 + 0.5 * q
        name, fill = "4(b) - old technology", GRAY
        word, pi_txt = "Loss", "π < 0"

    f = S.fig(5.90, 4.05, name=name)
    pn = Panel(f, 0.80, 3.42, 4.42, 3.00, ylabel="$", xlabel="Q")

    def u(q):
        return q / qmax * 100.0

    def v(y):
        return y / ymax * 100.0

    # the shaded area FIRST, so every curve is drawn on top of it
    pn.region(u(0), v(10), u(qstar), v(atcstar), fill=fill, alpha=22000,
              line=NAVY, name=word.lower())

    # MC is NAVY and the price line is DARK RED: for a price taker the
    # horizontal line at P is the demand curve the firm faces, so it takes
    # the deck's demand colour, label and axis value with it.  Module 4's
    # own price-taker panel says exactly this (2026-08-29, Nico).
    pn.polyline([(u(q), v(mc(q))) for q in (0.0, qmax)], color=NAVY,
                w_pt=2.0, name="MC")
    pn.polyline([(u(0), v(10)), (u(qmax), v(10))], color=DARKRED, w_pt=2.0,
                name="MR = P")
    # ATC, sampled finely enough that the U reads as a curve
    step = qmax / 220.0
    pts = []
    q = step
    while q <= qmax + 1e-9:
        y = atc(q)
        if y <= ymax:
            pts.append((u(q), v(y)))
        q += step
    pn.polyline(pts, color=GOLD, w_pt=2.0, name="ATC")

    # Curve labels, each at its own right-hand end -- but under the new
    # technology ATC arrives only $1 above the price line, so those two
    # labels would sit on top of each other.  Push them apart to one
    # label height, working outward from the lowest, rather than letting
    # them overlap.
    sep = 0.22 / pn.h * 100.0 + 0.5
    ends = sorted(((v(yv) + 3.0, lab, color) for lab, color, yv in
                   (("MC", NAVY, mc(qmax)),
                    ("MR = P", DARKRED, 10.0),
                    ("ATC", GOLD, atc(qmax)))))
    placed = []
    for vv, lab, color in ends:
        if placed and vv < placed[-1][0] + sep:
            vv = placed[-1][0] + sep
        placed.append((vv, lab, color))
    for vv, lab, color in placed:
        runs = [(lab, dict(bold=True, color=color, size=10.5))]
        pn.text(u(qmax) + 1.5, vv, runs, align="l")

    # the optimum
    pn.f.line(pn.x(u(qstar)), pn.y(v(max(10.0, atcstar))),
              pn.x(u(qstar)), pn.y(0), color=GRAY, w_pt=1.0, dash="dash",
              name="Q* guide")
    pn.f.line(pn.x(0), pn.y(v(atcstar)), pn.x(u(qstar)), pn.y(v(atcstar)),
              color=GRAY, w_pt=1.0, dash="dash", name="ATC* guide")
    pn.f.dot(pn.x(u(qstar)), pn.y(v(10)), d=0.075, color=NAVY, name="MR = MC")
    pn.f.dot(pn.x(u(qstar)), pn.y(v(atcstar)), d=0.075, color=GOLD,
             name="on ATC")
    pn.tick_x(u(qstar), "Q* = %g" % qstar, h=0.20)
    pn.tick_y(v(10), "10", h=0.19, color=DARKRED)
    pn.tick_y(v(atcstar), "%g" % atcstar, h=0.19)

    # Name the shaded area inside it, at a quantity where the ATC curve
    # is clear of the band.  Under the new technology ATC runs INSIDE the
    # rectangle over Q = 12.7 to 20, so the label is placed further right
    # (around Q = 32, where ATC has dipped to 6.3) rather than at the
    # midpoint of the box.
    runs = [("%s  (%s)" % (word, pi_txt),
             dict(bold=True, color=NAVY, size=10.5))]
    spot = _clear_spot(pn, runs, u, v, qstar, qmax, atcstar, (mc, atc))
    if spot is None:
        # No window anywhere in the rectangle holds the word clear of
        # both curves: the convention's own fallback is the short form
        # alone, since the panel heading already names the case.
        runs = [(pi_txt, dict(bold=True, color=NAVY, size=10.5))]
        spot = _clear_spot(pn, runs, u, v, qstar, qmax, atcstar, (mc, atc))
    if spot is None:
        spot = (0.5 * qstar, v((10.0 + atcstar) / 2.0) + 3.2)
    pn.text(u(spot[0]), spot[1], runs, align="c")
    return f


def _clear_spot(pn, runs, u, v, qstar, qmax, atcstar, curves, steps=120):
    """Where inside the shaded rectangle the label clears every curve.

    Returns (quantity, band top in logical v) for the label's CENTRE, or
    None if nothing clears.  The label is as wide as its own text, so a
    curve crossing the rectangle can cut it even when the rectangle is
    roomy: under the new technology ATC dips into the box over Q = 12.7
    to 20 and MC rises into it after Q = 40.  Searching both the
    quantity and the height finds the window instead of carrying a
    hand-picked fraction of `qstar`, which is what let MC cut the label
    when the numbers last changed.
    """
    # The label BOX is 0.22" tall around ~0.152" of glyphs, and it is
    # the glyphs that must clear the curves and sit inside the shaded
    # rectangle -- measuring the box instead rejected the loss panel,
    # whose rectangle is 0.225" deep, and cost it its "Loss" wording.
    hbox = 0.22 / pn.h * 100.0
    pad = 0.034 / pn.h * 100.0
    hv = hbox - 2.0 * pad
    half_q = (S.text_width(runs) + 0.08) / 2.0 / pn.w * qmax
    v_lo, v_hi = sorted((v(10.0), v(atcstar)))
    if v_hi - v_lo < hv + 0.6:                  # rectangle too thin
        return None

    best = None
    for k in range(13):
        top = v_hi - 0.3 - (v_hi - v_lo - hv - 0.6) * k / 12.0
        bottom = top - hv
        for i in range(steps + 1):
            cq = qstar * i / float(steps)
            if cq - half_q < 0 or cq + half_q > qstar:
                continue
            room = None
            for j in range(21):
                q = cq - half_q + 2.0 * half_q * j / 20.0
                if q <= 0:
                    continue
                for fn in curves:
                    vv = v(fn(q))
                    g = (vv - top if vv > top else
                         bottom - vv if vv < bottom else
                         -min(top - vv, vv - bottom))
                    if room is None or g < room:
                        room = g
            if room is not None and (best is None or room > best[0]):
                best = (room, cq, top)
    if best is None or best[0] < 0.5:
        return None
    return best[1], best[2] + pad          # back to the BOX's top edge


def fig_tax():
    """5(b): a tax levied on producers shifts the supply curve in.

    Notation is the deck's (slide 69): the buyer pays PB, the seller
    receives PS, and the wedge between them is the tax t.  That is what
    part (c) then reasons about.
    """
    f = S.fig(5.90, 4.05, name="5(b) - a tax on sugary drinks")
    pn = Panel(f, 0.90, 3.42, 4.32, 3.00, ylabel="P", xlabel="Q")

    # S and S' are drawn with the SAME slope -- a per-unit tax shifts the
    # supply curve UP by the tax, it does not tilt it.  With slopes of
    # +1 and -1 the burden splits evenly, which is the neutral case to
    # reason from in part (c).
    #   S :  v = u                 S' : v = u + 22          D : v = 100 - u
    #   E0 = (50, 50)   E1 = (39, 61)   PS = 39   tax = 22
    s0, s1 = (5, 5), (80, 80)
    t0, t1 = (5, 27), (70, 92)
    d0, d1 = (6, 94), (94, 6)

    pn.supply(s0, s1, label="S", lbl_dx=0.07, lbl_dy=-0.09)
    # naming the shifted curve "S' = S + t" carries the tax where there is
    # room for it, so the wedge itself only needs a one-character mark
    pn.supply(t0, t1, label="S′ = S + t", lbl_dx=0.07, lbl_dy=-0.09,
              name="S prime")
    pn.demand(d0, d1, label="D", lbl_dx=0.08, lbl_dy=-0.14)

    e0 = S.cross(s0, s1, d0, d1)              # no tax
    e1 = S.cross(t0, t1, d0, d1)              # with tax: the buyer's price
    ps = S.line_at(s0, s1, e1[0])             # what the seller keeps

    pn.equilibrium(e0[0], e0[1], ptick=("P", "0"), qtick=("Q", "0"))
    pn.equilibrium(e1[0], e1[1], ptick=("P", "B"), qtick=("Q", "1"))
    pn.f.dot(pn.x(e1[0]), pn.y(ps), d=0.075, color=GOLD, name="seller price")
    pn.f.line(pn.x(0), pn.y(ps), pn.x(e1[0]), pn.y(ps), color=GRAY,
              w_pt=1.0, dash="dash", name="PS guide")
    pn.tick_y(ps, ("P", "S"))

    # The wedge is measured where the two supply curves are clear of
    # everything else -- to the right of E0, where the demand curve lies
    # well below them.  A plain segment, not an arrow: an arrowhead would
    # suggest a direction of movement.
    wu = 60.0
    pn.curve((wu, S.line_at(s0, s1, wu)), (wu, S.line_at(t0, t1, wu)),
             GOLD, label=None, w_pt=1.75, name="tax wedge")
    # A one-character mark beside the wedge.  The band between S and S'
    # is diagonal, so any wider label set in, above or below it runs into
    # one of the two curves; a single "t" fits in the gap cleanly.
    pn.text(wu + 1.5,
            (S.line_at(s0, s1, wu) + S.line_at(t0, t1, wu)) / 2.0 + 3.5,
            [("t", dict(bold=True, italic=True, color=NAVY, size=11))],
            align="l")
    return f


# ==========================================================================
def main():
    doc = T.new_doc(margin_in=1.0)
    T.footer(doc)

    S.ps_masthead(doc, "Problem Set 3 – Solutions",
                  covers="Covers Modules 3 and 4")

    # ======================================================================
    # Problem 1
    # ======================================================================
    problem(doc, 1, "Tulip Production", 15, before=12)
    draws_on(doc, "Module 3 – The Production Function (marginal product, "
                  "diminishing returns)")

    p = part(doc, "a", 10)
    run(p, "To compute the marginal product of fertilizer we use ")
    T.equation_inline(p, MP_F + mrun("=") + mfrac(mrun("Δ", italic=False) + Q,
                                                  mrun("Δ", italic=False)
                                                  + mrun("F")))
    run(p, ", where ")
    run(p, "F", italic=True)
    run(p, " refers to fertilizer. ")
    dele(p, "The table gives the marginal product at each quantity, entered "
            "at the initial point of each step, and the figure below plots "
            "it.")
    # His wording, adopted from the copy he saved on 2026-09-13.  He cut
    # the sentence describing the column's half-row offset and moved the
    # figure to the midpoint of each interval.
    ins(p, "The marginal-product column corresponds to the step from the "
           "initial fertilizer level to the next. The figure below plots "
           "it at the mid-point of these intervals.")

    # Two rows per fertilizer level, merged: the tons and the tulip count
    # are centred across the pair, the marginal product across the pair
    # one row lower, so it lands on the border between the two levels it
    # comes from.  Widths measured off the wrapped HEADERS (the numbers
    # need barely a third of it) -- see S.stagger_table.
    S.stagger_table(
        doc,
        ["Tons of fertilizer per month", "Number of tulips per month",
         "Marginal product"],
        [["{:,}".format(t), "{:,}".format(n),
          "" if i >= len(MP) else _num(MP[i][1])]
         for i, (t, n) in enumerate(TULIPS)],
        (2,), widths_in=[1.35, 1.35, 1.15], size=10, change_fade=True)

    S.place(fig_mp(), doc, before=10, after=2)
    caption(doc, "Marginal product of fertilizer, plotted at the mid-point "
                 "of each interval. It rises to 600, then falls, and turns "
                 "negative beyond 6 tons.")

    p = part(doc, "b", 3)
    run(p, "Yes. As we move beyond 2 tons, the marginal return to fertilizer "
           "is decreasing. (It is also fine to say “as we start producing "
           "the 3rd ton”, since production of the 3rd ton starts as soon as "
           "we move beyond the initial point of 2 tons.)")

    p = part(doc, "c", 2)
    run(p, "No. From 6 to 8 tons, total output decreases as more fertilizer "
           "is added.")

    # ======================================================================
    # Problem 2
    # ======================================================================
    problem(doc, 2, "Sludge Production", 15)
    draws_on(doc, "Module 3 – Long Run: The Optimal Input Mix "
                  "(bang-for-the-buck rule)")

    # PROPOSED 4 + 5: one answer under (a), the "bang for the buck"
    # framing under (b); the duplicated restatement of the question and
    # the second, near-identical paragraph are dropped.
    p = part(doc, "a", 8)
    run(p, "When optimizing its input mix to produce a fixed output "
           "quantity, the firm should divide each input’s marginal product "
           "by that input’s price and choose inputs so that these ratios are "
           "equal:")
    T.equation(doc, mfrac(MP_L, W) + mrun("=") + mfrac(MP_K, P_K),
               before=3, after=5)
    p = body(doc)
    run(p, "Here ")
    T.equation_inline(p, mfrac(MP_L, W) + mrun("=") + mfrac(mrun("50"),
                                                            mrun("200"))
                      + mrun("=") + mrun("0.25"))
    run(p, " while ")
    T.equation_inline(p, mfrac(MP_K, P_K) + mrun("=") + mfrac(mrun("200"),
                                                              mrun("1000"))
                      + mrun("=") + mrun("0.20"))
    run(p, ", so the firm is ")
    run(p, "not", italic=True)
    run(p, " minimizing costs.")

    answer(doc, mfrac(MP_L, W) + mrun("=") + mrun("0.25") + mrun("  >  ")
           + mfrac(MP_K, P_K) + mrun("=") + mrun("0.20"))

    p = body(doc)
    run(p, "The firm should use ")
    run(p, "fewer machines", bold=True)
    run(p, " (which drives ")
    T.equation_inline(p, MP_K)
    run(p, " up) and ")
    run(p, "more labor", bold=True)
    run(p, " (which drives ")
    T.equation_inline(p, MP_L)
    run(p, " down), until the two ratios are equal. The condition says that "
           "the additional output you get per extra dollar spent should be "
           "the same no matter which input you spend the next dollar on. "
           "Otherwise you could reallocate spending toward the input that "
           "yields more extra output per dollar and produce the same output "
           "at lower cost. At the optimal input mix, no such cost-reducing "
           "substitution is left.")

    p = part(doc, "b", 7)
    run(p, "This is exactly the “bang for the buck” rule from class: the "
           "firm gets more bang for the buck from labor than from capital "
           "here, so it should use more labor and less capital until the "
           "bang for the buck is equalized across the two inputs.")

    S.note(doc,
           "Below is an example of how a student could use their own "
           "industry to illustrate the concept.", prefix="Grading:")

    body(doc,
         "A real-world example is the fast-food industry’s move from "
         "cashiers to self-service ordering kiosks. Restaurants compared the "
         "service each extra dollar on staff cashiers delivered with the "
         "service each extra dollar spent on kiosks delivered. When kiosks "
         "provided more “orders handled per dollar”, many chains reallocated "
         "resources from cashiers toward kiosks until the extra service per "
         "dollar spent became more balanced between the two. They did not "
         "switch entirely to kiosks because some tasks would be too costly "
         "to be performed only by machines — such as handling special "
         "requests, or solving problems when machines fail — and for those, "
         "the product of a human worker would still provide more bang for "
         "the buck.")

    # ======================================================================
    # Problem 3
    # ======================================================================
    problem(doc, 3, "Baseball Bats", 25)
    draws_on(doc, "Module 3 – Cost Concepts;  Module 4 – Profit "
                  "Maximization of a Price Taker in the Short Run")

    p = part(doc, "a", 10, after=6)
    run(p, "The completed table, at the market price of $6.40. The "
           "profit-maximizing row is highlighted.")
    bats_table(doc, 6.4)
    caption(doc, "Table 1 – costs and revenues at a price of $6.40",
            before=4, after=8)

    # Rewritten for the between-rows layout (2026-09-09).  It is NOT a
    # tracked change: `S.note` is a cream card, and a revision inside a
    # text box never reaches Word's review pane.
    S.note(doc,
           "MC and MR describe a STEP from one output level to the next, "
           "not either level on its own, so those two columns are set half "
           "a row lower than the rest — each value sits at the border "
           "between the two rows it comes from, the same convention the "
           "slides use for ΔQ, ΔL and MPL. The MC of 5.8 on the border "
           "between the 11,000 and 12,000 rows says that each additional "
           "bat costs $5.80 as output moves from 11,000 to 12,000.",
           prefix="Note:")

    body(doc, "The two figures below plot the cost curves (not required for "
              "full credit).", before=8)

    S.place(fig_bats_totals(), doc, before=6, after=2)
    caption(doc, "Figure 1 – TC, TVC, TFC, TR and profit")

    S.place(fig_bats_averages(), doc, before=8, after=2)
    caption(doc, "Figure 2 – AFC, AVC, ATC, MC and MR")

    p = part(doc, "b", 5)
    run(p, "From Table 1, profits are highest at an output of 12,000, which "
           "is therefore the optimal output. At that output, profits equal "
           "$8,480.")

    p = body(doc)
    run(p, "If production could be adjusted more finely, we would want to "
           "produce slightly ")
    run(p, "less", bold=True)
    run(p, " than 12,000. At the optimum MR = MC. Moving from 12,000 to "
           "13,000 costs ")
    ins(p, "$7.20 per bat (the MC on the border between those two rows), "
           "against an MR of $6.40")
    dele(p, "MR = $6.40 < $7.20 = MC")
    run(p, ", so the next block of output would contribute negatively to "
           "profit, and the exact optimum lies somewhat below 12,000.")

    p = part(doc, "c", 10)
    run(p, "We reconstruct the same table at a price of $5.50, adjusting the "
           "TR and MR columns (Table 2 below). At $5.50 the optimal output "
           "is 11,000, where profits are highest — that is, least negative.")

    p = body(doc)
    run(p, "At that output ATC = $5.70, which is above the price of $5.50, "
           "so the firm makes a loss after paying fixed costs. Nevertheless "
           "AVC = $3.40 is below the price, so ")
    # PROPOSED 2: the short-run verb
    dele(p, "the firm should operate")
    ins(p, "the firm should continue to produce")
    run(p, ", because doing so contributes to covering its fixed costs.")

    bats_table(doc, 5.5)
    caption(doc, "Table 2 – costs and revenues at a price of $5.50",
            before=4, after=8)

    # PROPOSED 1: the stale TR column
    p = body(doc)
    ins(p, "Note that the total-revenue column is now 5.50 × Q. In an "
           "earlier version of this key it still showed the $6.40 revenues, "
           "which contradicted the profit column on every row.")

    # ======================================================================
    # Problem 4
    # ======================================================================
    problem(doc, 4, "Basketball Production", 30)
    draws_on(doc, "Module 4 – Profit Maximization of a Price Taker in the "
                  "Short Run")

    p = part(doc, "a", 8)
    run(p, "Marginal cost under the old technology is")
    T.equation(doc, MC_O + mrun("=") + mrun("2") + mrun("+") + Q,
               before=3, after=4)
    body(doc, "To obtain optimal output, set marginal cost equal to the "
              "price, which for a price taker is also marginal revenue:",
         after=2)
    T.equation(doc, MC_O + mrun("=") + P + mrun("=") + MR + mrun("=")
               + mrun("10")
               + mrun("  ⇒  ") + mrun("2") + mrun("+") + Q + mrun("=")
               + mrun("10") + mrun("  ⇒  ") + Q_O + mrun("=") + mrun("8"),
               before=2, after=5)
    body(doc, "At this production volume,", after=2)
    T.equation(doc, TC_O + mrun("=") + mrun("50") + mrun("+") + mrun("2")
               + mrun("⋅") + mrun("8") + mrun("+") + mrun("0.5") + mrun("⋅")
               + msup(mrun("8"), mrun("2", italic=False)) + mrun("=")
               + mrun("98"), before=2, after=5)
    body(doc, "so profits are:", after=2)
    answer(doc, PI_O + mrun("=") + P + mrun("⋅") + Q_O + mrun("−") + TC_O
           + mrun("=") + mrun("80") + mrun("−") + mrun("98") + mrun("=")
           + mrun("−18"))

    body(doc,
         "To choose the profit-maximizing output, one should compare "
         "marginal revenue to marginal cost. In perfect competition MR "
         "equals the market price, so here MR = 10 for every unit produced. "
         "If MR is greater than MC, producing an additional basketball "
         "increases profit, so the firm should expand output. However, MC "
         "rises as production increases, so it is not optimal to produce "
         "indefinitely. Eventually MC catches up with MR. The point where "
         "MC = MR is exactly where the last unit produced adds just as much "
         "to revenue as it adds to cost. Beyond that point MC exceeds MR, "
         "and the extra units would reduce profit.")

    p = part(doc, "b", 3)
    run(p, "Profit is maximized where MR and MC intersect. Because Balding "
           "is a price taker, MR equals the market price of 10. To show the "
           "area representing profit we also need the average total cost "
           "curve:")
    T.equation(doc, ATC_O + mrun("=") + mfrac(TC_O, Q) + mrun("=")
               + mfrac(mrun("50"), Q) + mrun("+") + mrun("2") + mrun("+")
               + mrun("0.5") + Q, before=3, after=5)
    body(doc, "The graph below plots the exact functions; an illustration "
              "would have been fine.")

    S.place(fig_balding(False), doc, before=6, after=2)
    caption(doc, "Old technology. At Q* = 8 the price of $10 is below "
                 "ATC* = $12.25, so the shaded rectangle is a loss of "
                 "(10 − 12.25) × 8 = −$18.")

    p = part(doc, "c", 8)
    run(p, "Despite making negative profits, Balding should continue to "
           "produce: AVC is below the price, so production makes a positive "
           "contribution to covering fixed costs. Plugging Q = 8 into "
           "average variable cost:")
    T.equation(doc, AVC + mrun("=") + mfrac(mrun("2") + Q + mrun("+")
                                            + mrun("0.5")
                                            + msup(Q, mrun("2",
                                                           italic=False)), Q)
               + mrun("=") + mrun("2") + mrun("+") + mrun("0.5") + Q,
               before=3, after=4)
    answer(doc, AVC + mrun("(") + mrun("8") + mrun(")") + mrun("=")
           + mrun("2") + mrun("+") + mrun("4") + mrun("=") + mrun("6")
           + mrun("  <  ") + P + mrun("=") + mrun("10"))

    p = part(doc, "d", 8)
    run(p, "Marginal cost under the new technology is")
    T.equation(doc, MC_N + mrun("=") + mrun("0.2") + Q, before=3, after=4)
    body(doc, "Setting it equal to the price:", after=2)
    T.equation(doc, mrun("0.2") + Q + mrun("=") + mrun("10") + mrun("  ⇒  ")
               + Q_N + mrun("=") + mrun("50"), before=2, after=5)
    body(doc, "At this production volume,", after=2)
    T.equation(doc, TC_N + mrun("=") + mrun("100") + mrun("+") + mrun("0.1")
               + mrun("⋅") + msup(mrun("50"), mrun("2", italic=False))
               + mrun("=") + mrun("350"), before=2, after=5)
    body(doc, "so profits are:", after=2)
    answer(doc, PI_N + mrun("=") + P + mrun("⋅") + Q_N + mrun("−") + TC_N
           + mrun("=") + mrun("500") + mrun("−") + mrun("350") + mrun("=")
           + mrun("150"))

    p = part(doc, "e", 3)
    run(p, "Again profit is maximized where MR = MC, with MR equal to the "
           "market price of 10. The average total cost curve is now:")
    T.equation(doc, ATC_N + mrun("=") + mfrac(TC_N, Q) + mrun("=")
               + mfrac(mrun("100"), Q) + mrun("+") + mrun("0.1") + Q,
               before=3, after=5)

    S.place(fig_balding(True), doc, before=6, after=2)
    caption(doc, "New technology. At Q* = 50 the price of $10 is above "
                 "ATC* = $7, so the shaded rectangle is a profit of "
                 "(10 − 7) × 50 = $150.")

    # ======================================================================
    # Problem 5
    # ======================================================================
    problem(doc, 5, "Taxing Sugary Drinks", 25)
    draws_on(doc, "Module 4 – Externalities;  Market Distortions and "
                  "Regulations (taxes, tax incidence)")

    p = part(doc, "a", 10)
    run(p, "An externality is a cost or benefit that affects a third party "
           "not directly involved in an economic transaction. The two "
           "parties directly involved are consumers and sellers of sugary "
           "beverages. However, consuming sugary beverages is linked to "
           "diabetes and obesity, and these conditions impose additional "
           "healthcare costs on society. There is therefore an external "
           "cost — a negative externality — associated with consuming sugary "
           "beverages.")

    body(doc,
         "The goal of the tax is to internalize this externality, so that "
         "marginal cost encompasses both the internal marginal cost (the "
         "producers’ own) and the societal marginal cost. The tax per ounce "
         "should thus equal the external marginal cost: the dollar value of "
         "the externality created by consuming one additional ounce.")

    p = body(doc)
    run(p, "According to the article, diabetes cost the U.S. $327 billion in "
           "health care and lost productivity in 2017, and the average "
           "economic cost per person was around $13,240 for diagnosed "
           "diabetes in 2017 (")
    S.link(p, "source", NCBI)
    run(p, "). A recent study indicates that replacing one daily serving of "
           "a sugary beverage with water, coffee or tea was linked with a "
           "2–10% lower risk of diabetes (")
    S.link(p, "summary of studies", HSPH)
    run(p, ").")

    body(doc,
         "Assume a daily serving is the size of a regular Coca-Cola can, "
         "12 oz. Reducing sugary-beverage intake by 12 oz per day then "
         "reduces the risk of diabetes by 2 to 10%. Being conservative, "
         "assume it reduces the risk by 2%. The cost of one additional "
         "ounce is then:")
    answer(doc, mfrac(mrun("0.02") + mrun("×") + mrun("13,240"),
                      mrun("12") + mrun("×") + mrun("365")) + mrun("=")
           + mrun("$0.06"))
    body(doc,
         "The numerator is the yearly cost to society of increasing sugary "
         "drink consumption by 12 oz per day; the denominator is the total "
         "number of ounces this represents per year. Under these "
         "assumptions, the tax should be about 6 cents per ounce.")

    S.note(doc,
           "If students assume the tax is proportional to the price and "
           "write an argument consistent with that assumption, full credit "
           "is given.", prefix="Grading:")

    p = part(doc, "b", 8)
    run(p, "When the tax is added, the supply curve shifts up and to the "
           "left by the amount of the tax. The price paid by buyers rises "
           "from ")
    S.sub_inline(p, "P", "0")
    run(p, " to ")
    S.sub_inline(p, "P", "B")
    run(p, ", the price received by sellers falls to ")
    S.sub_inline(p, "P", "S")
    run(p, ", and the quantity traded falls from ")
    S.sub_inline(p, "Q", "0")
    run(p, " to ")
    S.sub_inline(p, "Q", "1")
    run(p, ". Note that the buyers’ price rises by ")
    run(p, "less", italic=True)
    run(p, " than the full tax.")

    S.place(fig_tax(), doc, before=8, after=2)
    caption(doc, "A tax levied on producers shifts the supply curve from S "
                 "to S′. The wedge between the price buyers pay and the "
                 "price sellers receive is the tax t.")

    p = part(doc, "c", 7)
    run(p, "The tax burden falls on both retailers and consumers, but the "
           "relative extent depends on the elasticities of supply and "
           "demand. If demand is highly elastic in absolute value, consumers "
           "are very responsive to a change in price, and a small percentage "
           "increase in price leads to a large percentage decrease in "
           "quantity; the burden then falls mostly on retailers. If instead "
           "demand is inelastic, consumers are not very responsive and are "
           "still willing to consume nearly as much after the price "
           "increase; the burden then falls mostly on consumers.")

    p = body(doc)
    run(p, "For instance, the ")
    S.link(p, "article", PBS)
    run(p, " (optional reading) indicates that the burden of the tax is "
           "likely to fall mainly on individuals living in low-income "
           "neighborhoods, “where residents do not have the option of "
           "shopping at stores outside of the city limits in order to avoid "
           "the tax”. Due to the lack of substitutes their demand is "
           "inelastic, and they are thus likely to bear a large part of the "
           "burden.")

    body(doc,
         "The answer is exactly the same if the tax is instead levied on "
         "consumers. Whether it is the consumers or the producers who send "
         "the check to the government does not affect who ultimately bears "
         "the burden. That depends only on the supply and demand "
         "elasticities.")

    S.save(doc, OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
