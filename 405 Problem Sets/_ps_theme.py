"""Problem-set layer on top of the teaching-note Word theme.

The engine is `_tn_theme.py` in
`405 Slide Revisions 2026/Teaching Notes/` -- palette, masthead, cream
card, native `Fig` figures, PIL text measurement, OMML and the
`ins` / `dele` tracked-change helpers.  Nothing there is modified; this
module only adds what a problem set needs on top:

  * `ps_masthead`  -- masthead plus the "covers Module N" line
  * `problem`      -- the numbered problem heading with its points chip
  * `part`         -- an (a) / (b) / (c) part label with its own points
  * `sd_axes`      -- a supply-and-demand panel: axes, arrow tips, titles
  * `sd_fig`       -- the whole panel wrapper used by both documents

Colour convention (Teaching CLAUDE.md, 2026-08-30): a DEMAND curve is
dark red `C00000`, curve and label alike; SUPPLY stays navy.  Guides are
thin dashed gray; equilibrium dots are gold.

Run nothing here directly -- `_build_PS1.py` and `_build_PS1_Solutions.py`
import it.
"""

import os
import re
import sys

_THEME_DIR = os.path.abspath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..",
    "405 Slide Revisions 2026", "Teaching Notes"))
if not os.path.isfile(os.path.join(_THEME_DIR, "_tn_theme.py")):
    raise ImportError(
        "_tn_theme.py not found at {}.  The problem sets reuse the teaching-"
        "note theme; if that folder moved, update _THEME_DIR."
        .format(_THEME_DIR))
sys.path.insert(0, _THEME_DIR)

import _tn_theme as T                                       # noqa: E402
from _tn_theme import (CREAM, DARKRED, GOLD, GRAY, LIGHT,   # noqa: E402,F401
                       NAVY, PALEGOLD, WHITE, Inches, Pt,
                       WD_ALIGN_PARAGRAPH, body, dele, equation,
                       equation_inline, ins, para, para_inserted, run,
                       text_width)

# The problem sets carry no term or year, so the same file is reusable
# next year -- same reasoning as the teaching notes' bare page number.
COURSE_TITLE = T.COURSE_TITLE


# --------------------------------------------------------------------------
# Chrome
# --------------------------------------------------------------------------
BYLINE = "Prof. Nico Voigtländer"


def ps_masthead(doc, subtitle, covers=None, due=None, byline=BYLINE):
    """Calendar-style masthead, plus an optional coverage / due line.

    `covers` names the modules the set draws on; `due` stays generic
    ("see the Class Website") so no date is baked into the file.

    The document names ITSELF first: "Problem Set N" is the large,
    CENTRED first line (2026-09-08, Nico -- this swapped the two header
    lines, which used to lead with the course title).  The course line
    follows underneath at subtitle size, with the byline at its right,
    above the gold rule, the way a handout header carries it.
    """
    p = para(doc, before=0, after=1)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, subtitle, bold=True, color=NAVY, size=20)

    p2 = para(doc, before=0, after=9)
    sec = doc.sections[0]
    width = sec.page_width - sec.left_margin - sec.right_margin
    p2.paragraph_format.tab_stops.add_tab_stop(
        width, T.WD_TAB_ALIGNMENT.RIGHT)
    run(p2, T.COURSE_TITLE, bold=True, color=NAVY, size=13)
    if byline:
        run(p2, "	")
        run(p2, byline, bold=True, color=GRAY, size=11)
    T.gold_rule(p2, size_eighths=18, space=6)
    if covers or due:
        p = para(doc, before=0, after=10)
        # `covers` / `due` are each either a plain string or a list of
        # (text, kind) pieces with kind in (None, "ins", "dele"), so a
        # wording change inside the meta line can be proposed as a
        # TRACKED revision instead of being applied silently.
        emit = {None: run, "ins": T.ins, "dele": T.dele}
        pieces = []
        for bit in (covers, due):
            if not bit:
                continue
            if pieces:
                pieces.append(("     ·     ", None))
            pieces.extend([(bit, None)] if isinstance(bit, str) else bit)
        for text, kind in pieces:
            emit[kind](p, text, color=GRAY, size=10.5, italic=True)
    return doc


def problem(doc, number, title, points, before=12):
    # before 17 -> 12 and part() 8/4 -> 6/3 on 2026-09-12.  PS 1
    # page 1 had been left with 3.1 pt of clearance by the new navy
    # banner, so any renderer that lays out a hair taller than this
    # one spilled a line onto page 2 and pushed Problem 4(c) to a
    # third page.  The document carried 4.53" of pure paragraph
    # spacing; this returns ~40 pt of it as slack.
    """Numbered problem heading: navy bold title, gold points chip.

    The points sit in their own right-aligned tab so the headings line up
    down the page instead of drifting with the title length.
    """
    p = para(doc, before=before, after=4, keep_next=True)
    sec = doc.sections[0]
    width = sec.page_width - sec.left_margin - sec.right_margin
    p.paragraph_format.tab_stops.add_tab_stop(
        width, T.WD_TAB_ALIGNMENT.RIGHT)
    run(p, "Problem {}.  {}".format(number, title),
        bold=True, color=NAVY, size=13)
    run(p, "\t")
    run(p, "({} points)".format(points), bold=True, color=GOLD, size=11)
    T.gold_rule(p, size_eighths=6, color=LIGHT, space=2)
    return p


def part(doc, letter, points=None, before=6, after=3, bonus=False, level=0):
    """An (a) / (b) / (c) label with its own point count, as a run-in head.

    `level=1` indents one step for a sub-part -- (i) / (ii) under a
    lettered part -- keeping the same hanging indent so the wrapped text
    lines up under the text, not under the label.

    Returns the paragraph so the caller keeps adding runs to it.
    """
    left = 0.30 + 0.34 * level
    p = para(doc, before=before, after=after, left=left, hang=0.30,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, "({})  ".format(letter), bold=True, color=NAVY, size=11)
    if points is not None:
        word = "point" if points == 1 else "points"
        label = "({} bonus {})  " if bonus else "({} {})  "
        run(p, label.format(points, word), bold=True, color=GOLD, size=10.5)
    return p


def quote(doc, text, attribution, width_in=6.5, size=11):
    """A verbatim quote in the cream card: italic quote, bold attribution.

    Teaching CLAUDE.md, "Quote callout": named executives, court
    opinions and press quotations get this treatment rather than being
    run into the body text.
    """
    def pop(cell):
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run(p, "“" + text + "”", italic=True, color=NAVY, size=size)
        a = cell.add_paragraph()
        a.paragraph_format.space_before = Pt(0)
        a.paragraph_format.space_after = Pt(0)
        a.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run(a, "— " + attribution, bold=True, color=NAVY,
            size=size - 1)

    return T.cream_card(doc, pop, width_in=width_in, before=7, after=7)


SUBMIT_LEAD = "Please upload "
SUBMIT_ONE = "one"
SUBMIT_TAIL = (" solution per group on the course BruinLearn site or email "
               "the group solution to the TA.")

SUBMIT_TA = ("Please address all questions about this Problem Set to the "
             "TA Rafael Macedo")


def submit_card(doc, width_in=6.5, size=10.0, height_in=None, ta_line=True):
    """The navy submission banner, above the AI-policy card (2026-09-12).

    Same rounded / shadowed card as `policy_card`, filled NAVY with white
    text, so the two read as one stacked pair at the top of page 1.  The
    word "one" is underlined -- the point of the line is that a group
    submits a single solution, not one per member.

    NO hyperlink on "BruinLearn": the same file is handed to both the EMBA
    and the FEMBA section, and the two have different BruinLearn sites, so
    a link would be wrong for one of them.

    `size` is 10 pt, not the policy card's 10.5: at 10.5 the sentence
    measures 6.23" against the card's 6.26" of inner width, which is
    inside `_measure_par`'s slack and would size the card for two lines
    while Word sets one.  At 10 pt it is 5.93" and fits with room.

    `height_in` is FIXED for the same reason `policy_card`'s is, plus a
    second one: `style_run` writes `<w:b w:val="0"/>` and `<w:i
    w:val="0"/>` on every plain run, and `_measure_par` tests only for
    the ELEMENT, so it measures every card in BOLD ITALIC Calibri.  That
    is 4 % wide here -- enough to report two lines for a sentence Word
    sets on one, which made the banner 0.55" tall and pushed Problem 2(c)
    onto a third page.  One 10 pt line is 0.169"; 0.34" holds it plus the
    shape's 0.12" of top and bottom inset.

    `ta_line` adds `SUBMIT_TA` as a centred BOLD second line, naming the
    TA to write to.  It went onto Problem Set 1 on 2026-09-12 and onto
    ALL FIVE the next day, so it is the DEFAULT -- pass `ta_line=False`
    only for a problem set that should not carry it.  The line measures
    4.43" bold against 6.26" of inner width, so it sets on one line.
    Leave `height_in` at None: the card is sized for one line or two from
    `ta_line`, since `_measure_par` cannot tell them apart.
    """
    # One Calibri line is 1.22 x size, plus the shape's 0.12" of top and
    # bottom inset and a little slack.  Computed, not guessed, because
    # `_measure_par` cannot be trusted here (see the docstring).
    if height_in is None:
        lines = 2 if ta_line else 1
        height_in = round(lines * 1.22 * size / 72.0 + 0.12 + 0.05, 3)

    def pop(cell):
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run(p, SUBMIT_LEAD, color=WHITE, size=size)
        run(p, SUBMIT_ONE, color=WHITE, size=size, underline=True)
        run(p, SUBMIT_TAIL, color=WHITE, size=size)

        if ta_line:
            q = cell.add_paragraph()
            q.paragraph_format.space_before = Pt(0)
            q.paragraph_format.space_after = Pt(0)
            q.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run(q, SUBMIT_TA, color=WHITE, size=size, bold=True)

    return T.cream_card(doc, pop, fill=NAVY, border=NAVY,
                        width_in=width_in, height_in=height_in,
                        before=5, after=0)


POLICY_HEADING = "Important information, please read carefully"

POLICY_BODY = (
    "For problem sets, you may use AI tools (e.g. ChatGPT) to help "
    "brainstorm or to revise existing work you have written. When you "
    "submit your problem set, we expect you to clearly attribute what text "
    "was generated by the AI tool (e.g., AI-generated text appears in a "
    "different colored font, quoted directly in the text, or use an "
    "in-text parenthetical citation). We will screen for undisclosed "
    "AI-generated content using advanced AI-detection programs. While AI "
    "tools can provide helpful insights, we want you to think critically "
    "about the information you receive. Think of AI tools as a study "
    "partner who can help you dive deeper into the concepts behind each "
    "problem set exercise. You will have to prove your understanding in "
    "the midterm and final exams, where the use of AI tools will be "
    "prohibited."
)


def policy_card(doc, width_in=6.5, size=10.5, height_in=1.9375):
    """The AI-tools policy, in Nico's own wording (2026-09-08).

    Two paragraphs inside the cream card: a centred heading line in BOLD
    ITALIC DARK RED, then the policy itself with a bold navy prefix.  The
    heading is the one place a problem set uses dark red for emphasis
    rather than for a demand curve -- it is a warning label, and he set it
    that way by hand.

    `height_in` is FIXED at the height he set by hand (2026-09-08,
    resized from the 2.197" `cream_card` measured for itself).  The
    measurement in `_measure_par` is conservative -- it assumes more
    wrapped lines than Word actually sets at this width -- so the card
    came out with a band of empty cream under the last line.  If the
    policy wording ever changes, re-measure against a Word render rather
    than trusting the estimate again.
    """
    def pop(cell):
        h = cell.add_paragraph()
        h.paragraph_format.space_before = Pt(0)
        h.paragraph_format.space_after = Pt(0)
        h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run(h, POLICY_HEADING, bold=True, italic=True, color=DARKRED,
            size=size)

        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run(p, "AI tools policy for problem sets: ", bold=True, color=NAVY,
            size=size)
        run(p, POLICY_BODY, color=NAVY, size=size)

    return T.cream_card(doc, pop, width_in=width_in, height_in=height_in,
                        before=6, after=6)


LINK_BLUE = "1F5AA8"


def link(p, text, url, size=11, color=LINK_BLUE, underline=True, **kw):
    """An external hyperlink run, keeping the original documents' links.

    The problem sets point at a PBS article and two research sources;
    dropping the URLs would lose information the student is told to go
    and read.
    """
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    rid = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = T.OxmlElement("w:hyperlink")
    h.set(T.qn("r:id"), rid)
    p._p.append(h)
    r = run(p, text, size=size, color=color, underline=underline, **kw)
    h.append(r._r)
    return r


ROW_BLUE = "0070C0"        # Module 7's row-player colour
COL_GOLD = GOLD            # Module 7's column-player colour


def payoff_matrix(doc, row_player, col_player, row_strats, col_strats,
                  payoffs, best_row=(), best_col=(), size=12,
                  widths=(1.55, 1.75, 1.75)):
    """A 2x2 payoff matrix in Module 7's house style.

    Teaching CLAUDE.md, "Rebuilding game-theory / payoff-matrix decks":
    the column player's name sits centred on top in the accent colour,
    the row player's name at the left in the reserved concept colour,
    strategy labels around a grid of WHITE cells with navy borders, and
    each cell reads "a , b" with the two players' payoffs in their own
    colours.  A caption below reads "Payoffs to (Row, Col)".

    `best_row` / `best_col` are (r, c) pairs whose row / column payoff is
    a best response; those numbers are underlined, which is how a Word
    document says what the deck says with a circle.  A cell in BOTH sets
    is a Nash equilibrium and takes the pale-gold fill as well.
    """
    n_r, n_c = len(row_strats), len(col_strats)
    tbl = doc.add_table(rows=n_r + 2, cols=n_c + 1)
    tbl.alignment = T.WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    T._strip_table_style(tbl)
    for c, w in enumerate(widths):
        tbl.columns[c].width = Inches(w)

    def put(r, c, runs, align=WD_ALIGN_PARAGRAPH.CENTER, fill=None):
        cell = tbl.cell(r, c)
        cell.width = Inches(widths[c])
        T._cell_margins(cell, 0.06)
        if fill:
            T._shade(cell, fill)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.alignment = align
        for t, pr in runs:
            kw = dict(size=size, color=NAVY)
            kw.update(pr)
            run(p, t, **kw)
        return cell

    # header band: the column player's name, spanning the payoff columns
    top = tbl.cell(0, 1).merge(tbl.cell(0, n_c))
    for extra in list(top.paragraphs)[1:]:
        extra._p.getparent().remove(extra._p)
    p = top.paragraphs[0]
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run(p, col_player, bold=True, color=COL_GOLD, size=size + 1)

    for c, s in enumerate(col_strats):
        put(1, c + 1, [(s, dict(bold=True))])
    put(1, 0, [("", {})])

    for r, s in enumerate(row_strats):
        runs = ([(row_player + "     ", dict(bold=True, color=ROW_BLUE,
                                             size=size + 1))] if r == 0
                else [])
        put(r + 2, 0, runs + [(s, dict(bold=True))],
            align=WD_ALIGN_PARAGRAPH.RIGHT)
        for c in range(n_c):
            a, b = payoffs[r][c]
            nash = (r, c) in best_row and (r, c) in best_col
            # A best response is UNDERLINED -- the Word equivalent of the
            # deck's circled number -- and a cell where both are best
            # responses (a Nash equilibrium) also takes the pale-gold fill.
            put(r + 2, c + 1,
                [(str(a), dict(bold=True, color=ROW_BLUE,
                               underline=(r, c) in best_row)),
                 ("  ,  ", {}),
                 (str(b), dict(bold=True, color=COL_GOLD,
                               underline=(r, c) in best_col))],
                fill=T.PALEGOLD if nash else WHITE)

    _payoff_borders(tbl, n_r, n_c)

    cap = para(doc, before=5, after=8)
    cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(cap, "Payoffs to (", color=NAVY, size=10)
    run(cap, row_player, bold=True, color=ROW_BLUE, size=10)
    run(cap, ", ", color=NAVY, size=10)
    run(cap, col_player, bold=True, color=COL_GOLD, size=10)
    run(cap, ")", color=NAVY, size=10)
    return tbl


def _payoff_borders(tbl, n_r, n_c):
    """Navy borders on the payoff cells only; the label band stays open."""
    for r in range(2, n_r + 2):
        for c in range(1, n_c + 1):
            tcPr = tbl.cell(r, c)._tc.get_or_add_tcPr()
            bd = T.OxmlElement("w:tcBorders")
            for side in ("top", "start", "bottom", "end"):
                e = T.OxmlElement("w:" + side)
                e.set(T.qn("w:val"), "single")
                e.set(T.qn("w:sz"), "12")
                e.set(T.qn("w:color"), NAVY)
                bd.append(e)
            tcPr.append(bd)


def cell_runs(tbl, r, c, runs, size=10, bold=False, color=NAVY):
    """Rewrite one table cell as explicit runs.

    `T.table` takes plain strings, but the course convention wants a real
    subscript on an indexed symbol -- italic letter, subscript index --
    never a Unicode subscript character.  This drops the cell's existing
    runs (keeping its fill, borders and paragraph properties) and writes
    the given (text, props) list in their place.
    """
    cell = tbl.cell(r, c)
    p = cell.paragraphs[0]
    for old in list(p._p.findall(T.qn("w:r"))):
        p._p.remove(old)
    for t, pr in runs:
        kw = dict(size=size, bold=bold, color=color)
        kw.update(pr)
        run(p, t, **kw)
    return cell


def sub_inline(p, base, idx, size=11, **kw):
    """An indexed symbol inside running text: italic letter, true
    subscript index -- P0, PB, Q1 and the like."""
    run(p, base, italic=True, size=size, **kw)
    run(p, idx, italic=True, subscript=True, size=size, **kw)
    return p


def sub_runs(base, idx, **pr):
    """(base, idx) as an indexed symbol: italic letter, true subscript."""
    a = dict(italic=True)
    a.update(pr)
    b = dict(a)
    b["subscript"] = True
    return [(base, a), (idx, b)]


def note(doc, text, prefix="Note:"):
    """The house cream aside, at the problem sets' width."""
    return T.callout(doc, prefix, text, width_in=6.5, size=10.5,
                     before=7, after=7)


def caption(doc, text, before=2, after=10):
    p = para(doc, before=before, after=after)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, text, italic=True, color=GRAY, size=9)
    return p


# --------------------------------------------------------------------------
# Supply-and-demand panels
# --------------------------------------------------------------------------
CURVE_W = 1.75          # pt, curve stroke
AXIS_W = 1.25           # pt, axis stroke
GUIDE_W = 1.00          # pt, dashed guide stroke
GUIDE = GRAY
DOT = GOLD

Y_TITLE_GAP = 0.08      # y-axis title clears the axis by this much
X_TITLE_GAP = 0.05      # x-axis title sits this far below the axis
TIP = 0.18              # arrowhead length beyond the plot bound


class Panel:
    """One supply-and-demand diagram on a logical 0-100 x 0-100 grid.

    Everything is addressed in logical units, so a panel can be moved or
    resized without re-deriving a single coordinate.  `x()` / `y()` map
    logical to figure inches.
    """

    def __init__(self, fig, ox, oy, w, h, ylabel="P", xlabel="Q"):
        self.f = fig
        self.ox, self.oy = ox, oy          # origin, in figure inches
        self.w, self.h = w, h              # plot extent, in figure inches
        self.xtip = ox + w + TIP
        self.ytip = oy - h - TIP
        self.axes(ylabel, xlabel)

    # -- coordinate transform ---------------------------------------------
    def x(self, u):
        return self.ox + self.w * u / 100.0

    def y(self, v):
        return self.oy - self.h * v / 100.0

    # -- chrome ------------------------------------------------------------
    def axes(self, ylabel, xlabel):
        f = self.f
        f.line(self.ox, self.oy, self.xtip, self.oy, color=NAVY,
               w_pt=AXIS_W, arrow=True, name="x-axis")
        f.line(self.ox, self.oy, self.ox, self.ytip, color=NAVY,
               w_pt=AXIS_W, arrow=True, name="y-axis")

        # Axis titles are anchored to the ARROW TIP in a box sized tight
        # to the label (Teaching CLAUDE.md, 2026-08-30).
        pr = dict(bold=True, italic=True, color=NAVY, size=12)
        yt = [(ylabel, pr)]
        wy = text_width(yt) + 0.06
        f.label(self.ox - Y_TITLE_GAP - wy, self.ytip - 0.11, yt,
                w=wy, h=0.22, align="l", name="y-axis title")
        xt = [(xlabel, pr)]
        wx = text_width(xt) + 0.06
        f.label(self.xtip - wx / 2.0, self.oy + X_TITLE_GAP, xt,
                w=wx, h=0.22, align="l", name="x-axis title")

    def title(self, text, size=10.5, dy=0.30):
        """A small navy panel heading, centred over the plot."""
        pr = dict(bold=True, color=NAVY, size=size)
        runs = [(text, pr)]
        w = text_width(runs) + 0.10
        self.f.label(self.ox + self.w / 2.0, self.ytip - dy, runs,
                     w=w, h=0.24, align="c", name="panel title")

    # -- curves ------------------------------------------------------------
    def curve(self, p0, p1, color, label=None, dash=None, lbl_dx=0.10,
              lbl_dy=-0.09, name="curve", w_pt=None, lbl_end=1, lbl_pr=None):
        """A straight curve between two logical points, labelled at its end.

        The label sits just past the end, in the curve's own colour, in a
        box measured to the text.  `lbl_end` picks WHICH end: 1 (the
        default) is `p1`, the lower end of a downward-sloping curve; 0 is
        `p0`, for a curve whose lower end is too crowded to label -- a
        reaction function in q1-q2 space, say, where a label set beside
        the line is cut through by it.  Either way the anchor comes from
        the curve itself, never from a typed coordinate.
        """
        (u0, v0), (u1, v1) = p0, p1
        self.f.line(self.x(u0), self.y(v0), self.x(u1), self.y(v1),
                    color=color, w_pt=w_pt or CURVE_W, dash=dash, name=name)
        if label:
            pr = dict(bold=True, color=color, size=11)
            pr.update(lbl_pr or {})
            runs = _label_runs(label, pr)
            w = text_width(runs) + 0.08
            au, av = (p1, p0)[lbl_end == 0]
            self.f.label(self.x(au) + lbl_dx, self.y(av) + lbl_dy, runs,
                         w=w, h=0.22, align="l", name=name + " label")

    def demand(self, p0, p1, label="D", name="D", **kw):
        return self.curve(p0, p1, DARKRED, label=label, name=name, **kw)

    def supply(self, p0, p1, label="S", name="S", **kw):
        return self.curve(p0, p1, NAVY, label=label, name=name, **kw)

    # -- annotation --------------------------------------------------------
    def equilibrium(self, u, v, label=None, dx=0.0, dy=-0.30, guides=True,
                    ptick=None, qtick=None, lbl_align="c"):
        """A dot at (u, v), with optional dashed guides and axis ticks."""
        if guides:
            self.f.line(self.x(u), self.y(v), self.x(u), self.oy,
                        color=GUIDE, w_pt=GUIDE_W, dash="dash", name="Q guide")
            self.f.line(self.ox, self.y(v), self.x(u), self.y(v),
                        color=GUIDE, w_pt=GUIDE_W, dash="dash", name="P guide")
        self.f.dot(self.x(u), self.y(v), d=0.075, color=DOT, name="equilibrium")
        if label:
            # CENTRED above the dot (2026-09-08).  Nico re-placed all six
            # E0 / E1 labels in the Problem 3(c) panels by hand and every
            # one of them landed within 0.07" of centred-above -- the
            # alternating left / right offsets this used before pushed the
            # label away from the point it names.
            pr = dict(bold=True, italic=True, color=NAVY, size=11)
            runs = _label_runs(label, pr)
            w = text_width(runs) + 0.08
            self.f.label(self.x(u) + dx, self.y(v) + dy, runs, w=w, h=0.22,
                         align=lbl_align, name="equilibrium label")
        if ptick:
            self.tick_y(v, ptick)
        if qtick:
            self.tick_x(u, qtick)

    def dot_open(self, u, v, d=0.085, color=DARKRED, w_pt=1.5,
                 name="open dot"):
        """A HOLLOW dot: white fill, coloured ring.

        `Fig.dot` is solid and `_tn_theme` is not to be modified, so the
        shape is emitted here.  Used to mark the end of an interval that
        is NOT the base, against the solid dot that is (2026-09-13).
        """
        self.f.shapes.append(
            '<wps:wsp><wps:cNvPr id="{i}" name="{n}"/><wps:cNvSpPr/>'
            '<wps:spPr><a:xfrm><a:off x="{x}" y="{y}"/>'
            '<a:ext cx="{d}" cy="{d}"/></a:xfrm>'
            '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            '<a:ln w="{w}"><a:solidFill><a:srgbClr val="{c}"/></a:solidFill>'
            '</a:ln></wps:spPr><wps:bodyPr/></wps:wsp>'.format(
                i=next(self.f._id), n=name,
                x=T.emu(self.x(u) - d / 2), y=T.emu(self.y(v) - d / 2),
                d=T.emu(d), w=int(round(w_pt * 12700)), c=color))

    def span(self, p0, p1, color=DARKRED, w_pt=4.0, name="interval"):
        """Redraw a stretch of a curve thicker, to mark an INTERVAL.

        Emitted BEFORE the dots and labels that sit on it, so it cannot
        paint over them -- the figure rules put shaded and highlighted
        things first, then curves, then labels.
        """
        self.f.line(self.x(p0[0]), self.y(p0[1]),
                    self.x(p1[0]), self.y(p1[1]),
                    color=color, w_pt=w_pt, name=name)

    def tick_y(self, v, label, size=10.5, h=0.22, italic=True, color=NAVY):
        """A tick label on the y-axis, right-aligned onto the axis.

        `h` shrinks the box when two ticks sit close together: the box
        hangs symmetrically about the tick, so two labels 0.19" apart
        would overlap at the default 0.22" height but not at 0.17".
        """
        pr = dict(bold=True, italic=italic, color=color, size=size)
        runs = _label_runs(label, pr)
        w = text_width(runs) + 0.08
        # Pass the ANCHOR, not a pre-shifted x: Fig.label already moves a
        # right-aligned box left by its own width.  Subtracting `w` here as
        # well put every y tick one box-width too far left (2026-09-08:
        # Nico dragged the w0 tick in Problem 4's figure back to exactly
        # ox - Y_TITLE_GAP - w, which is what this line always intended).
        self.f.label(self.ox - Y_TITLE_GAP, self.y(v) - h / 2.0, runs,
                     w=w, h=h, align="r", name="y tick")

    def tick_x(self, u, label, size=10.5, h=0.22, italic=True, color=NAVY):
        pr = dict(bold=True, italic=italic, color=color, size=size)
        runs = _label_runs(label, pr)
        w = text_width(runs) + 0.08
        # same double-shift as tick_y, half a box width here
        self.f.label(self.x(u), self.oy + X_TITLE_GAP, runs,
                     w=w, h=h, align="c", name="x tick")

    def arrow(self, p0, p1, color=GOLD, w_pt=1.25, name="shift arrow"):
        """An arrow from p0 to p1, drawn as an explicit path.

        NOT `Fig.line(arrow=True)`: that positions the line with flipH /
        flipV, and on a doubly-flipped (up-and-left) arrow Word puts the
        head on the wrong end -- the Problem 4 supply shift pointed
        down-right instead of up-left.  A `custGeom` addresses both
        endpoints inside an unflipped box, so the head always lands on
        the point the arrow is aimed at, in any direction.
        """
        x1, y1 = self.x(p0[0]), self.y(p0[1])
        x2, y2 = self.x(p1[0]), self.y(p1[1])
        ox, oy = min(x1, x2), min(y1, y2)
        cw = max(abs(x2 - x1), 0.01)
        ch = max(abs(y2 - y1), 0.01)
        W, H = T.emu(cw), T.emu(ch)

        def pt(px, py):
            return ('<a:pt x="{x}" y="{y}"/>'
                    .format(x=int(round((px - ox) / cw * W)),
                            y=int(round((py - oy) / ch * H))))

        self.f.shapes.append(
            '<wps:wsp><wps:cNvPr id="{i}" name="{n}"/><wps:cNvSpPr/>'
            '<wps:spPr>'
            '<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{W}" cy="{H}"/></a:xfrm>'
            '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
            '<a:rect l="0" t="0" r="{W}" b="{H}"/>'
            '<a:pathLst><a:path w="{W}" h="{H}">'
            '<a:moveTo>{a}</a:moveTo><a:lnTo>{b}</a:lnTo>'
            '</a:path></a:pathLst></a:custGeom>'
            '<a:noFill/>'
            '<a:ln w="{lw}" cap="rnd"><a:solidFill><a:srgbClr val="{c}"/>'
            '</a:solidFill><a:round/>'
            '<a:tailEnd type="triangle" w="med" len="med"/></a:ln>'
            '</wps:spPr><wps:bodyPr/></wps:wsp>'.format(
                i=next(self.f._id), n=name, x=T.emu(ox), y=T.emu(oy),
                W=W, H=H, a=pt(x1, y1), b=pt(x2, y2),
                lw=int(w_pt * 12700), c=color))

    def polyline(self, pts, color=NAVY, w_pt=None, dash=None, name="series"):
        """One editable freeform through a list of logical points.

        Used for a curve that is defined by tabulated values (a marginal
        product schedule, a total-cost series) or by a formula evaluated
        on a grid (an average-cost curve).  Emitted as a SINGLE custGeom
        shape rather than a chain of separate segments, so it can be
        selected and nudged in Word as one object, and so no rounding
        gap opens between segments.
        """
        xs = [self.x(u) for u, _ in pts]
        ys = [self.y(v) for _, v in pts]
        ox, oy = min(xs), min(ys)
        cw = max(max(xs) - ox, 0.01)
        ch = max(max(ys) - oy, 0.01)
        W, H = T.emu(cw), T.emu(ch)

        def pt(px, py):
            return ('<a:pt x="{x}" y="{y}"/>'
                    .format(x=int(round((px - ox) / cw * W)),
                            y=int(round((py - oy) / ch * H))))

        path = "<a:moveTo>{}</a:moveTo>".format(pt(xs[0], ys[0]))
        for px, py in zip(xs[1:], ys[1:]):
            path += "<a:lnTo>{}</a:lnTo>".format(pt(px, py))

        ln = ('<a:ln w="{lw}" cap="rnd"><a:solidFill><a:srgbClr val="{c}"/>'
              '</a:solidFill>'.format(lw=int((w_pt or CURVE_W) * 12700),
                                      c=color))
        if dash:
            ln += '<a:prstDash val="{}"/>'.format(dash)
        ln += "<a:round/></a:ln>"

        self.f.shapes.append(
            '<wps:wsp><wps:cNvPr id="{i}" name="{n}"/><wps:cNvSpPr/>'
            '<wps:spPr>'
            '<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{W}" cy="{H}"/></a:xfrm>'
            '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
            '<a:rect l="0" t="0" r="{W}" b="{H}"/>'
            '<a:pathLst><a:path w="{W}" h="{H}">{path}</a:path></a:pathLst>'
            '</a:custGeom><a:noFill/>{ln}</wps:spPr>'
            '<wps:bodyPr/></wps:wsp>'.format(
                i=next(self.f._id), n=name, x=T.emu(ox), y=T.emu(oy),
                W=W, H=H, path=path, ln=ln))

    def polygon(self, pts, fill=DARKRED, alpha=26000, line=None,
                name="area"):
        """A shaded closed area -- a surplus or deadweight-loss region.

        Build the polygon from the CURVES' own points, so a sloped edge
        coincides with the demand or supply line exactly rather than
        being eyeballed alongside it (Teaching CLAUDE.md: "A surplus
        triangle's hypotenuse IS the curve it lies on").
        """
        xs = [self.x(u) for u, _ in pts]
        ys = [self.y(v) for _, v in pts]
        ox, oy = min(xs), min(ys)
        cw = max(max(xs) - ox, 0.01)
        ch = max(max(ys) - oy, 0.01)
        W, H = T.emu(cw), T.emu(ch)

        def pt(px, py):
            return ('<a:pt x="{x}" y="{y}"/>'
                    .format(x=int(round((px - ox) / cw * W)),
                            y=int(round((py - oy) / ch * H))))

        path = "<a:moveTo>{}</a:moveTo>".format(pt(xs[0], ys[0]))
        for px, py in zip(xs[1:], ys[1:]):
            path += "<a:lnTo>{}</a:lnTo>".format(pt(px, py))
        path += "<a:close/>"

        lnxml = ('<a:ln w="9525"><a:solidFill><a:srgbClr val="{}"/>'
                 '</a:solidFill></a:ln>'.format(line) if line
                 else '<a:ln><a:noFill/></a:ln>')
        self.f.shapes.append(
            '<wps:wsp><wps:cNvPr id="{i}" name="{n}"/><wps:cNvSpPr/>'
            '<wps:spPr>'
            '<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{W}" cy="{H}"/></a:xfrm>'
            '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
            '<a:rect l="0" t="0" r="{W}" b="{H}"/>'
            '<a:pathLst><a:path w="{W}" h="{H}">{path}</a:path></a:pathLst>'
            '</a:custGeom>'
            '<a:solidFill><a:srgbClr val="{f}"><a:alpha val="{a}"/>'
            '</a:srgbClr></a:solidFill>{ln}</wps:spPr>'
            '<wps:bodyPr/></wps:wsp>'.format(
                i=next(self.f._id), n=name, x=T.emu(ox), y=T.emu(oy),
                W=W, H=H, path=path, f=fill, a=alpha, ln=lnxml))

    def region(self, u0, v0, u1, v1, fill=DARKRED, alpha=22000, line=None,
               name="region"):
        """A shaded rectangle -- the profit / loss area on a cost panel.

        Teaching CLAUDE.md, 2026-08-30: positive profit is dark red
        `C00000`, a loss the firm still produces through is gray
        `555B66`, and a loss large enough to stop production is a DARKER
        gray, each at about 20-25 % so the curves stay readable through
        it.  Pass `fill` accordingly; the caller writes "Profit" or
        "Loss" into the box with `Panel.text`.
        """
        x0, x1 = sorted((self.x(u0), self.x(u1)))
        y0, y1 = sorted((self.y(v0), self.y(v1)))
        lnxml = ('<a:ln w="9525"><a:solidFill><a:srgbClr val="{}"/>'
                 '</a:solidFill></a:ln>'.format(line) if line
                 else '<a:ln><a:noFill/></a:ln>')
        self.f.shapes.append(
            '<wps:wsp><wps:cNvPr id="{i}" name="{n}"/><wps:cNvSpPr/>'
            '<wps:spPr><a:xfrm><a:off x="{x}" y="{y}"/>'
            '<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="{f}"><a:alpha val="{a}"/>'
            '</a:srgbClr></a:solidFill>{ln}</wps:spPr>'
            '<wps:bodyPr/></wps:wsp>'.format(
                i=next(self.f._id), n=name, x=T.emu(x0), y=T.emu(y0),
                cx=max(T.emu(x1 - x0), 1), cy=max(T.emu(y1 - y0), 1),
                f=fill, a=alpha, ln=lnxml))

    def legend(self, entries, u, v, size=9.5, row=0.175, swatch=0.20,
               pad=0.07):
        """A legend badge inside the plot: white fill, thin navy border.

        For a chart whose series converge -- Problem Set 3's average-cost
        panel has ATC, MR and AVC within one logical unit of each other
        at the right-hand edge -- labelling each curve at its own end is
        not possible, and the deck's answer is a legend placed in a clear
        corner of the plot (Teaching CLAUDE.md, "Chart legends": stacked
        one per line, packed close, white fill, thin primary border).

        `u, v` is the badge's TOP-LEFT in logical coordinates.
        """
        wmax = max(text_width([(lab, dict(bold=True, size=size))])
                   for lab, _ in entries)
        w = pad * 2 + swatch + 0.06 + wmax
        h = pad * 2 + row * len(entries)
        x0, y0 = self.x(u), self.y(v)

        self.f.shapes.append(
            '<wps:wsp><wps:cNvPr id="{i}" name="legend"/><wps:cNvSpPr/>'
            '<wps:spPr><a:xfrm><a:off x="{x}" y="{y}"/>'
            '<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            '<a:ln w="9525"><a:solidFill><a:srgbClr val="{b}"/>'
            '</a:solidFill></a:ln></wps:spPr>'
            '<wps:bodyPr/></wps:wsp>'.format(
                i=next(self.f._id), x=T.emu(x0), y=T.emu(y0),
                cx=T.emu(w), cy=T.emu(h), b=NAVY))

        for k, (lab, color) in enumerate(entries):
            cy = y0 + pad + row * (k + 0.5)
            self.f.line(x0 + pad, cy, x0 + pad + swatch, cy, color=color,
                        w_pt=2.0, name="legend mark")
            runs = [(lab, dict(bold=True, color=color, size=size))]
            self.f.label(x0 + pad + swatch + 0.06, cy - row / 2.0, runs,
                         w=wmax + 0.06, h=row, align="l",
                         name="legend label")

    def text(self, u, v, runs, align="l", h=0.22):
        """A free label anchored at (u, v).

        `align` decides which edge of the box the anchor is: "l" its left,
        "c" its centre, "r" its right.  The shift is left to Fig.label --
        doing it here as well moved every centred label half a box width
        off its anchor.
        """
        w = text_width(runs) + 0.08
        self.f.label(self.x(u), self.y(v), runs, w=w, h=h, align=align,
                     name="note")


def _label_runs(label, pr):
    """Accept a plain string, a (base, subscript) pair, or a run list.

    A pair sets the index as a true subscript -- the course convention for
    P0, Q1, D', S2 and the like.
    """
    if isinstance(label, list):
        return label
    if isinstance(label, tuple):
        base, sub = label
        sp = dict(pr)
        sp["subscript"] = True
        return [(base, pr), (sub, sp)]
    return [(label, pr)]


def centroid(pts):
    """Area centroid of a simple polygon, in the same logical units.

    Teaching CLAUDE.md: a CS / PS label "sits INSIDE its own triangle
    (place it at the triangle's centroid, computed -- not eyeballed)".
    This is the computation, and it handles the four-sided producer
    surplus region as readily as a triangle.
    """
    a = cx = cy = 0.0
    n = len(pts)
    for i in range(n):
        x0, y0 = pts[i]
        x1, y1 = pts[(i + 1) % n]
        cross_ = x0 * y1 - x1 * y0
        a += cross_
        cx += (x0 + x1) * cross_
        cy += (y0 + y1) * cross_
    if abs(a) < 1e-12:                       # degenerate: fall back to mean
        return (sum(x for x, _ in pts) / n, sum(y for _, y in pts) / n)
    a *= 0.5
    return (cx / (6.0 * a), cy / (6.0 * a))


def line_at(p0, p1, u):
    """v on the straight line through logical p0 and p1 at abscissa u."""
    (u0, v0), (u1, v1) = p0, p1
    return v0 + (v1 - v0) * (u - u0) / float(u1 - u0)


def cross(a0, a1, b0, b1):
    """Logical intersection of two straight lines -- computed, never eyeballed.

    Every marked equilibrium in these documents comes from here, so a dot
    always sits exactly where the two curves actually meet.
    """
    (x1, y1), (x2, y2) = a0, a1
    (x3, y3), (x4, y4) = b0, b1
    den = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)
    if abs(den) < 1e-9:
        raise ValueError("parallel lines have no intersection")
    a = x1 * y2 - y1 * x2
    b = x3 * y4 - y3 * x4
    return ((a * (x3 - x4) - (x1 - x2) * b) / den,
            (a * (y3 - y4) - (y1 - y2) * b) / den)


def fig(w, h, name="Figure"):
    return T.Fig(w, h, name=name)


_SHAPE_XFRM = re.compile(
    r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>')


def content_box(f):
    """Bounding box of everything actually drawn in `f`, in inches.

    The Fig canvas is declared up front, before the labels are measured,
    so it almost always carries slack.  This reads the real extent back
    out of the emitted shapes.
    """
    xs, ys, xe, ye = [], [], [], []
    for shp in f.shapes:
        m = _SHAPE_XFRM.search(shp)
        if not m:
            continue
        x, y, cx, cy = (int(g) for g in m.groups())
        xs.append(x)
        ys.append(y)
        xe.append(x + cx)
        ye.append(y + cy)
    if not xs:
        return (0.0, 0.0, f.w, f.h)
    E = float(T.EMU_IN)
    return (min(xs) / E, min(ys) / E, max(xe) / E, max(ye) / E)


def place(f, doc, before=6, after=6, pad=0.0, crop=True):
    """Place a figure, CROPPED to what it actually draws.

    Nico tightened the group box on every figure in the Problem Set 1
    solutions by hand (2026-09-08): each one had been emitted at its
    declared canvas size with a band of dead margin around the content,
    which pushed the figures apart on the page and left the captions
    floating.  This does the same thing arithmetically -- the group's
    off / ext / chOff / chExt are set to the content's own bounding box,
    at 1:1 scale, so nothing inside moves or stretches.

    `pad` is 0 because his boxes measure the ink bounding box exactly
    (checked against every figure in his file: 0.040" -- twice the 0.02"
    pad an earlier version added -- taller and wider than his).  Those
    2.9 pt per figure accumulated down the page and pushed two lines of
    the store-brand bullet onto a seventh page.
    """
    if not crop:
        return f.place(doc, before=before, after=after)

    x0, y0, x1, y1 = content_box(f)
    x0, y0 = max(0.0, x0 - pad), max(0.0, y0 - pad)
    x1, y1 = x1 + pad, y1 + pad
    cx, cy = T.emu(x1 - x0), T.emu(y1 - y0)

    xml = (
        '<w:drawing {nsd}>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="{did}" name="{name}"/>'
        '<wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData'
        ' uri="http://schemas.microsoft.com/office/word/2010/'
        'wordprocessingGroup">'
        '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm>'
        '<a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/>'
        '<a:chOff x="{ox}" y="{oy}"/><a:chExt cx="{cx}" cy="{cy}"/>'
        '</a:xfrm></wpg:grpSpPr>{shapes}</wpg:wgp>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing>'
    ).format(nsd=T.nsdecls("w", "wp", "a", "wps", "wpg"),
             cx=cx, cy=cy, ox=T.emu(x0), oy=T.emu(y0),
             did=101, name=f.name, shapes="".join(f.shapes))

    p = para(doc, before=before, after=after)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run()._r.append(T.parse_xml(xml))
    return p


# ---------------------------------------------------------------- saving ----
# python-docx's default template declares Word's 2010 layout engine
# (`compatibilityMode` 14).  Word 2013 changed line breaking and
# justification, so a document built at 14 wraps identical text one line
# longer than the same file after Nico opens and saves it in Word (which
# upgrades it to 15).  That drift moved figures across page boundaries and
# added a page to the PS1 solutions.  Save through this helper so every
# problem set is laid out by the same engine Word itself will use.
_COMPAT_15 = (
    '<w:compatSetting w:name="compatibilityMode" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
    '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '<w:compatSetting w:name="enableOpenTypeFeatures" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '<w:compatSetting w:name="doNotFlipMirrorIndents" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '<w:compatSetting w:name="differentiateMultirowTableHeaders" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '<w:compatSetting w:name="useWord2013TrackBottomHyphenation" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="0"/>'
)


def save(doc, path):
    """Save `doc` to `path` with Word's 2013+ layout engine declared."""
    doc.save(path)
    _set_compat_15(path)
    return path


def _set_compat_15(path):
    import shutil
    import tempfile
    import zipfile

    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        parts = {n: z.read(n) for n in names}

    s = parts["word/settings.xml"].decode("utf-8")
    body_ = re.sub(r"<w:compat>.*?</w:compat>",
                   "<w:compat><w:useFELayout/>" + _COMPAT_15 + "</w:compat>",
                   s, count=1, flags=re.S)
    if body_ == s:                       # no <w:compat> block at all
        body_ = s.replace(
            "</w:settings>",
            "<w:compat><w:useFELayout/>" + _COMPAT_15 + "</w:compat>"
            "</w:settings>")
    if body_ == s:
        raise RuntimeError("could not set compatibilityMode in " + path)
    parts["word/settings.xml"] = body_.encode("utf-8")

    fd, tmp = tempfile.mkstemp(suffix=".docx",
                               dir=os.path.dirname(os.path.abspath(path)))
    os.close(fd)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, parts[n])
    shutil.move(tmp, path)

# ------------------------------------------------- change-based tables ----
# A marginal value is a statement about the STEP from one row to the next,
# not about either row on its own, so it belongs BETWEEN them.  The slides
# do this literally: on Module 3 slide 19 the DQ / DL / MPL columns of the
# MPL table are left EMPTY and each value is a floating box centred on the
# border between the two rows it comes from.
#
# In a Word table the equivalent is to drop the CHANGE COLUMNS half a row
# (2026-09-09, Nico -- an earlier version wove a thin row in between every
# pair of level rows, which doubled the number of horizontal rules and
# looked bad).  The value stays in its own row -- row `i` still holds the
# step from `i` to `i + 1`, as it always did -- and only its position
# inside the cell moves, so the table keeps its row count, its height and
# its banding, and gains no extra lines at all.
#
# The arithmetic that makes the offset exactly half a row: the text must
# still fit inside its own cell, so with no top or bottom cell margin,
# `row_h / 2 + line <= row_h`, i.e. `row_h >= 2 * line`.  A Calibri line
# box is about 1.22x the point size -- NOT the point size itself, which is
# the mistake that first put these rows at 0.24" and would have clipped
# every offset value at 8 pt (0.12 + 0.136 = 0.256 > 0.24).
LINE_FACTOR = 1.22


def change_row_h(size, slack=0.02):
    """The row height that leaves room for a half-row offset at `size` pt."""
    return 2.0 * LINE_FACTOR * size / 72.0 + slack


def _set_cell_margins(cell, top, side, bottom=None):
    """Replace a cell's margins.  `T._cell_margins` APPENDS a `w:tcMar`."""
    def dxa(v):
        return dict(w=int(round(v * 1440)), type="dxa")
    return _tcpr_put(cell, "tcMar", children=(
        ("top", dxa(top)), ("start", dxa(side)),
        ("bottom", dxa(top if bottom is None else bottom)),
        ("end", dxa(side))))


def change_table(doc, head, rows, change_cols, widths_in, size=9.5,
                 align_right=(), highlight=(), row_h=None):
    """A table whose change-based COLUMNS sit half a row lower.

    `rows` are the ordinary data rows, with each change value on the row
    it starts from (row `i` carries the step from `i` to `i + 1`, and the
    last row's change cells are blank -- there is no next row).
    `change_cols` are the column indices to drop by half a row, so each
    value lands on the border it describes.

    Everything else about the table is unchanged: same rows, same cream /
    white banding, same borders, no extra rules.
    """
    if row_h is None:
        row_h = change_row_h(size)
    tbl = T.table(doc, [head] + rows, widths_in=widths_in, size=size,
                  align_right=align_right, highlight=highlight)
    for r in range(1, len(rows) + 1):
        trPr = tbl.rows[r]._tr.get_or_add_trPr()
        for old in trPr.findall(T.qn("w:trHeight")):
            trPr.remove(old)
        h = T.OxmlElement("w:trHeight")
        # atLeast, never exact: if a cell ever needs more room Word grows
        # the row instead of clipping the offset value out of sight.
        h.set(T.qn("w:hRule"), "atLeast")
        h.set(T.qn("w:val"), str(int(round(row_h * 1440))))
        trPr.append(h)
        for c in change_cols:
            cell = tbl.cell(r, c)
            # no top / bottom margin, so the half-row push still leaves the
            # line inside the cell; the side margins stay as T.table set them
            _set_cell_margins(cell, 0.0, 0.085)
            pf = cell.paragraphs[0].paragraph_format
            pf.space_before = Pt(row_h / 2.0 * 72.0)
            pf.space_after = Pt(0)
    return tbl


# Light grey behind a marginal value (2026-09-10, Nico).  A step lighter
# than the palette's rule grey `C8CDD3`, which at full-cell coverage sits
# heavy under 8 pt navy digits; this keeps the change column clearly
# distinct from the cream / white banding without darkening the numbers.
CHANGE_GREY = "E8EBEE"


# The convention students are asked to follow when a problem has them
# compute a change-based series (2026-09-09, Nico).  Same wording in every
# problem that asks for one, so it reads as a course convention rather
# than a per-question instruction.  The second half matters as much as the
# first: a student working in Excel cannot drop a column half a row, and
# the point of the convention is that the reader can tell which interval
# each change belongs to.
CHANGE_CONVENTION = (
    "As in class, a change — a marginal product, a marginal cost, a "
    "marginal revenue — describes the STEP from one row to the next, not "
    "either row on its own. In the table above, the change columns (at the "
    "right) are therefore set half a row lower than the rest, so each "
    "value sits at the border between the two rows it comes from, shaded "
    "to fade from the one into the other. If your software "
    "cannot offset a column, label each change with the interval it "
    "refers to instead (for example “MC, 0 → 1,000”), so it is "
    "unambiguous which step each number belongs to."
)


def change_note(doc):
    """The half-row convention, as the house cream aside."""
    return note(doc, CHANGE_CONVENTION, prefix="Convention:")


ELASTICITY_CONVENTION = (
    "An elasticity, like any other change, describes an INTERVAL — here the "
    "move from one price-quantity pair to the other. To turn that change "
    "into a percentage we need a base, and we always use the START of the "
    "interval, (P\u2080, Q\u2080). So name the interval alongside the elasticity, "
    "exactly as you would for a marginal cost: \u201cE\u1d05, $45 \u2192 $40 = \u22121.50\u201d. "
    "On a straight-line demand curve this is also the elasticity AT the "
    "starting point, so the two readings agree."
)


def elasticity_note(doc):
    """The interval / base convention, as the house cream aside.

    The sibling of `change_note`.  Both say the same underlying thing: a
    change belongs to an interval.  They differ only in what is done with
    it -- a marginal value is PLOTTED at the middle of its interval, an
    elasticity is DIVIDED by the start of it.
    """
    return note(doc, ELASTICITY_CONVENTION, prefix="Convention:")


# ------------------------------------------- staggered (merged) tables ----
# 2026-09-10, Nico: build the between-rows offset out of MERGED half-rows
# rather than out of a paragraph indent.  Every output level gets TWO rows:
#
#   * a LEVEL column merges the pair and centres its number in the middle
#     of it, so the number sits where its row's single line used to;
#   * a CHANGE column merges one row LOWER -- the bottom half of one level
#     with the top half of the next -- so its number is centred exactly on
#     the border between the two levels it is derived from.
#
# Same reading as the `space_before` offset in `change_table`, but the
# value is genuinely centred in a cell of its own instead of pushed to the
# bottom of one, so Word's own vertical centring keeps it in place when a
# row grows, and the change column's cell boundaries fall half a level out
# of step with the rest -- which is the offset, made visible.
#
# The stagger leaves two half-cells over: the top half of the first level
# and the bottom half of the last, both empty, both shaded to continue the
# change column's own (offset) banding.
def half_row_h(size, slack=0.012):
    """Height of ONE half-row: a single Calibri line at `size` pt.

    Half of `change_row_h`, near enough -- see LINE_FACTOR above for why
    the line box is 1.22x the point size and not the point size.
    """
    return LINE_FACTOR * size / 72.0 + slack


# `w:tcPr`'s children are ORDERED by the schema.  Appending happens to
# work for shd / tcMar / vAlign because Word is tolerant, but `w:tcBorders`
# has to sit before `w:shd`, and a border Word decides to ignore is exactly
# the kind of defect that costs a round to find.  So put every hand-built
# element at its schema position.
_TCPR_ORDER = ("cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge",
               "tcBorders", "shd", "noWrap", "tcMar", "textDirection",
               "tcFitText", "vAlign", "hideMark")


def _tcpr_put(cell, tag, children=(), **attrs):
    """Create-or-replace `w:<tag>` in this cell's tcPr, in schema order."""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(T.qn("w:" + tag)):
        tcPr.remove(old)
    el = T.OxmlElement("w:" + tag)
    for k, v in attrs.items():
        el.set(T.qn("w:" + k), str(v))
    for ct, cattrs in children:
        ce = T.OxmlElement("w:" + ct)
        for k, v in cattrs.items():
            ce.set(T.qn("w:" + k), str(v))
        el.append(ce)
    after = _TCPR_ORDER[_TCPR_ORDER.index(tag) + 1:]
    for existing in tcPr:
        name = existing.tag.split("}")[-1]
        if name in after:
            existing.addprevious(el)
            return el
    tcPr.append(el)
    return el


def _valign(cell, how="center"):
    """Vertically centre a cell's text."""
    return _tcpr_put(cell, "vAlign", val=how)


def _shade(cell, fill):
    """Shade a cell in schema order.

    `T._shade` APPENDS, which lands `w:shd` after the `w:tcMar` this
    module writes first and leaves the `tcPr` out of the order the schema
    declares.  Word tolerates that, but `w:tcBorders` has to precede
    `w:shd`, so once one element is placed properly they all should be.
    """
    return _tcpr_put(cell, "shd", val="clear", color="auto", fill=fill)


def _no_hrules_tc(tr, col, top=True, bottom=True):
    """Drop the horizontal rules of the `col`-th `w:tc` of this row.

    Takes the raw `w:tr` rather than a `_Cell`: a vertically merged
    region is ONE `_Cell` in python-docx, so `tbl.cell(r, c)` cannot
    address the continue rows, and it is the continue row's own bottom
    border that Word actually draws at the foot of a merged region.

    Vertical borders are left alone, so the column still reads as a
    column; only its interior horizontal rules go.
    """
    tcs = tr.findall(T.qn("w:tc"))
    if col >= len(tcs):
        return
    tc = tcs[col]
    tcPr = tc.find(T.qn("w:tcPr"))
    if tcPr is None:
        tcPr = T.OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    for old in tcPr.findall(T.qn("w:tcBorders")):
        tcPr.remove(old)
    bd = T.OxmlElement("w:tcBorders")
    for side, on in (("top", top), ("bottom", bottom)):
        if not on:
            continue
        e = T.OxmlElement("w:" + side)
        e.set(T.qn("w:val"), "nil")
        bd.append(e)
    if len(bd) == 0:
        return
    after = _TCPR_ORDER[_TCPR_ORDER.index("tcBorders") + 1:]
    for existing in tcPr:
        if existing.tag.split("}")[-1] in after:
            existing.addprevious(bd)
            return
    tcPr.append(bd)


def _pmark_size(p, size):
    """Set the paragraph MARK's font size.

    A line box is as tall as the TALLEST thing on it, and the paragraph
    mark counts: left at the style's 11 pt it makes every row 0.186" high
    whatever the runs are set to, which is taller than a half-row.  An
    empty cell has nothing BUT its mark, so this is the only thing setting
    its height at all.

    `w:rPr` is the last-but-two child of `w:pPr` in schema order, so call
    this AFTER the paragraph's spacing and alignment are set.
    """
    pPr = p._p.get_or_add_pPr()
    for old in pPr.findall(T.qn("w:rPr")):
        pPr.remove(old)
    rPr = T.OxmlElement("w:rPr")
    for tag in ("w:sz", "w:szCs"):
        e = T.OxmlElement(tag)
        e.set(T.qn("w:val"), str(int(round(size * 2))))
        rPr.append(e)
    pPr.append(rPr)



def _keep_together_merged(tbl):
    """Keep a vertically merged table on one page.

    `T._keep_table_together` walks `row.cells`, and python-docx maps a
    whole vMerge region to the RESTART cell -- so the continue cells'
    own `w:p` elements never get `w:keepNext`, and every one of them is a
    page-break opportunity.  On the 35-row Bats table that left 198 of
    385 paragraphs flagged and Word broke the table across two pages.

    So walk the raw XML: `w:cantSplit` on every row, `w:keepNext` on
    every paragraph outside the LAST row.  `w:keepNext` is the second
    element of `w:pPr` in schema order, and none of these paragraphs
    carries a `w:pStyle`, so it goes at the front.
    """
    trs = tbl._tbl.findall(T.qn("w:tr"))
    for i, tr in enumerate(trs):
        trPr = tr.find(T.qn("w:trPr"))
        if trPr is None:
            trPr = T.OxmlElement("w:trPr")
            tr.insert(0, trPr)
        for old in trPr.findall(T.qn("w:cantSplit")):
            trPr.remove(old)
        trPr.insert(0, T.OxmlElement("w:cantSplit"))
        if i == len(trs) - 1:
            continue
        for tc in tr.findall(T.qn("w:tc")):
            for para_el in tc.findall(T.qn("w:p")):
                pPr = para_el.find(T.qn("w:pPr"))
                if pPr is None:
                    pPr = T.OxmlElement("w:pPr")
                    para_el.insert(0, pPr)
                for old in pPr.findall(T.qn("w:keepNext")):
                    pPr.remove(old)
                pPr.insert(0, T.OxmlElement("w:keepNext"))


def _flatten_continue_cells(tbl, size, side_margin):
    """Zero out the paragraphs of every vMerge CONTINUE cell.

    They are invisible -- the merged region's text lives in the restart
    cell -- but they still SET THE ROW HEIGHT.  Left at the Normal
    style's 11 pt paragraph mark plus its 8 pt space-after, each one is
    about 0.30" tall, so with `hRule="atLeast"` every row grows to fit
    it: the Bats table measured 7.74" against the 5.0" its declared
    half-rows come to.  python-docx cannot reach these cells (a whole
    vMerge region is one `_Cell`), so this walks the raw XML.
    """
    for tr in tbl._tbl.findall(T.qn("w:tr")):
        for tc in tr.findall(T.qn("w:tc")):
            vm = tc.find(T.qn("w:tcPr") + "/" + T.qn("w:vMerge"))
            if vm is None or (vm.get(T.qn("w:val")) or "continue") != "continue":
                continue
            tcPr = tc.find(T.qn("w:tcPr"))
            for old in tcPr.findall(T.qn("w:tcMar")):
                tcPr.remove(old)
            mar = T.OxmlElement("w:tcMar")
            for name, v in (("top", 0.0), ("start", side_margin),
                            ("bottom", 0.0), ("end", side_margin)):
                e = T.OxmlElement("w:" + name)
                e.set(T.qn("w:w"), str(int(round(v * 1440))))
                e.set(T.qn("w:type"), "dxa")
                mar.append(e)
            after = _TCPR_ORDER[_TCPR_ORDER.index("tcMar") + 1:]
            for existing in tcPr:
                if existing.tag.split("}")[-1] in after:
                    existing.addprevious(mar)
                    break
            else:
                tcPr.append(mar)
            for para_el in tc.findall(T.qn("w:p")):
                pPr = para_el.find(T.qn("w:pPr"))
                if pPr is None:
                    pPr = T.OxmlElement("w:pPr")
                    para_el.insert(0, pPr)
                for old in pPr.findall(T.qn("w:spacing")):
                    pPr.remove(old)
                sp = T.OxmlElement("w:spacing")
                sp.set(T.qn("w:before"), "0")
                sp.set(T.qn("w:after"), "0")
                sp.set(T.qn("w:line"), "240")
                sp.set(T.qn("w:lineRule"), "auto")
                # w:spacing sits after w:keepNext and before w:rPr
                rpr = pPr.find(T.qn("w:rPr"))
                if rpr is not None:
                    rpr.addprevious(sp)
                else:
                    pPr.append(sp)
                for old in pPr.findall(T.qn("w:rPr")):
                    pPr.remove(old)
                rPr = T.OxmlElement("w:rPr")
                for tag in ("w:sz", "w:szCs"):
                    e = T.OxmlElement(tag)
                    e.set(T.qn("w:val"), str(int(round(size * 2))))
                    rPr.append(e)
                pPr.append(rPr)


# A change box's fill FADES between the two levels it is derived from --
# its top stop is the band above it, its bottom stop the band below
# (2026-09-10, Nico: "illustrate that it stems from the values in those
# two boxes").  A Word table cell cannot do that: `w:shd` is a solid fill
# plus an optional pattern and there is no gradient in `CT_Shd`.  So the
# cell is left plain and an INLINE SHAPE inside it carries both the
# gradient and the number.
_GRAD_XML = (
    '<w:drawing {nsd}>'
    '<wp:inline distT="0" distB="0" distL="0" distR="0">'
    '<wp:extent cx="{cx}" cy="{cy}"/>'
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
    '<wp:docPr id="{did}" name="{name}"/>'
    '<wp:cNvGraphicFramePr/>'
    '<a:graphic><a:graphicData'
    ' uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
    '<wps:wsp><wps:cNvSpPr/><wps:spPr bwMode="auto">'
    '<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:srgbClr val="{top}"/></a:gs>'
    '<a:gs pos="100000"><a:srgbClr val="{bottom}"/></a:gs>'
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    '<a:ln><a:noFill/></a:ln>'
    '</wps:spPr><wps:txbx><w:txbxContent>{para}</w:txbxContent></wps:txbx>'
    '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0"'
    ' rIns="0" bIns="0" anchor="ctr" anchorCtr="0"><a:noAutofit/>'
    '</wps:bodyPr></wps:wsp>'
    '</a:graphicData></a:graphic></wp:inline></w:drawing>'
)

_GRAD_PARA = (
    '<w:p><w:pPr>'
    '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
    '<w:jc w:val="center"/>'
    '<w:rPr><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    '</w:pPr><w:r><w:rPr>'
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
    '<w:color w:val="{color}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    '</w:rPr><w:t xml:space="preserve">{txt}</w:t></w:r></w:p>'
)

_grad_id = [3000]


def _gradient_box(cell, txt, w_in, h_in, top, bottom, size, color=NAVY):
    """Put a vertical white/cream fade, with `txt` centred on it, in `cell`.

    The shape fills the WHOLE cell: this zeroes the cell's own side
    margins, because a fade that stops short of the borders reads as a
    small tinted panel floating in a white cell rather than as the cell's
    own fill (2026-09-10, Nico).  `w_in` / `h_in` are therefore the full
    cell, less a few thousandths of an inch so an inline drawing can
    never be the thing that grows the row.

    The cell's paragraph holds nothing but the shape, at exact line
    spacing, so it cannot push the row past the two half-rows the merged
    cell already spans.
    """
    _set_cell_margins(cell, 0.0, 0.0)
    _grad_id[0] += 1
    para = _GRAD_PARA.format(sz=int(round(size * 2)), color=color,
                             txt=T.esc(txt) if hasattr(T, "esc") else txt)
    xml = _GRAD_XML.format(
        nsd=T.nsdecls("w", "wp", "a", "wps"),
        cx=T.emu(w_in), cy=T.emu(h_in), did=_grad_id[0],
        name="change %d" % _grad_id[0], top=top, bottom=bottom, para=para)

    p = cell.paragraphs[0]
    for r in list(p._p.findall(T.qn("w:r"))):
        p._p.remove(r)
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(h_in * 72.0)
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _pmark_size(p, size)
    p.add_run()._r.append(T.parse_xml(xml))

_GRAD_ANCHOR_XML = (
    '<w:drawing {nsd}>'
    # behindDoc: the number is real cell text and is painted ON TOP.
    # layoutInCell: without it Word positions the shape against the PAGE
    # and every fade lands in the same spot outside the table.
    '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"'
    ' relativeHeight="{z}" behindDoc="1" locked="0" layoutInCell="1"'
    ' allowOverlap="1">'
    '<wp:simplePos x="0" y="0"/>'
    '<wp:positionH relativeFrom="column">'
    '<wp:posOffset>{dx}</wp:posOffset></wp:positionH>'
    '<wp:positionV relativeFrom="paragraph">'
    '<wp:posOffset>{dy}</wp:posOffset></wp:positionV>'
    '<wp:extent cx="{cx}" cy="{cy}"/>'
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
    '<wp:wrapNone/>'
    '<wp:docPr id="{did}" name="{name}"/>'
    '<wp:cNvGraphicFramePr/>'
    '<a:graphic><a:graphicData'
    ' uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
    '<wps:wsp><wps:cNvSpPr/><wps:spPr bwMode="auto">'
    '<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:srgbClr val="{top}"/></a:gs>'
    '<a:gs pos="100000"><a:srgbClr val="{bottom}"/></a:gs>'
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    '<a:ln><a:noFill/></a:ln>'
    '</wps:spPr>'
    '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0"'
    ' rIns="0" bIns="0" anchor="ctr" anchorCtr="0"><a:noAutofit/>'
    '</wps:bodyPr></wps:wsp>'
    '</a:graphicData></a:graphic></wp:anchor></w:drawing>'
)


def _grad_anchor(cell, w_in, h_in, top, bottom, dx_in, dy_in):
    """A full-cell vertical fade, floating BEHIND the cell's own text.

    `dx_in` / `dy_in` place the shape's top-left on the cell's top-left:
    the horizontal anchor is the cell's text column, so dx is minus the
    left cell margin; the vertical anchor is the paragraph, which in a
    vertically-centred cell starts half the leftover height down, so dy
    is minus that.  Both are computed by the caller and then CHECKED
    against the rendered shape with Word COM -- see the module docstring
    of `_check_fade.py`.

    Returns the shape's `w:drawing` element.
    """
    _grad_id[0] += 1
    xml = _GRAD_ANCHOR_XML.format(
        nsd=T.nsdecls("w", "wp", "a", "wps"),
        cx=T.emu(w_in), cy=T.emu(h_in),
        dx=T.emu(dx_in), dy=T.emu(dy_in),
        z=251650000 + _grad_id[0],
        did=_grad_id[0], name="fade %d" % _grad_id[0],
        top=top, bottom=bottom)
    el = T.parse_xml(xml)
    # The anchor hangs off a RUN in the cell's first paragraph, placed
    # AFTER `w:pPr`.  `w:pPr` must be the first child of `w:p`: inserting
    # the run at index 0 pushes it down, and Word then ignores the
    # paragraph properties entirely -- the centring is lost and the shape
    # does not paint (2026-09-13).
    p = cell.paragraphs[0]
    r = T.OxmlElement("w:r")
    r.append(el)
    idx = 1 if p._p.find(T.qn("w:pPr")) is not None else 0
    p._p.insert(idx, r)
    return el


def stagger_table(doc, head, rows, change_cols, widths_in, size=9.5,
                  highlight=(), row_h=None, keep_together=True,
                  side_margin=0.05, change_fill=None,
                  change_rules=True, change_fade=False,
                  fade_slack=0.008, fade_inset=0.01):
    """A change-based table built from merged half-rows.

    `rows`        one list per output LEVEL, with each change value on the
                  level its step starts from (level `i` carries the step
                  `i -> i+1`; the last level's change cells are blank).
    `change_cols` the column indices to stagger.
    `highlight`   LEVEL indices to mark with the house pale-gold fill.

    Numbers are centred in their columns; the caller sizes the columns.

    The change column keeps its own boxes -- a rule above and below each
    merged cell, the same as every other box in the table but shifted
    down half a level -- filled SOLID in `change_fill`, by default the
    same cream the value boxes are banded with.  The two leftover
    half-boxes at top and bottom are left empty, since there is no step
    above the first level or below the last.

    `change_fade=True` fills each change cell with a vertical FADE
    instead, from the colour of the level the change is derived FROM to
    the colour of the level it runs TO (2026-09-13, Nico), so the cell
    says visually which two rows made it.  Word still has no gradient for
    a table cell -- `CT_Shd` is a solid fill plus an optional dither
    pattern -- so it is a shape; what changed is that the shape now
    FLOATS (`_grad_anchor`) instead of sitting inline.  An inline drawing
    lives on a text line, which is why the 2026-09-10 attempt could never
    cover the cell: exact line spacing clipped it and automatic spacing
    grew the row.  An anchored one takes no part in line layout, so it
    can be sized to the whole cell -- measured edge to edge, within
    0.3 pt of all four borders.  Two things are load-bearing: the cell
    must be shaded `auto`, since a `behindDoc` shape is painted behind
    the cell's own `w:shd`, and the run carrying the anchor must sit
    AFTER `w:pPr`.  `_gradient_box` keeps the old inline route for
    reference only.

    `change_rules=False` drops the interior rules for one continuous
    strip.
    """
    n = len(rows)
    ncols = len(widths_in)
    chg = tuple(change_cols)
    lvl = [c for c in range(ncols) if c not in set(chg)]
    if row_h is None:
        row_h = half_row_h(size)
    if change_fill is None:
        change_fill = CREAM
    hi = set(highlight)

    tbl = doc.add_table(rows=1 + 2 * n, cols=ncols)
    tbl.alignment = T.WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    T._strip_table_style(tbl)
    for c, w in enumerate(widths_in):
        tbl.columns[c].width = Inches(w)

    # MERGE FIRST, while every cell is still empty: python-docx carries the
    # paragraphs of both cells into the merged one, and a second paragraph
    # is a second line, which a half-height row has no room for.
    for i in range(n):
        r = 1 + 2 * i
        for c in lvl:
            tbl.cell(r, c).merge(tbl.cell(r + 1, c))
    for i in range(n - 1):
        r = 2 + 2 * i
        for c in chg:
            tbl.cell(r, c).merge(tbl.cell(r + 1, c))

    # before the row heights, so w:cantSplit precedes w:trHeight in trPr
    if keep_together:
        _keep_together_merged(tbl)

    def band(i):
        """The banding fill of level `i`."""
        return CREAM if i % 2 else WHITE

    def fill_of(i):
        """The fill a LEVEL actually renders in, highlight included.

        The fade runs from the fill of the level a change is derived FROM
        to the fill of the level it runs TO (2026-09-13, Nico), so it has
        to read the real colour, not just the banding.
        """
        return PALEGOLD if i in hi else band(i)

    def put(cell, txt, fill, bold=False, color=NAVY, sz=None):
        # the caller sets `cell.width` right after: a merged cell keeps one
        # grid column, so its tcW is still its own column's width
        _set_cell_margins(cell, 0.0, side_margin)
        _shade(cell, fill)
        _valign(cell)
        p = cell.paragraphs[0]
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if txt:
            run(p, txt, bold=bold, color=color, size=sz or size)
        _pmark_size(p, sz or size)

    # header: its own margins, and free to wrap onto two lines
    for c in range(ncols):
        cell = tbl.cell(0, c)
        _set_cell_margins(cell, 0.035, side_margin)
        _shade(cell, NAVY)
        _valign(cell)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run(p, head[c], bold=True, color=WHITE, size=size)
        _pmark_size(p, size)
        cell.width = Inches(widths_in[c])

    for i in range(n):
        fill = PALEGOLD if i in hi else band(i)
        for c in lvl:
            cell = tbl.cell(1 + 2 * i, c)
            txt = rows[i][c] if c < len(rows[i]) else ""
            put(cell, txt, fill, bold=i in hi)
            cell.width = Inches(widths_in[c])
        if i < n - 1:
            for c in chg:
                cell = tbl.cell(2 + 2 * i, c)
                txt = rows[i][c] if c < len(rows[i]) else ""
                if change_fade:
                    # A table cell still has no gradient of its own, so
                    # the fade is a shape -- but a FLOATING one now
                    # (2026-09-13).  `_gradient_box` put it INLINE, on a
                    # text line, which is why it never covered the cell.
                    # This writes the number as ordinary cell text and
                    # hangs the fade behind it at the cell's full size.
                    put(cell, txt, WHITE, bold=i in hi)
                    # TRANSPARENT, not white: a `behindDoc` shape is
                    # painted behind the table's own cell shading, so an
                    # opaque `w:shd` hides the fade completely (2026-09-13
                    # -- the second reason the first attempt showed
                    # nothing).  `fill="auto"` lets the shape through.
                    _shade(cell, "auto")
                    cell.width = Inches(widths_in[c])
                    _grad_anchor(cell,
                                 widths_in[c], 2 * row_h,
                                 fill_of(i), fill_of(i + 1),
                                 # minus the LEFT CELL MARGIN, which this
                                 # module's `_set_cell_margins` sets to
                                 # `side_margin` flat.  `T._cell_margins`
                                 # adds 0.04 to the sides and borrowing
                                 # that here put the shape 2.88 pt left of
                                 # the cell, where Word clipped it -- so it
                                 # fell exactly that far short on the right.
                                 -side_margin,
                                 -(2 * row_h - 1.22 * size / 72.0) / 2.0)
                else:
                    # A SOLID `w:shd` fills the cell exactly, edge to
                    # edge, with no shape involved -- which is what the
                    # box needs (2026-09-10, Nico, twice).
                    put(cell, txt, change_fill, bold=i in hi)
                    cell.width = Inches(widths_in[c])

    # The two half-boxes the stagger leaves over, top and bottom: EMPTY
    # (2026-09-10, Nico).  They are not marginal values -- there is no
    # step above the first level or below the last -- so they carry no
    # fill of their own and read as blank.
    for c in chg:
        for r in (1, 2 * n):
            cell = tbl.cell(r, c)
            put(cell, "", WHITE)
            cell.width = Inches(widths_in[c])

    # Optional: drop the change column's interior rules for a continuous
    # strip instead of boxes.  Off by default -- Nico wants the boxes.
    if not change_rules:
        for c in chg:
            for r in range(1, 1 + 2 * n):
                # NOT tbl.cell(r, c): python-docx maps a whole vMerge
                # region to ONE _Cell, so writing through it lands on the
                # restart row every time and the CONTINUE row keeps the
                # table's insideH rule -- which is the rule that shows.
                _no_hrules_tc(tbl.rows[r]._tr, c, top=(r != 1),
                              bottom=(r != 2 * n))

    for r in range(1, 1 + 2 * n):
        trPr = tbl.rows[r]._tr.get_or_add_trPr()
        for old in trPr.findall(T.qn("w:trHeight")):
            trPr.remove(old)
        h = T.OxmlElement("w:trHeight")
        h.set(T.qn("w:hRule"), "atLeast")
        h.set(T.qn("w:val"), str(int(round(row_h * 1440))))
        trPr.append(h)

    # after every cell is written: the invisible continue cells set the
    # row height, so they have to be flattened too
    _flatten_continue_cells(tbl, size, side_margin)

    T._thin_borders(tbl)
    return tbl
