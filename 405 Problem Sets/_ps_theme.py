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
def ps_masthead(doc, subtitle, covers=None, due=None):
    """Calendar-style masthead, plus an optional coverage / due line.

    `covers` names the modules the set draws on; `due` stays generic
    ("see the Course Calendar") so no date is baked into the file.
    """
    T.masthead(doc, subtitle)
    if covers or due:
        p = para(doc, before=0, after=10)
        bits = [b for b in (covers, due) if b]
        run(p, "     ·     ".join(bits), color=GRAY, size=10.5, italic=True)
    return doc


def problem(doc, number, title, points, before=17):
    """Numbered problem heading: navy bold title, gold points chip.

    The points sit in their own right-aligned tab so the headings line up
    down the page instead of drifting with the title length.
    """
    p = para(doc, before=before, after=6, keep_next=True)
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


def part(doc, letter, points=None, before=8, after=4, bonus=False):
    """An (a) / (b) / (c) label with its own point count, as a run-in head.

    Returns the paragraph so the caller keeps adding runs to it.
    """
    p = para(doc, before=before, after=after, left=0.30, hang=0.30,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, "({})  ".format(letter), bold=True, color=NAVY, size=11)
    if points is not None:
        label = "({} bonus points)  " if bonus else "({} points)  "
        run(p, label.format(points), bold=True, color=GOLD, size=10.5)
    return p


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
        self.f.label(self.ox + self.w / 2.0 - w / 2.0,
                     self.ytip - dy, runs, w=w, h=0.24, align="c",
                     name="panel title")

    # -- curves ------------------------------------------------------------
    def curve(self, p0, p1, color, label=None, dash=None, lbl_dx=0.10,
              lbl_dy=-0.09, name="curve", w_pt=None):
        """A straight curve between two logical points, labelled at its end.

        The label sits just past `p1`, in the curve's own colour, in a box
        measured to the text.
        """
        (u0, v0), (u1, v1) = p0, p1
        self.f.line(self.x(u0), self.y(v0), self.x(u1), self.y(v1),
                    color=color, w_pt=w_pt or CURVE_W, dash=dash, name=name)
        if label:
            pr = dict(bold=True, color=color, size=11)
            runs = _label_runs(label, pr)
            w = text_width(runs) + 0.08
            self.f.label(self.x(u1) + lbl_dx, self.y(v1) + lbl_dy, runs,
                         w=w, h=0.22, align="l", name=name + " label")

    def demand(self, p0, p1, label="D", name="D", **kw):
        return self.curve(p0, p1, DARKRED, label=label, name=name, **kw)

    def supply(self, p0, p1, label="S", name="S", **kw):
        return self.curve(p0, p1, NAVY, label=label, name=name, **kw)

    # -- annotation --------------------------------------------------------
    def equilibrium(self, u, v, label=None, dx=0.07, dy=-0.30, guides=True,
                    ptick=None, qtick=None):
        """A dot at (u, v), with optional dashed guides and axis ticks."""
        if guides:
            self.f.line(self.x(u), self.y(v), self.x(u), self.oy,
                        color=GUIDE, w_pt=GUIDE_W, dash="dash", name="Q guide")
            self.f.line(self.ox, self.y(v), self.x(u), self.y(v),
                        color=GUIDE, w_pt=GUIDE_W, dash="dash", name="P guide")
        self.f.dot(self.x(u), self.y(v), d=0.075, color=DOT, name="equilibrium")
        if label:
            pr = dict(bold=True, italic=True, color=NAVY, size=11)
            runs = _label_runs(label, pr)
            w = text_width(runs) + 0.08
            self.f.label(self.x(u) + dx, self.y(v) + dy, runs, w=w, h=0.22,
                         align="l", name="equilibrium label")
        if ptick:
            self.tick_y(v, ptick)
        if qtick:
            self.tick_x(u, qtick)

    def tick_y(self, v, label, size=10.5):
        pr = dict(bold=True, italic=True, color=NAVY, size=size)
        runs = _label_runs(label, pr)
        w = text_width(runs) + 0.08
        self.f.label(self.ox - Y_TITLE_GAP - w, self.y(v) - 0.11, runs,
                     w=w, h=0.22, align="r", name="y tick")

    def tick_x(self, u, label, size=10.5):
        pr = dict(bold=True, italic=True, color=NAVY, size=size)
        runs = _label_runs(label, pr)
        w = text_width(runs) + 0.08
        self.f.label(self.x(u) - w / 2.0, self.oy + X_TITLE_GAP, runs,
                     w=w, h=0.22, align="c", name="x tick")

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

    def text(self, u, v, runs, align="l", h=0.22):
        w = text_width(runs) + 0.08
        x = self.x(u)
        if align == "c":
            x -= w / 2.0
        elif align == "r":
            x -= w
        self.f.label(x, self.y(v), runs, w=w, h=h, align=align, name="note")


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
