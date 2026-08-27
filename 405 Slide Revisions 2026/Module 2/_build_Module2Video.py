# -*- coding: utf-8 -*-
# ==========================================================================
#  _build_Module2Video.py — phase-1 scaffold for
#  "Module 2 - Video Part Revised.pptx" (Videos 1–3, 57 slides per the
#  approved outline in "Module 2 - Video Part Revised - outline.md").
#
#  Reuses the ENTIRE In-Class helper layer via
#  `from _build_Module2InClass import *` (chrome, boxes, bullets, OMML,
#  SimpleFig, badges, make_m2_outline with the 2026-08-16 agenda
#  convention). Verbatim speaker notes come from _video_notes.py
#  (auto-extracted, keyed by OLD slide number).
#
#  Pipeline: _build_Module2Video.py -> _splice_video.py ->
#            _group_pass.py <deck> --spliced 9,28 ->
#            _animate_video.py all apply
# ==========================================================================

import _build_Module2InClass as _M

# pull the ENTIRE In-Class helper namespace (import * skips _-names)
globals().update({k: v for k, v in vars(_M).items()
                  if not k.startswith('__')})
from _video_notes import NOTES as VNOTES

VDECK = "Module 2 - Video Part Revised.pptx"

TAG_V1 = "Module 2 · Video 1 · Demand and Revenue"
TAG_V2 = "Module 2 · Video 2 · Marginal Revenue"
TAG_V3 = "Module 2 · Video 3 · Demand Estimation"
TAG_VOUT = "Module 2 · Outline"

ROSE = RGBColor(0xF2, 0xC4, 0xC4)        # pale red (revenue lost)
PALE_GREEN = RGBColor(0xC9, 0xE3, 0xC9)  # pale green (revenue gained)
CBLUE = RGBColor(0x00, 0x70, 0xC0)       # concept blue

# 2026-08-26 (Nico): the deck-wide rule is that a solution slide's FINAL
# line is dark red.  The standalone video deck reproduces the three
# recorded video decks slide for slide, so it stays as taped; the merged
# "Module 2 - Revised.pptx" builds with this flag ON (_merge_Module2.py).
RED_SOLUTIONS = False


def _sol(opts):
    """Paint a run dark red when the red-solution mode is on."""
    if not RED_SOLUTIONS:
        return opts
    out = dict(opts)
    out['color'] = RED
    return out


def _video_title_slide(prs, main, video_line):
    slide = _blank_slide(prs)
    _add_text(slide, 0, Inches(2.10), SLIDE_W, Inches(1.1), main,
              size=60, bold=True, color=NAVY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_text(slide, 0, Inches(3.25), SLIDE_W, Inches(0.75), video_line,
              size=40, bold=True, color=GOLD, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_rect(slide, int((SLIDE_W - Inches(4.0)) / 2), Inches(4.28),
              Inches(4.0), 54864, GOLD)
    _add_text(slide, 0, Inches(4.62), SLIDE_W, Inches(0.55),
              "Management 405", size=26, bold=True, color=GRAY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _add_text(slide, 0, Inches(5.32), SLIDE_W, Inches(0.5),
              "Prof. Nico Voigtländer  ·  UCLA Anderson",
              size=22, color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
    _add_rect(slide, 0, Inches(7.15), SLIDE_W, Inches(0.02), RULE)
    _add_rect(slide, MARGIN, Inches(7.135), GOLD_W, Inches(0.05), GOLD)
    return slide


def _add_groupdiscussion_badge(slide):
    """Nico's badge, relabeled 'Group Discussion' and stretched wider
    (group scaling widens the boxes; font size is unaffected)."""
    grp = _inject_handoff_group(slide, "_handoff_pollbreak.xml",
                                id_base=9700)
    for t in grp.iter(qn('a:t')):
        if t.text == "Poll Break":
            t.text = "Group Discussion"
    xf = grp.find(qn('p:grpSpPr') + '/' + qn('a:xfrm'))
    off = xf.find(qn('a:off'))
    ext = xf.find(qn('a:ext'))
    scale = 1.55
    old_w = int(ext.get('cx'))
    new_w = int(old_w * scale)
    off.set('x', str(int(off.get('x')) - (new_w - old_w)))
    ext.set('cx', str(new_w))
    return grp


def _vid_media(slide, fname, **kw):
    """Like _add_media_image but from _source_images_video/."""
    path = OUT_DIR / "_source_images_video" / fname
    kwargs = {"left": int(kw["left"]), "top": int(kw["top"])}
    if kw.get("width") is not None:
        kwargs["width"] = int(kw["width"])
    if kw.get("height") is not None:
        kwargs["height"] = int(kw["height"])
    pic = slide.shapes.add_picture(str(path), **kwargs)
    if kw.get("rounded", True):
        _apply_picture_style(pic, corner_pct=kw.get("corner_pct", 8))
    elif kw.get("shadow", True):
        _add_drop_shadow(pic)
    return pic


def _vnote(slide, old_no):
    if old_no in VNOTES:
        _set_notes(slide, VNOTES[old_no])


# --------------------------------------------------------------------------
# Aligned D-panel + TR-panel figures (slides 4, 6, 26): exact geometry —
# D: P = 400 − Q/4; TR = 400Q − Q²/4 (parabola drawn as an EXACT quadratic
# via its cubic-Bézier equivalent; peak at Q = 800 aligns across panels).
# --------------------------------------------------------------------------

def _dtr_figs(slide):
    figD = SimpleFig(1.75, 4.05, 6.3, 2.15, 1800, 450)
    figT = SimpleFig(1.75, 6.85, 6.3, 2.15, 1800, 180000)
    for fig, ylab in ((figD, "P"), (figT, "TR")):
        _fig_axes(slide, fig, weight_pt=1.75)
        _add_text(slide, Inches(fig.l - 0.55),
                  Inches(fig.b - fig.h - 0.42), Inches(0.7), Inches(0.32),
                  ylab, size=16, bold=True, italic=True, color=NAVY,
                  font="Calibri")
        _add_text(slide, Inches(fig.l + fig.w + 0.02),
                  Inches(fig.b - 0.05), Inches(0.6), Inches(0.32), "Q",
                  size=16, bold=True, italic=True, color=NAVY,
                  font="Calibri")
    return figD, figT


PALE_GOLD = RGBColor(0xF6, 0xE8, 0xC9)      # revenue-rectangle fill
CT_RED = RGBColor(0xC0, 0x50, 0x4D)         # CT's revenue-lost red
CT_GREEN = RGBColor(0x6E, 0x8B, 0x3D)       # CT's revenue-gained green


def _rev_rect(slide, fig, *, q, p, fill=None, line=GOLD,
              weight_pt=1.0):
    """The P x Q revenue rectangle under a demand curve: pale-gold fill
    with a gold edge (CT's treatment, adopted 2026-08-24)."""
    x0, y0 = int(fig.x(0)), int(fig.y(p))
    w = int(fig.x(q) - fig.x(0))
    h = int(fig.y(0) - fig.y(p))
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x0, y0, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill if fill is not None else PALE_GOLD
    shp.line.color.rgb = line
    shp.line.width = Pt(weight_pt)
    shp.shadow.inherit = False
    return shp


def _rot_brace(slide, x, y, cx, cy, rot, *, color=GOLD, weight_pt=2.0):
    """A rotated right-brace, the way CT spans a stretch of the demand
    line.  EMU + 60000ths-of-a-degree, copied straight from CT so the
    brace lands on the same segment."""
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_BRACE, int(x), int(y),
                                 int(cx), int(cy))
    shp.fill.background()
    shp.line.color.rgb = color
    shp.line.width = Pt(weight_pt)
    shp.shadow.inherit = False
    xf = shp._element.spPr.find(qn('a:xfrm'))
    xf.set('rot', str(int(rot)))
    return shp


def _brace_along(slide, p0, p1, *, width_in=0.38, offset_in=0.22,
                 color=GOLD, weight_pt=2.0):
    """A right-brace laid ALONG the segment p0 -> p1 (inches), opening
    away from it on the upper side.

    A rightBrace's spine runs down its height with the point out to the
    right, so rotating by (theta - 90) puts the spine on the segment and
    throws the point up and away from it.  PowerPoint rotates about the
    shape centre, so the box is centred on the segment midpoint.
    """
    import math
    dx, dy = p1[0] - p0[0], p1[1] - p0[1]
    length = math.hypot(dx, dy)
    theta = math.degrees(math.atan2(dy, dx))
    # turn the brace so it CURLS DOWN onto the segment, then lift the
    # whole thing clear of it along the upward normal
    rot = int(((theta + 90.0) % 360.0) * 60000)
    th = math.radians(theta)
    nx, ny = math.sin(th), -math.cos(th)
    cx = (p0[0] + p1[0]) / 2.0 + offset_in * nx
    cy = (p0[1] + p1[1]) / 2.0 + offset_in * ny
    return _rot_brace(slide, Inches(cx - width_in / 2.0),
                      Inches(cy - length / 2.0), Inches(width_in),
                      Inches(length), rot, color=color,
                      weight_pt=weight_pt)


def _tr_parabola(slide, figT, *, color=GOLD, weight_pt=3.0):
    # quadratic TR = 400Q − Q²/4 == Bézier P0=(0,0), ctrl=(800,320000),
    # P2=(1600,0); exact cubic equivalent controls at 2/3 toward ctrl
    c1 = (figT.x(1600 / 3), figT.y(320000 * 2 / 3))
    c2 = (figT.x(3200 / 3), figT.y(320000 * 2 / 3))
    return _add_cubic_curve(slide, (figT.x(0), figT.y(0)), c1, c2,
                            (figT.x(1600), figT.y(0)), color=color,
                            weight_pt=weight_pt)


# ==========================================================================
#  VIDEO 1 — Demand and Revenue
# ==========================================================================

def v02_outline(prs):
    slide = make_m2_outline(prs, 2, section_tag=TAG_VOUT,
                            highlight_set={2, 3})
    # back to the convention position 2026-08-25: at 6.272 the box ran
    # into the new coverage pill on item 4
    _add_outlined_box(slide, Inches(8.20), Inches(6.680), Inches(4.9),
                      Inches(0.72),
                      "▶  Teaching Note – Demand Elasticity and Total "
                      "Revenue\nOn BL under “Module 2 Post-Work”",
                      line=GOLD, text_color=NAVY, size=14, bold=True,
                      rounded=True, shadow=True, corner_pct=0.20)
    return slide


def v03_plot_demand(prs):
    """CT slide 4, "Plotting the Demand Curve" (adopted 2026-08-24, Nico:
    take the CT slides for 3-7 exactly).  The note about whose demand
    this is now sits as an italic caption ABOVE the graph, and the cream
    box carries both the demand and the inverse demand function."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    # retitled by hand 2026-08-24
    _draw_action_title(slide,
                       "Plotting the (Inverse) Demand Curve")
    # 2026-08-25 (Nico): the note moved below the chart and now names
    # the INVERSE demand function
    _add_hierarchical_bullets(
        slide, Inches(0.86), Inches(6.42), Inches(12.1), Inches(0.39),
        [([("Note: this is the (inverse) demand function that a "
            "single ", {'italic': True, 'color': GRAY}),
           ("firm", {'italic': True, 'bold': True, 'color': NAVY}),
           (" faces", {'italic': True, 'color': GRAY})], 0,
          {'bullet_style': 'none'})],
        size=23)
    # scaled so Q = 1600 lands at x 6.289" and P = 400 at y 3.166",
    # exactly where CT's demand line starts and ends
    # chart lifted 0.55" on 2026-08-25 to make room for the note
    fig = SimpleFig(1.10, 5.85, 6.30, 3.98, 1942.6, 492.3)
    _fig_axes(slide, fig)
    _add_text(slide, Inches(0.77), Inches(1.47), Inches(0.66),
              Inches(0.34), "P", size=19, bold=True, italic=True,
              color=NAVY, font="Calibri")
    _add_text(slide, Inches(7.07), Inches(5.95), Inches(0.66),
              Inches(0.32), "Q", size=19, bold=True, italic=True,
              color=NAVY, font="Calibri")
    _add_arrow(slide, (fig.x(0), fig.y(400)), (fig.x(1600), fig.y(0)),
               color=NAVY, weight_pt=3.0, head=False)
    _fig_ytick(slide, fig, 400, "$400", size=18)
    _fig_xtick(slide, fig, 1600, "1600", size=18)
    _add_text(slide, Inches(6.43), Inches(5.51), Inches(1.6),
              Inches(0.34), "D", size=20, bold=True, color=NAVY,
              font="Calibri")
    _add_convention_box(
        # hand-resized 2026-08-24 (was 8.20, 3.20, 4.00 x 2.00)
        slide, Inches(8.20), Inches(1.95), Inches(4.00), Inches(3.25),
        # hand-rewritten 2026-08-24: the rearrangement step is spelled
        # out, and the inverse demand function is set bold
        runs=[("Demand function", {'bold': True, 'size': 19}),
              ("\nQ = 1,600 − 4P", {'italic': True, 'size': 22}),
              ("\n", {'size': 12}),
              ("\nRearrange for P: ",
               {'bold': True, 'italic': True, 'size': 19}),
              ("\nInverse demand function", {'bold': True, 'size': 19}),
              ("\nP = 400 − Q/4",
               {'bold': True, 'italic': True, 'size': 22}),
              # closing line added by hand 2026-08-25
              ("[this is what we plot]",
               {'size': 14, 'newline': True})],
        size=19, align=PP_ALIGN.CENTER)
    _vnote(slide, 3)
    _draw_footer(slide, FOOTER_TEXT, 3)
    return slide


def v04_demand_tr(prs):
    """From demand to total revenue.

    2026-08-24 (Nico, hand-edited): the box on the right is now three
    bullets - inverse demand, total revenue, plug in for P - and each has
    to be able to appear on its own click.  So the cream rectangle is a
    plain background shape and the three points live in their own text
    box as three paragraphs; _group_pass is told NOT to merge the two
    (see NO_GROUP_BOXES), because a merged box+text can only animate as a
    single object.
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "From Demand to Total Revenue")
    figD = SimpleFig(1.45, 3.85, 5.05, 2.08, 1966.9, 514.5)
    figT = SimpleFig(1.45, 6.55, 5.05, 2.23, 1966.9, 201240.0)
    _rev_rect(slide, figD, q=800, p=200)
    for fig, ylab in ((figD, "P"), (figT, "TR")):
        _fig_axes(slide, fig, weight_pt=2.5)
        _add_text(slide, Inches(fig.l - 0.33),
                  Inches(fig.b - fig.h - 0.40), Inches(0.66),
                  Inches(0.34), ylab, size=17, bold=True, italic=True,
                  color=NAVY, font="Calibri")
        _add_text(slide, Inches(fig.l + fig.w + 0.02),
                  Inches(fig.b + 0.10), Inches(0.66), Inches(0.32), "Q",
                  size=17, bold=True, italic=True, color=NAVY,
                  font="Calibri")
    _add_text(slide, Inches(1.450), Inches(3.276), Inches(2.054),
              Inches(0.34), "Total revenue", size=16, bold=True,
              color=NAVY, font="Calibri", align=PP_ALIGN.CENTER)
    _fig_ytick(slide, figD, 200, "200", size=15)
    _fig_xtick(slide, figD, 800, "800", size=15)
    # demand line + its label (hand-moved 2026-08-24; grouped with the
    # $400 / 1600 ticks in _group_pass so the whole curve reveals as one)
    _add_arrow(slide, (figD.x(0), figD.y(400)), (figD.x(1600), figD.y(0)),
               color=NAVY, weight_pt=3.0, head=False)
    _add_text(slide, Inches(5.476), Inches(3.471), Inches(1.600),
              Inches(0.286), "D", size=17, bold=True, color=NAVY,
              font="Calibri")
    _fig_ytick(slide, figD, 400, "$400", size=15)
    _fig_xtick(slide, figD, 1600, "1600", size=15)
    _add_arrow(slide, (figD.x(0), figD.y(200)), (figD.x(800), figD.y(200)),
               color=GRAY, weight_pt=1.4, head=False, dash="dash")
    _add_arrow(slide, (figD.x(800), figD.y(200)), (figD.x(800), figD.y(0)),
               color=GRAY, weight_pt=1.4, head=False, dash="dash")
    _tr_parabola(slide, figT)
    _fig_ytick(slide, figT, 160000, "$160,000", size=15)
    _fig_xtick(slide, figT, 800, "800", size=15)
    _fig_xtick(slide, figT, 1600, "1600", size=15)
    _add_arrow(slide, (figT.x(0), figT.y(160000)),
               (figT.x(800), figT.y(160000)), color=GRAY, weight_pt=1.4,
               head=False, dash="dash")
    _add_arrow(slide, (figT.x(800), figT.y(160000)),
               (figT.x(800), figT.y(0)), color=GRAY, weight_pt=1.4,
               head=False, dash="dash")
    # --- the three points, as three separately animatable paragraphs ----
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 int(Inches(7.923)), int(Inches(2.500)),
                                 int(Inches(5.030)), int(Inches(3.040)))
    box.fill.solid()
    box.fill.fore_color.rgb = CREAM
    box.line.color.rgb = NAVY
    box.line.width = Pt(1.0)
    box.shadow.inherit = False
    try:
        box.adjustments[0] = 0.12
    except Exception:
        pass
    _add_drop_shadow(box)
    _add_hierarchical_bullets(
        slide, Inches(8.178), Inches(2.791), Inches(4.521), Inches(2.457),
        # five paragraphs, three of them bulleted: PowerPoint reads a
        # literal newline inside a run as a paragraph break, so the
        # continuation lines have to be their own un-bulleted paragraphs
        # (the animation reveals each bullet together with its
        # continuation - see PLANS[4])
        [
            ([("Inverse demand:  ", {'bold': True}),
              ("P = 400 − Q/4", {'italic': True})], 0, {}),
            ([("Total revenue:  ", {'bold': True}),
              ("TR = P · Q", {'italic': True})], 0, {}),
            ([("(illustrate for P=200)", {'size': 20})], 1,
             {'bullet_style': 'none', 'space_before_pts': 2}),
            ([("Plug in for P:  ", {'bold': True}),
              ("TR = (400 − Q/4) · Q", {'italic': True})], 0, {}),
            ([("TR = 400Q − Q²/4",
               {'italic': True, 'bold': True, 'color': GOLD})], 1,
             {'bullet_style': 'none', 'space_before_pts': 2}),
        ],
        size=21, line_spacing_pts=10)
    _vnote(slide, 4)
    _draw_footer(slide, FOOTER_TEXT, 4)
    return slide


def v05_price_change_tr(prs):
    """CT slide 6, "Effect of a Price Change on Total Revenue" (adopted
    2026-08-24).  One demand line, the base revenue rectangle, the red
    slice lost to the lower price and the green slice gained from the
    higher volume, a two-entry legend and the gold question box.  All
    four corners sit exactly on the demand line."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "Effect of a Price Change on Total Revenue")
    # CT geometry, in inches: demand runs (1.20, 2.44) -> (6.60, 6.40),
    # so P0 = 3.760 at Q0 = 3.000 and P1 = 4.200 at Q1 = 3.600 lie ON it
    X0, Y_TOP, X1, Y_BOT = 1.200, 1.820, 7.500, 6.400
    P0, P1, Q0, Q1 = 3.760, 4.200, 3.000, 3.600
    # regions first, so the axes and the demand line draw over them
    _add_rect(slide, Inches(X0), Inches(P1), Inches(Q1 - X0),
              Inches(Y_BOT - P1), PALE_GOLD)                 # TR base
    _add_rect(slide, Inches(X0), Inches(P0), Inches(Q0 - X0),
              Inches(P1 - P0), CT_RED)                       # lost to P
    _add_rect(slide, Inches(Q0), Inches(P1), Inches(Q1 - Q0),
              Inches(Y_BOT - P1), CT_GREEN)                  # gained on Q
    _add_arrow(slide, (Inches(X0), Inches(Y_BOT)),
               (Inches(X0), Inches(Y_TOP)), color=NAVY,
               weight_pt=2.5, head=True)
    _add_arrow(slide, (Inches(X0), Inches(Y_BOT)),
               (Inches(X0 + 6.300), Inches(Y_BOT)), color=NAVY,
               weight_pt=2.5, head=True)
    _add_text(slide, Inches(0.87), Inches(1.42), Inches(0.66),
              Inches(0.34), "P", size=19, bold=True, italic=True,
              color=NAVY, font="Calibri")
    _add_text(slide, Inches(7.17), Inches(6.50), Inches(0.66),
              Inches(0.32), "Q", size=19, bold=True, italic=True,
              color=NAVY, font="Calibri")
    _add_arrow(slide, (Inches(1.200), Inches(2.440)),
               (Inches(6.600), Inches(6.400)), color=NAVY,
               weight_pt=3.0, head=False)
    for txt, x, y, sz in (("P₀", 0.510, 3.590, 22),
                          ("P₁", 0.510, 4.030, 22),
                          ("Q₀", 2.700, 6.500, 20),
                          ("Q₁", 3.300, 6.500, 20)):
        _add_text(slide, Inches(x), Inches(y), Inches(0.60),
                  Inches(0.34), txt, size=sz, bold=True, color=NAVY,
                  font="Calibri")
    # 2026-08-24 (Nico): show the direction of the change - price down,
    # quantity up.  Down-arrow to the LEFT of P0 / P1, right-arrow UNDER
    # Q0 / Q1.
    _add_arrow(slide, (Inches(0.410), Inches(3.720)),
               (Inches(0.410), Inches(4.185)), color=GOLD,
               weight_pt=2.25, head=True)
    _add_arrow(slide, (Inches(2.870), Inches(6.905)),
               (Inches(3.500), Inches(6.905)), color=GOLD,
               weight_pt=2.25, head=True)
    _add_text(slide, Inches(5.05), Inches(4.91), Inches(2.0),
              Inches(0.34), "Demand", size=20, bold=True, color=NAVY,
              font="Calibri")
    _add_text(slide, Inches(1.20), Inches(5.10), Inches(1.8),
              Inches(0.40), "TR", size=22, bold=True, italic=True,
              color=NAVY, font="Calibri", align=PP_ALIGN.CENTER)
    # legend
    for fill, y_sw, y_tx, runs in (
            (CT_RED, 2.400, 2.300,
             [("Reduction in ", {}), ("TR", {'italic': True}),
              (" from a lower price", {})]),
            (CT_GREEN, 3.500, 3.400,
             [("Increase in ", {}), ("TR", {'italic': True}),
              (" from higher volume", {})])):
        _add_rect(slide, Inches(8.00), Inches(y_sw), Inches(0.55),
                  Inches(0.42), fill)
        _add_hierarchical_bullets(
            slide, Inches(8.62), Inches(y_tx), Inches(4.30),
            Inches(0.60), [(runs, 0, {'bullet_style': 'none'})],
            size=20)
    _add_rounded_filled_box(
        slide, Inches(8.60), Inches(4.90), Inches(3.60), Inches(0.95),
        # reworded 2026-08-24 (Nico): it is the AREAS that are compared
        "Which area is bigger?", fill=GOLD, text_color=NAVY, size=24,
        bold=True)
    _vnote(slide, 5)
    _draw_footer(slide, FOOTER_TEXT, 5)
    return slide


def v06_elasticity_tr(prs):
    """CT slide 7, "Effect of a Price Change on Total Revenue: It Depends
    on E_D" (adopted 2026-08-24).  Two stacked panels: the demand line
    braced into its elastic and inelastic stretches with the unit-elastic
    midpoint marked, and the TR hill below sharing the same Q scale.

    Nico's addition (his own video slide 6): the BOTTOM panel also gets
    the two rising / falling arrows with the region descriptions, so the
    TR consequence is spelled out inside the graph rather than beside it.
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "Effect of a Price Change on Total Revenue: "
                              "It Depends on Eᴅ")
    # ---- top panel: demand with the elasticity regions ------------------
    _add_arrow(slide, (Inches(2.80), Inches(4.20)),
               (Inches(2.80), Inches(1.62)), color=NAVY, weight_pt=2.5,
               head=True)
    _add_arrow(slide, (Inches(2.80), Inches(4.20)),
               (Inches(9.20), Inches(4.20)), color=NAVY, weight_pt=2.5,
               head=True)
    _add_text(slide, Inches(2.48), Inches(1.32), Inches(0.66),
              Inches(0.34), "P", size=17, bold=True, italic=True,
              color=NAVY, font="Calibri")
    _add_text(slide, Inches(8.87), Inches(4.30), Inches(0.66),
              Inches(0.32), "Q", size=17, bold=True, italic=True,
              color=NAVY, font="Calibri")
    _add_arrow(slide, (Inches(2.800), Inches(2.040)),
               (Inches(8.290), Inches(4.200)), color=NAVY,
               weight_pt=3.0, head=False)
    _add_arrow(slide, (Inches(5.545), Inches(4.200)),
               (Inches(5.545), Inches(3.120)), color=NAVY,
               weight_pt=1.4, head=False)
    _add_text(slide, Inches(8.41), Inches(3.90), Inches(1.6),
              Inches(0.34), "D", size=17, bold=True, color=NAVY,
              font="Calibri")
    # CT's two rotated braces spanning the elastic / inelastic stretches
    _rot_brace(slide, 3717590, 854826, 295148, 2668634, 17488607)
    _add_text(slide, Inches(3.85), Inches(1.84), Inches(3.0),
              Inches(0.34), "Eᴅ < −1   elastic", size=16, bold=True,
              color=CBLUE, font="Calibri")
    _rot_brace(slide, 6257346, 1879504, 292608, 2602951, 17488607)
    _add_text(slide, Inches(6.744), Inches(2.957), Inches(3.6),
              Inches(0.34), "−1 < Eᴅ < 0   inelastic", size=16,
              bold=True, color=CBLUE, font="Calibri")
    _add_text(slide, Inches(5.60), Inches(2.425), Inches(1.8),
              Inches(0.34), "Eᴅ = −1", size=16, bold=True, color=CT_RED,
              font="Calibri")
    _add_arrow(slide, (Inches(5.663), Inches(2.835)),
               (Inches(5.575), Inches(3.091)), color=CT_RED,
               weight_pt=1.75, head=False)
    # ---- bottom panel: total revenue ------------------------------------
    _add_arrow(slide, (Inches(2.80), Inches(6.75)),
               (Inches(2.80), Inches(4.32)), color=NAVY, weight_pt=2.5,
               head=True)
    _add_arrow(slide, (Inches(2.80), Inches(6.75)),
               (Inches(9.20), Inches(6.75)), color=NAVY, weight_pt=2.5,
               head=True)
    _add_text(slide, Inches(8.87), Inches(6.85), Inches(0.66),
              Inches(0.32), "Q", size=17, bold=True, italic=True,
              color=NAVY, font="Calibri")
    figT = SimpleFig(2.80, 6.75, 5.49, 1.898, 1600.0, 160000.0)
    _tr_parabola(slide, figT)
    _add_arrow(slide, (Inches(5.545), Inches(6.750)),
               (Inches(5.545), Inches(4.852)), color=NAVY,
               weight_pt=1.4, head=False)
    # hand-moved 2026-08-24 (was 3.532, 4.828)
    _add_text(slide, Inches(4.165), Inches(4.775), Inches(1.6),
              Inches(0.34), "TR", size=18, bold=True, color=GOLD,
              font="Calibri")
    # Nico's in-graph region annotations (from his own video slide 6):
    # a short rising arrow on the left, a falling one on the right
    # arrows + region notes all hand-placed 2026-08-24
    _add_arrow(slide, (Inches(4.015), Inches(5.825)),
               (Inches(4.765), Inches(5.225)), color=GOLD,
               weight_pt=2.0, head=True)
    _add_arrow(slide, (Inches(6.653), Inches(5.367)),
               (Inches(7.353), Inches(5.917)), color=GOLD,
               weight_pt=2.0, head=True)
    _add_hierarchical_bullets(
        slide, Inches(3.504), Inches(6.105), Inches(1.846),
        Inches(0.471),
        [([("Elastic region", {'bold': True}),
           (":  P falls\n→ Q rises → ", {}),
           ("TR rises", {'bold': True, 'color': GOLD})], 0,
          {'bullet_style': 'none'})],
        size=14, line_spacing_pts=0)
    _add_hierarchical_bullets(
        slide, Inches(5.792), Inches(6.098), Inches(1.860),
        Inches(0.471),
        [([("Inelastic region", {'bold': True}),
           (":  P falls\n→ Q rises → ", {}),
           ("TR falls", {'bold': True, 'color': GOLD})], 0,
          {'bullet_style': 'none'})],
        size=14, line_spacing_pts=0)
    _vnote(slide, 6)
    _draw_footer(slide, FOOTER_TEXT, 6)
    return slide


def v07_depends_on_ed(prs):
    """Nico's own slide from Module 2 Video 1, reinstated 2026-08-24 in
    place of the CT version: the two formulas carry the argument, and the
    two questions sit in cream cards underneath.  His hand-edits from the
    CT round are folded in - "price decrease" in red, and P and Q italic
    inside the questions."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(
        slide, "Effect of a Price Change on Total Revenue: "
               "It Depends on Eᴅ")
    _add_mixed_textbox(
        slide, MARGIN + Inches(0.15), Inches(1.75), Inches(12.4),
        Inches(1.35),
        [
            ("text", "Total Revenues:   ", {'size': 24}),
            ("omml", _omml_text('TR') + _omml_text(' = ')
             + _omml_run('P') + _omml_text(' ⋅ ') + _omml_run('Q'),
             {'size': 24}),
            ("break", None, None),
            # "In % changes" set bold by hand, 2026-08-25
            ("text", "In % changes", {'size': 24, 'bold': True}),
            ("text", ", this implies:   ", {'size': 24}),
            ("omml", _omml_text('%Δ') + _omml_text('TR')
             + _omml_text(' = %Δ') + _omml_run('P')
             + _omml_text(' + %Δ') + _omml_run('Q'),
             {'size': 24, 'bold': True}),
        ])
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(3.15), Inches(12.4),
        Inches(0.9),
        [
            # the note sits on its own line under the question
            # (hand-moved 2026-08-24)
            ([("To assess the effect of a ", {}),
              ("price decrease", {'bold': True, 'color': RED}),
              (" on revenue, ask:", {})], 0, {}),
            # its own paragraph, no bullet, no space before it, indented
            # by hand with spaces — exactly as Nico set it 2026-08-24
            ([("        ", {'italic': True, 'size': 18}),
              ("[note: if P decreases, Q must increase]",
               {'italic': True, 'size': 18})], 0,
             {'bullet_style': 'arrow', 'mar_l': 0, 'indent': 0,
              'space_before_pts': 0}),
        ],
        size=24)
    # cards hand-resized again 2026-08-25 (were 0.726 / 6.950, 5.9
    # wide, h 1.452): the left one now starts at the slide margin and
    # both are shorter
    _add_convention_box(
        slide, Inches(0.280), Inches(4.300), Inches(6.35), Inches(1.230),
        pad_h=Inches(0.21), pad_v=Inches(0.10),
        runs=[("Is the % reduction in ", {}), ("P", {'italic': True}),
              (" ", {}),
              ("smaller", {'bold': True, 'underline': True}),
              (" than the % increase in ", {}),
              ("Q", {'italic': True}), ("?\n", {}),
              ("If yes: total revenue rises\n", {}),
              ("→ Demand is ", {}),
              ("elastic", {'bold': True, 'color': CBLUE})],
        size=18)
    _add_convention_box(
        slide, Inches(6.950), Inches(4.300), Inches(6.15), Inches(1.230),
        pad_h=Inches(0.21), pad_v=Inches(0.10),
        runs=[("Is the % reduction in ", {}), ("P", {'italic': True}),
              (" ", {}),
              ("larger", {'bold': True, 'underline': True}),
              (" than the % increase in ", {}),
              ("Q", {'italic': True}), ("?\n", {}),
              ("If yes: total revenue declines\n", {}),
              ("→ Demand is ", {}),
              ("inelastic", {'bold': True, 'color': CBLUE})],
        size=18)
    _vnote(slide, 7)
    _draw_footer(slide, FOOTER_TEXT, 7)
    return slide


def v08_gum(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "Example: Demand for Chewing Gum")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.95), Inches(7.6),
        Inches(4.4),
        [
            ("Information", 0, {'bold': True, 'bullet_style': 'none'}),
            ("Retail prices of Wrigley’s chewing gum increased "
             "(slightly)", 0),
            ([("As a result, the company's ", {}),
              ("sales revenues", {'underline': True}),
              (" increased (slightly)", {})], 0, {}),
            ("What can we say about the elasticity of demand?", 0,
             {'bold': True, 'bullet_style': 'none'}),
        ],
        size=24, line_spacing_pts=16)
    _vid_media(slide, "image13.jpg", left=Inches(8.55), top=Inches(1.90),
               width=Inches(3.35))
    _draw_footer(slide, FOOTER_TEXT, 8)
    _add_pollbreak_badge(slide)
    return slide


def v10_gum_solution(prs):
    return make_content_bulleted(
        prs, 10, TAG_V1, "Solution",
        [
            ([("Demand is ", {}),
              ("inelastic", {'bold': True, 'color': CBLUE})], 0, {}),
            ("A price increase caused total revenue to increase", 1),
            ("This means that the %change in price was larger than the "
             "%decline in quantity  →  inelastic response of quantity "
             "to the price change", 1),
        ],
        size=26, sub_size=24)


def v11_ozempic(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    # retitled by hand 2026-08-24
    _draw_action_title(slide, "What can a Price Cut and Revenues tell "
                              "us about the Elasticity?")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(2.00), Inches(6.9),
        Inches(3.9),
        [
            ([("Novo Nordisk", {'bold': True}),
              (" lowered its U.S. price for Ozempic and Wegovy under a "
               "deal with the White House", {})], 0, {}),
            ("As a result, revenues are forecasted to fall 5–13%", 0),
            ([("What does this imply about the ", {}),
              ("elasticity of demand", {'bold': True, 'color': CBLUE}),
              ("?", {})], 0, {}),
            # the options he reads out, added by hand 2026-08-25
            ("Elastic?", 1),
            ("Inelastic?", 1),
            # 2026-08-25: his own file said "Eᴅ=0"; unit elasticity
            # is Eᴅ = −1, corrected with his approval
            ("Eᴅ=−1 (unit elastic)?", 1),
            ("Need more information?", 1),
        ],
        size=24, line_spacing_pts=16)
    _vid_media(slide, "ct_ozempic_image1.png", left=Inches(7.65),
               top=Inches(2.30), width=Inches(4.65))
    _add_text(slide, Inches(7.65), Inches(5.55), Inches(4.65),
              Inches(0.32), "Source: Pharmaceutical Technology",
              size=12, italic=True, color=GRAY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _set_notes(slide, (
        "A current example, from early 2026: under a deal with the White "
        "House, Novo Nordisk cut the U.S. self-pay price of Ozempic and "
        "Wegovy from $499 to $349 a month. The company then guided 2026 "
        "sales revenues down by 5 to 13%. Ask yourself what a price cut "
        "that "
        "REDUCES revenue tells you about the elasticity of demand."))
    _draw_footer(slide, FOOTER_TEXT, 8)
    # 2026-08-24 (Nico): use the deck-standard Poll Break badge here,
    # not the 'Group Discussion' relabel
    _add_pollbreak_badge(slide)
    return slide


def v12_ozempic_solution(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "Solution")
    _add_mixed_textbox(
        slide, Inches(4.05), Inches(2.05), Inches(6.0), Inches(0.8),
        [("text", "Demand is  ", {'size': 34}),
         ("text", "inelastic", {'size': 34, 'bold': True,
                                'color': CBLUE})])
    _add_hierarchical_bullets(
        slide, Inches(1.10), Inches(3.55), Inches(11.2), Inches(2.4),
        [
            ("A price cut caused TR to fall", 0),
            ("The quantity response was too small to offset the price "
             "cut", 0),
            ([("Inelastic:  ", {'bold': True, 'color': CBLUE}),
              ("|Eᴅ| < 1", {'italic': True})], 0, {}),
        ],
        size=28, line_spacing_pts=18)
    _draw_footer(slide, FOOTER_TEXT, 9)
    return slide


def v13_profits(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "What About Profits?")
    _add_rounded_filled_box(slide, Inches(3.55), Inches(1.80),
                            Inches(6.25), Inches(0.95),
                            "Profit  =  Revenue  −  Costs", fill=NAVY,
                            text_color=WHITE, size=26, bold=True,
                            corner_pct=0.12, shadow=True)
    _add_hierarchical_bullets(
        slide, Inches(1.50), Inches(3.15), Inches(10.3), Inches(3.4),
        [
            ([("We know: ", {}),
              ("The price decrease caused TR to decrease", {})], 0, {}),
            # "produce/sold" in his file is a slip for "produced/sold"
            ("A declining price also raises quantity produced/sold "
             "Q  →  costs increase", 0),
            ("Thus: Profits fell", 0),
            ([("Markets agreed:  ", {}),
              ("Novo Nordisk shares tumbled ~18%",
               {'underline': True, 'color': CBLUE})], 0, {}),
        ],
        size=24, line_spacing_pts=16)
    _set_notes(slide, (
        "One step further: profit is revenue minus costs. The price cut "
        "reduced revenue, and because quantity sold rises, production "
        "costs rise too — so profits fall on both ends. The market "
        "agreed: when Novo Nordisk warned in February 2026 that lower "
        "U.S. prices would push 2026 sales and operating profit down 5 "
        "to 13%, its shares tumbled about 18%."))
    _draw_footer(slide, FOOTER_TEXT, 10)
    return slide


def v14_four_cases(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(
        slide, "Summary: Effect of a Price Change on TR — "
               "It Depends on Eᴅ")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.70), Inches(12.4),
        Inches(1.0),
        [
            ("Is the percent reduction in price greater or smaller than "
             "the percent increase in quantity sold?", 0),
            ("More precisely: is demand elastic or inelastic?", 0),
        ],
        size=22, line_spacing_pts=10)
    rows = [
        ["% change in P", "Resulting % change in Q", "Demand is…",
         "Resulting change in revenue"],
        # middle column reworded by hand 2026-08-25
        ["If P decreases", "Relatively large increase in Q",
         "Elastic", "Revenue increases"],
        ["If P decreases", "Relatively small increase in Q",
         "Inelastic", "Revenue decreases"],
        ["If P increases", "Relatively large decline in Q",
         "Elastic", "Revenue decreases"],
        ["If P increases", "Relatively small decline in Q",
         "Inelastic", "Revenue increases"],
    ]
    _add_styled_table(slide, Inches(0.85), Inches(3.05), Inches(11.6),
                      Inches(3.6), rows, font_size=16, header_size=16,
                      first_col_bold=True)
    _draw_footer(slide, FOOTER_TEXT, 11)
    return slide


def v15_netflix(prs):
    """
    NOT BUILT since 2026-08-25: this slide moved to
    "Module 2 - Potential Practice Exercises.pptx".
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "The Netflix 2014 Price Increase")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.35), Inches(1.95), Inches(11.6),
        Inches(4.3),
        [
            ("In 2014, Netflix raised its monthly price by $1, from "
             "$7.99 to $8.99", 0),
            ([("Then: 50 million subscribers, own-price elasticity of "
               "demand ", {}),
              ("≈ −2", {'bold': True, 'color': CBLUE})], 0, {}),
            ("Implications of the price increase for:", 0),
            ("Number of subscriptions (quantitative)?", 1),
            ("Revenue (qualitative)?", 1),
            ("Profits (speculative)?", 1),
        ],
        size=26, sub_size=24)
    _draw_footer(slide, FOOTER_TEXT, 47)
    _add_groupdiscussion_badge(slide)
    return slide


def v16_netflix_solution(prs):
    """
    NOT BUILT since 2026-08-25: this slide moved to
    "Module 2 - Potential Practice Exercises.pptx".
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "Solution: Netflix")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.75), Inches(12.4),
        Inches(0.6),
        [([("Subscriptions decline", {'bold': True}),
           (" — by how much? Use Method 1 with ", {}),
           ("Eᴅ = −2", {'bold': True, 'color': CBLUE}),
           (" :", {})], 0, {'bullet_style': 'none'})],
        size=22)
    _add_math_equation(
        slide, Inches(2.65), Inches(2.55), Inches(8.0), Inches(1.15),
        _oED() + _omml_text(' = ')
        + _omml_frac(_o_pct('Q'), _o_pct('P')) + _omml_text(' = ')
        + _omml_frac(_o_pct('Q'), _omml_text('12.5%'))
        + _omml_text(' = −2'),
        size_pt=26, color=NAVY)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.55), Inches(3.90), Inches(12.0),
        Inches(0.5),
        [("where  %ΔP = (8.99 − 7.99) / 7.99 = +12.5%,   so   "
          "%ΔQ = −25%", 0, {'bullet_style': 'none', 'size': 20})],
        size=20)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(4.70), Inches(12.4),
        Inches(2.1),
        [
            ([("→  a ", {}),
              ("25% decline", {'bold': True, 'color': CBLUE}),
              (" in subscriptions", {})], 0, {'bullet_style': 'none'}),
            ([("Revenue:", {'bold': True}), (" demand is ", {}),
              ("elastic", {'bold': True, 'color': CBLUE}),
              (", so quantity outweighs price  →  revenue falls", {})],
             0, {}),
            ([("Profits:", {'bold': True}),
              (" fewer subscribers, near-zero cost savings  →  profits "
               "likely fell too", {})], 0, {}),
        ],
        size=22, line_spacing_pts=12)
    _draw_footer(slide, FOOTER_TEXT, 48)
    return slide


def v17_mcdonalds(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "Why Is McDonald's Cutting Prices?")
    _vid_media(slide, "ct_mcdonalds_image3.png", left=Inches(3.42),
               top=Inches(1.50), width=Inches(6.5), rounded=False,
               shadow=False)
    _add_hierarchical_bullets(
        slide, Inches(1.28), Inches(3.70), Inches(11.0), Inches(3.32),
        [
            # 2026-08-25 (Nico): two questions, each followed by the
            # answer options he reads out
            ("What does this suggest about its demand elasticity?", 0),
            ("Elastic? ", 1),
            ("Inelastic?", 1),
            ("Need more information?", 1),
            ("What are the implications for 1) revenue and 2) profit "
             "given this elasticity?", 0),
            ("Rise?", 1),
            ("Fall?", 1),
            ("Unchanged?", 1),
            ("Need more information?", 1),
        ],
        size=22, line_spacing_pts=12)
    _set_notes(slide, (
        "A live example from 2025: McDonald's and its U.S. franchisees "
        "agreed to cut combo-meal prices by about 15% relative to buying "
        "the items separately, and extended the value menu in 2026. "
        "Notable detail for the franchisee question: McDonald's set "
        "aside roughly $35 million to compensate operators hit by the "
        "lower prices — side payments to align incentives across "
        "locations with different demand elasticities."))
    _draw_footer(slide, FOOTER_TEXT, 12)
    # 2026-08-25 (Nico): Poll Break here, not Group Discussion
    _add_pollbreak_badge(slide)
    return slide


def v18_megamillions_revisited(prs):
    """
    NOT BUILT since 2026-08-25: this slide moved to
    "Module 2 - Potential Practice Exercises.pptx".
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "Mega Millions, Revisited")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.35), Inches(1.95), Inches(11.9),
        Inches(3.6),
        [
            ("April 2025: Mega Millions raised its ticket price from $2 "
             "to $5  (+150%)", 0),
            ("In New York, tickets sold per drawing fell from ~1.9 "
             "million to ~560,000  (≈ −70%)", 0),
            ([("We computed  ", {}),
              ("Eᴅ ≈ −0.47", {'bold': True, 'color': CBLUE}),
              ("  →  inelastic", {})], 0, {}),
            ("Inelastic demand suggests revenue should RISE when the "
             "price rises. Did it?", 0, {'bold': True}),
        ],
        size=24, line_spacing_pts=16)
    _add_convention_box(
        slide, Inches(1.55), Inches(5.75), Inches(10.2), Inches(0.85),
        prefix="Caution: ",
        body="Method 1 approximates % changes – for LARGE changes like "
             "this one, the elasticity-revenue rule can mislead",
        size=15)
    _set_notes(slide, (
        "Back to Mega Millions. Method 1 gave us an elasticity of about "
        "−0.47 — inelastic — which suggests the price increase should "
        "raise revenue. In fact, revenue per drawing FELL from roughly "
        "$3.8 million to $2.8 million. The catch: the price rose 150%, "
        "and Method 1 is a small-change approximation. With changes this "
        "large, the initial-point elasticity and the revenue rule based "
        "on it can point the wrong way — a good reason to be careful "
        "before applying the rule to big price moves."))
    _draw_footer(slide, FOOTER_TEXT, 49)
    _add_groupdiscussion_badge(slide)
    return slide


# ==========================================================================
#  VIDEO 2 — Marginal Revenue
# ==========================================================================

def v20_outline(prs):
    return make_m2_outline(prs, 14, section_tag=TAG_VOUT,
                           highlight_set={4})


def v21_why_mr(prs):
    """CT video slide 21 (adopted 2026-08-24, Nico): the two rules get
    their own navy bars instead of sitting as sub-bullets.  Kept close to
    CT, except the bars are rounded with a soft shade, per the deck's
    filled-box convention."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V2)
    _draw_action_title(slide, "Context: Why Do We Care About Marginal Revenue?")
    _add_hierarchical_bullets(
        slide, Inches(1.10), Inches(1.75), Inches(11.1), Inches(1.05),
        [([("General objective: maximize ", {}),
           ("net benefits", {'bold': True})], 0,
          {'bullet_style': 'none'}),
         ("Use marginal analysis", 1, {'size': 20})],
        size=23, sub_size=20)
    _add_rounded_filled_box(
        slide, Inches(1.60), Inches(2.76), Inches(10.13), Inches(0.72),
        "Net benefits are maximized where MB = MC",
        fill=NAVY, text_color=WHITE, size=22, bold=True,
        corner_pct=0.18)
    _add_hierarchical_bullets(
        slide, Inches(1.10), Inches(3.95), Inches(11.1), Inches(0.55),
        [([("Firms’ objective: maximize ", {}),
           ("profits", {'bold': True})], 0,
          {'bullet_style': 'none'})],
        size=23)
    _add_rounded_filled_box(
        slide, Inches(1.60), Inches(4.50), Inches(10.13), Inches(0.72),
        "Produce where Marginal Revenue (MR) = Marginal Cost (MC)",
        fill=NAVY, text_color=WHITE, size=22, bold=True,
        corner_pct=0.18)
    _add_hierarchical_bullets(
        slide, Inches(1.10), Inches(5.75), Inches(11.1), Inches(0.60),
        [([("Today: how to compute ", {}),
           ("MR", {'bold': True, 'italic': True})], 0,
          {'bullet_style': 'none'})],
        size=23)
    _vnote(slide, 13)
    _draw_footer(slide, FOOTER_TEXT, 15)
    return slide


def v22_mr_definition(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V2)
    _draw_action_title(slide, "Marginal Revenue: Definition")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.90), Inches(12.4),
        Inches(1.5),
        [
            ([("Intuitively:", {'bold': True}),
              ("  Additional revenue from selling one more unit", {})],
             0, {}),
            ([("Mathematically:", {'bold': True}),
              ("  How much total revenue changes when quantity sold "
               "increases by one unit", {})], 0, {}),
        ],
        size=24, line_spacing_pts=14)
    _add_math_equation(
        slide, Inches(4.35), Inches(3.75), Inches(4.6), Inches(1.35),
        _omml_text('MR') + _omml_text(' = ')
        + _omml_frac(_omml_text('Δ') + _omml_text('TR'),
                     _omml_text('Δ') + _omml_run('Q'))
        + _omml_text(' = ')
        + _omml_frac(_omml_run('d') + _omml_text('TR'),
                     _omml_run('d') + _omml_run('Q')),
        size_pt=30, color=NAVY, fill=CREAM, line=NAVY, rounded=True,
        shadow=True)
    _add_text(slide, Inches(9.35), Inches(4.05), Inches(2.0),
              Inches(0.4), "Derivative", size=18, bold=True, color=CBLUE,
              font="Calibri")
    # same blue as the word "Derivative" it points at (Nico, 2026-08-24)
    _add_arrow(slide, (Inches(9.30), Inches(4.25)),
               (Inches(8.65), Inches(4.35)), color=CBLUE, weight_pt=1.75,
               head=True)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(5.65), Inches(12.4),
        Inches(0.8),
        [([("How to compute ", {'bold': True}),
           ("MR", {'bold': True, 'italic': True}),
           (" : start from the demand function and use a 3-step method",
            {'bold': True})], 0, {})],
        size=24)
    _draw_footer(slide, FOOTER_TEXT, 16)
    return slide


def v23_calculus(prs):
    """Calculus refresher.

    2026-08-24 (Nico): the whole thing used to be crammed into the top
    left.  Now it breathes - the definition sits in its own cream card at
    the top, the general rule and the worked example are two columns
    underneath, and each formula gets its own line at a readable size.
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V2)
    _draw_action_title(slide, "Calculus Refresher: Derivatives")
    # --- what a derivative IS -------------------------------------------
    _add_convention_box(
        # kept under 10" wide so _group_pass rule 1 treats it as a
        # callout (a wider filled roundRect reads as a layout band)
        slide, Inches(1.767), Inches(1.62), Inches(9.80), Inches(1.30),
        runs=[("A derivative measures a rate of change.", {'bold': True}),
              ("\n", {}),
              ("dY/dX is the change in ", {}), ("Y", {'italic': True}),
              (" caused by a small (marginal) change in ", {}),
              ("X", {'italic': True})],
        size=22, align=PP_ALIGN.CENTER)
    # --- left column: the general rule ----------------------------------
    _add_text(slide, Inches(0.75), Inches(3.15), Inches(5.7),
              Inches(0.40), "The general rule", size=22, bold=True,
              color=NAVY, font="Calibri")
    _add_mixed_textbox(
        slide, Inches(0.75), Inches(3.68), Inches(5.7), Inches(2.30),
        [
            ("text", "If   ", {'size': 21}),
            ("omml", _omml_run('Y') + _omml_text(' = ') + _omml_run('a')
             + _omml_text(' + ') + _omml_run('bX') + _omml_text(' + ')
             + _omml_run('c') + _omml_sup(_omml_run('X'),
                                          _omml_text('2')),
             {'size': 21}),
            ("break", None, None),
            ("text", "with a, b, c constant,", {'size': 18}),
            ("break", None, None),
            ("break", None, None),
            ("text", "then   ", {'size': 21}),
            ("omml", _omml_frac(_omml_run('dY'), _omml_run('dX'))
             + _omml_text(' = ') + _omml_run('b') + _omml_text(' + 2')
             + _omml_run('cX'), {'size': 21, 'bold': True}),
        ])
    # --- right column: the same rule on numbers -------------------------
    _add_text(slide, Inches(7.05), Inches(3.15), Inches(5.55),
              Inches(0.40), "A worked example", size=22, bold=True,
              color=NAVY, font="Calibri")
    _add_mixed_textbox(
        slide, Inches(7.05), Inches(3.68), Inches(5.55), Inches(2.30),
        [
            ("text", "With a = 1, b = 3, c = 2:", {'size': 18}),
            ("break", None, None),
            ("omml", _omml_run('Y') + _omml_text(' = 1 + 3')
             + _omml_run('X') + _omml_text(' + 2')
             + _omml_sup(_omml_run('X'), _omml_text('2')),
             {'size': 21}),
            ("break", None, None),
            ("break", None, None),
            ("omml", _omml_frac(_omml_run('dY'), _omml_run('dX'))
             + _omml_text(' = 3 + 2⋅2⋅') + _omml_run('X')
             + _omml_text(' = 3 + 4') + _omml_run('X'),
             {'size': 21, 'bold': True}),
        ])
    _add_outlined_box(slide, Inches(4.55), Inches(6.35), Inches(4.25),
                      Inches(0.55), "→  See TA’s Math Review Videos",
                      line=GOLD, text_color=NAVY, size=18, bold=True,
                      rounded=True, shadow=True, corner_pct=0.25)
    _draw_footer(slide, FOOTER_TEXT, 17)
    return slide


def v24_three_step(prs):
    """The 3-step method worked on Q = 1,600 - 4P.

    2026-08-24 (Nico): each step now carries CT's explanatory sub-line
    (CT video slide 23) under the bold step label, so the summary slide
    that used to follow could be dropped.  The parenthetical after each
    label is his hand-edit and stays un-bolded.
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V2)
    _draw_action_title(slide, "3-Step Method to Compute Marginal Revenue")
    _add_convention_box(
        slide, MARGIN + Inches(0.15), Inches(1.55), Inches(5.9),
        Inches(0.75),
        runs=[("Suppose demand is:   Q = 1,600 − 4P", {'bold': True})],
        size=18)
    steps = [
        ("Step 1:  Calculate inverse demand", " (solve for P)",
         [("Rearrange demand so ", {}), ("P", {'italic': True}),
          (" is a function of ", {}), ("Q", {'italic': True}),
          (":  P(Q)", {'italic': True})],
         _omml_run('P') + _omml_text(' = 400 − ')
         + _omml_frac(_omml_text('1'), _omml_text('4'))
         + _omml_run('Q')),
        ("Step 2:  Calculate total revenue ", "(multiply P(Q) by Q)",
         [("TR = P · Q", {'italic': True}),
          ("; replace P with inverse demand so TR is a function "
           "of ", {}), ("Q", {'italic': True}), (" only", {})],
         _omml_text('TR') + _omml_text(' = (400 − ')
         + _omml_frac(_omml_text('1'), _omml_text('4'))
         + _omml_run('Q') + _omml_text(') ∙ ') + _omml_run('Q')
         + _omml_text(' = 400∙') + _omml_run('Q') + _omml_text(' − ')
         + _omml_frac(_omml_text('1'), _omml_text('4'))
         + _omml_sup(_omml_run('Q'), _omml_text('2'))),
        ("Step 3:  Compute Marginal Revenue from total revenue ",
         "(MR = dTR/dQ)",
         [("Take the derivative of ", {}), ("TR", {'italic': True}),
          (" with respect to ", {}), ("Q", {'italic': True})],
         _omml_text('MR') + _omml_text(' = 400 − (2⋅')
         + _omml_frac(_omml_text('1'), _omml_text('4'))
         + _omml_text(')∙') + _omml_run('Q') + _omml_text(' = 400 − ')
         + _omml_frac(_omml_text('1'), _omml_text('2'))
         + _omml_run('Q')),
    ]
    y = Inches(2.42)
    for bold_part, tail, sub_runs, omml in steps:
        _add_hierarchical_bullets(
            slide, MARGIN + Inches(0.15), int(y + Inches(0.10)),
            Inches(5.85), Inches(1.20),
            [([(bold_part, {'bold': True}), (tail, {})], 0,
              {'bullet_style': 'none'}),
             ([(t, dict(o, color=GRAY)) for t, o in sub_runs], 0,
              {'bullet_style': 'none', 'size': 17,
               'space_before_pts': 5})],
            size=19)
        _add_math_equation(slide, Inches(6.45), y, Inches(6.3),
                           Inches(1.15), omml, size_pt=21, color=NAVY,
                           fill=CREAM, line=NAVY, rounded=True)
        y = int(y + Inches(1.40))
    # split into two pointers on 2026-08-24 (Nico)
    # exercise numbers dropped by hand 2026-08-24 - the reference names
    # the problem set only, so it survives next year's re-numbering
    _add_reference_box(slide, Inches(3.40), Inches(6.52), Inches(2.60),
                       Inches(0.50), "Problem Set 2", kind="ps")
    _add_reference_box(slide, Inches(6.30), Inches(6.52), Inches(3.60),
                       Inches(0.50), "Teaching Note: Marginal Revenue",
                       kind="tn")
    _vnote(slide, 16)
    _draw_footer(slide, FOOTER_TEXT, 18)
    return slide


def v25_three_step_summary(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V2)
    _draw_action_title(slide, "The “3-Step Method”: Summary Notes")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.35), Inches(1.85), Inches(12.0),
        Inches(4.2),
        [
            ("1.  Calculate inverse demand", 0,
             {'bold': True, 'bullet_style': 'none'}),
            ("Rearrange demand such that P is expressed as a function "
             "of Q  (see Module 2, Video 1)", 1),
            ("2.  Calculate total revenue   TR = P ∙ Q", 0,
             {'bold': True, 'bullet_style': 'none'}),
            ("Replace P by inverse demand so that TR is expressed as a "
             "function of Q only", 1),
            ("3.  Calculate marginal revenue from total revenue", 0,
             {'bold': True, 'bullet_style': 'none'}),
            ("Derivation of TR with respect to Q", 1),
        ],
        size=24, sub_size=22, line_spacing_pts=14)
    _add_outlined_box(slide, MARGIN, Inches(6.42), Inches(5.2),
                      Inches(0.5),
                      "→  PS 2 + Teaching Note Marginal Revenue",
                      line=GOLD, text_color=NAVY, size=18, bold=True,
                      rounded=True, shadow=True, corner_pct=0.25)
    _draw_footer(slide, FOOTER_TEXT, 21)
    return slide


def v26_mr_graph(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V2)
    _draw_action_title(slide, "MR: Graphical Representation")
    figD, figT = _dtr_figs(slide)
    # navy, matching the inverse demand on slide 6 (Nico, 2026-08-24)
    _add_arrow(slide, (figD.x(0), figD.y(400)), (figD.x(1600), figD.y(0)),
               color=NAVY, weight_pt=2.75, head=False)
    _add_text(slide, figD.x(1620), figD.y(90), Inches(0.5), Inches(0.35),
              "D", size=18, bold=True, color=NAVY, font="Calibri")
    # MR = 400 − Q/2: exactly from (0,400) to (800,0)
    _add_arrow(slide, (figD.x(0), figD.y(400)), (figD.x(800), figD.y(0)),
               color=CBLUE, weight_pt=2.75, head=False)
    _add_text(slide, figD.x(390), figD.y(155), Inches(0.7), Inches(0.35),
              "MR", size=16, bold=True, italic=True, color=CBLUE,
              font="Calibri")
    _fig_ytick(slide, figD, 400, "400", size=15)
    _fig_xtick(slide, figD, 800, "800", size=15)
    _fig_xtick(slide, figD, 1600, "1600", size=15)
    # braces over the two stretches of the demand curve, from Nico's
    # own Module 2 Video 2 slide 8 (2026-08-24)
    _brace_along(slide, (figD.x(0) / 914400.0, figD.y(400) / 914400.0),
                 (figD.x(800) / 914400.0, figD.y(200) / 914400.0))
    _brace_along(slide, (figD.x(800) / 914400.0, figD.y(200) / 914400.0),
                 (figD.x(1600) / 914400.0, figD.y(0) / 914400.0))
    _fig_point(slide, figD, 800, 200, fill=NAVY, r_in=0.055)
    _add_text(slide, figD.x(690), figD.y(255), Inches(1.5), Inches(0.32),
              "Eᴅ = −1", size=16, bold=True, color=NAVY,
              font="Calibri")
    # CT slide 26's treatment: 16 pt bold concept blue, sitting on the
    # stretch of the demand curve each one names
    # both sit in clear zones ABOVE the demand line, the way CT places
    # them on its slide 26
    _add_text(slide, figD.x(150), figD.y(440), Inches(2.4),
              Inches(0.34), "Elastic portion", size=16, bold=True,
              color=CBLUE, font="Calibri")
    _add_text(slide, figD.x(1150), figD.y(230), Inches(2.2),
              Inches(0.34), "Inelastic portion", size=16, bold=True,
              color=CBLUE, font="Calibri")
    _tr_parabola(slide, figT)
    _add_text(slide, figT.x(1250), figT.y(130000), Inches(0.8),
              Inches(0.35), "TR", size=16, bold=True, italic=True,
              color=GOLD, font="Calibri")
    _fig_xtick(slide, figT, 800, "800", size=15)
    _add_arrow(slide, (figD.x(800), figD.y(200)),
               (figT.x(800), figT.y(0)), color=GRAY, weight_pt=1.25,
               head=False, dash="dash")
    _add_convention_box(
        slide, Inches(8.75), Inches(1.70), Inches(4.25), Inches(2.15),
        runs=[("Q = 1,600 − 4P", {'italic': True}),
              ("\nP = 400 − Q/4", {'italic': True}),
              ("\nTR = 400Q − Q²/4", {'italic': True}),
              ("\nMR = 400 − Q/2",
               {'italic': True, 'bold': True, 'color': CBLUE})],
        size=17)
    _add_convention_box(
        slide, Inches(8.75), Inches(4.35), Inches(4.25), Inches(1.0),
        runs=[("TR is max. where MR = 0",
               {'bold': True, 'color': CBLUE})], size=17)
    _add_arrow(slide, (Inches(8.70), Inches(4.85)),
               (figT.x(830), figT.y(158000)), color=GOLD, weight_pt=1.75,
               head=True)
    _vnote(slide, 18)
    _draw_footer(slide, FOOTER_TEXT, 21)
    return slide


def v27_mr_vs_price(prs):
    return make_content_bulleted(
        prs, 22, TAG_V2, "Why Is MR Different From the Price?",
        [
            ([("MR", {'italic': True}),
              (" : additional revenue from selling one more unit. "
               "Sounds like the price! But…", {})], 0, {}),
            ("Law of demand: you need to lower the price to sell one "
             "more unit", 0),
            ([("The lower price will apply to ", {}),
              ("all", {'italic': True}),
              (" units, not just the extra one", {})], 0, {}),
            ([("MR", {'italic': True}),
              (" from selling one additional unit takes into account "
               "the loss on every other unit  →  ", {}),
              ("MR < P", {'bold': True, 'italic': True})], 0, {}),
            # rewritten by hand 2026-08-25 to state the MR < 0 case
            ([("Note", {'underline': True}),
              (": ", {}), ("When ", {}),
              ("MR < 0, ", {'bold': True, 'italic': True, 'size': 20}),
              ("Total", {'bold': True}), (" Revenues (", {}),
              ("TR", {'italic': True}), (") decrease", {}),
              (" as price falls (and ", {}), ("Q", {'italic': True}),
              (" increases", {}), (")", {})], 0, {'size': 22}),
        ],
        # column narrowed to 11.49" by hand, 2026-08-25
        size=24, sub_size=22, bullets_width=Inches(11.49))


def v29_mr_solution(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V2)
    _custom_title_runs(slide, [
        ("Solution: If  ", {}),
        ("Q = 10 − 0.5P", {'italic': True}),
        (" , What Is  ", {}), ("MR", {'italic': True}), (" ?", {})])
    steps = [
        ("Inverse Demand:",
         _omml_run('P') + _omml_text(' = 20 − 2') + _omml_run('Q')),
        ("Total revenue:  TR = P ∙ Q",
         _omml_text('TR') + _omml_text(' = (20 − 2') + _omml_run('Q')
         + _omml_text(') ∙ ') + _omml_run('Q')
         + _omml_text(' = 20') + _omml_run('Q') + _omml_text(' − 2')
         + _omml_sup(_omml_run('Q'), _omml_text('2'))),
        ("Derivative of TR w.r.t. Q:",
         _omml_text('MR') + _omml_text(' = ')
         + _omml_frac(_omml_run('d') + _omml_text('TR'),
                      _omml_run('d') + _omml_run('Q'))
         + _omml_text(' = 20 − 4') + _omml_run('Q')),
    ]
    y = Inches(2.15)
    for i, (label, omml) in enumerate(steps):
        _add_hierarchical_bullets(
            slide, MARGIN + Inches(0.15), int(y + Inches(0.22)),
            Inches(5.4), Inches(0.9),
            [("%d.  %s" % (i + 1, label), 0,
              {'bold': True, 'bullet_style': 'none'})],
            size=20)
        _add_math_equation(slide, Inches(6.05), y, Inches(6.7),
                           Inches(1.1), omml, size_pt=22,
                           color=RED if i == 2 else NAVY,
                           fill=CREAM, line=NAVY, rounded=True)
        y = int(y + Inches(1.45))
    _draw_footer(slide, FOOTER_TEXT, 20)
    return slide


def v30_insideout(prs):
    """
    NOT BUILT since 2026-08-25: this slide moved to
    "Module 2 - Potential Practice Exercises.pptx".
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V2)
    _draw_action_title(slide, "Inside Out 2's Marginal Revenue")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.95), Inches(7.3),
        Inches(1.2),
        [([("Demand for streaming ", {}),
           ("Inside Out 2", {'italic': True}),
           (" on Disney+   (Q: viewers, in millions):", {})], 0, {})],
        size=24)
    _add_math_equation(
        slide, Inches(1.45), Inches(3.15), Inches(4.4), Inches(1.1),
        _omml_run('Q') + _omml_text(' = 20 − ')
        + _omml_frac(_omml_text('1'), _omml_text('2')) + _omml_run('P'),
        size_pt=28, color=NAVY, fill=CREAM, line=NAVY, rounded=True,
        shadow=True)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(4.75), Inches(7.3),
        Inches(1.7),
        [
            ("What is the revenue-maximizing price?", 0, {'bold': True}),
            ([("Why might Disney care about ", {}),
              ("revenue", {'italic': True}),
              (" maximization?", {})], 0, {}),
        ],
        size=24, line_spacing_pts=14)
    _vid_media(slide, "ct_insideout_image7.png", left=Inches(8.35),
               top=Inches(1.90), width=Inches(4.2))
    _draw_footer(slide, FOOTER_TEXT, 50)
    # 2026-08-24 (Nico): Poll Break, not Group Discussion
    _add_pollbreak_badge(slide)
    return slide


def v31_insideout_solution(prs):
    """
    NOT BUILT since 2026-08-25: this slide moved to
    "Module 2 - Potential Practice Exercises.pptx".
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V2)
    _draw_action_title(
        slide, "Solution: Inside Out 2's Revenue-Maximizing Price")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.75), Inches(12.4),
        Inches(0.6),
        [([("Revenue is maximized where ", {}),
           ("MR = 0", {'bold': True, 'color': CBLUE}),
           (" .  Compute MR using the 3-step method:", {})], 0,
          {'bullet_style': 'none'})],
        size=22)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.35), Inches(2.55), Inches(6.4),
        Inches(2.5),
        [
            ([("1.  Inverse demand:   ", {'bold': True}),
              ("P = 40 − 2Q", {'italic': True})], 0,
             {'bullet_style': 'none'}),
            ([("2.  Total revenue:   ", {'bold': True}),
              ("TR = 40Q − 2Q²", {'italic': True})], 0,
             {'bullet_style': 'none'}),
            ([("3.  Derivative:   ", {'bold': True}),
              ("MR = 40 − 4Q", {'italic': True})], 0,
             {'bullet_style': 'none'}),
        ],
        size=22, line_spacing_pts=16)
    _add_math_equation(
        slide, Inches(7.45), Inches(2.60), Inches(5.2), Inches(2.0),
        _omml_text('MR') + _omml_text(' = 0:  40 − 4')
        + _omml_run('Q') + _omml_text(' = 0  ⇒  ') + _omml_run('Q')
        + _omml_text(' = 10  ⇒  ') + _omml_run('P')
        + _omml_text(' = $20'),
        size_pt=24, color=RED)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(5.35), Inches(12.4),
        Inches(1.3),
        [([("Why revenue maximization? The marginal cost of one more "
            "stream is ", {}),
           ("close to zero", {'bold': True}),
           (" – so maximizing revenue is (almost) maximizing profit",
            {})], 0, {})],
        size=22)
    _draw_footer(slide, FOOTER_TEXT, 51)
    return slide


# ==========================================================================
#  VIDEO 3 — Demand Estimation
# ==========================================================================

def v33_outline(prs):
    slide = make_m2_outline(prs, 24, section_tag=TAG_VOUT,
                            highlight_set={5})
    return slide


def v34_how_estimate(prs):
    """CT slide 34."""
    slide = make_content_bulleted(
        prs, 25, TAG_V3, "How to Estimate a Demand Curve?",
        [
            ("Use regression analysis to estimate the relationship "
             "between price and quantity", 0),
            ("Need data linking prices to quantities", 0),
            ("Several approaches", 0),
            ("Transaction data (marketing firms, internet shopping, "
             "store club cards)", 1),
            ("Surveys", 1),
            ("Market experimentation", 1, {'bold': True}),
        ],
        size=26, sub_size=25)
    _vnote(slide, 23)
    return slide


def v35_abtest(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(
        slide, "A/B Testing: What Is the Demand Function for This Product?")
    _vid_media(slide, "ct_abtest_image8.png", left=Inches(4.85),
               top=Inches(1.60), height=Inches(5.35), width=None)
    _draw_footer(slide, FOOTER_TEXT, 26)
    return slide


def v36_amazon_exp(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "A/B Testing: Amazon “Experiments”")
    _vid_media(slide, "ct_amazonexp_image9.png", left=Inches(2.55),
               top=Inches(1.70), width=Inches(8.2))
    _draw_footer(slide, FOOTER_TEXT, 27)
    return slide


# Nico asked (2026-08-24) what the camelcamelcamel chart is doing here.
S33_NOTE = (
    "camelcamelcamel.com is a free Amazon price tracker. It has followed "
    "Amazon listings since 2008 and plots the price history of a single "
    "product, marking the highest and lowest price it has ever recorded. "
    "The blue line on this slide is one listing over about five months, "
    "and the point is how restless it is: the price moves almost daily "
    "and swings between roughly $35 and $84 for the same item.\n\n"
    "That is the illustration. Amazon is not running one A/B test and "
    "then settling on an answer, the way the previous slide's classic "
    "experiment does. It re-prices continuously, so the price variation "
    "we would need to estimate a demand curve is being generated all the "
    "time. The flip side is the warning that comes a few slides later: "
    "these price moves are not randomized, so quantity differences across "
    "high- and low-price days also reflect whatever else changed on those "
    "days (a holiday, a competitor's promotion, a stock-out).\n\n"
    "The product above the chart is a Labubu, in case that name means "
    "nothing to you. Labubu is a character - an elf-like creature with "
    "pointed ears and a wide, serrated grin - created by the Hong Kong "
    "illustrator Kasing Lung for his picture-book series THE MONSTERS, "
    "which he started in 2015 and drew from Nordic folklore. In 2019 "
    "Lung signed an exclusive licence with Pop Mart, a Chinese "
    "designer-toy company, which turned the character into small vinyl "
    "and plush collectibles sold in BLIND BOXES: you pay for a sealed "
    "box and only find out which variant of the series you got when you "
    "open it, with one rare 'secret' version per case.\n\n"
    "From about 2024 the toys became a global craze, helped by "
    "celebrities photographed with one clipped to a handbag, and Pop "
    "Mart's Labubu lines have been a large part of its revenue growth. "
    "The one on this slide is from the Coca-Cola collaboration series. "
    "That background is exactly why the price line is so restless: "
    "supply comes in limited drops, demand is faddish, and the listing "
    "gets re-priced constantly. Nothing in the economics turns on the "
    "product, though - any item with a volatile price would make the "
    "same point."
)




# --------------------------------------------------------------------------
# Slide 33's price history.  The screenshot's own axis labels come out at
# roughly 4.6 pt on the slide, so the plot area is cropped out of it
# (_mk_pricechart.py) and the labels are drawn natively at 17 pt.
#
# Crop is 942 x 431 px starting at (41, 12) of the original, where
#     price -> row   y = 440 - 6.9835 * (price - 30)
#     May..Sep ticks x = 180, 343, 500, 663, 826
# --------------------------------------------------------------------------

_S33_X, _S33_Y = 3.400, 3.450          # top-left of the plot image
_S33_W, _S33_H = 6.230, 2.850          # its size on the slide
_S33_SX = _S33_W / 942.0               # inches per source pixel
_S33_SY = _S33_H / 431.0


def _s33_row(price):
    """Source row of a price, then slide y (crop starts at row 12)."""
    return _S33_Y + (440.0 - 6.9835 * (price - 30.0) - 12.0) * _S33_SY


def _s33_col(src_x):
    return _S33_X + (src_x - 41.0) * _S33_SX


def _s33_price_chart(slide):
    _vid_media(slide, "camel_plot.png", left=Inches(_S33_X),
               top=Inches(_S33_Y), width=Inches(_S33_W),
               rounded=False, shadow=False)
    # y axis: the four tick labels, right-aligned against the axis
    for price in (90, 70, 50, 30):
        _add_text(slide, Inches(2.640), Inches(_s33_row(price) - 0.150),
                  Inches(0.680), Inches(0.300), "$%d" % price,
                  size=17, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.RIGHT)
    # x axis: the month ticks
    for src_x, month in ((180, "May"), (343, "Jun"), (500, "Jul"),
                         (663, "Aug"), (826, "Sep")):
        _add_text(slide, Inches(_s33_col(src_x) - 0.450), Inches(6.330),
                  Inches(0.900), Inches(0.300), month, size=17, bold=True,
                  color=NAVY, font="Calibri", align=PP_ALIGN.CENTER)
    # the three prices camelcamelcamel calls out: highest, current, lowest
    for price, label, color in ((83.99, "$83.99", CT_RED),
                                (62.98, "$62.98", CBLUE),
                                (35.00, "$35.00", CT_GREEN)):
        _add_text(slide, Inches(9.720), Inches(_s33_row(price) - 0.140),
                  Inches(1.100), Inches(0.280), label, size=16, bold=True,
                  color=color, font="Calibri", align=PP_ALIGN.LEFT)

def v37_amazon_recent(prs):
    """CT slide 37: the product shot sits above the price history, both
    centred, with the source link underneath."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Recent Amazon \u201cExperiments\u201d")
    # 2026-08-24 (Nico): name the product.  It is Pop Mart's blind-box
    # plush keyring from THE MONSTERS x Coca-Cola series - the character
    # is Labubu, and the packaging in the shot reads "THE MONSTERS".
    # The photo moved down and the price chart shrank a little to make
    # room for the caption above the photo.
    _add_text(slide, Inches(5.076), Inches(1.350), Inches(3.200),
              Inches(0.300),
              "Pop Mart \u00d7 Coca-Cola Labubu keychain",
              size=13, bold=True, italic=True, color=NAVY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _vid_media(slide, "ct_amazonrecent_image10.png", left=Inches(5.593),
               top=Inches(1.680), width=Inches(2.165))
    _s33_price_chart(slide)
    # narrow box centred under the plot, so the source line groups
    # TIGHTLY with the chart instead of dragging a slide-wide box along
    _add_text(slide, Inches(5.015), Inches(6.700), Inches(3.000),
              Inches(0.34), "camelcamelcamel.com", size=14, italic=True,
              color=CBLUE, font="Calibri", align=PP_ALIGN.CENTER)
    _set_notes(slide, S33_NOTE)
    _draw_footer(slide, FOOTER_TEXT, 28)
    return slide


def v38_econometrics(prs):
    """CT slide 38."""
    slide = make_content_bulleted(
        prs, 29, TAG_V3, "Econometric Estimates",
        [
            ([("Econometrics:  ", {'bold': True}),
              ("combine economics, statistics, and mathematical model "
               "building", {})], 0, {}),
            ("Develop hypotheses linking dependent variables (e.g., "
             "sales) to explanatory variables (e.g., price)", 0),
            ([("How:  ", {'bold': True}),
              ("Least Squares regression analysis", {})], 0, {}),
            ([("Covered in detail in your ", {}),
              ("Stats class", {'bold': True})], 1, {}),
            ("Today: A broad overview", 1),
            ([("What you need for our class: ", {}),
              ("Interpret and use the estimated coefficients",
               {'bold': True})], 1, {}),
        ],
        size=25, sub_size=24)
    _vnote(slide, 24)
    return slide


def v39_ols(prs):
    """CT slide 39: the estimating equation on a navy bar, what each
    coefficient means, and the elasticity that falls out of b.  CT sets
    that last formula as a pile of little text boxes; ours is native
    OMML in the cream box, per the deck's equation rule."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Ordinary Least Squares: Estimating the Demand Function")
    _add_rounded_filled_box(
        slide, Inches(4.170), Inches(1.550), Inches(5.000), Inches(1.000),
        "Q = a + b \u00b7 P", fill=NAVY, text_color=WHITE, size=32,
        bold=True, corner_pct=0.14)
    _add_hierarchical_bullets(
        slide, Inches(0.750), Inches(2.650), Inches(11.900), Inches(3.300),
        [
            ([("Q", {'bold': True, 'italic': True}),
              (" = dependent variable", {})], 0, {}),
            ([("P", {'bold': True, 'italic': True}),
              (" = independent variable", {})], 0, {}),
            ([("a", {'bold': True, 'italic': True}),
              (" = intercept with the y-axis (Q on y-axis)", {})], 0, {}),
            ([("b", {'bold': True, 'italic': True}),
              (" = slope of the regression line (demand curve)",
               {})], 0, {}),
        ],
        size=25, line_spacing_pts=14)
    _add_math_equation(
        slide, Inches(4.360), Inches(5.250), Inches(4.610), Inches(1.070),
        _oED() + _omml_text(' = ')
        + _omml_frac(_omml_text('\u0394') + _omml_run('Q'),
                     _omml_text('\u0394') + _omml_run('P'))
        + _omml_text(' \u00b7 ')
        + _omml_frac(_omml_run('P'), _omml_run('Q'))
        + _omml_text(' = ') + _omml_run('b') + _omml_text(' \u00b7 ')
        + _omml_frac(_omml_run('P'), _omml_run('Q')),
        size_pt=24, color=NAVY, fill=CREAM, line=NAVY, rounded=True)
    _vnote(slide, 25)
    _draw_footer(slide, FOOTER_TEXT, 30)
    return slide


AIRLINE_DATA = [(250, 64), (265, 33), (265, 37), (240, 83), (230, 111),
                (225, 137), (225, 109), (220, 96), (230, 59), (235, 83),
                (245, 90), (240, 105), (250, 75), (240, 91), (240, 112),
                (235, 102)]


def v40_airline_data(prs):
    """CT slide 40: the raw price / seats-sold table on the left, with
    the aircraft shot beside it."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Airline Ticket Experimentation: Data")
    rows = [("Price ($)", "Seats sold")] + [
        ("%.2f" % p_, str(q_)) for p_, q_ in AIRLINE_DATA]
    # 2026-08-24 (Nico): 17 rows at 16 pt ran past the footer rule.
    # Row height is line height + the two vertical cell margins, so the
    # table is squeezed with a smaller font AND thinner margins.
    _add_styled_table(slide, Inches(2.950), Inches(1.560), Inches(4.350),
                      Inches(4.700), rows,
                      col_widths=[Inches(2.175), Inches(2.175)],
                      font_size=14, header_size=14,
                      margin_v=Inches(0.015))
    _vid_media(slide, "ct_airline_plane.png", left=Inches(8.150),
               top=Inches(2.900), width=Inches(4.200),
               rounded=False, shadow=False)
    _draw_footer(slide, FOOTER_TEXT, 31)
    return slide


def _scatter_fig(slide, *, x_left=3.400, y_bottom=6.300, w=6.72,
                 h=4.0, dot_fill=None, show_dots=True):
    """Shape-built scatter of the airline data, on CT's geometry: the
    x-axis starts at P = 210 and Q runs 0-150, so P = 220 lands at
    x 4.415" and Q = 150 at y 2.300" exactly as on CT's slides 41/42.
    Returns (fig, fx) where fx maps a PRICE to a slide x."""
    fig = SimpleFig(x_left, y_bottom, w, h, 66.2, 150.0)

    def fx(p_):
        return fig.x(p_ - 210)

    _fig_axes(slide, fig)
    _add_text(slide, Inches(x_left - 0.35), Inches(y_bottom - h - 0.62),
              Inches(1.9), Inches(0.34), "Quantity", size=17, bold=True,
              italic=True, color=NAVY, font="Calibri")
    _add_text(slide, Inches(x_left + w + 0.17), Inches(y_bottom + 0.125),
              Inches(1.6), Inches(0.34), "Price ($)", size=17, bold=True,
              italic=True, color=NAVY, font="Calibri")
    for p_ in (220, 240, 260):
        _fig_xtick(slide, fig, p_ - 210, str(p_), size=15)
    for q_ in (50, 100, 150):
        _fig_ytick(slide, fig, q_, str(q_), size=15)
    if show_dots:
        for p_, q_ in AIRLINE_DATA:
            _fig_point(slide, fig, p_ - 210, q_,
                       fill=dot_fill if dot_fill is not None else GOLD,
                       r_in=0.06, line=NAVY)
    return fig, fx


def v41_scatter(prs):
    """CT slide 41: the same data as a scatter, with two candidate lines
    drawn through it - which one fits best is the next slide."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "First Pass: Eyeball the Relationship")
    _add_text(slide, Inches(3.000), Inches(1.500), Inches(7.300),
              Inches(0.45), "Scatterplot of Q versus P", size=20,
              bold=True, italic=True, color=NAVY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _scatter_fig(slide)
    _add_arrow(slide, (Inches(4.415), Inches(2.753)),
               (Inches(9.086), Inches(5.500)), color=CBLUE,
               weight_pt=3.0, head=False)
    _add_arrow(slide, (Inches(4.415), Inches(3.180)),
               (Inches(9.289), Inches(4.967)), color=GOLD,
               weight_pt=3.0, head=False)
    # 2026-08-25 (Nico): the question the two candidate lines pose.  His
    # own copy is a hand-scaled group; the padding here reproduces the
    # rendered geometry (box 8.48/3.41 6.35x0.56, text 8.59/3.25).
    _draw_footer(slide, FOOTER_TEXT, 32)
    _add_convention_box(
        slide, Inches(8.480), Inches(3.410), Inches(3.620),
        Inches(0.560),
        runs=[("Which line provides the best fit?", {'italic': True})],
        size=18, align=PP_ALIGN.CENTER,
        pad_h=Inches(0.110), pad_v=Inches(-0.165))
    return slide


def v42_least_squares(prs):
    """CT slide 42: the fitted line with two of the vertical residuals
    marked, and the algorithm spelled out beside it."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Least Squares Estimation")
    _scatter_fig(slide, x_left=1.000, y_bottom=6.300, w=6.32, h=4.10)
    _add_arrow(slide, (Inches(2.240), Inches(3.205)),
               (Inches(7.200), Inches(5.536)), color=NAVY,
               weight_pt=3.0, head=False)
    for x_, y0, y1 in ((2.908, 4.687, 3.519), (4.338, 3.840, 4.191)):
        _add_arrow(slide, (Inches(x_), Inches(y0)), (Inches(x_), Inches(y1)),
                   color=CT_RED, weight_pt=1.8, head=False)
    _add_outlined_box(slide, Inches(7.950), Inches(2.500), Inches(4.850),
                      Inches(2.950), "", line=GOLD, fill=WHITE,
                      line_w=1.5, rounded=True, shadow=True,
                      corner_pct=0.10)
    _add_hierarchical_bullets(
        slide, Inches(8.200), Inches(2.620), Inches(4.450), Inches(2.700),
        [
            ("Least Squares Algorithm", 0,
             {'bold': True, 'bullet_style': 'none'}),
            ("1.  Calculate all vertical distances", 0,
             {'bullet_style': 'none'}),
            ("2.  Square them", 0, {'bullet_style': 'none'}),
            ("3.  Sum them over all data points", 0,
             {'bullet_style': 'none'}),
            ("4.  Choose the line that minimizes the sum of squares", 0,
             {'bullet_style': 'none'}),
        ],
        size=19, line_spacing_pts=10)
    _draw_footer(slide, FOOTER_TEXT, 33)
    return slide


def v43_regression1(prs):
    """CT slide 43: the estimated coefficients, the demand equation they
    give, and the hand-off to marginal revenue.  CT folds into one slide
    what we had split across two, so the old "Regression Results II"
    slide is gone."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Regression Results for the Airline Data")
    _add_styled_table(
        slide, Inches(3.570), Inches(1.600), Inches(6.200), Inches(2.460),
        [("Coefficient", "Value"),
         ("Constant (intercept)", "478.95"),
         ("Own price", "\u22121.64")],
        col_widths=[Inches(3.100), Inches(3.100)], font_size=26,
        header_size=24)
    _add_convention_box(
        slide, Inches(3.850), Inches(4.500), Inches(5.650), Inches(1.050),
        runs=[("Q = 478.95 \u2212 1.64 P", {'italic': True})],
        size=30, align=PP_ALIGN.CENTER)
    # 2026-08-25 (Nico, second pass): the line now points back at the
    # 3-step method, and the parenthetical is set regular, not bold
    _add_hierarchical_bullets(
        slide, Inches(1.000), Inches(6.100), Inches(11.330), Inches(0.71),
        [([("Now that we have an estimated demand curve, we can "
            "compute MR ", {'bold': True, 'italic': True}),
           ("(Following the 3 steps we’ve seen above)", {})], 0,
          {'bullet_style': 'none', 'align': PP_ALIGN.CENTER})],
        size=21)
    _vnote(slide, 29)
    _draw_footer(slide, FOOTER_TEXT, 34)
    return slide


def v44_regression2(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Regression Results II")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.95), Inches(12.4),
        Inches(3.0),
        [
            ([("R Square: Coefficient of determination (0 < ", {}),
              ("R", {'italic': True}), ("² < 1)", {})], 0, {}),
            ("R² = 0.56", 1),
            ("t-statistics (rule of thumb: |t| > 2 statistically "
             "significant)", 0),
            ("Own-price |t-stat| = 4.45", 1),
            ("Predicted values: Plug in for price to see predicted "
             "quantity", 0),
        ],
        size=24, sub_size=22)
    _add_math_equation(
        slide, Inches(2.35), Inches(5.15), Inches(8.6), Inches(1.0),
        _omml_text('e.g., ') + _omml_run('P') + _omml_text(' = 200.')
        + _omml_text('  Then:  ') + _omml_run('Q') + _omml_text(' = ')
        + _omml_run('a') + _omml_text(' + ') + _omml_run('b')
        + _omml_text(' ⋅ ') + _omml_run('P')
        + _omml_text(' = 479 − 1.64 ⋅ 200  ⇒  ') + _omml_run('Q')
        + _omml_text(' = 151'),
        size_pt=22, color=NAVY),
    _vnote(slide, 30)
    _draw_footer(slide, FOOTER_TEXT, 39)
    return slide


def v45_elasticity_from_est(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(
        slide, "Obtain Elasticity from Estimated Demand Curve")
    _add_mixed_textbox(
        slide, MARGIN + Inches(0.35), Inches(1.85), Inches(12.2),
        Inches(2.6),
        [
            ("text", "We have estimated:   ", {'size': 24}),
            ("omml", _omml_run('Q') + _omml_text(' = ') + _omml_run('a')
             + _omml_text(' + ') + _omml_run('b') + _omml_text(' ∙ ')
             + _omml_run('P') + _omml_text(' = 479 − 1.64')
             + _omml_run('P'), {'size': 24}),
            ("break", None, None),
            ("text", "Derive demand elasticity:   ", {'size': 24}),
            ("omml", _omml_sub(_omml_run('E'), _omml_run('d'))
             + _omml_text(' = ')
             + _omml_frac(_omml_text('Δ') + _omml_run('Q'),
                          _omml_text('Δ') + _omml_run('P'))
             + _omml_text(' ∙ ')
             + _omml_frac(_omml_run('P'), _omml_run('Q'))
             + _omml_text(' = ') + _omml_run('b') + _omml_text(' ∙ ')
             + _omml_frac(_omml_run('P'), _omml_run('Q')),
             {'size': 24}),
            ("break", None, None),
            ("text", "Example: Compute elasticity at the price P = $200",
             {'size': 24, 'bold': True}),
        ])
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.55), Inches(4.55), Inches(12.0),
        Inches(0.5),
        [([("Q = 151", {'italic': True}),
           ("  (see previous slide)", {'size': 20})], 0,
          {'bullet_style': 'none'})],
        size=24)
    _add_math_equation(
        slide, Inches(2.85), Inches(5.25), Inches(7.6), Inches(1.1),
        _omml_sub(_omml_run('E'), _omml_run('d')) + _omml_text(' = −1.64 ⋅ ')
        + _omml_frac(_omml_text('200'), _omml_text('151'))
        + _omml_text(' = −2.17   (elastic demand)'),
        size_pt=24, color=RED)
    _vnote(slide, 31)
    _draw_footer(slide, FOOTER_TEXT, 40)
    return slide


def v46_multivariate(prs):
    """CT slide 44."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Note: We can also Run Multivariate Regressions")
    _add_text(slide, Inches(0.300), Inches(1.400), Inches(12.730),
              Inches(0.50), "Incorporating other determinants of demand",
              size=23, bold=True, italic=True, color=NAVY,
              font="Calibri", align=PP_ALIGN.CENTER)
    _add_math_equation(
        slide, Inches(3.170), Inches(2.050), Inches(7.000), Inches(1.050),
        _omml_run('Q') + _omml_text(' = ') + _omml_run('a')
        + _omml_text(' + ') + _omml_run('b') + _omml_text(' \u00b7 ')
        + _omml_run('P') + _omml_text(' + ') + _omml_run('c')
        + _omml_text(' \u00b7 ') + _omml_sub(_omml_run('P'),
                                              _omml_text('c'))
        + _omml_text(' + ') + _omml_run('d') + _omml_text(' \u00b7 ')
        + _omml_run('Y'),
        size_pt=31, color=WHITE, fill=NAVY, line=None, rounded=True)
    _add_hierarchical_bullets(
        slide, Inches(3.550), Inches(3.300), Inches(6.900), Inches(3.500),
        [
            ([("Q", {'bold': True, 'italic': True}),
              (" = amount purchased", {})], 0, {}),
            ([("P", {'bold': True, 'italic': True}),
              (" = own price", {})], 0, {}),
            ([("P\u1d04", {'bold': True, 'italic': True}),
              (" = competitor\u2019s price", {})], 0, {}),
            ([("Y", {'bold': True, 'italic': True}),
              (" = income", {})], 0, {}),
        ],
        size=29, line_spacing_pts=16)
    _vnote(slide, 30)
    _draw_footer(slide, FOOTER_TEXT, 39)
    return slide


def v47_added_vars(prs):
    """CT slide 45."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Results With Added Variables")
    _add_styled_table(
        slide, Inches(3.420), Inches(2.000), Inches(6.500), Inches(3.300),
        [("Coefficient", "Value"),
         ("Constant (intercept)", "26.70"),
         ("Own price", "\u22122.13"),
         ("Competitor\u2019s price", "1.04"),
         ("Income", "3.10")],
        col_widths=[Inches(3.900), Inches(2.600)], font_size=23,
        header_size=22)
    _add_convention_box(
        slide, Inches(2.850), Inches(5.650), Inches(7.650), Inches(0.980),
        runs=[("Q = 26.70 \u2212 2.13 P + 1.04 P\u1d04 + 3.10 Y",
               {'italic': True})],
        size=26, align=PP_ALIGN.CENTER)
    _vnote(slide, 31)
    _draw_footer(slide, FOOTER_TEXT, 40)
    return slide


def v48_application(prs):
    """CT slide 46: the estimated demand curve the polls work from."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Application Using the Airline Demand Function")
    # 2026-08-25 (Nico): the equation moved up and the three questions
    # the polls ask are now named on the slide, one paragraph each so
    # they can be revealed one at a time
    _add_convention_box(
        slide, Inches(3.400), Inches(1.790), Inches(6.500), Inches(1.150),
        runs=[("Q = 478.95 \u2212 1.64 P", {'italic': True})],
        size=28, align=PP_ALIGN.CENTER)
    _add_hierarchical_bullets(
        slide, Inches(1.000), Inches(3.390), Inches(11.330), Inches(2.120),
        [
            ("Use the estimated airline demand to compute:", 0,
             {'bullet_style': 'none', 'align': PP_ALIGN.CENTER,
              'bold': True}),
            ("What Is E\u1d05 at P = 140?", 0,
             {'bullet_style': 'none', 'align': PP_ALIGN.CENTER,
              'bold': True, 'mar_l': 457200, 'indent': -457200}),
            ("What Is MR at P = 140?", 0,
             {'bullet_style': 'none', 'align': PP_ALIGN.CENTER,
              'bold': True, 'mar_l': 457200, 'indent': -457200}),
            ("Can you provide a Pricing Recommendation?", 0,
             {'bullet_style': 'none', 'align': PP_ALIGN.CENTER,
              'bold': True, 'mar_l': 457200, 'indent': -457200}),
        ],
        size=24, line_spacing_pts=14)
    # 2026-08-24 (Nico): the Problem Set 2 pointer sits on THIS slide,
    # bottom-left, opposite the Poll Break badge (moved off slide 29)
    _add_reference_box(slide, Inches(1.000), Inches(6.520), Inches(2.600),
                       Inches(0.500), "Problem Set 2", kind="ps")
    _add_pollbreak_badge(slide)
    _vnote(slide, 32)
    _draw_footer(slide, FOOTER_TEXT, 35)
    return slide


def v49_ed_solution(prs):
    """CT slide 49."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Solution: E\u1d05 at P=140")
    _add_hierarchical_bullets(
        slide, Inches(1.000), Inches(1.450), Inches(11.800), Inches(5.550),
        [
            ("Q = 478.95 \u2212 1.64 P", 0,
             {'bold': True, 'italic': True}),
            ([("Recall  ", {}),
              ("E\u1d05 = (\u0394Q/\u0394P) \u00b7 (P/Q)",
               {'italic': True})], 0, {}),
            ([("\u0394Q/\u0394P", {'italic': True}),
              (" is the slope of demand  \u21d2  \u0394Q/\u0394P = ", {}),
              ("\u22121.64", {'bold': True})], 0, {}),
            ([("Plug ", {}), ("P", {'italic': True}),
              (" in to get ", {}), ("Q", {'italic': True}), (":  ", {}),
              ("Q = 478.95 \u2212 1.64 \u00b7 140 = 249.35",
               {'bold': True, 'italic': True})], 0, {}),
            # the line that delivers the answer - dark red when the
            # red-solution mode is on (2026-08-26, Nico)
            ([("E\u1d05 = ", _sol({'italic': True})),
              ("\u22121.64 \u00b7 (140 / 249.35) = ", _sol({})),
              ("\u22120.92", _sol({'bold': True}))], 0, {}),
            ([("Demand is ", {}),
              ("inelastic", {'bold': True, 'color': CBLUE})], 0, {}),
        ],
        size=27, line_spacing_pts=16)
    _vnote(slide, 33)
    _draw_footer(slide, FOOTER_TEXT, 36)
    return slide


def v50_mr_solution(prs):
    """CT slide 52."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Solution: MR at P=140")
    _add_hierarchical_bullets(
        slide, Inches(1.000), Inches(1.450), Inches(11.800), Inches(5.550),
        [
            ("3-step method", 0, {'bold': True}),
            ([("1.  Inverse demand:   ", {'bold': True}),
              ("P = 292 \u2212 (1/1.64) Q", {'italic': True})], 1,
             {'bullet_style': 'none', 'size': 26}),
            ([("2.  Total revenue:   ", {'bold': True}),
              ("TR = 292 Q \u2212 (1/1.64) Q\u00b2", {'italic': True})], 1,
             {'bullet_style': 'none', 'size': 26}),
            ([("3.  Derivative:   ", {'bold': True}),
              ("MR = 292 \u2212 1.22 Q", {'italic': True})], 1,
             {'bullet_style': 'none', 'size': 26}),
            ("MR at P = 140?", 0, {'bold': True}),
            # the line that delivers the answer (2026-08-26, Nico)
            ([("P = 140  \u2192  Q = 249.35  \u2192  ", _sol({})),
              ("MR = \u221212",
               _sol({'bold': True, 'italic': True}))], 1,
             {'bullet_style': 'none', 'size': 26}),
        ],
        size=27, sub_size=26, line_spacing_pts=14)
    _vnote(slide, 34)
    _draw_footer(slide, FOOTER_TEXT, 37)
    return slide


def v51_raise_price(prs):
    """CT slide 55: the verdict on the left, and the picture of why -
    at P = 140 the airline sits on the inelastic stretch, where MR is
    already negative."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Solution: Raise the Price")
    _add_hierarchical_bullets(
        slide, Inches(1.000), Inches(1.700), Inches(6.500), Inches(3.600),
        [
            ([("MR", {'bold': True, 'italic': True}),
              (" is negative", {})], 0, {}),
            ([("So you want to sell ", {}),
              ("less", {'bold': True})], 0, {}),
            ([("By ", {}),
              ("increasing the price", {'bold': True})], 0, {}),
        ],
        size=28, line_spacing_pts=18)
    _add_arrow(slide, (Inches(7.600), Inches(6.300)),
               (Inches(7.600), Inches(1.820)), color=NAVY,
               weight_pt=2.5, head=True)
    _add_arrow(slide, (Inches(7.600), Inches(6.300)),
               (Inches(12.800), Inches(6.300)), color=NAVY,
               weight_pt=2.5, head=True)
    _add_text(slide, Inches(7.270), Inches(1.420), Inches(0.660),
              Inches(0.34), "P", size=18, bold=True, italic=True,
              color=NAVY, font="Calibri")
    _add_text(slide, Inches(12.470), Inches(6.400), Inches(0.660),
              Inches(0.32), "Q", size=18, bold=True, italic=True,
              color=NAVY, font="Calibri")
    _add_arrow(slide, (Inches(7.600), Inches(2.495)),
               (Inches(12.114), Inches(6.300)), color=NAVY,
               weight_pt=3.0, head=False)
    _add_arrow(slide, (Inches(7.600), Inches(2.495)),
               (Inches(10.144), Inches(6.787)), color=GOLD,
               weight_pt=3.0, head=False)
    _add_arrow(slide, (Inches(9.852), Inches(6.300)),
               (Inches(9.852), Inches(4.398)), color=NAVY,
               weight_pt=1.4, head=False)
    _add_text(slide, Inches(8.467), Inches(4.710), Inches(1.300),
              Inches(0.34), "MR", size=22, bold=True, color=GOLD,
              font="Calibri")
    _add_text(slide, Inches(12.083), Inches(5.839), Inches(1.000),
              Inches(0.34), "D", size=22, bold=True, color=NAVY,
              font="Calibri")
    _add_text(slide, Inches(9.833), Inches(4.018), Inches(2.300),
              Inches(0.34), "E\u1d05 = \u22121", size=18, bold=True,
              color=NAVY, font="Calibri")
    _add_text(slide, Inches(8.587), Inches(2.770), Inches(1.900),
              Inches(0.34), "Elastic", size=18, bold=True, color=CBLUE,
              font="Calibri")
    _add_text(slide, Inches(11.011), Inches(4.856), Inches(2.100),
              Inches(0.34), "Inelastic", size=18, bold=True,
              color=CBLUE, font="Calibri")
    _vnote(slide, 35)
    _draw_footer(slide, FOOTER_TEXT, 38)
    return slide


def v52_transaction_issues(prs):
    """CT slide 56."""
    slide = make_content_bulleted(
        prs, 41, TAG_V3,
        "Issues With Transaction Data: Why Randomization Is Key",
        [
            ([("Transaction data:  ", {'bold': True}),
              ("prices and quantities evolving over time, across "
               "locations and stores", {})], 0, {}),
            ("Can we be sure the change in quantity is due to the price "
             "change?", 0),
            # 2026-08-25 (Nico): his own Hawaii numbers, with the two
            # observations on their own lines and a blank line before
            # each of the two questions
            ([("Airline example:  ", {'bold': True})], 0, {}),
            ("August: Sell 5000 tickets to Hawaii for $600", 1),
            ("September: Sell 3,500 tickets at $500", 1),
            ("", 1, {'bullet_style': 'none'}),
            ("Should we conclude demand slopes upward?", 1),
            ("", 1, {'bullet_style': 'none'}),
            ("What else changed between those two periods?", 1),
        ],
        size=25, sub_size=24)
    _vnote(slide, 36)
    return slide


def v53_coffee(prs):
    """CT slide 57: the wine-and-health correlation, with the causal
    reading spelled out so the next slide can knock it down."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Correlation Does Not Mean Causation")
    _add_hierarchical_bullets(
        slide, Inches(0.800), Inches(1.550), Inches(11.730), Inches(1.000),
        [([("Example:  ", {'bold': True}),
           ("on average, people who drink two glasses of wine a day are "
            "healthier than those who drink none", {})], 0,
          {'bullet_style': 'none'})],
        size=27)
    _add_hierarchical_bullets(
        slide, Inches(0.800), Inches(2.950), Inches(11.730), Inches(0.600),
        [([("Causal interpretation:  ", {'bold': True}),
           ("wine improves health (I wish!)", {'italic': True})], 0,
          {'bullet_style': 'none'})],
        size=26)
    _vid_media(slide, "ct_s57_image19.png", left=Inches(3.400),
               top=Inches(4.000), width=Inches(2.400))
    _add_arrow(slide, (Inches(6.100), Inches(4.850)),
               (Inches(7.400), Inches(4.850)), color=NAVY,
               weight_pt=3.0, head=True)
    _vid_media(slide, "ct_s57_image20.png", left=Inches(7.700),
               top=Inches(4.000), width=Inches(2.400))
    _vnote(slide, 37)
    _draw_footer(slide, FOOTER_TEXT, 42)
    return slide


def v54_omitted(prs):
    """CT slide 58: the third factor that drives both."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Correlation Does Not Mean Causation")
    _add_hierarchical_bullets(
        slide, Inches(0.800), Inches(1.550), Inches(11.730), Inches(0.800),
        [([("Potential problem:  ", {'bold': True}),
           ("some other factor is causing the correlation", {})], 0,
          {'bullet_style': 'none'})],
        size=27)
    _vid_media(slide, "ct_s58_image21.png", left=Inches(5.500),
               top=Inches(2.500), width=Inches(2.350))
    _add_arrow(slide, (Inches(5.650), Inches(4.200)),
               (Inches(4.650), Inches(5.000)), color=CT_RED,
               weight_pt=2.5, head=True)
    _add_arrow(slide, (Inches(7.550), Inches(4.200)),
               (Inches(8.550), Inches(5.000)), color=CT_RED,
               weight_pt=2.5, head=True)
    _vid_media(slide, "ct_s58_image19.png", left=Inches(3.300),
               top=Inches(5.050), width=Inches(2.300))
    _vid_media(slide, "ct_s58_image20.png", left=Inches(7.700),
               top=Inches(5.050), width=Inches(2.300))
    _vnote(slide, 38)
    _draw_footer(slide, FOOTER_TEXT, 43)
    return slide


def v55_spurious(prs):
    """CT slide 59: two of Tyler Vigen's spurious correlations."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Spurious Correlations")
    _vid_media(slide, "ct_s59_image22.png", left=Inches(0.464),
               top=Inches(2.000), width=Inches(5.822),
               rounded=False, shadow=True)
    _vid_media(slide, "ct_s59_image23.png", left=Inches(6.543),
               top=Inches(2.000), width=Inches(6.414),
               rounded=False, shadow=True)
    _add_text(slide, Inches(0.300), Inches(6.750), Inches(12.730),
              Inches(0.34), "tylervigen.com/spurious-correlations",
              size=14, italic=True, color=CBLUE, font="Calibri",
              align=PP_ALIGN.CENTER)
    _vnote(slide, 39)
    _draw_footer(slide, FOOTER_TEXT, 44)
    return slide


def v56_randomization(prs):
    """CT slide 60."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Why Randomization Is Key")
    _add_hierarchical_bullets(
        slide, Inches(1.200), Inches(1.550), Inches(11.000), Inches(5.200),
        [
            ([("Randomizing prices ensures that ", {}),
              ("\u201call else is equal\u201d", {'bold': True}),
              (" other than the price change", {})], 0, {}),
            ([("Allows us to estimate the ", {}),
              ("causal effect", {'bold': True, 'color': CBLUE}),
              (" of price on quantity demanded  \u2192  the demand "
               "curve!", {})], 0, {}),
        ],
        size=27, line_spacing_pts=18)
    _vnote(slide, 40)
    _draw_footer(slide, FOOTER_TEXT, 45)
    return slide


def v57_summary(prs):
    """
    NOT BUILT since 2026-08-25: this slide moved to
    "Module 2 - Potential Practice Exercises.pptx".
    """
    return make_m2_outline(prs, 58, section_tag="Module 2 · Summary",
                           title="Module 2: Summary", descriptions=True)


# ==========================================================================
#  build_video() — 57-slide registry
# ==========================================================================

def _backup_divider_slide(prs):
    """
    NOT BUILT since 2026-08-25: this slide moved to
    "Module 2 - Potential Practice Exercises.pptx".

Divider introducing the slides that did NOT go into the videos.

    Same family as the three video title slides - no top bar, no page
    number - so it reads as a section break rather than a content slide.
    """
    slide = _blank_slide(prs)
    _add_text(slide, 0, Inches(2.55), SLIDE_W, Inches(1.0),
              "Slides not Used in Videos",
              size=48, bold=True, color=NAVY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_rect(slide, int((SLIDE_W - Inches(4.0)) / 2), Inches(3.72),
              Inches(4.0), 54864, GOLD)
    _add_text(slide, 0, Inches(4.06), SLIDE_W, Inches(0.55),
              "Kept for reference: poll placeholders, further examples "
              "and the module summary",
              size=22, color=GRAY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_rect(slide, 0, Inches(7.15), SLIDE_W, Inches(0.02), RULE)
    _add_rect(slide, MARGIN, Inches(7.135), GOLD_W, Inches(0.05), GOLD)
    return slide


def build_video(out_path=None, red_solutions=False):
    global RED_SOLUTIONS
    RED_SOLUTIONS = bool(red_solutions)
    prs = Presentation()
    prs.slide_width = int(SLIDE_W)
    prs.slide_height = int(SLIDE_H)
    _video_title_slide(prs, "Elasticity and Revenue",
                       "Module 2  ·  Video 1")                     #  1
    v02_outline(prs)                                               #  2
    v03_plot_demand(prs)                                           #  3
    v04_demand_tr(prs)                                             #  4
    v05_price_change_tr(prs)                                       #  5
    v06_elasticity_tr(prs)                                         #  6
    v07_depends_on_ed(prs)                                         #  7
    v11_ozempic(prs)                                               #  8
    v12_ozempic_solution(prs)                                      #  9
    v13_profits(prs)                                               # 10
    v14_four_cases(prs)                                            # 11
    v17_mcdonalds(prs)                                             # 12
    _video_title_slide(prs, "Marginal Revenue",
                       "Module 2  ·  Video 2")                     # 13
    v20_outline(prs)                                               # 14
    v21_why_mr(prs)                                                # 15
    v22_mr_definition(prs)                                         # 16
    v23_calculus(prs)                                              # 17
    v24_three_step(prs)                                            # 18
    make_stub(prs, 19, TAG_V2, "Poll: MR when Q = 10 − 0.5P",
              STUB_POLL)                                           # 19
    v29_mr_solution(prs)                                           # 20
    v26_mr_graph(prs)                                              # 21
    v27_mr_vs_price(prs)                                           # 22
    _video_title_slide(prs, "Demand Estimation",
                       "Module 2  ·  Video 3")                     # 23
    v33_outline(prs)                                               # 24
    v34_how_estimate(prs)                                          # 25
    v35_abtest(prs)                                                # 26
    v36_amazon_exp(prs)                                            # 27
    v37_amazon_recent(prs)                                         # 28
    v38_econometrics(prs)                                          # 29
    v39_ols(prs)                                                   # 30
    v40_airline_data(prs)                                          # 31
    v41_scatter(prs)                                               # 32
    v42_least_squares(prs)                                         # 33
    v43_regression1(prs)                                           # 34
    v48_application(prs)                                           # 35
    v49_ed_solution(prs)                                           # 36
    v50_mr_solution(prs)                                           # 37
    v51_raise_price(prs)                                           # 38
    v46_multivariate(prs)                                          # 39
    v47_added_vars(prs)                                            # 40
    v52_transaction_issues(prs)                                    # 41
    v53_coffee(prs)                                                # 42
    v54_omitted(prs)                                               # 43
    v55_spurious(prs)                                              # 44
    v56_randomization(prs)                                         # 45
    # 2026-08-25: the deck ENDS here, at 45 slides - Videos 1, 2 and 3
    # laid end to end, matching "Videos Final" exactly.  Nico moved the
    # slides he did not use into "Module 2 - Potential Practice
    # Exercises.pptx": the Netflix pair, Mega Millions, the Inside Out 2
    # pair, the six PollEverywhere stubs and the module summary.  Their
    # builders are still in this file, uncalled, so that deck can be
    # regenerated - see _backup_divider_slide and the functions it used
    # to introduce.
    # CT's own source links, restored on the runs we adopted
    apply_ct_source_links(prs)

    # deck-wide subscript pass (2026-08-24, shared with the In-Class
    # build): every faked index becomes a real subscript run
    apply_subscripts(prs)

    out = Path(out_path) if out_path else OUT_DIR / VDECK
    prs.save(str(out))
    print(f"saved {out} — {len(prs.slides._sldIdLst)} slides")
    return out


if __name__ == "__main__":
    import sys as _s
    _args = [a for a in _s.argv[1:] if a != "--red-solutions"]
    build_video(_args[0] if _args else None,
                red_solutions="--red-solutions" in _s.argv[1:])
