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
    _add_outlined_box(slide, Inches(8.15), Inches(6.68), Inches(4.9),
                      Inches(0.72),
                      "▶  Teaching Note – Demand Elasticity and Total "
                      "Revenue\nOn BL under “Module 2 Post-Work”",
                      line=GOLD, text_color=NAVY, size=14, bold=True,
                      rounded=True, shadow=True, corner_pct=0.20)
    return slide


def v03_plot_demand(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "Plotting (Inverse) Demand")
    fig = SimpleFig(1.95, 6.30, 6.2, 4.2, 1800, 450)
    _fig_axes(slide, fig)
    _add_text(slide, Inches(fig.l - 0.55), Inches(fig.b - fig.h - 0.45),
              Inches(0.7), Inches(0.35), "P", size=20, bold=True,
              italic=True, color=NAVY, font="Calibri")
    _add_text(slide, Inches(fig.l + fig.w + 0.05), Inches(fig.b - 0.05),
              Inches(0.7), Inches(0.35), "Q", size=20, bold=True,
              italic=True, color=NAVY, font="Calibri")
    _add_arrow(slide, (fig.x(0), fig.y(400)), (fig.x(1600), fig.y(0)),
               color=NAVY, weight_pt=3.0, head=False)
    _fig_ytick(slide, fig, 400, "$400", size=18)
    _fig_xtick(slide, fig, 1600, "1600", size=18)
    _add_convention_box(
        slide, Inches(8.75), Inches(2.05), Inches(4.15), Inches(2.35),
        runs=[("Demand function:", {'underline': True}),
              ("\nQ = 1,600 – 4P", {'italic': True}),
              ("\nRearrange to get inverse demand function:",
               {'underline': True}),
              ("\nP = 400 – Q/4", {'italic': True, 'color': RED,
                                   'bold': True})],
        size=18)
    _add_hierarchical_bullets(
        slide, Inches(8.75), Inches(4.80), Inches(4.3), Inches(1.2),
        [([("This is the demand function that a ", {}),
           ("firm", {'underline': True}), (" faces", {})], 0,
          {'bullet_style': 'none', 'color': RED, 'size': 20})],
        size=20)
    _vnote(slide, 3)
    _draw_footer(slide, FOOTER_TEXT, 3)
    return slide


def v04_demand_tr(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "Relationship Between Demand and Revenue")
    figD, figT = _dtr_figs(slide)
    # demand + ticks
    _add_arrow(slide, (figD.x(0), figD.y(400)), (figD.x(1600), figD.y(0)),
               color=RED, weight_pt=2.75, head=False)
    _add_text(slide, figD.x(1620), figD.y(90), Inches(0.5), Inches(0.35),
              "D", size=18, bold=True, italic=True, color=RED,
              font="Calibri")
    _fig_ytick(slide, figD, 400, "$400", size=15)
    _fig_ytick(slide, figD, 200, "200", size=15)
    _fig_xtick(slide, figD, 800, "800", size=15)
    _fig_xtick(slide, figD, 1600, "1600", size=15)
    # guides: (800, 200) on D — exact
    _add_arrow(slide, (figD.x(0), figD.y(200)), (figD.x(800), figD.y(200)),
               color=GRAY, weight_pt=1.25, head=False, dash="dash")
    _add_arrow(slide, (figD.x(800), figD.y(200)), (figD.x(800), figD.y(0)),
               color=GRAY, weight_pt=1.25, head=False, dash="dash")
    # TR parabola + peak guide at Q=800 (exact peak)
    _tr_parabola(slide, figT)
    _add_text(slide, figT.x(1050), figT.y(150000), Inches(1.8),
              Inches(0.35), "Total revenue", size=16, bold=True,
              italic=True, color=GOLD, font="Calibri")
    _fig_ytick(slide, figT, 160000, "$160,000", size=15)
    _fig_xtick(slide, figT, 800, "800", size=15)
    _fig_xtick(slide, figT, 1600, "1600", size=15)
    _add_arrow(slide, (figT.x(0), figT.y(160000)),
               (figT.x(800), figT.y(160000)), color=GRAY, weight_pt=1.25,
               head=False, dash="dash")
    _add_arrow(slide, (figT.x(800), figT.y(160000)),
               (figT.x(800), figT.y(0)), color=GRAY, weight_pt=1.25,
               head=False, dash="dash")
    _add_convention_box(
        slide, Inches(8.75), Inches(1.85), Inches(4.25), Inches(2.75),
        runs=[("Inverse demand function:", {'underline': True}),
              ("\nP = 400 – Q/4", {'italic': True, 'color': RED}),
              ("\nTotal Revenue:  TR = P∙Q", {'underline': True}),
              ("\nPlug in for P:", {'underline': True}),
              ("\nTR = (400 – Q/4) ∙ Q", {'italic': True}),
              ("\n      = 400Q – Q²/4", {'italic': True})],
        size=17)
    _vnote(slide, 4)
    _draw_footer(slide, FOOTER_TEXT, 4)
    return slide


def v05_price_change_tr(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(
        slide, "Effect of Price Changes on Total Revenue: "
               "A Graphical Approach")
    fig = SimpleFig(3.15, 6.55, 6.0, 4.5, 10, 10)
    # rectangles first (behind), on exact D: P = 10 − Q
    p0, q0, p1, q1 = 7.0, 3.0, 5.0, 5.0
    lost = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, fig.x(0), fig.y(p0),
        fig.x(q0) - fig.x(0), fig.y(p1) - fig.y(p0))
    lost.fill.solid()
    lost.fill.fore_color.rgb = ROSE
    lost.line.color.rgb = RED
    lost.line.width = Pt(1.0)
    lost.shadow.inherit = False
    gain = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, fig.x(q0), fig.y(p1),
        fig.x(q1) - fig.x(q0), fig.y(0) - fig.y(p1))
    gain.fill.solid()
    gain.fill.fore_color.rgb = PALE_GREEN
    gain.line.color.rgb = GREEN
    gain.line.width = Pt(1.0)
    gain.shadow.inherit = False
    _fig_axes(slide, fig)
    _add_text(slide, Inches(fig.l - 0.55), Inches(fig.b - fig.h - 0.45),
              Inches(0.7), Inches(0.35), "P", size=20, bold=True,
              italic=True, color=NAVY, font="Calibri")
    _add_text(slide, Inches(fig.l + fig.w + 0.05), Inches(fig.b - 0.05),
              Inches(0.7), Inches(0.35), "Q", size=20, bold=True,
              italic=True, color=NAVY, font="Calibri")
    _add_arrow(slide, (fig.x(0.6), fig.y(9.4)), (fig.x(9.4), fig.y(0.6)),
               color=NAVY, weight_pt=2.75, head=False)
    _add_text(slide, fig.x(8.2), fig.y(2.6), Inches(1.6), Inches(0.4),
              "Demand", size=18, bold=True, italic=True, color=NAVY,
              font="Calibri")
    for pv, lbl in ((p0, "P₀"), (p1, "P₁")):
        _fig_ytick(slide, fig, pv, lbl, size=18)
        _add_arrow(slide, (fig.x(0), fig.y(pv)),
                   (fig.x(10 - pv), fig.y(pv)), color=GRAY,
                   weight_pt=1.0, head=False, dash="dash")
    for qv, lbl in ((q0, "Q₀"), (q1, "Q₁")):
        _fig_xtick(slide, fig, qv, lbl, size=18)
        _add_arrow(slide, (fig.x(qv), fig.y(10 - qv)),
                   (fig.x(qv), fig.y(0)), color=GRAY, weight_pt=1.0,
                   head=False, dash="dash")
    _add_text(slide, fig.x(0.4), fig.y(4.4), Inches(1.4), Inches(0.4),
              "TR₀", size=16, bold=True, italic=True, color=RED,
              font="Calibri")
    _add_text(slide, fig.x(1.6), fig.y(3.0), Inches(1.4), Inches(0.4),
              "TR₁", size=16, bold=True, italic=True, color=GREEN,
              font="Calibri")
    # callouts
    _add_convention_box(
        slide, Inches(9.55), Inches(1.85), Inches(3.4), Inches(0.95),
        runs=[("Reduction in TR from lower price", {'color': RED})],
        size=15)
    _add_arrow(slide, (Inches(9.50), Inches(2.35)),
               (fig.x(1.5), fig.y(6.0)), color=RED, weight_pt=1.75,
               head=True)
    _add_convention_box(
        slide, Inches(9.55), Inches(3.15), Inches(3.4), Inches(0.95),
        runs=[("Increase in TR from higher volume", {'color': GREEN})],
        size=15)
    _add_arrow(slide, (Inches(9.50), Inches(3.65)),
               (fig.x(4.0), fig.y(2.5)), color=GREEN, weight_pt=1.75,
               head=True)
    _add_rounded_filled_box(slide, Inches(9.55), Inches(4.75),
                            Inches(3.4), Inches(0.85),
                            "Which area is bigger?", fill=GOLD,
                            text_color=NAVY, size=20, bold=True,
                            corner_pct=0.15, shadow=True)
    _add_hierarchical_bullets(
        slide, Inches(9.55), Inches(5.95), Inches(3.5), Inches(0.9),
        [([("Total revenue  ", {}),
           ("TR = P · Q", {'italic': True, 'bold': True})], 0,
          {'bullet_style': 'none'})],
        size=16)
    _draw_footer(slide, FOOTER_TEXT, 5)
    return slide


def v06_elasticity_tr(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "Own-Price Elasticity (Eᴅ) and "
                              "Total Revenue")
    figD, figT = _dtr_figs(slide)
    _add_arrow(slide, (figD.x(0), figD.y(400)), (figD.x(1600), figD.y(0)),
               color=RED, weight_pt=2.75, head=False)
    _add_text(slide, figD.x(1620), figD.y(90), Inches(0.5), Inches(0.35),
              "D", size=18, bold=True, italic=True, color=RED,
              font="Calibri")
    _fig_point(slide, figD, 800, 200, fill=NAVY, r_in=0.06)
    _add_text(slide, figD.x(850), figD.y(255), Inches(1.6), Inches(0.35),
              "Eᴅ = −1", size=15, bold=True, color=RED, font="Calibri")
    _add_text(slide, figD.x(60), figD.y(430), Inches(2.4), Inches(0.32),
              "Eᴅ < −1   “elastic”", size=14, bold=True, color=NAVY,
              font="Calibri")
    _add_text(slide, figD.x(900), figD.y(140), Inches(2.8), Inches(0.32),
              "−1 < Eᴅ < 0   “inelastic”", size=14, bold=True,
              color=NAVY, font="Calibri")
    _tr_parabola(slide, figT)
    _add_text(slide, figT.x(1250), figT.y(130000), Inches(0.8),
              Inches(0.35), "TR", size=16, bold=True, italic=True,
              color=GOLD, font="Calibri")
    _add_arrow(slide, (figD.x(800), figD.y(200)),
               (figT.x(800), figT.y(0)), color=GRAY, weight_pt=1.25,
               head=False, dash="dash")
    _add_convention_box(
        slide, Inches(8.75), Inches(1.70), Inches(4.25), Inches(1.05),
        runs=[("Inverse demand function:", {'underline': True}),
              ("\nP = 400 – Q/4", {'italic': True, 'color': RED})],
        size=16)
    _add_convention_box(
        slide, Inches(8.75), Inches(3.10), Inches(4.25), Inches(1.5),
        runs=[("Elastic region:", {'bold': True}),
              ("  P falls → Q rises\n→ ", {}),
              ("TR rises", {'bold': True, 'color': GREEN})],
        size=16)
    _add_convention_box(
        slide, Inches(8.75), Inches(4.90), Inches(4.25), Inches(1.5),
        runs=[("Inelastic region:", {'bold': True}),
              ("  P falls → Q rises\n→ ", {}),
              ("TR falls", {'bold': True, 'color': RED})],
        size=16)
    _vnote(slide, 6)
    _draw_footer(slide, FOOTER_TEXT, 6)
    return slide


def v07_depends_on_ed(prs):
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
            ("text", "In % changes, this implies:   ", {'size': 24}),
            ("omml", _omml_text('%Δ') + _omml_text('TR')
             + _omml_text(' = %Δ') + _omml_run('P')
             + _omml_text(' + %Δ') + _omml_run('Q'),
             {'size': 24, 'bold': True}),
        ])
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(3.15), Inches(12.4),
        Inches(0.9),
        [
            ([("To assess the effect of a ", {}),
              ("price decrease", {'bold': True}),
              (" on revenue, ask:  ", {}),
              ("[note: if P decreases, Q must increase]",
               {'italic': True, 'size': 18})], 0, {}),
        ],
        size=24)
    _add_convention_box(
        slide, Inches(0.75), Inches(4.15), Inches(5.9), Inches(2.3),
        runs=[("Is the % reduction in P ", {}),
              ("smaller", {'bold': True, 'underline': True}),
              (" than the % increase in Q?\n", {}),
              ("If yes: total revenue rises\n", {}),
              ("→ Demand is ", {}),
              ("elastic", {'bold': True, 'color': CBLUE})],
        size=18)
    _add_convention_box(
        slide, Inches(6.95), Inches(4.15), Inches(5.9), Inches(2.3),
        runs=[("Is the % reduction in P ", {}),
              ("larger", {'bold': True, 'underline': True}),
              (" than the % increase in Q?\n", {}),
              ("If yes: total revenue declines\n", {}),
              ("→ Demand is ", {}),
              ("inelastic", {'bold': True, 'color': CBLUE})],
        size=18)
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
    _draw_action_title(slide, "Does a Price Cut Raise Revenue?")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(2.00), Inches(6.9),
        Inches(3.9),
        [
            ([("Novo Nordisk", {'bold': True}),
              (" lowered its U.S. price for Ozempic and Wegovy under a "
               "deal with the White House", {})], 0, {}),
            ("As a result, sales are forecasted to fall 5–13%", 0),
            ([("What does this imply about the ", {}),
              ("elasticity of demand", {'bold': True, 'color': CBLUE}),
              ("?", {})], 0, {}),
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
        "sales down by 5 to 13%. Ask yourself what a price cut that "
        "REDUCES revenue tells you about the elasticity of demand."))
    _draw_footer(slide, FOOTER_TEXT, 11)
    _add_groupdiscussion_badge(slide)
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
            ("A price cut caused total revenue to fall", 0),
            ("The quantity response was too small to offset the price "
             "cut", 0),
            ([("Inelastic:  ", {'bold': True, 'color': CBLUE}),
              ("|Eᴅ| < 1", {'italic': True})], 0, {}),
        ],
        size=28, line_spacing_pts=18)
    _draw_footer(slide, FOOTER_TEXT, 12)
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
            ("The price decrease caused total revenue to decrease", 0),
            ("It also raised the quantity produced/sold  →  costs "
             "increase", 0),
            ("So profits fell", 0),
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
    _draw_footer(slide, FOOTER_TEXT, 13)
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
        ["If P decreases", "Q increases by MORE than P falls",
         "Elastic", "Revenue increases"],
        ["If P decreases", "Q increases by LESS than P falls",
         "Inelastic", "Revenue decreases"],
        ["If P increases", "Q decreases by MORE than P rises",
         "Elastic", "Revenue decreases"],
        ["If P increases", "Q decreases by LESS than P rises",
         "Inelastic", "Revenue increases"],
    ]
    _add_styled_table(slide, Inches(0.85), Inches(3.05), Inches(11.6),
                      Inches(3.6), rows, font_size=16, header_size=16,
                      first_col_bold=True)
    _draw_footer(slide, FOOTER_TEXT, 14)
    return slide


def v15_netflix(prs):
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
    _draw_footer(slide, FOOTER_TEXT, 15)
    _add_groupdiscussion_badge(slide)
    return slide


def v16_netflix_solution(prs):
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
    _draw_footer(slide, FOOTER_TEXT, 16)
    return slide


def v17_mcdonalds(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V1)
    _draw_action_title(slide, "Why Is McDonald's Cutting Prices?")
    _vid_media(slide, "ct_mcdonalds_image3.png", left=Inches(3.40),
               top=Inches(1.60), width=Inches(6.5), rounded=False,
               shadow=False)
    _add_hierarchical_bullets(
        slide, Inches(1.20), Inches(4.00), Inches(11.0), Inches(2.8),
        [
            ("What does this suggest about its demand?", 0),
            ("How did it persuade franchisees to adopt the cuts?", 0),
            ("What does this reveal about demand across locations?", 0),
            ([("What are the implications for revenue and profit given "
               "a franchise's ", {}),
              ("Eᴅ", {'bold': True, 'color': CBLUE}), ("?", {})], 0, {}),
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
    _draw_footer(slide, FOOTER_TEXT, 17)
    _add_groupdiscussion_badge(slide)
    return slide


def v18_megamillions_revisited(prs):
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
    _draw_footer(slide, FOOTER_TEXT, 18)
    _add_groupdiscussion_badge(slide)
    return slide


# ==========================================================================
#  VIDEO 2 — Marginal Revenue
# ==========================================================================

def v20_outline(prs):
    return make_m2_outline(prs, 20, section_tag=TAG_VOUT,
                           highlight_set={4})


def v21_why_mr(prs):
    return make_content_bulleted(
        prs, 21, TAG_V2, "Why Do We Care About Marginal Revenue?",
        [
            ([("General objective: Maximize ", {}),
              ("net benefits", {'bold': True})], 0, {}),
            ("Use marginal analysis", 1),
            ([("Rule: net benefits are maximized where ", {}),
              ("MB = MC", {'italic': True, 'bold': True})], 1, {}),
            ("Example: Optimal amount of running", 1),
            ([("Firms’ objective: ", {}),
              ("Maximize profits", {'bold': True, 'color': RED})], 0,
             {}),
            ([("Rule: produce where  ", {}),
              ("Marginal Revenue = Marginal Cost   (MR = MC)",
               {'bold': True, 'italic': True})], 1, {}),
            ([("Today: how to compute ", {}),
              ("MR", {'bold': True, 'italic': True, 'color': RED})], 0,
             {}),
        ],
        size=24, sub_size=22)


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
               "increases by 1 unit", {})], 0, {}),
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
    _add_arrow(slide, (Inches(9.30), Inches(4.25)),
               (Inches(8.65), Inches(4.35)), color=GOLD, weight_pt=1.75,
               head=True)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(5.65), Inches(12.4),
        Inches(0.8),
        [([("How to compute ", {'bold': True}),
           ("MR", {'bold': True, 'italic': True}),
           (" : start from the demand function and use a 3-step method",
            {'bold': True})], 0, {})],
        size=24)
    _draw_footer(slide, FOOTER_TEXT, 22)
    return slide


def v23_calculus(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V2)
    _draw_action_title(slide, "Calculus Refresher: Derivatives")
    _add_mixed_textbox(
        slide, MARGIN + Inches(0.35), Inches(1.80), Inches(12.2),
        Inches(4.4),
        [
            ("omml", _omml_frac(_omml_run('dY'), _omml_run('dX')),
             {'size': 22}),
            ("text", " :  derivative of Y with respect to X",
             {'size': 22}),
            ("break", None, None),
            ("text", "Corresponds to the change in Y due to a small "
             "(marginal) change in X", {'size': 22}),
            ("break", None, None),
            ("text", "Suppose   ", {'size': 22}),
            ("omml", _omml_run('Y') + _omml_text(' = ') + _omml_run('a')
             + _omml_text(' + ') + _omml_run('bX') + _omml_text(' + ')
             + _omml_run('c') + _omml_sup(_omml_run('X'),
                                          _omml_text('2')),
             {'size': 22}),
            ("text", "    where a, b, c are constants", {'size': 20}),
            ("break", None, None),
            ("text", "E.g., a=1, b=3, c=2:   ", {'size': 22}),
            ("omml", _omml_run('Y') + _omml_text(' = 1 + 3')
             + _omml_run('X') + _omml_text(' + 2')
             + _omml_sup(_omml_run('X'), _omml_text('2')),
             {'size': 22}),
            ("break", None, None),
            ("text", "Then:   ", {'size': 22}),
            ("omml", _omml_frac(_omml_run('dY'), _omml_run('dX'))
             + _omml_text(' = ') + _omml_run('b') + _omml_text(' + 2')
             + _omml_run('cX'), {'size': 22, 'bold': True}),
            ("break", None, None),
            ("text", "In the above example:   ", {'size': 22}),
            ("omml", _omml_frac(_omml_run('dY'), _omml_run('dX'))
             + _omml_text(' = 3 + 2⋅2⋅') + _omml_run('X')
             + _omml_text(' = 3 + 4') + _omml_run('X'), {'size': 22}),
        ])
    _add_outlined_box(slide, MARGIN, Inches(6.42), Inches(4.1),
                      Inches(0.5), "→  See TA’s Math Review Videos",
                      line=GOLD, text_color=NAVY, size=18, bold=True,
                      rounded=True, shadow=True, corner_pct=0.25)
    _draw_footer(slide, FOOTER_TEXT, 23)
    return slide


def v24_three_step(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V2)
    _draw_action_title(slide, "3-Step Method in Action")
    _add_convention_box(
        slide, MARGIN + Inches(0.15), Inches(1.65), Inches(5.9),
        Inches(0.75),
        runs=[("Suppose demand is:   Q = 1,600 − 4P", {'bold': True})],
        size=18)
    steps = [
        ("Step 1:  Calculate inverse demand (solve for P)",
         _omml_run('P') + _omml_text(' = 400 − ')
         + _omml_frac(_omml_text('1'), _omml_text('4'))
         + _omml_run('Q')),
        ("Step 2:  Calculate total revenue (multiply P by Q)",
         _omml_text('TR') + _omml_text(' = (400 − ')
         + _omml_frac(_omml_text('1'), _omml_text('4'))
         + _omml_run('Q') + _omml_text(') ∙ ') + _omml_run('Q')
         + _omml_text(' = 400∙') + _omml_run('Q') + _omml_text(' − ')
         + _omml_frac(_omml_text('1'), _omml_text('4'))
         + _omml_sup(_omml_run('Q'), _omml_text('2'))),
        ("Step 3:  Marginal revenue from total revenue (MR = dTR/dQ)",
         _omml_text('MR') + _omml_text(' = 400 − (2⋅')
         + _omml_frac(_omml_text('1'), _omml_text('4'))
         + _omml_text(')∙') + _omml_run('Q') + _omml_text(' = 400 − ')
         + _omml_frac(_omml_text('1'), _omml_text('2'))
         + _omml_run('Q')),
    ]
    y = Inches(2.75)
    for label, omml in steps:
        _add_hierarchical_bullets(
            slide, MARGIN + Inches(0.15), int(y + Inches(0.22)),
            Inches(5.9), Inches(1.0),
            [(label, 0, {'bold': True, 'bullet_style': 'none'})],
            size=19)
        _add_math_equation(slide, Inches(6.45), y, Inches(6.3),
                           Inches(1.15), omml, size_pt=21, color=NAVY,
                           fill=CREAM, line=NAVY, rounded=True)
        y = int(y + Inches(1.35))
    _vnote(slide, 16)
    _draw_footer(slide, FOOTER_TEXT, 24)
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
    _draw_footer(slide, FOOTER_TEXT, 25)
    return slide


def v26_mr_graph(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V2)
    _draw_action_title(slide, "MR: Graphical Representation")
    figD, figT = _dtr_figs(slide)
    _add_arrow(slide, (figD.x(0), figD.y(400)), (figD.x(1600), figD.y(0)),
               color=RED, weight_pt=2.75, head=False)
    _add_text(slide, figD.x(1620), figD.y(90), Inches(0.5), Inches(0.35),
              "D", size=18, bold=True, italic=True, color=RED,
              font="Calibri")
    # MR = 400 − Q/2: exactly from (0,400) to (800,0)
    _add_arrow(slide, (figD.x(0), figD.y(400)), (figD.x(800), figD.y(0)),
               color=CBLUE, weight_pt=2.75, head=False)
    _add_text(slide, figD.x(390), figD.y(155), Inches(0.7), Inches(0.35),
              "MR", size=16, bold=True, italic=True, color=CBLUE,
              font="Calibri")
    _fig_ytick(slide, figD, 400, "400", size=15)
    _fig_xtick(slide, figD, 800, "800", size=15)
    _fig_xtick(slide, figD, 1600, "1600", size=15)
    _fig_point(slide, figD, 800, 200, fill=RED, r_in=0.055)
    _add_text(slide, figD.x(850), figD.y(255), Inches(1.5), Inches(0.32),
              "Eᴅ = −1", size=14, bold=True, color=RED, font="Calibri")
    _add_text(slide, figD.x(60), figD.y(435), Inches(2.0), Inches(0.3),
              "Elastic portion", size=13, italic=True, color=GRAY,
              font="Calibri")
    _add_text(slide, figD.x(950), figD.y(115), Inches(2.2), Inches(0.3),
              "Inelastic portion", size=13, italic=True, color=GRAY,
              font="Calibri")
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
    _draw_footer(slide, FOOTER_TEXT, 26)
    return slide


def v27_mr_vs_price(prs):
    return make_content_bulleted(
        prs, 27, TAG_V2, "Why Is MR Different From the Price?",
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
            ([("Note", {'underline': True}),
              (": Total Revenues (", {}), ("TR", {'italic': True}),
              (") may either increase or decrease when ", {}),
              ("Q", {'italic': True}),
              (" increases (depends on elastic vs. inelastic demand)",
               {})], 0, {'size': 22}),
        ],
        size=24, sub_size=22)


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
    _draw_footer(slide, FOOTER_TEXT, 29)
    return slide


def v30_insideout(prs):
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
    _draw_footer(slide, FOOTER_TEXT, 30)
    _add_groupdiscussion_badge(slide)
    return slide


def v31_insideout_solution(prs):
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
    _draw_footer(slide, FOOTER_TEXT, 31)
    return slide


# ==========================================================================
#  VIDEO 3 — Demand Estimation
# ==========================================================================

def v33_outline(prs):
    slide = make_m2_outline(prs, 33, section_tag=TAG_VOUT,
                            highlight_set={5})
    _add_outlined_box(slide, Inches(8.95), Inches(5.30), Inches(4.1),
                      Inches(0.72),
                      "✎  Problem Set 2\nOn BL under "
                      "“Module 2 Post-Work”",
                      line=GOLD, text_color=NAVY, size=15, bold=True,
                      rounded=True, shadow=True, corner_pct=0.20)
    return slide


def v34_how_estimate(prs):
    slide = make_content_bulleted(
        prs, 34, TAG_V3, "How to Estimate a Demand Curve?",
        [
            ("Use regression analysis to estimate the relationship "
             "between price and quantity", 0),
            ("Need data linking prices and quantities", 0),
            ("Several approaches", 0),
            ("Transaction data (marketing research firms, internet, "
             "« Ralph cards »)", 1),
            ("Surveys", 1),
            ("Market experimentation", 1, {'bold': True}),
        ],
        size=26, sub_size=24)
    _vnote(slide, 23)
    return slide


def v35_abtest(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(
        slide, "A/B Testing: What Is the Demand for This Product?")
    _vid_media(slide, "ct_abtest_image8.png", left=Inches(4.85),
               top=Inches(1.60), height=Inches(5.35), width=None)
    _draw_footer(slide, FOOTER_TEXT, 35)
    return slide


def v36_amazon_exp(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "A/B Testing: Amazon “Experiments”")
    _vid_media(slide, "ct_amazonexp_image9.png", left=Inches(2.55),
               top=Inches(1.70), width=Inches(8.2))
    _draw_footer(slide, FOOTER_TEXT, 36)
    return slide


def v37_amazon_recent(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Recent Amazon “Experiments”")
    _vid_media(slide, "ct_amazonrecent_image10.png", left=Inches(0.85),
               top=Inches(1.85), width=Inches(5.7))
    _vid_media(slide, "ct_amazonrecent_image11.png", left=Inches(6.85),
               top=Inches(1.85), width=Inches(5.7))
    _add_text(slide, Inches(0.85), Inches(6.60), Inches(11.7),
              Inches(0.35), "camelcamelcamel.com", size=13, italic=True,
              color=GRAY, font="Calibri", align=PP_ALIGN.CENTER)
    _draw_footer(slide, FOOTER_TEXT, 37)
    return slide


def v38_econometrics(prs):
    return make_content_bulleted(
        prs, 38, TAG_V3, "Econometric Estimates",
        [
            ([("Description:", {'bold': True}),
              (" Combination of economics, statistics, and mathematical "
               "model building", {})], 0, {}),
            ([("Develop hypotheses about relationships between ", {}),
              ("dependent variables", {'color': RED}),
              (" (quantity sold) and ", {}),
              ("independent/explanatory variables", {'color': RED}),
              (" (e.g., price)", {})], 0, {}),
            ([("Common method:", {'bold': True}), (" ", {}),
              ("Least squares regression", {'color': RED})], 0, {}),
        ],
        size=26)


def v39_ols(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Ordinary Least Squares: The Basic Equation")
    _add_math_equation(
        slide, Inches(4.35), Inches(1.70), Inches(4.6), Inches(1.15),
        _omml_run('Q') + _omml_text(' = ') + _omml_run('a')
        + _omml_text(' + ') + _omml_run('b') + _omml_text(' ∙ ')
        + _omml_run('P'),
        size_pt=32, color=NAVY, fill=CREAM, line=NAVY, rounded=True,
        shadow=True)
    _add_text(slide, Inches(9.25), Inches(2.00), Inches(1.4),
              Inches(0.4), "Slope", size=16, bold=True, color=RED,
              font="Calibri")
    _add_arrow(slide, (Inches(9.20), Inches(2.20)),
               (Inches(8.05), Inches(2.28)), color=GOLD, weight_pt=1.75,
               head=True)
    _add_hierarchical_bullets(
        slide, Inches(2.05), Inches(3.30), Inches(9.0), Inches(2.6),
        [
            ("where:", 0, {'bullet_style': 'none'}),
            ([("Q", {'italic': True}),
              (" = Quantity sold (dependent variable)", {})], 0,
             {'bullet_style': 'none'}),
            ([("P", {'italic': True}),
              (" = Price (explanatory variable)", {})], 0,
             {'bullet_style': 'none'}),
            ([("a", {'italic': True}),
              (" = Intercept with y-axis (", {}),
              ("Q", {'italic': True}), (" is on y-axis)", {})], 0,
             {'bullet_style': 'none'}),
            ([("b", {'italic': True}), (" = Slope of ", {}),
              ("regression line", {'color': RED})], 0,
             {'bullet_style': 'none'}),
        ],
        size=24, line_spacing_pts=12)
    _add_math_equation(
        slide, Inches(7.85), Inches(5.85), Inches(5.0), Inches(1.0),
        _omml_sub(_omml_run('E'), _omml_run('d')) + _omml_text(' = ')
        + _omml_frac(_omml_text('Δ') + _omml_run('Q'),
                     _omml_text('Δ') + _omml_run('P'))
        + _omml_text(' ∙ ')
        + _omml_frac(_omml_run('P'), _omml_run('Q'))
        + _omml_text(' = ') + _omml_run('b') + _omml_text(' ∙ ')
        + _omml_frac(_omml_run('P'), _omml_run('Q')),
        size_pt=20, color=NAVY)
    _vnote(slide, 25)
    _draw_footer(slide, FOOTER_TEXT, 39)
    return slide


AIRLINE_DATA = [(250, 64), (265, 33), (265, 37), (240, 83), (230, 111),
                (225, 137), (225, 109), (220, 96), (230, 59), (235, 83),
                (245, 90), (240, 105), (250, 75), (240, 91), (240, 112),
                (235, 102)]


def v40_airline_data(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Airline Ticket Experimentation")
    rows = [["Price ($)", "Seats sold"]]
    for p_, q_ in AIRLINE_DATA:
        rows.append(["%.2f" % p_, str(q_)])
    _add_styled_table(slide, Inches(4.55), Inches(1.62), Inches(4.2),
                      Inches(5.3), rows, font_size=13, header_size=14,
                      first_col_bold=False, first_col_align_left=False)
    _draw_footer(slide, FOOTER_TEXT, 40)
    return slide


def _scatter_fig(slide):
    """Shape-built scatter of the airline data: P on x (210–270),
    Q on y (0–150). Returns the fig for overlays."""
    fig = SimpleFig(2.45, 6.35, 8.2, 4.3, 66, 155)

    def fx(p_):
        return fig.x(p_ - 209)

    _fig_axes(slide, fig)
    _add_text(slide, Inches(fig.l - 1.15), Inches(fig.b - fig.h - 0.5),
              Inches(1.5), Inches(0.32), "Quantity", size=16, bold=True,
              italic=True, color=NAVY, font="Calibri")
    _add_text(slide, Inches(fig.l + fig.w - 0.3), Inches(fig.b + 0.08),
              Inches(1.4), Inches(0.32), "Price ($)", size=16, bold=True,
              italic=True, color=NAVY, font="Calibri")
    for p_ in (220, 230, 240, 250, 260):
        _fig_xtick(slide, fig, p_ - 209, str(p_), size=14)
    for q_ in (50, 100, 150):
        _fig_ytick(slide, fig, q_, str(q_), size=14)
    for p_, q_ in AIRLINE_DATA:
        _fig_point(slide, fig, p_ - 209, q_, fill=NAVY, r_in=0.05)
    return fig, fx


def v41_scatter(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(
        slide, "First Pass: Eyeball the Relationship — "
               "Is It Approximately Linear?")
    fig, fx = _scatter_fig(slide)
    _add_text(slide, Inches(fig.l + 1.2), Inches(1.55), Inches(4.5),
              Inches(0.4), "Scatterplot of Q versus P", size=16,
              italic=True, color=GRAY, font="Calibri")
    _draw_footer(slide, FOOTER_TEXT, 41)
    return slide


def v42_least_squares(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(
        slide, "Least Squares Estimation: Which Line Fits Best?")
    fig, fx = _scatter_fig(slide)
    # fitted line Q = 479 − 1.64 P (exact at both ends)
    _add_arrow(slide, (fx(214), fig.y(479 - 1.64 * 214)),
               (fx(268), fig.y(479 - 1.64 * 268)), color=RED,
               weight_pt=2.75, head=False)
    # two deliberately-worse candidates (dashed gray)
    _add_arrow(slide, (fx(214), fig.y(150)), (fx(268), fig.y(45)),
               color=GRAY, weight_pt=1.5, head=False, dash="dash")
    _add_arrow(slide, (fx(214), fig.y(95)), (fx(268), fig.y(75)),
               color=GRAY, weight_pt=1.5, head=False, dash="dash")
    # a few vertical distances to the fitted line
    for p_, q_ in ((225, 137), (230, 59), (245, 90)):
        _add_arrow(slide, (fx(p_), fig.y(q_)),
                   (fx(p_), fig.y(479 - 1.64 * p_)), color=GOLD,
                   weight_pt=1.5, head=False)
    _add_convention_box(
        slide, Inches(8.95), Inches(1.70), Inches(4.05), Inches(2.25),
        runs=[("Least Squares Algorithm:", {'bold': True}),
              ("\n1. Calculate all vertical distances", {}),
              ("\n2. Square them", {}),
              ("\n3. Sum them over all data points", {}),
              ("\n4. Choose the line that minimizes the sum of squares",
               {})],
        size=15)
    _draw_footer(slide, FOOTER_TEXT, 42)
    return slide


def v43_regression1(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(
        slide, "Regression Results I (From Excel Using “Data Analysis”)")
    rows = [["coefficient", "value"],
            ["Constant (intercept)", "479"],
            ["Own price", "−1.64"]]
    _add_styled_table(slide, Inches(2.05), Inches(1.75), Inches(6.4),
                      Inches(1.9), rows, font_size=18, header_size=18)
    _add_mixed_textbox(
        slide, Inches(2.55), Inches(4.15), Inches(9.0), Inches(0.8),
        [("text", "Demand Equation:   ",
          {'size': 26, 'bold': True, 'color': RED}),
         ("omml", _omml_run('Q') + _omml_text(' = 479 − 1.64')
          + _omml_run('P'), {'size': 26, 'color': RED})])
    _add_rounded_filled_box(slide, Inches(2.05), Inches(5.25),
                            Inches(9.2), Inches(0.8),
                            "Note: Now that we have demand, we can "
                            "compute MR!", fill=GOLD, text_color=NAVY,
                            size=20, bold=True, corner_pct=0.15,
                            shadow=True)
    _add_outlined_box(slide, MARGIN, Inches(6.42), Inches(5.6),
                      Inches(0.5),
                      "→  More detail: “Teaching Note Regressions”",
                      line=GOLD, text_color=NAVY, size=18, bold=True,
                      rounded=True, shadow=True, corner_pct=0.25)
    _vnote(slide, 29)
    _draw_footer(slide, FOOTER_TEXT, 43)
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
    _draw_footer(slide, FOOTER_TEXT, 44)
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
    _draw_footer(slide, FOOTER_TEXT, 45)
    return slide


def v46_multivariate(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Multivariate Regression")
    _add_math_equation(
        slide, Inches(3.15), Inches(1.75), Inches(7.0), Inches(1.15),
        _omml_run('Q') + _omml_text(' = ') + _omml_run('a')
        + _omml_text(' + ') + _omml_run('b') + _omml_text(' ∙ ')
        + _omml_run('P') + _omml_text(' + ') + _omml_run('c')
        + _omml_text(' ∙ ') + _omml_sub(_omml_run('P'), _omml_run('c'))
        + _omml_text(' + ') + _omml_run('d') + _omml_text(' ∙ ')
        + _omml_run('Y'),
        size_pt=30, color=NAVY, fill=CREAM, line=NAVY, rounded=True,
        shadow=True)
    _add_hierarchical_bullets(
        slide, Inches(3.15), Inches(3.35), Inches(8.0), Inches(2.8),
        [
            ("where:", 0, {'bullet_style': 'none'}),
            ([("Q", {'italic': True}), (" = amount purchased", {})], 0,
             {'bullet_style': 'none'}),
            ([("P", {'italic': True}), (" = own price", {})], 0,
             {'bullet_style': 'none'}),
            ([("Pᴄ", {'italic': True}), (" = competitor's price", {})],
             0, {'bullet_style': 'none'}),
            ([("Y", {'italic': True}), (" = income", {})], 0,
             {'bullet_style': 'none'}),
        ],
        size=24, line_spacing_pts=12)
    _vnote(slide, 32)
    _draw_footer(slide, FOOTER_TEXT, 46)
    return slide


def v47_added_vars(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Results With Added Variables")
    rows = [["coefficient", "value"],
            ["Constant (intercept)", "28.8"],
            ["Own price", "−2.12"],
            ["Competitor’s price", "1.03"],
            ["Income", "3.09"]]
    _add_styled_table(slide, Inches(2.05), Inches(1.70), Inches(6.4),
                      Inches(2.9), rows, font_size=18, header_size=18)
    _add_mixed_textbox(
        slide, Inches(1.55), Inches(5.35), Inches(11.0), Inches(0.8),
        [("text", "Demand Equation:   ",
          {'size': 24, 'bold': True, 'color': RED}),
         ("omml", _omml_run('Q') + _omml_text(' = 28.8 − 2.12')
          + _omml_run('P') + _omml_text(' + 1.03')
          + _omml_sub(_omml_run('P'), _omml_run('c'))
          + _omml_text(' + 3.09') + _omml_run('Y'),
          {'size': 24, 'color': RED})])
    _vnote(slide, 33)
    _draw_footer(slide, FOOTER_TEXT, 47)
    return slide


def v48_application(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Application")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.85), Inches(12.4),
        Inches(1.1),
        [("The airline company has run a market experimentation and "
          "estimated the demand function:", 0)],
        size=24)
    _add_math_equation(
        slide, Inches(4.05), Inches(2.75), Inches(5.2), Inches(1.0),
        _omml_run('Q') + _omml_text(' = 479 − 1.64') + _omml_run('P'),
        size_pt=28, color=NAVY, fill=CREAM, line=NAVY, rounded=True,
        shadow=True)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(4.05), Inches(12.4),
        Inches(2.4),
        [
            ("Currently, the company is charging $140 for a seat", 0),
            ([("What is ", {}), ("Eᴅ", {'bold': True, 'italic': True}),
              (" at this price?", {})], 1),
            ([("What is ", {}), ("MR", {'bold': True, 'italic': True}),
              (" at this price?", {})], 1),
            ("Should the airline raise or lower its price?", 1),
        ],
        size=24, sub_size=22)
    _draw_footer(slide, FOOTER_TEXT, 48)
    _add_pollbreak_badge(slide)
    return slide


def v49_ed_solution(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Solution: Eᴅ at P = 140")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.95), Inches(12.4),
        Inches(2.2),
        [
            ([("Q = 479 − 1.64P", {'italic': True})], 0, {}),
            ([("Recall  ", {}),
              ("Eᴅ = (ΔQ/ΔP) ∙ (P/Q)", {'italic': True}),
              ("  and  ΔQ/ΔP  is the slope of demand  ⇒  ΔQ/ΔP = −1.64",
               {})], 0, {}),
            ([("Plug ", {}), ("P", {'italic': True}),
              (" in to get ", {}), ("Q", {'italic': True}),
              (" :   Q = 479 − 1.64 ⋅ 140 = 249", {})], 0, {}),
        ],
        size=24, line_spacing_pts=16)
    _add_math_equation(
        slide, Inches(2.85), Inches(4.45), Inches(7.6), Inches(1.1),
        _omml_sub(_omml_run('E'), _omml_run('d'))
        + _omml_text(' = −1.64 ⋅ ')
        + _omml_frac(_omml_text('140'), _omml_text('249'))
        + _omml_text(' = −0.92'),
        size_pt=26, color=RED)
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(5.85), Inches(12.4),
        Inches(0.6),
        [([("→  Demand is ", {}),
           ("inelastic", {'bold': True, 'color': CBLUE}),
           ("   (|Eᴅ| < 1)", {})], 0, {'bullet_style': 'none'})],
        size=26)
    _draw_footer(slide, FOOTER_TEXT, 49)
    return slide


def v50_mr_solution(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Solution: MR at P = 140")
    steps = [
        ("1.  Inverse demand:",
         _omml_run('P') + _omml_text(' = 292 − ')
         + _omml_frac(_omml_text('1'), _omml_text('1.64'))
         + _omml_run('Q')),
        ("2.  Total revenue:",
         _omml_text('TR') + _omml_text(' = 292') + _omml_run('Q')
         + _omml_text(' − ')
         + _omml_frac(_omml_text('1'), _omml_text('1.64'))
         + _omml_sup(_omml_run('Q'), _omml_text('2'))),
        ("3.  Derivative:",
         _omml_text('MR') + _omml_text(' = 292 − 1.22') + _omml_run('Q')),
    ]
    y = Inches(1.85)
    for label, omml in steps:
        _add_hierarchical_bullets(
            slide, MARGIN + Inches(0.35), int(y + Inches(0.18)),
            Inches(4.4), Inches(0.8),
            [(label, 0, {'bold': True, 'bullet_style': 'none'})],
            size=20)
        _add_math_equation(slide, Inches(5.15), y, Inches(6.6),
                           Inches(0.95), omml, size_pt=21, color=NAVY,
                           fill=CREAM, line=NAVY, rounded=True)
        y = int(y + Inches(1.15))
    _add_math_equation(
        slide, Inches(1.15), Inches(5.45), Inches(11.0), Inches(1.0),
        _omml_text('MR') + _omml_text(' at ') + _omml_run('P')
        + _omml_text(' = 140:  ') + _omml_run('P')
        + _omml_text(' = 140  ⇒  ') + _omml_run('Q')
        + _omml_text(' = 249  ⇒  ') + _omml_text('MR')
        + _omml_text(' = −12.3'),
        size_pt=24, color=RED)
    _draw_footer(slide, FOOTER_TEXT, 50)
    return slide


def v51_raise_price(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Solution: Raise the Price")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.90), Inches(6.2),
        Inches(2.0),
        [
            ([("MR", {'italic': True}),
              (" is negative, so you want to sell ", {}),
              ("less", {'italic': True})], 0, {}),
            ([("How do we sell less? By ", {}),
              ("increasing the price", {'bold': True})], 0, {}),
        ],
        size=26, line_spacing_pts=16)
    # generic D/MR figure — MR hits zero exactly at half the D intercept,
    # Eᴅ = −1 at the D midpoint directly above it
    fig = SimpleFig(7.35, 6.35, 5.0, 4.1, 10, 10)
    _fig_axes(slide, fig)
    _add_text(slide, Inches(fig.l - 0.5), Inches(fig.b - fig.h - 0.45),
              Inches(0.6), Inches(0.32), "P", size=18, bold=True,
              italic=True, color=NAVY, font="Calibri")
    _add_text(slide, Inches(fig.l + fig.w + 0.05), Inches(fig.b - 0.05),
              Inches(0.6), Inches(0.32), "Q", size=18, bold=True,
              italic=True, color=NAVY, font="Calibri")
    _add_arrow(slide, (fig.x(0), fig.y(9)), (fig.x(9), fig.y(0)),
               color=RED, weight_pt=2.75, head=False)
    _add_text(slide, fig.x(9.15), fig.y(0.7), Inches(0.5), Inches(0.32),
              "D", size=18, bold=True, italic=True, color=RED,
              font="Calibri")
    _add_arrow(slide, (fig.x(0), fig.y(9)), (fig.x(4.5), fig.y(0)),
               color=CBLUE, weight_pt=2.75, head=False)
    _add_text(slide, fig.x(2.1), fig.y(3.6), Inches(0.8), Inches(0.32),
              "MR", size=16, bold=True, italic=True, color=CBLUE,
              font="Calibri")
    _fig_point(slide, fig, 4.5, 4.5, fill=RED, r_in=0.055)
    _add_text(slide, fig.x(4.8), fig.y(5.3), Inches(1.5), Inches(0.32),
              "Eᴅ = −1", size=14, bold=True, color=RED, font="Calibri")
    _add_text(slide, fig.x(0.3), fig.y(9.7), Inches(1.9), Inches(0.3),
              "Elastic portion", size=13, italic=True, color=GRAY,
              font="Calibri")
    _add_text(slide, fig.x(5.4), fig.y(2.9), Inches(2.2), Inches(0.3),
              "Inelastic portion", size=13, italic=True, color=GRAY,
              font="Calibri")
    _vnote(slide, 36)
    _draw_footer(slide, FOOTER_TEXT, 51)
    return slide


def v52_transaction_issues(prs):
    return make_content_bulleted(
        prs, 52, TAG_V3,
        "Issues With Transaction Data: Why Randomization Is Key",
        [
            ("Transaction data: evolution of prices and quantities over "
             "time, across locations, stores…", 0),
            ("Can we be sure that the change in quantity purchased is "
             "due to price changes?", 0),
            ("Example: in August 2019, the airline sold 135 tickets for "
             "$200; in October, 90 tickets for $180", 1),
            ("Should we conclude that the demand curve is upward "
             "sloping?", 1),
            ("What can confound the analysis?", 1),
        ],
        size=26, sub_size=24)


def v53_coffee(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Correlation Does Not Mean Causation")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.85), Inches(12.4),
        Inches(1.5),
        [([("Newspaper headline:  ", {'bold': True}),
           ("\"Drinking Three Cups of Coffee a Day Reduces Risk of "
            "Heart Attack”  ", {'italic': True}),
           ("(causal statement)", {'size': 20})], 0,
          {'bullet_style': 'none'})],
        size=26)
    _vid_media(slide, "image16.png", left=Inches(5.35), top=Inches(3.30),
               width=Inches(2.6))
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(5.55), Inches(12.4),
        Inches(1.3),
        [([("Actual finding in study: ", {'bold': True, 'color': CBLUE}),
           ("\"Among coffee drinkers, there is a smaller risk of "
            "serious cardiovascular diseases”",
            {'italic': True, 'color': CBLUE})], 0,
          {'bullet_style': 'none'})],
        size=24)
    _draw_footer(slide, FOOTER_TEXT, 53)
    return slide


def v54_omitted(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Correlation Does Not Mean Causation")
    _add_hierarchical_bullets(
        slide, MARGIN + Inches(0.15), Inches(1.80), Inches(12.4),
        Inches(0.8),
        [([("Potential Problem: ", {'bold': True}),
           ("other factors (e.g., lifestyle) are driving the "
            "correlation", {})], 0, {'bullet_style': 'none'})],
        size=26)
    _vid_media(slide, "image18.png", left=Inches(5.15), top=Inches(2.80),
               width=Inches(3.1))
    _vid_media(slide, "image19.jpg", left=Inches(3.05), top=Inches(5.05),
               width=Inches(1.9))
    _vid_media(slide, "image16.png", left=Inches(8.25), top=Inches(5.00),
               width=Inches(2.0))
    _vnote(slide, 39)
    _draw_footer(slide, FOOTER_TEXT, 54)
    return slide


def v55_spurious(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Spurious Correlations")
    _vid_media(slide, "image20.png", left=Inches(2.65), top=Inches(1.75),
               width=Inches(8.0))
    _add_text(slide, Inches(2.65), Inches(6.55), Inches(8.0),
              Inches(0.35), "tylervigen.com/spurious-correlations",
              size=13, italic=True, color=GRAY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _draw_footer(slide, FOOTER_TEXT, 55)
    return slide


def v56_randomization(prs):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, TAG_V3)
    _draw_action_title(slide, "Why Randomization Is Key")
    _add_hierarchical_bullets(
        slide, Inches(1.55), Inches(2.55), Inches(10.3), Inches(2.6),
        [
            ([("Randomizing prices ensures that ", {}),
              ("“all else is equal”", {'bold': True}),
              (" other than the price change", {})], 0, {}),
            ([("Allows us to estimate the ", {}),
              ("causal effect", {'bold': True, 'color': CBLUE}),
              (" of price on quantity demanded  →  the demand curve!",
               {})], 0, {}),
        ],
        size=28, line_spacing_pts=20)
    _draw_footer(slide, FOOTER_TEXT, 56)
    return slide


def v57_summary(prs):
    return make_m2_outline(prs, 57, section_tag="Module 2 · Summary",
                           title="Module 2: Summary", descriptions=True)


# ==========================================================================
#  build_video() — 57-slide registry
# ==========================================================================

def build_video(out_path=None):
    prs = Presentation()
    prs.slide_width = int(SLIDE_W)
    prs.slide_height = int(SLIDE_H)
    _video_title_slide(prs, "Demand and Revenue",
                       "Module 2  ·  Video 1")                     #  1
    v02_outline(prs)                                               #  2
    v03_plot_demand(prs)                                           #  3
    v04_demand_tr(prs)                                             #  4
    v05_price_change_tr(prs)                                       #  5
    v06_elasticity_tr(prs)                                         #  6
    v07_depends_on_ed(prs)                                         #  7
    v08_gum(prs)                                                   #  8
    make_stub(prs, 9, TAG_V1, "Poll: gum demand is…", STUB_POLL)   #  9
    v10_gum_solution(prs)                                          # 10
    v11_ozempic(prs)                                               # 11
    v12_ozempic_solution(prs)                                      # 12
    v13_profits(prs)                                               # 13
    v14_four_cases(prs)                                            # 14
    v15_netflix(prs)                                               # 15
    v16_netflix_solution(prs)                                      # 16
    v17_mcdonalds(prs)                                             # 17
    v18_megamillions_revisited(prs)                                # 18
    _video_title_slide(prs, "Marginal Revenue",
                       "Module 2  ·  Video 2")                     # 19
    v20_outline(prs)                                               # 20
    v21_why_mr(prs)                                                # 21
    v22_mr_definition(prs)                                         # 22
    v23_calculus(prs)                                              # 23
    v24_three_step(prs)                                            # 24
    v25_three_step_summary(prs)                                    # 25
    v26_mr_graph(prs)                                              # 26
    v27_mr_vs_price(prs)                                           # 27
    make_stub(prs, 28, TAG_V2, "Poll: MR when Q = 10 − 0.5P",
              STUB_POLL)                                           # 28
    v29_mr_solution(prs)                                           # 29
    v30_insideout(prs)                                             # 30
    v31_insideout_solution(prs)                                    # 31
    _video_title_slide(prs, "Demand Estimation",
                       "Module 2  ·  Video 3")                     # 32
    v33_outline(prs)                                               # 33
    v34_how_estimate(prs)                                          # 34
    v35_abtest(prs)                                                # 35
    v36_amazon_exp(prs)                                            # 36
    v37_amazon_recent(prs)                                         # 37
    v38_econometrics(prs)                                          # 38
    v39_ols(prs)                                                   # 39
    v40_airline_data(prs)                                          # 40
    v41_scatter(prs)                                               # 41
    v42_least_squares(prs)                                         # 42
    v43_regression1(prs)                                           # 43
    v44_regression2(prs)                                           # 44
    v45_elasticity_from_est(prs)                                   # 45
    v46_multivariate(prs)                                          # 46
    v47_added_vars(prs)                                            # 47
    v48_application(prs)                                           # 48
    v49_ed_solution(prs)                                           # 49
    v50_mr_solution(prs)                                           # 50
    v51_raise_price(prs)                                           # 51
    v52_transaction_issues(prs)                                    # 52
    v53_coffee(prs)                                                # 53
    v54_omitted(prs)                                               # 54
    v55_spurious(prs)                                              # 55
    v56_randomization(prs)                                         # 56
    v57_summary(prs)                                               # 57

    out = Path(out_path) if out_path else OUT_DIR / VDECK
    prs.save(str(out))
    print(f"saved {out} — {len(prs.slides._sldIdLst)} slides")
    return out


if __name__ == "__main__":
    import sys as _s
    build_video(_s.argv[1] if len(_s.argv) > 1 else None)
