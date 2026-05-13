"""
Build a clean Module 3 deck from scratch, using ONLY the six template layouts
defined in `_build_template_samples.py`.

Goal: every slide in this deck uses one of six layout types (title, section
header, content bulleted, content two-column, poll, closing synthesis), all
on the Blank layout, so PowerPoint's Layout dropdown stays clean.

Build is by batches – front matter (1-6), then §1.1 Short Run (7-22), etc.

Output: `Module 3_clean.pptx`
"""

import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree as ET
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# Reuse all primitives from the template script (single source of truth).
from _build_template_samples import (
    FADED,
    FOOTER_TEXT,
    GOLD,
    GOLD_W,
    GRAY,
    MARGIN,
    MODULE_AGENDA,
    NAVY,
    RULE,
    RULE_W,
    SLIDE_H,
    SLIDE_W,
    WHITE,
    _add_bulleted_list,
    _add_rect,
    _add_text,
    _blank_slide,
    _draw_action_title,
    _set_bullet_char,
)

OUT_DIR = Path(__file__).parent


# --------------------------------------------------------------------------
# Title-case top bar – replaces the all-caps default from the template script.
# The user prefers academic-paper title case (each major word capitalized,
# small connectors like "and", "of", "for", "the" stay lowercase).
# --------------------------------------------------------------------------

def _draw_top_bar_tc(slide, section_tag):
    """Navy top bar with title-cased section tag (no uppercase forcing)."""
    bar_h = Inches(0.42)
    _add_rect(slide, 0, 0, SLIDE_W, bar_h, NAVY)
    _add_text(slide, MARGIN, 0, Inches(12), bar_h,
              section_tag, size=16, bold=True,
              color=WHITE, font="Calibri",
              anchor=MSO_ANCHOR.MIDDLE)


def _draw_footer(slide, footer_text, page_num):
    """Footer rule + gold accent + footer text/page number, with larger type
    than the template default so handout printouts remain legible."""
    _add_rect(slide, 0, Inches(7.15), SLIDE_W, Inches(0.02), RULE)
    _add_rect(slide, MARGIN, Inches(7.135), GOLD_W, Inches(0.05), GOLD)
    _add_text(slide, MARGIN, Inches(7.20), Inches(11), Inches(0.32),
              footer_text, size=12, color=GRAY)
    _add_text(slide, Inches(12.5), Inches(7.20), Inches(0.6), Inches(0.32),
              str(page_num), size=12, color=GRAY, align=PP_ALIGN.RIGHT)


# --------------------------------------------------------------------------
# Speaker-notes helper
# --------------------------------------------------------------------------

def _set_notes(slide, text):
    """Replace the slide's speaker notes with *text*."""
    notes_tf = slide.notes_slide.notes_text_frame
    notes_tf.clear()
    notes_tf.text = text


# --------------------------------------------------------------------------
# Reusable shape primitives for diagrams
# --------------------------------------------------------------------------

def _add_filled_box(slide, left, top, width, height, label, *,
                    fill=NAVY, text_color=WHITE, line=None,
                    size=18, bold=True, font="Calibri"):
    """Filled rectangle with centered text."""
    left, top, width, height = int(left), int(top), int(width), int(height)
    shp = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top, width, height,
    )
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = label
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = text_color
    return shp


def _add_outlined_box(slide, left, top, width, height, label, *,
                      line=NAVY, text_color=NAVY, fill=WHITE,
                      size=18, bold=True, line_w=1.25, font="Calibri"):
    """Outlined rectangle (white fill) with centered text."""
    left, top, width, height = int(left), int(top), int(width), int(height)
    shp = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top, width, height,
    )
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.color.rgb = line
    shp.line.width = Pt(line_w)
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = label
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = text_color
    return shp


def _add_arrow(slide, start_xy, end_xy, *, color=NAVY, weight_pt=1.5,
               head=True):
    """Draw a line/arrow from start to end (in EMU/Inches values).

    EMU coordinates MUST be integers — PowerPoint rejects decimal values
    in <a:off>/<a:ext> and refuses to open the file. Cast to int defensively.
    """
    sx, sy = int(start_xy[0]), int(start_xy[1])
    ex, ey = int(end_xy[0]), int(end_xy[1])
    line = slide.shapes.add_connector(1, sx, sy, ex, ey)  # 1 = STRAIGHT
    line.line.color.rgb = color
    line.line.width = Pt(weight_pt)
    if head:
        ln = line.line._get_or_add_ln()
        tailEnd = ET.SubElement(ln, qn('a:tailEnd'))
        tailEnd.set('type', 'triangle')
        tailEnd.set('w', 'med')
        tailEnd.set('h', 'med')
    return line


def _add_arrow_shape(slide, left, top, width, height, *,
                     direction="right", fill=GOLD, line=None):
    """Block arrow shape (the 'you-are-here' indicator).

    direction: "right" (default) or "left".
    """
    left, top, width, height = int(left), int(top), int(width), int(height)
    geom = MSO_SHAPE.LEFT_ARROW if direction == "left" else MSO_SHAPE.RIGHT_ARROW
    shp = slide.shapes.add_shape(geom, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


# --------------------------------------------------------------------------
# Layout 1 — Title slide
# --------------------------------------------------------------------------

def make_title_slide(prs):
    slide = _blank_slide(prs)
    _add_text(slide, MARGIN, Inches(2.0), RULE_W, Inches(1.3),
              "Production and Costs",
              size=60, bold=True, color=NAVY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_text(slide, MARGIN, Inches(3.35), RULE_W, Inches(0.75),
              "Module 3",
              size=40, bold=True, color=GOLD, font="Calibri",
              align=PP_ALIGN.CENTER)
    accent_w = Inches(4.0)
    accent_x = (SLIDE_W - accent_w) // 2
    _add_rect(slide, accent_x, Inches(4.4), accent_w, Inches(0.06), GOLD)
    _add_text(slide, MARGIN, Inches(4.8), RULE_W, Inches(0.55),
              "Management 405  ·  EMBA",
              size=26, bold=True, color=GRAY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_text(slide, MARGIN, Inches(5.5), RULE_W, Inches(0.5),
              "Prof. Nico Voigtlaender  ·  UCLA Anderson",
              size=22, color=GRAY, font="Calibri",
              align=PP_ALIGN.CENTER)
    _add_rect(slide, 0, Inches(7.15), SLIDE_W, Inches(0.02), RULE)
    _add_rect(slide, MARGIN, Inches(7.135), GOLD_W, Inches(0.05), GOLD)
    return slide


# --------------------------------------------------------------------------
# Layout 2 — Section Header / Agenda (parameterized)
# --------------------------------------------------------------------------

def make_section_agenda(prs, page_num, *, current_part_idx=None,
                        section_tag="Module 3 · Agenda",
                        title="Agenda"):
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, section_tag)
    _draw_action_title(slide, title)

    y = Inches(1.85)
    part_title_h = Inches(0.6)
    sub_list_h = Inches(1.35)
    block_gap = Inches(0.15)

    for idx, part in enumerate(MODULE_AGENDA):
        if current_part_idx is None:
            color = NAVY  # preview: highlight all Parts
        else:
            color = NAVY if idx == current_part_idx else FADED

        _add_text(slide, MARGIN, y, RULE_W, part_title_h,
                  part["title"], size=30, bold=True, color=color,
                  font="Calibri")
        y += part_title_h

        _add_bulleted_list(
            slide,
            left=MARGIN + Inches(0.4),
            top=y,
            width=RULE_W - Inches(0.4),
            height=sub_list_h,
            items=part["subs"],
            size=24, color=color, bullet_color=color,
            line_spacing_pts=10,
            autonum_scheme='alphaLcPeriod',
        )
        y += sub_list_h + block_gap

    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


# --------------------------------------------------------------------------
# Layout 3 — Content bulleted
# --------------------------------------------------------------------------

def make_content_bulleted(prs, page_num, section_tag, title, bullets, *,
                          size=24, sub_size=None, line_spacing_pts=18,
                          extras=None):
    """bullets: list of (text, level) tuples OR plain strings (level=0)."""
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, section_tag)
    _draw_action_title(slide, title)

    normalized = [(b, 0) if isinstance(b, str) else b for b in bullets]

    _add_hierarchical_bullets(
        slide,
        left=MARGIN,
        top=Inches(1.85),
        width=RULE_W,
        height=Inches(5.0),
        items=normalized,
        size=size,
        sub_size=sub_size,
        line_spacing_pts=line_spacing_pts,
    )

    if extras is not None:
        extras(slide)

    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def _add_hierarchical_bullets(slide, left, top, width, height, items,
                              *, size=24, sub_size=None, line_spacing_pts=18):
    """Render bullets with indent levels (0 = top-level navy ▪,
    1+ = sub-bullets smaller and grey with – marker).

    size: font size for level-0 bullets.
    sub_size: font size for level-1+ bullets (default: size - 4).
    """
    if sub_size is None:
        sub_size = size - 4

    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    for i, (text, level) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        if i > 0:
            pPr = p._p.get_or_add_pPr()
            sp = line_spacing_pts if level == 0 else max(6, line_spacing_pts - 8)
            spcBef = ET.SubElement(pPr, qn('a:spcBef'))
            pts = ET.SubElement(spcBef, qn('a:spcPts'))
            pts.set('val', str(sp * 100))

        run = p.add_run()
        run.text = text
        run.font.name = 'Calibri'
        if level == 0:
            run.font.size = Pt(size)
            run.font.bold = False
            run.font.color.rgb = NAVY
            _set_bullet_char(p, char='▪', color=NAVY,
                              mar_l=342900, indent=-342900, size_pct=100)
        else:
            run.font.size = Pt(sub_size)
            run.font.color.rgb = GRAY
            mar = 342900 + level * 342900
            _set_bullet_char(p, char='–', color=GRAY,
                              mar_l=mar, indent=-228600, size_pct=100)
    return box


# --------------------------------------------------------------------------
# Layout-3 variant — diagram canvas (action title + free-form shapes below).
# Used for the agenda flowchart and the Big Picture diagram so they live
# inside the same visual chrome as content slides but render a diagram
# instead of bullets.
# --------------------------------------------------------------------------

def make_diagram_slide(prs, page_num, section_tag, title, draw_diagram):
    """Action-title slide with free-form diagram region below.

    draw_diagram: callable(slide) that renders the diagram in the body
    region (approximately y = 1.85 to 6.95).
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, section_tag)
    _draw_action_title(slide, title)
    draw_diagram(slide)
    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


# --------------------------------------------------------------------------
# Layout 5 — Poll slide (A./B./C./D. options + POLL pill, no QR box)
# --------------------------------------------------------------------------

def _draw_poll_pill(slide):
    """Small navy 'POLL' pill, top-right, sitting just below the top bar."""
    pill_w = Inches(1.05)
    pill_h = Inches(0.42)
    pill_x = SLIDE_W - MARGIN - pill_w
    pill_y = Inches(0.55)

    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  pill_x, pill_y, pill_w, pill_h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = NAVY
    shp.line.fill.background()
    shp.shadow.inherit = False
    try:
        shp.adjustments[0] = 0.5
    except Exception:
        pass
    _add_text(slide, pill_x, pill_y, pill_w, pill_h,
              "POLL", size=14, bold=True, color=WHITE, font="Calibri",
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    dot_size = Inches(0.16)
    dot_x = pill_x - Inches(0.16) - dot_size
    dot_y = pill_y + (pill_h - dot_size) // 2
    _add_rect(slide, dot_x, dot_y, dot_size, dot_size, GOLD)


def make_poll_slide(prs, page_num, section_tag, title, options, *,
                    instructions="Respond at PollEv.com/nvoigtlaender",
                    size=30, line_spacing_pts=22):
    """Layout 5 – Poll slide.

    options: list of strings (A./B./C./D. auto-numbering applied).
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, section_tag)
    _draw_action_title(slide, title)
    _draw_poll_pill(slide)

    _add_bulleted_list(
        slide,
        left=MARGIN,
        top=Inches(2.0),
        width=RULE_W,
        height=Inches(4.4),
        items=options,
        size=size, color=NAVY, bullet_color=NAVY,
        line_spacing_pts=line_spacing_pts,
        autonum_scheme='alphaUcPeriod',
    )

    _add_text(slide, MARGIN, Inches(6.55), RULE_W, Inches(0.4),
              instructions, size=16, italic=True, color=GRAY,
              align=PP_ALIGN.RIGHT)

    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


# --------------------------------------------------------------------------
# Image helpers — embed source-deck images into the new deck.
# Images are pre-extracted to _source_images/slide{N}_{rId}.{ext}.
# --------------------------------------------------------------------------

SRC_IMG_DIR = Path(__file__).parent / "_source_images"


def _add_source_image(slide, src_slide_no, rid, *, left, top, width=None,
                      height=None):
    """Place a source-deck image on the new slide."""
    candidates = list(SRC_IMG_DIR.glob(f"slide{src_slide_no}_{rid}.*"))
    if not candidates:
        return None
    img = candidates[0]
    kwargs = {"left": int(left), "top": int(top)}
    if width is not None:
        kwargs["width"] = int(width)
    if height is not None:
        kwargs["height"] = int(height)
    return slide.shapes.add_picture(str(img), **kwargs)


# --------------------------------------------------------------------------
# Illustrative callout boxes — the "major-concept" badges, "Teaching Note"
# bars, and bottom-of-slide takeaway bands that recur throughout the source
# deck.  Replicating them faithfully (in template colors) is what makes the
# slides feel like the original, not a stripped-down rewrite.
# --------------------------------------------------------------------------

def _add_takeaway_bar(slide, text, *, top=Inches(6.4), width=None,
                       height=Inches(0.55), fill=GOLD, text_color=WHITE,
                       size=20, font="Calibri", bold=True):
    """Bottom-of-slide takeaway band — the 'major concept' callout."""
    if width is None:
        width = Inches(9.6)
    left = (SLIDE_W - width) // 2
    return _add_filled_box(slide, left, top, width, height, text,
                            fill=fill, text_color=text_color,
                            size=size, bold=bold, font=font)


def _add_teaching_note(slide, text, *, top=Inches(6.6), width=None,
                        height=Inches(0.6)):
    """External-document reference card.

    Visually distinct from in-slide callouts: cream/parchment fill, navy
    DASHED border, italic navy text, with a small page-icon glyph on the
    left — clearly signals 'this links to an external Teaching Note doc'.
    """
    if width is None:
        width = Inches(8.0)
    left = int((SLIDE_W - width) // 2)
    top, width, height = int(top), int(width), int(height)

    # Card: cream fill, dashed navy border
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(0xF4, 0xF1, 0xEA)  # parchment cream
    card.line.color.rgb = NAVY
    card.line.width = Pt(1.25)
    card.shadow.inherit = False
    # Dashed line style
    ln = card.line._get_or_add_ln()
    # Remove any existing prstDash
    for el in ln.findall(qn('a:prstDash')):
        ln.remove(el)
    prstDash = ET.SubElement(ln, qn('a:prstDash'))
    prstDash.set('val', 'dash')

    # Empty text on the card; we add the icon + text as separate textboxes
    tf = card.text_frame
    tf.text = ""

    # Small "page" icon on the left – a small navy rectangle with a folded
    # corner, made from a folded-corner shape if available, else a styled
    # mini-rectangle with a triangle.
    icon_size = Inches(0.4)
    icon_x = left + Inches(0.2)
    icon_y = top + (height - icon_size) // 2
    page = slide.shapes.add_shape(MSO_SHAPE.FOLDED_CORNER,
                                   int(icon_x), int(icon_y),
                                   int(icon_size), int(icon_size))
    page.fill.solid()
    page.fill.fore_color.rgb = NAVY
    page.line.fill.background()
    page.shadow.inherit = False

    # Label text (italic, navy) inside the card to the right of the icon
    txt_left = int(icon_x + icon_size + Inches(0.2))
    txt_w = int(width - (icon_x + icon_size + Inches(0.4) - left))
    label_box = slide.shapes.add_textbox(txt_left, top,
                                          txt_w, height)
    label_tf = label_box.text_frame
    label_tf.word_wrap = True
    label_tf.margin_left = 0
    label_tf.margin_right = 0
    label_tf.margin_top = 0
    label_tf.margin_bottom = 0
    label_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = label_tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    # First run: "See teaching note:" in smaller all-caps navy
    r1 = p.add_run()
    r1.text = "SEE TEACHING NOTE  →  "
    r1.font.name = "Calibri"
    r1.font.size = Pt(12)
    r1.font.bold = True
    r1.font.color.rgb = GOLD
    # Second run: the title of the note, italic navy
    r2 = p.add_run()
    r2.text = text
    r2.font.name = "Calibri"
    r2.font.size = Pt(16)
    r2.font.italic = True
    r2.font.bold = True
    r2.font.color.rgb = NAVY
    return card


def _add_discussion_break(slide, *, top=Inches(6.6), width=Inches(4.8),
                           text="Discussion break"):
    """The slanted parallelogram 'discussion break' badge (bottom-right)."""
    left = SLIDE_W - MARGIN - width
    height = Inches(0.65)
    left, top, width, height = int(left), int(top), int(width), int(height)
    shp = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = GOLD
    shp.line.fill.background()
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.name = "Calibri"
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = WHITE
    return shp


def _add_callout_box(slide, left, top, width, height, text, *,
                      fill=GOLD, text_color=WHITE, size=14, bold=True):
    """Small free-form annotation/callout (e.g., 'plot the slope', 'Revenue
    per car net of material cost').  Used to mark a graph or sub-region."""
    return _add_filled_box(slide, left, top, width, height, text,
                            fill=fill, text_color=text_color,
                            size=size, bold=bold, font="Calibri")


def _add_anchor_burst(slide, left, top, width, height,
                       top_text, bottom_text=None,
                       *, fill=GOLD, text_color=NAVY,
                       top_size=14, bottom_size=10):
    """12-point star background + a separate text-box overlay.

    The star is purely decorative; the text lives in a normal
    rectangular text box layered on top so the text isn't constrained
    by the star's geometry (no risk of bleeding into the points).
    Reusable on every slide where MB = MC is being invoked, so the same
    visual pattern carries over.
    """
    left, top, width, height = int(left), int(top), int(width), int(height)

    # 1. Decorative star background (no text)
    shp = slide.shapes.add_shape(MSO_SHAPE.STAR_12_POINT, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.color.rgb = NAVY
    shp.line.width = Pt(1.0)
    shp.shadow.inherit = False
    # Suppress any auto-inserted text frame contents on the shape itself.
    shp.text_frame.text = ""

    # 2. Overlay text box, sized to the star's inscribed body
    inner_w = int(width * 0.65)
    inner_h = int(height * 0.55)
    inner_x = left + (width - inner_w) // 2
    inner_y = top + (height - inner_h) // 2

    box = slide.shapes.add_textbox(inner_x, inner_y, inner_w, inner_h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = top_text
    r1.font.name = 'Calibri'
    r1.font.size = Pt(top_size)
    r1.font.bold = True
    r1.font.color.rgb = text_color
    if bottom_text:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = bottom_text
        r2.font.name = 'Calibri'
        r2.font.size = Pt(bottom_size)
        r2.font.italic = True
        r2.font.bold = True
        r2.font.color.rgb = text_color
    return shp


# --------------------------------------------------------------------------
# OMML (Office Math Markup Language) equation helper – gives formulas a
# proper TeX-style render with italic variables, stacked fractions, real
# subscripts/superscripts.  Uses Cambria Math (the standard PPT math font).
# --------------------------------------------------------------------------

M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main'


def _omml_run(text):
    """OMML run for an italic variable (default math style).

    Inside an oMath, italic style is the math default for Latin letters;
    we leave m:rPr out entirely so the Cambria Math italic comes through.
    The a:rPr applies drawing-level font sizing/coloring.
    """
    return (
        f'<m:r xmlns:m="{M_NS}">'
        f'<a:rPr xmlns:a="{A_NS}" lang="en-US" b="0" i="1">'
        f'<a:latin typeface="Cambria Math"/>'
        f'<a:ea typeface="Cambria Math"/>'
        f'</a:rPr>'
        f'<m:t>{text}</m:t>'
        f'</m:r>'
    )


def _omml_text(text):
    """Upright-style OMML run (for operators, numbers, acronyms).

    Force plain (upright) style via <m:rPr><m:sty m:val="p"/></m:rPr> – this
    is the documented way to disable the math-default italics for the
    enclosed run.
    """
    return (
        f'<m:r xmlns:m="{M_NS}">'
        f'<m:rPr><m:sty m:val="p"/></m:rPr>'
        f'<a:rPr xmlns:a="{A_NS}" lang="en-US" b="0" i="0">'
        f'<a:latin typeface="Cambria Math"/>'
        f'</a:rPr>'
        f'<m:t xml:space="preserve">{text}</m:t>'
        f'</m:r>'
    )


def _omml_sub(base, sub):
    """OMML subscript: base with subscript expression."""
    return (
        f'<m:sSub xmlns:m="{M_NS}">'
        f'<m:sSubPr><m:ctrlPr>'
        f'<a:rPr xmlns:a="{A_NS}" lang="en-US" i="1">'
        f'<a:latin typeface="Cambria Math"/></a:rPr>'
        f'</m:ctrlPr></m:sSubPr>'
        f'<m:e>{base}</m:e>'
        f'<m:sub>{sub}</m:sub>'
        f'</m:sSub>'
    )


def _omml_frac(num, den):
    """OMML stacked fraction: num / den."""
    return (
        f'<m:f xmlns:m="{M_NS}">'
        f'<m:fPr><m:ctrlPr>'
        f'<a:rPr xmlns:a="{A_NS}" lang="en-US" i="1">'
        f'<a:latin typeface="Cambria Math"/></a:rPr>'
        f'</m:ctrlPr></m:fPr>'
        f'<m:num>{num}</m:num>'
        f'<m:den>{den}</m:den>'
        f'</m:f>'
    )


def _omml_sup(base, sup):
    """OMML superscript: base^sup (e.g. Q²)."""
    return (
        f'<m:sSup xmlns:m="{M_NS}">'
        f'<m:sSupPr><m:ctrlPr>'
        f'<a:rPr xmlns:a="{A_NS}" lang="en-US" i="1">'
        f'<a:latin typeface="Cambria Math"/></a:rPr>'
        f'</m:ctrlPr></m:sSupPr>'
        f'<m:e>{base}</m:e>'
        f'<m:sup>{sup}</m:sup>'
        f'</m:sSup>'
    )


def _add_math_equation(slide, left, top, width, height, omml_content, *,
                       size_pt=32, color=NAVY, fill=None, line=None):
    """Place an OMML equation in a textbox on the slide.

    omml_content: a string built from _omml_* helpers (without the outer
    <m:oMathPara> wrapper).
    """
    left, top, width, height = int(left), int(top), int(width), int(height)
    box = slide.shapes.add_textbox(left, top, width, height)

    if fill is not None:
        box.fill.solid()
        box.fill.fore_color.rgb = fill
    if line is not None:
        box.line.color.rgb = line
        box.line.width = Pt(0.75)

    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    # Replace the default <a:p> with one that hosts a math zone (<a14:m>)
    # containing the oMathPara.  The <a14:m> wrapper is REQUIRED by
    # PowerPoint to recognise OMML inside a textbox – without it PPT just
    # shows empty boxes.
    txBody = tf._txBody
    for p in list(txBody.findall(qn('a:p'))):
        txBody.remove(p)

    sz = int(size_pt * 100)
    clr_hex = '{:02X}{:02X}{:02X}'.format(color[0], color[1], color[2])

    # Use unique namespace prefixes; lxml will resolve them when parsing.
    p_xml = (
        f'<a:p xmlns:a="{A_NS}" xmlns:m="{M_NS}" xmlns:a14="{A14_NS}">'
        f'<a:pPr algn="ctr">'
        f'<a:defRPr sz="{sz}" b="0" i="1">'
        f'<a:solidFill><a:srgbClr val="{clr_hex}"/></a:solidFill>'
        f'<a:latin typeface="Cambria Math"/>'
        f'</a:defRPr>'
        f'</a:pPr>'
        f'<a14:m>'
        f'<m:oMathPara>'
        f'<m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>'
        f'<m:oMath>{omml_content}</m:oMath>'
        f'</m:oMathPara>'
        f'</a14:m>'
        f'<a:endParaRPr lang="en-US" sz="{sz}"/>'
        f'</a:p>'
    )

    new_p = ET.fromstring(p_xml)
    txBody.append(new_p)

    # Set size and color on every OMML run's <a:rPr> so the text renders at
    # the right size.  (The defRPr above is the fallback.)
    for r in new_p.iter(qn('m:r')):
        arPr = r.find(qn('a:rPr'))
        if arPr is None:
            arPr = ET.Element(qn('a:rPr'))
            r.insert(0, arPr)
        arPr.set('sz', str(sz))
        if arPr.get('lang') is None:
            arPr.set('lang', 'en-US')
        for sf in arPr.findall(qn('a:solidFill')):
            arPr.remove(sf)
        sf = ET.SubElement(arPr, qn('a:solidFill'))
        srgb = ET.SubElement(sf, qn('a:srgbClr'))
        srgb.set('val', clr_hex)
    return box


# --------------------------------------------------------------------------
# Convenience: build typical formula structures
# --------------------------------------------------------------------------

def _formula_bang_for_buck(op=' = '):
    """MP_K / p_K   [op]   MP_L / w  with stacked fractions and subscripts."""
    mp_k = _omml_sub(_omml_run('MP'), _omml_run('K'))
    p_k  = _omml_sub(_omml_run('p'),  _omml_run('K'))
    mp_l = _omml_sub(_omml_run('MP'), _omml_run('L'))
    w    = _omml_run('w')
    frac1 = _omml_frac(mp_k, p_k)
    frac2 = _omml_frac(mp_l, w)
    return frac1 + _omml_text(op) + frac2


def _formula_optimal_inputs():
    """The combined optimum condition with both fractions."""
    return _formula_bang_for_buck()


def _formula_mp_ratio(input_name='K', price_symbol='p'):
    """MP_X / p_X  (one side of the rule)."""
    base = _omml_sub(_omml_run('MP'), _omml_run(input_name))
    if price_symbol == 'w':
        den = _omml_run('w')
    else:
        den = _omml_sub(_omml_run(price_symbol), _omml_run(input_name))
    return _omml_frac(base, den)


def _add_half_textbox(slide, left, top, width, height, items, *,
                      size=22, line_spacing_pts=14, color=NAVY):
    """A simple half-page text block with bulleted or unbulleted lines."""
    box = slide.shapes.add_textbox(int(left), int(top), int(width), int(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    tf.margin_top = Inches(0.1)
    tf.margin_bottom = Inches(0.1)
    for i, item in enumerate(items):
        if isinstance(item, tuple):
            text, is_bullet = item
        else:
            text, is_bullet = item, True
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        if i > 0:
            pPr = p._p.get_or_add_pPr()
            spcBef = ET.SubElement(pPr, qn('a:spcBef'))
            pts = ET.SubElement(spcBef, qn('a:spcPts'))
            pts.set('val', str(line_spacing_pts * 100))
        run = p.add_run()
        run.text = text
        run.font.name = "Calibri"
        run.font.size = Pt(size)
        run.font.bold = False
        run.font.color.rgb = color
        if is_bullet:
            _set_bullet_char(p, char='▪', color=NAVY,
                              mar_l=342900, indent=-342900, size_pct=100)
    return box


# --------------------------------------------------------------------------
# Slide builders
# --------------------------------------------------------------------------

SECTION_TAG_FRONT = "Module 3 · Front Matter"


def slide_1(prs):
    s = make_title_slide(prs)
    _set_notes(s, (
        "Welcome – this is Module 3, Production and Costs. Last time we "
        "wrapped up the demand side; tonight we tackle the supply side and "
        "how output depends on inputs. By the end you'll have all the "
        "pieces you need for Module 4, where we put demand and costs "
        "together to find profit-maximizing decisions."
    ))


def slide_2(prs):
    bullets = [
        "If your internet permits, keep video on, mic muted",
        ("Class discussions: raise hand", 0),
        ("Will call on you to unmute", 1),
        ("Chat on Zoom: only clarifying questions to TA", 0),
        ("TA will filter questions and raise hand if needed", 1),
        "Group discussions / exercises: breakout rooms",
    ]

    def add_zoom_logo(slide):
        logo_path = OUT_DIR / "_zoom_logo.png"
        if logo_path.exists():
            slide.shapes.add_picture(
                str(logo_path),
                Inches(11.3), Inches(0.65),
                width=Inches(1.6),
            )

    s = make_content_bulleted(
        prs,
        page_num=2,
        section_tag="Module 3 · Logistics",
        title="Zoom-Specific Logistics",
        bullets=bullets,
        size=32, sub_size=28,
        line_spacing_pts=22,
        extras=add_zoom_logo,
    )
    _set_notes(s, (
        "Quick housekeeping before we dive in: Zoom rules. Video on if your "
        "bandwidth allows, mic muted by default. Raise hand to be called on; "
        "use the TA chat for clarifying questions. Group exercises run in "
        "breakout rooms – I'll move between them. Then we move on."
    ))


def slide_3(prs):
    bullets = [
        ("The law of demand", 0),
        ("Holding everything else constant, if P falls, Q rises", 1),
        ("Elasticities", 0),
        ("Responsiveness of demand to own-price, income, competitors' price", 1),
        ("Demand and revenue", 0),
        ("Own-price elasticity drives total revenue; marginal revenue via the 3-step method", 1),
        ("Demand estimation", 0),
        ("Market experimentation and regression analysis", 1),
    ]
    s = make_content_bulleted(
        prs,
        page_num=3,
        section_tag="Module 3 · Recap",
        title="Recap of Module 2: Demand and Revenue",
        bullets=bullets,
        size=32, sub_size=24,
        line_spacing_pts=10,
    )
    _set_notes(s, (
        "A 60-second reminder of where Module 2 left us: demand curves, "
        "price elasticity, and marginal revenue. The revenue side is settled; "
        "tonight we crack open the cost side. Once both sides are in hand, "
        "profit-maximization in Module 4 falls out almost mechanically."
    ))


def slide_4(prs):
    """Course-roadmap flowchart preserving the original 4-module structure.

    Layout (mirroring the source slide, beautified in template colors):

        ┌───────────────────────────────────────────┐
        │ 1. Basic Principles and Economic Way ...  │  (faded)
        └─────────────────┬─────────────────────────┘
              ┌───────────┴──────────────┐
              ▼                          ▼
        ┌──────────────────┐    ┌──────────────────┐
        │ 2. Value & Demand│    │ 3. Supply & Cost │  (←navy, current)
        └─────────┬────────┘    └────────┬─────────┘   ←┐ "You are here"
                  └─────────┬────────────┘             │
                            ▼                          │
                  ┌───────────────────────────┐        │
                  │ 4. Markets, Pricing, ...  │  (faded)
                  └───────────────────────────┘

    The current module (3) is highlighted in navy with a gold "you are here"
    arrow; past/future modules render in faded grey.
    """

    def draw(slide):
        # Geometry — taller boxes to accommodate larger type
        box_h = Inches(0.85)
        narrow_w = Inches(4.6)
        wide_w = Inches(8.6)
        gap = Inches(0.3)

        slide_mid = SLIDE_W // 2

        # Row 1 – top module (faded)
        top_x = slide_mid - wide_w // 2
        top_y = Inches(2.0)
        _add_filled_box(slide, top_x, top_y, wide_w, box_h,
                         "1. Basic Principles and Economic Way of Thinking",
                         fill=FADED, text_color=WHITE, size=24, bold=True)

        # Row 2 – two parallel modules
        row2_y = Inches(3.65)
        left_x = slide_mid - gap // 2 - narrow_w
        right_x = slide_mid + gap // 2
        _add_filled_box(slide, left_x, row2_y, narrow_w, box_h,
                         "2. Value and Demand",
                         fill=FADED, text_color=WHITE, size=26, bold=True)
        # Current module (navy)
        _add_filled_box(slide, right_x, row2_y, narrow_w, box_h,
                         "3. Supply and Cost",
                         fill=NAVY, text_color=WHITE, size=26, bold=True)

        # Row 3 – bottom module (faded)
        bot_x = slide_mid - wide_w // 2
        bot_y = Inches(5.5)
        _add_filled_box(slide, bot_x, bot_y, wide_w, box_h,
                         "4. Markets, Pricing, and Strategy",
                         fill=FADED, text_color=WHITE, size=24, bold=True)

        # Connectors — top down to row 2 (faded grey lines)
        top_bottom_y = top_y + box_h
        row2_top_y = row2_y
        _add_arrow(slide,
                    (top_x + wide_w // 2, top_bottom_y),
                    (left_x + narrow_w // 2, row2_top_y),
                    color=FADED, weight_pt=1.5, head=True)
        _add_arrow(slide,
                    (top_x + wide_w // 2, top_bottom_y),
                    (right_x + narrow_w // 2, row2_top_y),
                    color=NAVY, weight_pt=2.0, head=True)

        # Row 2 down to row 3
        row2_bottom_y = row2_y + box_h
        row3_top_y = bot_y
        _add_arrow(slide,
                    (left_x + narrow_w // 2, row2_bottom_y),
                    (bot_x + wide_w // 2, row3_top_y),
                    color=FADED, weight_pt=1.5, head=True)
        _add_arrow(slide,
                    (right_x + narrow_w // 2, row2_bottom_y),
                    (bot_x + wide_w // 2, row3_top_y),
                    color=FADED, weight_pt=1.5, head=True)

        # "You are here" — gold LEFT-pointing arrow positioned to the right
        # of box 3, so the arrow head points INTO the box.
        arrow_w = Inches(1.0)
        arrow_h = Inches(0.55)
        arrow_left = right_x + narrow_w + Inches(0.2)
        arrow_top = row2_y + (box_h - arrow_h) // 2
        _add_arrow_shape(slide, arrow_left, arrow_top, arrow_w, arrow_h,
                          direction="left", fill=GOLD)
        # Label just to the right of the arrow
        _add_text(slide, arrow_left + arrow_w + Inches(0.1),
                   row2_y + box_h // 2 - Inches(0.2),
                   Inches(2.4), Inches(0.4),
                   "you are here", size=18, italic=True, bold=True,
                   color=GOLD, font="Calibri")

    s = make_diagram_slide(
        prs, page_num=4,
        section_tag="Module 3 · Course Roadmap",
        title="Agenda for the Class",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "This is where Module 3 fits in the 405 course as a whole. We "
        "started with the Basic Principles and the Economic Way of Thinking. "
        "Module 2 covered Value and Demand – the customer side. Tonight is "
        "Module 3, Supply and Cost – the firm side. Module 4 ties it all "
        "together in Markets, Pricing, and Strategy. So tonight you're "
        "filling in the right-hand box on this map."
    ))


def slide_5(prs):
    """'Big Picture' diagram – Production Functions + Costs (Module 3) feed
    into Output Decisions (Module 4); Demand (Module 2) also feeds in.

    Layout:

        ┌─────────────────────┐         ┌──────────────┐
        │ Production Functions│         │   Demand     │  (faded, M2)
        │  inputs → output    │         │   (Module 2) │
        └─────────────────────┘         └──────┬───────┘
                  │                            │
                  ▼                            │
        ┌─────────────────────┐                │
        │       Costs         │                │
        │ what each unit costs│                │
        └─────────────────────┘                │
                  │                            │
                  └────────────┬───────────────┘
                               ▼
                  ┌───────────────────────────┐
                  │   Output Decisions        │  (gold, M4 preview)
                  │   (Module 4: pricing &    │
                  │    profit-maximization)   │
                  └───────────────────────────┘
    """

    def draw(slide):
        # Geometry: 3-column grid, left (M3 stack), right (M2), bottom-center (M4)
        box_h = Inches(1.35)
        small_h = Inches(1.2)
        m3_w = Inches(4.8)
        m2_w = Inches(4.2)
        m4_w = Inches(7.0)

        # Module 3 left stack
        m3_x = Inches(0.5)
        prod_y = Inches(1.95)
        costs_y = Inches(3.55)
        _add_filled_box(slide, m3_x, prod_y, m3_w, box_h,
                         "Production Functions\n(inputs → output)",
                         fill=NAVY, text_color=WHITE, size=26, bold=True)
        _add_filled_box(slide, m3_x, costs_y, m3_w, box_h,
                         "Costs\n(what each unit costs)",
                         fill=NAVY, text_color=WHITE, size=26, bold=True)

        # Module 2 right (faded — already covered)
        m2_x = Inches(8.6)
        m2_y = Inches(1.95)
        _add_filled_box(slide, m2_x, m2_y, m2_w, small_h,
                         "Demand\n(Module 2)",
                         fill=FADED, text_color=WHITE, size=24, bold=True)

        # Module 4 bottom-center (gold — coming up)
        m4_x = (SLIDE_W - m4_w) // 2
        m4_y = Inches(5.5)
        _add_filled_box(slide, m4_x, m4_y, m4_w, box_h,
                         "Output Decisions\n(Module 4: pricing & profit-maximization)",
                         fill=GOLD, text_color=WHITE, size=24, bold=True)

        # Arrows
        # Production → Costs (vertical, inside the M3 stack)
        _add_arrow(slide,
                    (m3_x + m3_w // 2, prod_y + box_h),
                    (m3_x + m3_w // 2, costs_y),
                    color=NAVY, weight_pt=2.0, head=True)
        # Costs → Output Decisions (diagonal down to centre)
        _add_arrow(slide,
                    (m3_x + m3_w // 2, costs_y + box_h),
                    (int(m4_x + m4_w * 0.30), m4_y),
                    color=NAVY, weight_pt=2.0, head=True)
        # Demand → Output Decisions (diagonal down to centre-right)
        _add_arrow(slide,
                    (m2_x + m2_w // 2, m2_y + small_h),
                    (int(m4_x + m4_w * 0.70), m4_y),
                    color=FADED, weight_pt=1.5, head=True)

        # Side legend
        legend_x = Inches(0.5)
        legend_y = Inches(6.95)
        _add_text(slide, legend_x, legend_y, Inches(12.5), Inches(0.25),
                   "Tonight: build the left-hand chain   ·   "
                   "M2 (Demand) is already in hand   ·   "
                   "M4: combine them to set price and quantity",
                   size=14, italic=True, color=GRAY, font="Calibri")

    s = make_diagram_slide(
        prs, page_num=5,
        section_tag="Module 3 · Production · Big Picture",
        title="Every Executive Decision Is a Production-and-Cost Decision",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The big-picture frame for the night. Every executive decision "
        "you've ever made – pricing, hiring, capacity, sourcing, "
        "outsourcing – is at some level a production-and-cost decision. "
        "Tonight builds the left-hand chain in this diagram: Production "
        "Functions describe how inputs become output, and Costs translate "
        "those inputs into dollars. Combined with Demand from Module 2, "
        "they give us Output Decisions in Module 4 – pricing and profit-"
        "maximization. Click through the boxes one at a time to walk the "
        "students through the flow."
    ))


# --------------------------------------------------------------------------
# Batch 2 – Part 1 (Production) §1.1 Short Run
# (Old textual slide_6 outline dropped – replaced by slide_concept_map.)
# --------------------------------------------------------------------------

SECTION_TAG_P1 = "Module 3 · Production · Short Run"
SECTION_TAG_P1_DIV = "Module 3 · Agenda"


# --------------------------------------------------------------------------
# Slide 7 (NEW): Concept-map / network graph
# Visualises how the components from the Outline (slide 6) relate.
# --------------------------------------------------------------------------

def slide_concept_map(prs):
    """Two parallel trees showing how Module 3's pieces fit together.

    Production rules use OMML (Cambria Math) for proper TeX-style
    formula rendering with real subscripts and stacked fractions.

    Visual story for EMBA pacing:
      • Production function → two time horizons → two decision rules
      • Each rule IS the optimization condition in its time horizon
      • Following these rules minimizes cost for a given Q  →  THAT is
        the bridge to the Costs side, which describes those costs
      • A small gold callout flags the cleanest MB = MC example
        (short-run hiring); bang-for-the-buck is a related MB / $ rule
    """

    def _formula_box(slide, x, y, w, h, label, omml_expr,
                      *, label_size=12, formula_size=22):
        """Navy filled box with a Calibri label on top + OMML formula below."""
        _add_filled_box(slide, x, y, w, h, "", fill=NAVY)
        _add_text(slide, x, y + Inches(0.06), w, Inches(0.30),
                   label, size=label_size, bold=True, color=WHITE,
                   font="Calibri", align=PP_ALIGN.CENTER)
        _add_math_equation(
            slide, x, y + Inches(0.36), w, h - Inches(0.42),
            omml_expr, size_pt=formula_size, color=WHITE,
        )

    def draw(slide):
        # Cluster header labels (section headers for the two halves)
        _add_text(slide, Inches(0.3), Inches(1.35), Inches(6.0), Inches(0.45),
                   "PRODUCTION", size=22, bold=True, color=GRAY,
                   font="Calibri", align=PP_ALIGN.CENTER)
        _add_text(slide, Inches(7.0), Inches(1.35), Inches(6.0), Inches(0.45),
                   "COSTS", size=22, bold=True, color=GRAY,
                   font="Calibri", align=PP_ALIGN.CENTER)

        # ---- PRODUCTION cluster (left) ---------------------------------

        # N1: Production function — label on top + OMML Q = f(K, L)
        n1_x, n1_y = Inches(1.0), Inches(1.90)
        n1_w, n1_h = Inches(4.6), Inches(0.95)
        eq_pf = (
            _omml_run('Q') + _omml_text(' = ') + _omml_run('f') +
            _omml_text('(') + _omml_run('K') + _omml_text(', ') +
            _omml_run('L') + _omml_text(')')
        )
        _formula_box(slide, n1_x, n1_y, n1_w, n1_h,
                      "Production function", eq_pf,
                      label_size=16, formula_size=24)

        # N2 / N3: Short Run | Long Run
        sr_x, lr_x = Inches(0.3), Inches(3.5)
        sr_y = Inches(3.05)
        col_w, col_h = Inches(2.8), Inches(0.65)
        _add_filled_box(slide, sr_x, sr_y, col_w, col_h,
                         "Short Run\n(K fixed, L flexible)",
                         fill=NAVY, text_color=WHITE, size=14, bold=True)
        _add_filled_box(slide, lr_x, sr_y, col_w, col_h,
                         "Long Run\n(Both K and L flexible)",
                         fill=NAVY, text_color=WHITE, size=14, bold=True)

        # N4 / N5: decision rules with OMML formulas
        rule_y, rule_h = Inches(3.90), Inches(1.05)

        # Hire until  MRPL = w
        eq_hire = (
            _omml_text('MRPL') + _omml_text(' = ') + _omml_run('w')
        )
        _formula_box(slide, sr_x, rule_y, col_w, rule_h,
                      "Hire until", eq_hire,
                      label_size=12, formula_size=24)

        # Bang-for-the-buck:  MP_K / p_K  =  MP_L / w
        mp_k = _omml_sub(_omml_text('MP'), _omml_run('K'))
        p_k  = _omml_sub(_omml_run('p'),  _omml_run('K'))
        mp_l = _omml_sub(_omml_text('MP'), _omml_run('L'))
        wvar = _omml_run('w')
        eq_bfb = (
            _omml_frac(mp_k, p_k) + _omml_text(' = ') + _omml_frac(mp_l, wvar)
        )
        _formula_box(slide, lr_x, rule_y, col_w, rule_h,
                      "Bang-for-the-buck", eq_bfb,
                      label_size=12, formula_size=17)

        # Small italic labels under each rule — tying the rule to
        # optimization in its time horizon
        opt_y = rule_y + rule_h + Inches(0.05)
        _add_text(slide, sr_x, opt_y, col_w, Inches(0.28),
                   "↑  short-run optimization",
                   size=12, italic=True, color=GOLD, bold=True,
                   font="Calibri", align=PP_ALIGN.CENTER)
        _add_text(slide, lr_x, opt_y, col_w, Inches(0.28),
                   "↑  long-run optimization",
                   size=12, italic=True, color=GOLD, bold=True,
                   font="Calibri", align=PP_ALIGN.CENTER)

        # ---- COSTS cluster (right) -------------------------------------

        c1_x, c1_y = Inches(7.4), Inches(1.90)
        c1_w, c1_h = Inches(5.6), Inches(0.95)
        # Cost-types header: label on top + OMML acronyms below in
        # upright Cambria Math (matches the deck's TeX-style convention).
        # Six acronyms drawn from the cost-concepts section: total fixed,
        # total variable, average fixed, average variable, average total,
        # and marginal cost.
        eq_costs = (
            _omml_text('TFC') + _omml_text('  /  ') +
            _omml_text('TVC') + _omml_text('  /  ') +
            _omml_text('AFC') + _omml_text('  /  ') +
            _omml_text('AVC') + _omml_text('  /  ') +
            _omml_text('ATC') + _omml_text('  /  ') +
            _omml_text('MC')
        )
        _formula_box(slide, c1_x, c1_y, c1_w, c1_h,
                      "Cost types", eq_costs,
                      label_size=16, formula_size=20)

        # Three parallel children of C1 — Fixed / Marginal / Average costs.
        # Each is a key decision-relevant view of costs.
        child_y, child_h = Inches(3.05), Inches(1.90)
        gap = Inches(0.2)
        cw = (c1_w - 2 * gap) // 3
        c2_x = c1_x
        c3_x = c1_x + cw + gap
        c4_x = c1_x + 2 * (cw + gap)
        _add_filled_box(slide, c2_x, child_y, cw, child_h,
                         "Fixed costs\n\n⇒ ignore if\nthey are sunk",
                         fill=NAVY, text_color=WHITE, size=14, bold=True)
        _add_filled_box(slide, c3_x, child_y, cw, child_h,
                         "Marginal Costs\n\ncrucial for optimal\noutput decisions:\nMR = MC\n(In Module 4)",
                         fill=NAVY, text_color=WHITE, size=13, bold=True)
        _add_filled_box(slide, c4_x, child_y, cw, child_h,
                         "Average Costs\n\nmostly for\naccounting purposes",
                         fill=NAVY, text_color=WHITE, size=13, bold=True)

        # ---- Within-cluster arrows (NAVY) ------------------------------

        # Production
        _add_arrow(slide,
                    (n1_x + n1_w // 2, n1_y + n1_h),
                    (sr_x + col_w // 2, sr_y),
                    color=NAVY, weight_pt=2.0, head=True)
        _add_arrow(slide,
                    (n1_x + n1_w // 2, n1_y + n1_h),
                    (lr_x + col_w // 2, sr_y),
                    color=NAVY, weight_pt=2.0, head=True)
        _add_arrow(slide,
                    (sr_x + col_w // 2, sr_y + col_h),
                    (sr_x + col_w // 2, rule_y),
                    color=NAVY, weight_pt=2.0, head=True)
        _add_arrow(slide,
                    (lr_x + col_w // 2, sr_y + col_h),
                    (lr_x + col_w // 2, rule_y),
                    color=NAVY, weight_pt=2.0, head=True)

        # Costs – C1 fans out into the three parallel children
        for cx in (c2_x, c3_x, c4_x):
            _add_arrow(slide,
                        (c1_x + c1_w // 2, c1_y + c1_h),
                        (cx + cw // 2, child_y),
                        color=NAVY, weight_pt=2.0, head=True)

        # ---- MB = MC anchor (12-point star, distinctive shape) ---------

        # The MB = MC anchor uses a 12-point star (continuous outline) so
        # it stands out from every other rectangular box on the slide.
        # The same shape is reused on every slide where MB = MC appears.
        sun_w, sun_h = Inches(1.8), Inches(1.35)
        sun_x = sr_x + (col_w - sun_w) // 2          # centred under MRPL = w
        sun_y = Inches(5.45)
        _add_anchor_burst(
            slide, sun_x, sun_y, sun_w, sun_h,
            top_text="MB = MC",
            bottom_text="(of labor)",
            top_size=16, bottom_size=12,
        )
        # Arrow from burst UP to MRPL = w (the rule where MB = MC lives)
        _add_arrow(slide,
                    (sun_x + sun_w // 2, sun_y),
                    (sr_x + col_w // 2, rule_y + rule_h),
                    color=GOLD, weight_pt=2.0, head=True)

        # ---- Min-cost bridge: ONE inflow (long-run optimization) -------

        # Centred under Bang-for-the-buck (the long-run optimization rule
        # is the cleanest cost-minimization condition – it chooses BOTH
        # K and L for a given Q).
        bridge_x = lr_x - Inches(0.1)
        bridge_y = Inches(5.60)
        bridge_w = Inches(3.0)
        bridge_h = Inches(1.05)
        _add_outlined_box(
            slide, bridge_x, bridge_y, bridge_w, bridge_h,
            "Minimum cost\nfor any given Q",
            fill=WHITE, line=GOLD, text_color=NAVY,
            size=18, bold=True, line_w=2.0,
        )

        # Inflow arrow: Bang-for-the-buck rule  →  Bridge
        # (long-run optimization gives the minimum cost for any Q)
        _add_arrow(slide,
                    (lr_x + col_w // 2, rule_y + rule_h + Inches(0.32)),
                    (bridge_x + bridge_w // 2, bridge_y),
                    color=GOLD, weight_pt=2.5, head=True)

        # ---- Outflow: bridge → cost cluster ----------------------------

        # Arrow departs from the "Minimum cost for any Q" bridge and
        # points up-right at the cost cluster (lands inside the cluster's
        # bottom-left area, NOT into any specific cost child).
        arrow_start = (bridge_x + bridge_w - Inches(0.3),
                        bridge_y + Inches(0.15))
        arrow_end = (c1_x + Inches(0.4), child_y + child_h)
        _add_arrow(slide, arrow_start, arrow_end,
                    color=GOLD, weight_pt=2.5, head=True)

        # Compact label sitting slightly to the right of the outflow
        # arrow (offset so the arrow line is no longer covered by text).
        _add_text(slide, Inches(7.10), Inches(5.45),
                   Inches(1.70), Inches(0.80),
                   "Optimized production leads to minimum-possible costs",
                   size=14, italic=True, bold=True, color=GOLD,
                   font="Calibri", align=PP_ALIGN.CENTER)

        # ---- Second MB = MC star (anchored under Marginal Costs) -------

        # The "MR = MC" rule inside the Marginal Costs box is a specific
        # instance of MB = MC (here marginal benefit is marginal revenue,
        # over OUTPUT choice). Same star pattern as the production-side
        # anchor, keeping the visual convention consistent.
        mc_sun_w, mc_sun_h = Inches(1.8), Inches(1.35)
        mc_sun_x = c3_x + (cw - mc_sun_w) // 2
        mc_sun_y = Inches(5.45)
        _add_anchor_burst(
            slide, mc_sun_x, mc_sun_y, mc_sun_w, mc_sun_h,
            top_text="MB = MC",
            bottom_text="(of output)",
            top_size=16, bottom_size=12,
        )
        # Arrow from this star UP to the Marginal Costs box (Box 2)
        _add_arrow(slide,
                    (mc_sun_x + mc_sun_w // 2, mc_sun_y),
                    (c3_x + cw // 2, child_y + child_h),
                    color=GOLD, weight_pt=2.0, head=True)

        # ---- Scale-implication annotation under Average Costs ----------

        # A small gold-outlined box that names the long-run AC-falls-with-Q
        # condition for economies of scale.
        scale_y = Inches(5.55)
        scale_h = Inches(1.05)
        _add_outlined_box(
            slide, c4_x, scale_y, cw, scale_h,
            "If long-run AC\nfalls with Q\n⇒ Economies\nof Scale",
            fill=WHITE, line=GOLD, text_color=NAVY,
            size=12, bold=True, line_w=1.5,
        )
        # Arrow: Average Costs box DOWN to the Economies-of-Scale annotation
        _add_arrow(slide,
                    (c4_x + cw // 2, child_y + child_h),
                    (c4_x + cw // 2, scale_y),
                    color=NAVY, weight_pt=2.0, head=True)

    s = make_diagram_slide(
        prs, page_num=6,
        section_tag="Module 3 · Concept Map",
        title="How the Pieces of Module 3 Connect",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Two parallel trees – Production and Costs. On the Production "
        "side, the production function Q = f(K, L, M) branches into two "
        "time horizons (short run, long run), each with its decision "
        "rule rendered in proper math notation: MRPL = w for short-run "
        "hiring and MP_K/p_K = MP_L/w (bang-for-the-buck) for long-run "
        "input mix. The italic gold labels under each rule remind "
        "students that these rules ARE the optimization conditions in "
        "their time horizons. The gold callout flags the cleanest "
        "MB = MC instance (hiring); bang-for-the-buck is the related "
        "marginal-benefit-per-dollar rule, so the callout doesn't "
        "extend there. The gold-outlined bridge band at the bottom "
        "names the link to Costs: solving the optimization conditions "
        "minimizes cost for any given output, and the cost concepts on "
        "the right describe what those minimized costs look like. Use "
        "this slide as the roadmap to return to at section transitions."
    ))


def slide_7(prs):
    """Section divider – Part 1: Production. Layout 2 (Agenda) with Part 1
    highlighted in navy, Part 2 faded grey."""
    s = make_section_agenda(
        prs, page_num=7,
        current_part_idx=0,
        section_tag=SECTION_TAG_P1_DIV,
        title="Part 1: Production – Picking the Right Inputs",
    )
    _set_notes(s, (
        "Entering Part 1 – Production. The core question for the next "
        "40 minutes: how does output depend on inputs? We'll do short-run "
        "hiring decisions first, then long-run input choice."
    ))


def slide_8(prs):
    """Output depends on inputs: Q = f(K, L, M)."""
    def draw(slide):
        # Big equation, centred near the top of the body region
        _add_text(slide, MARGIN, Inches(2.0), RULE_W, Inches(1.0),
                   "Q = f (K, L, M)",
                   size=54, bold=True, color=NAVY, font="Calibri",
                   align=PP_ALIGN.CENTER)
        # Variable legend on the LEFT
        legend = [
            ("Q  =  Output", 0),
            ("f   =  a function of inputs:", 0),
            ("K  =  Capital  (physical: factories, machinery, software, IP)", 1),
            ("L  =  Labor", 1),
            ("M  =  Raw materials", 1),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN + Inches(0.5),
            top=Inches(3.3),
            width=Inches(8.5),
            height=Inches(2.6),
            items=legend,
            size=24, sub_size=20, line_spacing_pts=8,
        )
        # The Karl Marx book / "Das Kapital" – source slide had this image at
        # (5.98, 3.13) 1.82x1.82.  Restore it on the right of the legend.
        _add_source_image(slide, 8, "rId4",
                           left=Inches(10.3), top=Inches(3.3),
                           width=Inches(2.4))
        _add_text(slide, Inches(10.3), Inches(5.85), Inches(2.4), Inches(0.25),
                   "Marx, Das Kapital  (1867)",
                   size=12, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        # Bottom takeaway band — the "major concept" callout
        _add_takeaway_bar(slide, "Output depends on inputs",
                           top=Inches(6.4), fill=NAVY)

    s = make_diagram_slide(
        prs, page_num=8,
        section_tag=SECTION_TAG_P1,
        title="Output Depends on Inputs:  Q = f (K, L, M)",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The production function in its most basic form: output Q is a "
        "function of capital K, labor L, and materials M. Capital is more "
        "than buildings – it includes machinery, software, IP, AI systems – "
        "anything you've already paid for that keeps producing. Everything "
        "else in the module is built on this expression."
    ))


def slide_9(prs):
    """In the short run, you're stuck with your capacity."""
    bullets = [
        "In the short run, some inputs (fixed factors) cannot be increased or decreased",
        "The long run is a period long enough for all inputs to be variable",
    ]

    def draw_pictures(slide):
        # Two images side-by-side at the bottom – short run vs long run
        # (source slide had two grouped pictures; the second image of each
        # group is what we actually want).
        # slide9_rId3.png is the short-run image; slide9_rId5.jpg the long-run.
        _add_source_image(slide, 9, "rId3",
                           left=Inches(1.0), top=Inches(4.4),
                           width=Inches(5.0))
        _add_source_image(slide, 9, "rId5",
                           left=Inches(7.3), top=Inches(4.4),
                           width=Inches(5.0))
        # Captions under each image
        _add_text(slide, Inches(1.0), Inches(6.8), Inches(5.0), Inches(0.3),
                   "Short run: capacity is fixed",
                   size=14, italic=True, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        _add_text(slide, Inches(7.3), Inches(6.8), Inches(5.0), Inches(0.3),
                   "Long run: build new plant",
                   size=14, italic=True, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")

    s = make_content_bulleted(
        prs, page_num=9,
        section_tag=SECTION_TAG_P1,
        title="In the Short Run, You're Stuck with Your Capacity",
        bullets=bullets,
        size=28, sub_size=22, line_spacing_pts=14,
        extras=draw_pictures,
    )
    # Shrink the bullet box so it doesn't overlap the images
    _set_notes(s, (
        "The single most important time-scale distinction in this course. "
        "Short run = your capacity (K) is fixed; you can only adjust labor "
        "and materials. Long run = everything is variable, including the "
        "plant itself. The factory walls are literally what defines short run."
    ))


# --------------------------------------------------------------------------
# Shared production-function table used by slides 10 (table) and 11 (chart).
# Cobb-Douglas Q = 0.5 · √(K · L)  (i.e., A=0.5, α=β=0.5 – constant returns
# to scale, but BOTH MPK and MPL strictly diminishing).
#
# K and L grids match the ORIGINAL Tesla-Gigafactory table (K = 100 / 200 /
# 300 / 400 robots; L = 0 … 10,000 workers in steps of 1,000).
#
# Integer-rounded values from this exact formula give strictly diminishing
# MPL down every column AND strictly diminishing MPK along every row –
# fixing the bug the original deck had a "CORRECTION" slide about (where
# rounding had inadvertently produced increasing marginal returns).
# --------------------------------------------------------------------------

PF_A, PF_ALPHA, PF_BETA = 0.5, 0.5, 0.5
PF_K_VALS = [100, 200, 300, 400]
PF_L_VALS = [0, 1000, 2000, 3000, 4000, 5000, 6000, 7000, 8000, 9000, 10000]


def _pf_value(K, L):
    """Cobb-Douglas production function (integer-rounded cars per week)."""
    if K == 0 or L == 0:
        return 0
    return int(round(PF_A * K ** PF_ALPHA * L ** PF_BETA))


def _pf_table():
    """Full Q matrix indexed [row=L_index][col=K_index]."""
    return [[_pf_value(K, L) for K in PF_K_VALS] for L in PF_L_VALS]


def slide_10(prs):
    """Tesla's production function: output of the Gigafactory per week.

    A 2D table showing weekly output as a function of (workers, robots).
    Values come from Cobb-Douglas Q = 4 · K^0.3 · L^0.5 – chosen so that
    BOTH MPL (across L for fixed K) AND MPK (across K for fixed L) are
    STRICTLY diminishing. (The original-deck table had constant MPK across
    the first three K-steps, which technically violated diminishing
    returns to K.)
    """
    def draw(slide):
        Q = _pf_table()
        # Build display table: header row + data rows. Each cell renders
        # workforce sizes with thousands separator for legibility.
        header = [""] + [str(K) for K in PF_K_VALS]
        rows_data = [header]
        for ri, L in enumerate(PF_L_VALS):
            rows_data.append([f"{L:,}"] + [str(v) for v in Q[ri]])

        rows = len(rows_data)
        cols = len(rows_data[0])
        tbl_left = Inches(2.4)
        tbl_top = Inches(2.4)
        tbl_w = Inches(9.0)
        tbl_h = Inches(4.4)
        table_shape = slide.shapes.add_table(rows, cols, tbl_left, tbl_top,
                                              tbl_w, tbl_h)
        tbl = table_shape.table
        for r, row in enumerate(rows_data):
            for c, val in enumerate(row):
                cell = tbl.cell(r, c)
                cell.text = str(val)
                tf = cell.text_frame
                for p in tf.paragraphs:
                    p.alignment = PP_ALIGN.CENTER
                    for run in p.runs:
                        run.font.name = "Calibri"
                        run.font.size = Pt(14)
                        if r == 0 or c == 0:
                            run.font.bold = True
                            run.font.color.rgb = WHITE
                        else:
                            run.font.color.rgb = NAVY
                if r == 0 or c == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = NAVY
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = WHITE
        # Axis labels: workers (left) and robots (top)
        _add_text(slide, Inches(0.5), Inches(4.2), Inches(1.9), Inches(0.7),
                   "Number of\nworkers (L)",
                   size=18, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                   font="Calibri")
        _add_text(slide, Inches(4.5), Inches(1.85), Inches(4.0), Inches(0.4),
                   "Number of robots (K)",
                   size=18, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        # Caption below table
        _add_text(slide, MARGIN, Inches(6.9), RULE_W, Inches(0.3),
                   "Output = cars per week.  MPL falls down each column;  MPK falls along each row.",
                   size=13, italic=True, color=GRAY, font="Calibri")

    s = make_diagram_slide(
        prs, page_num=10,
        section_tag=SECTION_TAG_P1,
        title="Tesla's Production Function: Output of the Gigafactory per Week",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Tesla's Gigafactory weekly output as a function of workers (L) "
        "and robots (K). Built from Q = 4 · K^0.3 · L^0.5 – a Cobb-"
        "Douglas with decreasing returns to scale. Look down a column: "
        "MPL falls as you add labor with capital fixed. Look along a "
        "row: MPK falls as you add robots with labor fixed. Strict "
        "diminishing returns in both directions – the textbook story."
    ))


def slide_11(prs):
    """Plot total output Q vs. L for four plant-capital levels.

    Native line-with-markers chart driven by the same Cobb-Douglas table
    on slide 10. Four K series (100 / 200 / 300 / 400 robots) – matches
    the original deck's chart layout, including marker shapes per series
    and the in-plot legend in the top-left corner.
    """
    SER_COLORS = [
        RGBColor(0x2E, 0x75, 0xB6),  # blue   – K=100
        RGBColor(0xC0, 0x50, 0x4D),  # red    – K=200
        RGBColor(0x80, 0x80, 0x80),  # gray   – K=300
        RGBColor(0xE6, 0xB8, 0x00),  # gold   – K=400
    ]
    SER_MARKERS = ['circle', 'triangle', 'square', 'diamond']

    def draw(slide):
        chart_data = CategoryChartData()
        chart_data.categories = [f"{L:,}" for L in PF_L_VALS]
        for K in PF_K_VALS:
            series_vals = [_pf_value(K, L) for L in PF_L_VALS]
            chart_data.add_series(f"K = {K}", series_vals)

        chart_left = Inches(1.4)
        chart_top = Inches(1.85)
        chart_w = Inches(10.5)
        chart_h = Inches(4.8)
        chart_shape = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE,  # markers added per series below
            chart_left, chart_top, chart_w, chart_h,
            chart_data,
        )
        chart = chart_shape.chart

        # Title: "Output per Week"
        chart.has_title = True
        chart.chart_title.text_frame.text = "Output per Week"
        title_run = chart.chart_title.text_frame.paragraphs[0].runs[0]
        title_run.font.name = "Calibri"
        title_run.font.size = Pt(20)
        title_run.font.bold = True
        title_run.font.color.rgb = NAVY

        # Native legend, positioned inside the plot area (top-left).
        chart.has_legend = True
        chart.legend.font.name = "Calibri"
        chart.legend.font.size = Pt(12)
        chart.legend.font.color.rgb = NAVY
        chart.legend.include_in_layout = False
        # Force legend to a manual layout inside the plot area (top-left).
        # python-pptx doesn't expose this, so write the layout XML directly.
        leg_el = chart.legend._element
        for old in leg_el.findall(qn('c:layout')):
            leg_el.remove(old)
        for old in leg_el.findall(qn('c:legendPos')):
            leg_el.remove(old)
        # Insert legendPos = 'tr' (anything works, layout below overrides)
        pos_el = ET.SubElement(leg_el, qn('c:legendPos'))
        pos_el.set('val', 'tr')
        leg_el.remove(pos_el)
        leg_el.insert(0, pos_el)
        # Insert <c:layout><c:manualLayout>… positioning legend in upper-left.
        layout = ET.SubElement(leg_el, qn('c:layout'))
        ml = ET.SubElement(layout, qn('c:manualLayout'))
        xMode = ET.SubElement(ml, qn('c:xMode')); xMode.set('val', 'edge')
        yMode = ET.SubElement(ml, qn('c:yMode')); yMode.set('val', 'edge')
        x_el = ET.SubElement(ml, qn('c:x')); x_el.set('val', '0.08')
        y_el = ET.SubElement(ml, qn('c:y')); y_el.set('val', '0.18')
        w_el = ET.SubElement(ml, qn('c:w')); w_el.set('val', '0.13')
        h_el = ET.SubElement(ml, qn('c:h')); h_el.set('val', '0.32')
        # Re-order: legendPos must precede layout (already done by insert(0)).
        # Move <c:layout> right after <c:legendPos>.
        leg_el.remove(layout)
        leg_el.insert(list(leg_el).index(pos_el) + 1, layout)

        # Axes
        cat = chart.category_axis
        cat.tick_labels.font.name = "Calibri"
        cat.tick_labels.font.size = Pt(11)
        cat.tick_labels.font.color.rgb = NAVY
        cat.has_title = True
        cat.axis_title.text_frame.text = "Number of Workers"
        ar = cat.axis_title.text_frame.paragraphs[0].runs[0]
        ar.font.name = "Calibri"; ar.font.size = Pt(14)
        ar.font.bold = True; ar.font.color.rgb = NAVY

        val = chart.value_axis
        val.tick_labels.font.name = "Calibri"
        val.tick_labels.font.size = Pt(11)
        val.tick_labels.font.color.rgb = NAVY
        val.minimum_scale = 0
        val.maximum_scale = 1000
        val.major_unit = 100
        val.has_title = True
        val.axis_title.text_frame.text = "Cars per Week"
        ar = val.axis_title.text_frame.paragraphs[0].runs[0]
        ar.font.name = "Calibri"; ar.font.size = Pt(14)
        ar.font.bold = True; ar.font.color.rgb = NAVY

        # Style each series: distinct color + marker shape.
        for idx, series in enumerate(chart.series):
            line = series.format.line
            line.color.rgb = SER_COLORS[idx]
            line.width = Pt(2.0)
            ser_xml = series._element
            # Marker block: <c:marker><c:symbol val="…"/><c:size val="7"/>
            #               <c:spPr>(solid fill + outline)</c:spPr></c:marker>
            for old in ser_xml.findall(qn('c:marker')):
                ser_xml.remove(old)
            marker = ET.SubElement(ser_xml, qn('c:marker'))
            sym = ET.SubElement(marker, qn('c:symbol'))
            sym.set('val', SER_MARKERS[idx])
            size_el = ET.SubElement(marker, qn('c:size'))
            size_el.set('val', '7')
            spPr = ET.SubElement(marker, qn('c:spPr'))
            fill = ET.SubElement(spPr, qn('a:solidFill'))
            rgb = ET.SubElement(fill, qn('a:srgbClr'))
            r, g, b = SER_COLORS[idx][0], SER_COLORS[idx][1], SER_COLORS[idx][2]
            rgb.set('val', f'{r:02X}{g:02X}{b:02X}')
            ln = ET.SubElement(spPr, qn('a:ln'))
            ln_fill = ET.SubElement(ln, qn('a:solidFill'))
            ln_rgb = ET.SubElement(ln_fill, qn('a:srgbClr'))
            ln_rgb.set('val', f'{r:02X}{g:02X}{b:02X}')
            # Move marker block before c:smooth / after c:spPr ordering
            # (schema: order, idx, tx, spPr, marker, …)
            # python-pptx adds elements in normal sub-element order;
            # if needed, reorder explicitly:
            order = ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker',
                     'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars',
                     'c:cat', 'c:val', 'c:smooth']
            children = list(ser_xml)
            children.sort(key=lambda el: order.index(el.tag.replace(
                '{http://schemas.openxmlformats.org/drawingml/2006/chart}',
                'c:')) if el.tag.replace(
                '{http://schemas.openxmlformats.org/drawingml/2006/chart}',
                'c:') in order else 999)
            for c in children:
                ser_xml.remove(c)
            for c in children:
                ser_xml.append(c)
            # Disable smoothing so curves are straight segments between points
            # (matches the original chart's piecewise-linear look).
            for sm in ser_xml.findall(qn('c:smooth')):
                ser_xml.remove(sm)
            smooth = ET.SubElement(ser_xml, qn('c:smooth'))
            smooth.set('val', '0')

        _add_takeaway_bar(
            slide,
            "Curves flatten with L  (diminishing MPL)  and bunch together with K  (diminishing MPK)",
            top=Inches(6.85), fill=GOLD, text_color=NAVY,
            width=Inches(12.5), size=16, height=Inches(0.35),
        )

    s = make_diagram_slide(
        prs, page_num=11,
        section_tag=SECTION_TAG_P1,
        title="Plotting Total Output:  Q vs. L for Four Plant Sizes",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Same numbers as the table on the previous slide, now plotted. "
        "Four lines for four plant-capital levels (100, 200, 300, 400 "
        "robots). Each line rises but flattens – diminishing MPL. The "
        "vertical spacing between adjacent lines narrows as K grows – "
        "diminishing MPK. Memorise this shape: it shows up in every "
        "production context, from factories to consulting teams."
    ))


def slide_12(prs):
    """Marginal Product of Labor – the slope of the output curve."""
    bullets = [
        "MPL = the extra output you get from one more worker",
        ("It is the slope of the total-output curve", 1),
        ("With capital K fixed, MPL falls as L rises – the curve flattens", 1),
        "Formally:  MPL = ΔQ / ΔL   (or, for tiny steps:  dQ / dL)",
        "Read the table down a column: MPL is the row-to-row difference",
    ]
    s = make_content_bulleted(
        prs, page_num=12,
        section_tag=SECTION_TAG_P1,
        title="Marginal Product of Labor – the Slope of the Output Curve",
        bullets=bullets,
        size=28, sub_size=24, line_spacing_pts=14,
    )
    _set_notes(s, (
        "MPL – marginal product of labor – is simply the extra output you "
        "get from one more worker. It's the slope of the total-output curve "
        "from the last slide. With K fixed, MPL falls as L rises: the slope "
        "flattens. Mechanically: read the production-function table down a "
        "column and look at row-to-row differences."
    ))


def slide_13(prs):
    """Hire more labor, get less per worker: diminishing MPL."""
    bullets = [
        ("Hold one input fixed (capital) and use more of the variable input (labor)", 0),
        ("Then total output increases by less and less", 1),
        ("i.e., the Marginal Product of Labor (MPL) falls", 1),
    ]

    def draw_pictures(slide):
        _add_source_image(slide, 13, "rId3",
                           left=Inches(0.5), top=Inches(3.9),
                           width=Inches(5.8))
        _add_source_image(slide, 13, "rId4",
                           left=Inches(7.0), top=Inches(3.9),
                           width=Inches(5.8))
        _add_text(slide, Inches(0.5), Inches(6.15), Inches(5.8), Inches(0.25),
                   "Total output  (rising, flattening)",
                   size=13, italic=True, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        _add_text(slide, Inches(7.0), Inches(6.15), Inches(5.8), Inches(0.25),
                   "Slope = MPL  (falling)",
                   size=13, italic=True, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        # "plot the slope" callout between the two graphs (source had this
        # animated to appear on click)
        _add_callout_box(slide,
                          left=Inches(6.2), top=Inches(4.6),
                          width=Inches(0.95), height=Inches(0.5),
                          text="plot the slope",
                          fill=GOLD, text_color=NAVY, size=11, bold=True)
        # Small connector arrow from callout to right graph
        _add_arrow(slide,
                    start_xy=(Inches(7.15), Inches(4.85)),
                    end_xy=(Inches(7.5), Inches(4.85)),
                    color=GOLD, weight_pt=1.5, head=True)
        # Bottom takeaway: "Note: MPL is the slope (dQ/dL) of the output curve"
        _add_takeaway_bar(slide,
                           "Note:  MPL is the slope  (dQ / dL)  of the output curve",
                           top=Inches(6.5), fill=NAVY,
                           width=Inches(9.5), size=18)

    s = make_content_bulleted(
        prs, page_num=13,
        section_tag=SECTION_TAG_P1,
        title="Hire More Labor, Get Less per Worker:  Diminishing MPL",
        bullets=bullets,
        size=22, sub_size=18, line_spacing_pts=6,
        extras=draw_pictures,
    )
    _set_notes(s, (
        "The headline of this section. Diminishing MPL is a near-universal "
        "feature of short-run production: each additional worker has to "
        "share the same fixed capital, so the marginal contribution shrinks. "
        "This isn't a quirk of Tesla – it's nearly always true."
    ))


def slide_14(prs):
    """The Black Death and the return to labor.

    Layout matches the source: a half-page setup textbox on the left (the
    pre-1800 economy + 1348 question), wage-and-population chart on the right.
    """
    def draw(slide):
        # Top setup (full-width row 1) – context bullets
        top_bullets = [
            ("The (agriculture-based) economy before 1800:", 0),
            ("Land was the fixed factor; labor was variable", 1),
            ("Q = f (labor, land);  no capital", 1),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=RULE_W, height=Inches(1.55),
            items=top_bullets,
            size=22, sub_size=18, line_spacing_pts=8,
        )

        # Half-page setup textbox on the LEFT (the question framing)
        left_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            int(MARGIN), int(Inches(3.6)),
            int(Inches(5.8)), int(Inches(3.0)),
        )
        left_box.fill.solid()
        left_box.fill.fore_color.rgb = RGBColor(0xF4, 0xF1, 0xEA)  # warm parchment cream
        left_box.line.color.rgb = NAVY
        left_box.line.width = Pt(1.0)
        left_box.shadow.inherit = False
        tf = left_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.25)
        tf.margin_right = Inches(0.25)
        tf.margin_top = Inches(0.2)
        tf.margin_bottom = Inches(0.2)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.LEFT
        r1 = p1.add_run()
        r1.text = "In 1348, the Black Death killed almost half the population (labor)."
        r1.font.name = "Calibri"
        r1.font.size = Pt(20)
        r1.font.color.rgb = NAVY
        # Blank line
        p_blank = tf.add_paragraph()
        # Highlighted question
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.LEFT
        r2 = p2.add_run()
        r2.text = "What happened to the return to labor?"
        r2.font.name = "Calibri"
        r2.font.size = Pt(22)
        r2.font.bold = True
        r2.font.color.rgb = GOLD

        # Wages-and-population picture on the RIGHT
        _add_source_image(slide, 14, "rId3",
                           left=Inches(6.7), top=Inches(3.5),
                           width=Inches(6.3))
        _add_text(slide, Inches(6.7), Inches(6.55), Inches(6.3), Inches(0.25),
                   "Wages and population, England 1300-1500",
                   size=13, italic=True, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")

    s = make_diagram_slide(
        prs, page_num=14,
        section_tag=SECTION_TAG_P1,
        title="The Black Death and the Return to Labor",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "A favorite historical example. The Black Death killed roughly 40% "
        "of Europe's labor force in the 14th century. Wages of survivors "
        "rose sharply – consistent with their (now higher) marginal product. "
        "Real-world proof of marginal-product reasoning, 600 years before "
        "economists named it."
    ))


def slide_15(prs):
    """Tesla hiring scenario: $90k cars, $40k materials, fixed K.

    Layout (matches source): pictures across the TOP, bullets below, gold
    takeaway bar at the bottom.
    """
    def draw(slide):
        # Pictures across the top band
        _add_source_image(slide, 15, "rId3",
                           left=Inches(0.5), top=Inches(1.85),
                           width=Inches(8.3))
        _add_source_image(slide, 15, "rId4",
                           left=Inches(9.1), top=Inches(1.85),
                           width=Inches(3.8))

        # Bullets below the pictures
        bullets = [
            ("Demand and output price are given", 0),
            ("Large number of Model S ordered at price of $90k", 1),
            ("Raw material cost per car: $40k", 1),
            ("Short run:  capital (factory size, robots) is fixed", 0),
            ("The only way to expand production is to hire more workers", 0),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(3.7),
            width=RULE_W, height=Inches(2.6),
            items=bullets,
            size=22, sub_size=18, line_spacing_pts=8,
        )

        # Gold takeaway bar at the bottom – the key question
        _add_takeaway_bar(slide,
                           "How many workers should Tesla optimally hire?",
                           top=Inches(6.45), fill=GOLD, text_color=NAVY,
                           width=Inches(10.5))

    s = make_diagram_slide(
        prs, page_num=15,
        section_tag=SECTION_TAG_P1,
        title="Tesla Hiring Scenario:  $90k Cars, $40k Materials, Fixed K",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Setup for the next eight slides. We're going to derive Tesla's "
        "optimal hiring level. Given: cars sell at $90K, materials cost "
        "$40K per car, capital is fixed. Question: how many workers should "
        "Tesla hire?"
    ))


def slide_16(prs):
    """Concept: Marginal Revenue Product of Labor (MRPL).

    Source had only a title placeholder and an animated 'Teaching Note – Hiring
    Decisions in the Short Run' callout at the bottom.  Rebuild with the
    concept content in a major-concept box and keep the Teaching Note.
    """
    def draw(slide):
        # Major-concept callout box centered upper half – "Optimal Hiring"
        # framing.  This is the illustrative box for the major concept on
        # this slide.
        _add_filled_box(slide,
                         left=Inches(2.5), top=Inches(2.1),
                         width=Inches(8.3), height=Inches(0.85),
                         label="Optimal Hiring  →  How many workers should we hire?",
                         fill=GOLD, text_color=NAVY,
                         size=24, bold=True)

        # Definition stack below
        _add_text(slide, MARGIN, Inches(3.4), RULE_W, Inches(0.5),
                   "MRPL  =  Marginal Revenue Product of Labor",
                   size=28, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        _add_text(slide, MARGIN, Inches(4.0), RULE_W, Inches(0.4),
                   "=  the extra revenue from one more worker",
                   size=22, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        # Formula highlighted
        _add_filled_box(slide,
                         left=Inches(2.5), top=Inches(4.7),
                         width=Inches(8.3), height=Inches(0.85),
                         label="MRPL  =  MPL × (Price − Materials cost per unit)",
                         fill=NAVY, text_color=WHITE,
                         size=24, bold=True)
        _add_text(slide, MARGIN, Inches(5.75), RULE_W, Inches(0.4),
                   "Plain English: a worker's dollar value to the firm",
                   size=18, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")

        # Teaching Note bar at the bottom (matches source's animated callout)
        _add_teaching_note(slide,
                            "Hiring Decisions in the Short Run",
                            top=Inches(6.5), width=Inches(8.0))

    s = make_diagram_slide(
        prs, page_num=16,
        section_tag=SECTION_TAG_P1,
        title="Concept:  Marginal Revenue Product of Labor (MRPL)",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "MRPL = Marginal Revenue Product of Labor. In plain terms: how much "
        "extra revenue does one more worker produce? It equals MPL times "
        "the price per unit, net of materials cost. This is what a worker "
        "is 'worth' to the firm in dollar terms – and it's the right "
        "benchmark to compare against the wage."
    ))


def slide_17(prs):
    """MRPL – detail."""
    def draw(slide):
        # Major-concept formula highlighted at top – proper OMML rendering.
        # MRPL = MPL × (P − MC)  (with MRPL, MPL, MC upright acronyms; P, MC
        # italic-variable styling handled by Cambria Math).
        formula = (
            _omml_text('MRPL') + _omml_text(' = ') +
            _omml_text('MPL') + _omml_text(' × ') +
            _omml_text('(') + _omml_run('P') + _omml_text(' − ') +
            _omml_text('MC') + _omml_text(')')
        )
        _add_math_equation(
            slide,
            left=Inches(1.5), top=Inches(2.0),
            width=Inches(10.3), height=Inches(1.0),
            omml_content=formula,
            size_pt=36, color=NAVY,
        )
        _add_text(slide, MARGIN, Inches(3.05), RULE_W, Inches(0.3),
                   "(net of materials cost per unit)",
                   size=16, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")

        # Detail bullets
        bullets = [
            ("When MPL falls, MRPL falls", 0),
            ("Even if price stays constant, each additional worker is worth less", 1),
            ("The economic value of a marginal hire shrinks as you scale up", 0),
            ("Implication: there is a finite optimal number of workers", 1),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(3.4),
            width=RULE_W, height=Inches(3.0),
            items=bullets,
            size=26, sub_size=22, line_spacing_pts=14,
        )

        # Bottom takeaway: the optimal hiring rule preview
        _add_takeaway_bar(slide,
                           "Optimal hiring stops where MRPL just covers the wage",
                           top=Inches(6.5), fill=GOLD, text_color=NAVY,
                           width=Inches(10.0))

    s = make_diagram_slide(
        prs, page_num=17,
        section_tag=SECTION_TAG_P1,
        title="MRPL – Detail",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Same definition, more carefully. Notice: when MPL falls, MRPL "
        "falls. So even if the price stays constant, each additional "
        "worker is worth less than the previous one. The economic value of "
        "a marginal hire shrinks as you scale up – which is why every firm "
        "has a finite optimal hiring level."
    ))


def slide_18(prs):
    """Example: MRPL at 6,000 employees and 100 robots.

    Source had: title + bullets + small image (group with picture & oval) +
    'Discussion break' parallelogram badge at bottom-right.
    """
    def draw(slide):
        bullets = [
            "Tesla currently has 100 robots and 6,000 employees in its Freemont Gigafactory",
            ("What is MRPL?  (in $ per week)", 0),
            ("Hints:", 0),
            ("Recall sales price is $90k; material cost is $40k", 1),
            ("Use marginal revenue net of material input costs", 1),
        ]
        normalized = [(b, 0) if isinstance(b, str) else b for b in bullets]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(7.8), height=Inches(4.4),
            items=normalized,
            size=22, sub_size=18, line_spacing_pts=10,
        )

        # Small annotated image on the right (the production-function table
        # excerpt source slide had as a visual reference)
        _add_source_image(slide, 18, "rId4",
                           left=Inches(8.4), top=Inches(2.7),
                           width=Inches(4.6))

        # "Discussion break" parallelogram (bottom-right)
        _add_discussion_break(slide, top=Inches(6.55), width=Inches(4.8))

    s = make_diagram_slide(
        prs, page_num=18,
        section_tag=SECTION_TAG_P1,
        title="Example:  Calculate MRPL at 6,000 Employees and 100 Robots",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "A specific number to anchor the concept. Walk through it: at 6,000 "
        "workers and 100 robots, find MPL from the table, then multiply by "
        "$50K (price minus materials)."
    ))


def slide_19(prs):
    """Poll: MRPL at 6,000 employees?

    Source slide is a full-bleed PollEv screenshot (single picture).  Keep
    that picture so the on-screen poll content matches what students see.
    """
    def draw(slide):
        # Render the source poll picture, centered in the body region.
        # Source pic: 9.58 x 7.08, aspect ~1.35:1.  Fit by height (5.1") so
        # width becomes ~6.9", centered horizontally.
        _add_source_image(slide, 19, "rId4",
                           left=Inches(3.2), top=Inches(1.85),
                           height=Inches(5.1))
        _add_text(slide, MARGIN, Inches(7.0), RULE_W, Inches(0.3),
                   "Respond at PollEv.com/nvoigtlaender",
                   size=14, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")

    s = make_diagram_slide(
        prs, page_num=19,
        section_tag=SECTION_TAG_P1,
        title="What Is Tesla's MRPL at 6,000 Employees?",
        draw_diagram=draw,
    )
    _draw_poll_pill(s)
    _set_notes(s, (
        "Quick PollEv – compute the MRPL at 6,000 employees and submit. "
        "Give them 30 seconds. The point isn't to get the number perfectly, "
        "it's to make sure everyone is doing the calculation in their head."
    ))


def slide_20(prs):
    """Solution: MRPL of Tesla."""
    bullets = [
        ("At 6,000 workers (and 100 robots), MPL ≈ 60 cars per week", 0),
        ("Per car: revenue $90k − materials $40k = $50k", 0),
        ("MRPL = MPL × $50k ≈ 60 × $50,000 ≈ $3,000,000 per week", 0),
        ("Compare to the weekly wage bill of one worker (not the whole workforce)", 1),
        ("Common slip:  forgetting to net out materials cost", 0),
    ]
    s = make_content_bulleted(
        prs, page_num=20,
        section_tag=SECTION_TAG_P1,
        title="Solution:  MRPL of Tesla",
        bullets=bullets,
        size=26, sub_size=22, line_spacing_pts=12,
    )
    _set_notes(s, (
        "Reveal the answer. Walk through anyone's confusion – the most "
        "common slip is forgetting to net out materials cost. Note the "
        "MPL number is illustrative – read it off the table for your "
        "actual classroom version."
    ))


def slide_21(prs):
    """Hire when MRPL > wage; stop when MRPL = wage."""
    bullets = [
        ("Should Tesla hire more workers?", 0),
        ("Suppose the weekly gross wage is $1,400 per worker", 1),
        ("Yes — MRPL > wage", 1),
        ("Hiring one more worker:", 0),
        ("Revenue (net of materials) rises by MRPL", 1),
        ("Wage bill rises by w", 1),
        ("Profit rises whenever MRPL > w", 1),
    ]
    s = make_content_bulleted(
        prs, page_num=21,
        section_tag=SECTION_TAG_P1,
        title="Hire When MRPL > Wage;  Stop When MRPL = Wage",
        bullets=bullets,
        size=26, sub_size=22, line_spacing_pts=10,
    )
    _set_notes(s, (
        "The hiring rule in one sentence. Hire as long as the next worker "
        "brings in more than they cost. Stop the moment the next worker "
        "just breaks even. That's it – the rest is just applying this in "
        "different settings."
    ))


def slide_22(prs):
    """The optimal hiring rule: MRPL = w.

    Layout: bullets on the left, MRPL/wage graph on the right with the
    'Revenue per car net of material cost' annotation callout pointing at
    the MRPL curve.
    """
    def draw(slide):
        # Bullets on the left
        bullets = [
            ("Compute MRPL as a function of L", 0),
            ("Use the MPL function above:  MRPL = MPL × $50k", 1),
            ("Plot the wage line  (constant at  w )", 0),
            ("Optimum: hire L*  where  MRPL = w", 0),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(2.0),
            width=Inches(6.0), height=Inches(4.0),
            items=bullets,
            size=22, sub_size=18, line_spacing_pts=14,
        )

        # MRPL / wage graph on the right
        _add_source_image(slide, 22, "rId4",
                           left=Inches(6.6), top=Inches(2.0),
                           width=Inches(6.4))

        # "Revenue per car net of material cost" callout pointing at the
        # MRPL curve (mirrors source's annotation group)
        _add_callout_box(slide,
                          left=Inches(10.5), top=Inches(2.1),
                          width=Inches(2.5), height=Inches(0.7),
                          text="Revenue per car\nnet of material cost",
                          fill=GOLD, text_color=NAVY,
                          size=11, bold=True)
        _add_arrow(slide,
                    start_xy=(Inches(11.4), Inches(2.8)),
                    end_xy=(Inches(10.6), Inches(3.4)),
                    color=GOLD, weight_pt=1.5, head=True)

        # Bottom anchor: MB = MC star next to the rule statement.
        # Same star pattern as the concept map – this rule IS the labor
        # case of MB = MC, and the visual recurs across the deck.
        star_w = Inches(1.6)
        star_h = Inches(1.05)
        star_x = MARGIN
        star_y = Inches(6.20)
        _add_anchor_burst(
            slide, star_x, star_y, star_w, star_h,
            top_text="MB = MC",
            bottom_text="(of labor)",
            top_size=14, bottom_size=11,
        )

        bar_x = star_x + star_w + Inches(0.25)
        bar_w = Inches(10.6)
        _add_filled_box(
            slide, bar_x, Inches(6.45), bar_w, Inches(0.55),
            "Optimum:  L*  where  MRPL  =  w",
            fill=NAVY, text_color=WHITE, size=20, bold=True,
        )
        _add_arrow(slide,
                    (star_x + star_w, star_y + star_h // 2),
                    (bar_x, Inches(6.45) + Inches(0.275)),
                    color=GOLD, weight_pt=2.0, head=True)

    s = make_diagram_slide(
        prs, page_num=22,
        section_tag=SECTION_TAG_P1,
        title="The Optimal Hiring Rule:  MRPL = w",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The same rule, in algebra: MRPL = w at the optimum. Burn this "
        "into your brain. Every short-run hiring problem you encounter as "
        "an executive is some version of this comparison."
    ))


# --------------------------------------------------------------------------
# Batch 3 – §1.1b Wage Searchers (23-29), §1.2 Long Run (30-41),
#           Part 2 section divider (42)
# --------------------------------------------------------------------------

SECTION_TAG_WAGE = "Module 3 · Production · Wage Searchers"
SECTION_TAG_LR   = "Module 3 · Production · Long Run"
SECTION_TAG_DIV  = "Module 3 · Agenda"


def slide_23(prs):
    """Caution: wages are not always constant."""
    bullets = [
        ("So far, we've assumed wages are constant", 0),
        ("Realistic for a small firm hiring at the market wage", 1),
        ("Reality check: for a big enough employer, hiring more workers can push the wage up", 0),
        ("Example: a frontier AI lab adding 100 senior researchers in one year", 1),
        ("Term:  the firm is a wage searcher  (not a wage taker)", 0),
    ]
    s = make_content_bulleted(
        prs, page_num=23,
        section_tag=SECTION_TAG_WAGE,
        title="Caution:  Wages Are Not Always Constant",
        bullets=bullets,
        size=26, sub_size=22, line_spacing_pts=14,
    )
    _set_notes(s, (
        "We've been assuming wages are constant. Reality check: for big "
        "enough employers, hiring more workers can push the wage up. A "
        "frontier AI lab can't just pay the market wage when it adds 100 "
        "senior researchers in one year."
    ))


def slide_24(prs):
    """Big employers bid their own wages up."""
    bullets = [
        ("Large firm  (relative to the labor market)", 0),
        ("To recruit more labor, the firm must increase the wage", 0),
        ("For wage searchers, the wage rate is upward-sloping in employment", 0),
        ("Example: a hospital hiring highly specialized surgeons", 1),
        ("Example: Anthropic, OpenAI, DeepMind hiring senior AI researchers", 1),
    ]

    def draw_extras(slide):
        _add_takeaway_bar(slide,
                           "The true marginal cost of labor includes the wage-bid-up effect",
                           top=Inches(6.5), fill=NAVY, width=Inches(10.5))

    s = make_content_bulleted(
        prs, page_num=24,
        section_tag=SECTION_TAG_WAGE,
        title="Big Employers Bid Their Own Wages Up",
        bullets=bullets,
        size=26, sub_size=22, line_spacing_pts=12,
        extras=draw_extras,
    )
    _set_notes(s, (
        "The technical term is monopsony, but you don't need the word. The "
        "intuition: as a big employer hires more, the local talent pool "
        "tightens and you pay more for everyone, not just the new hire. The "
        "'true' marginal cost of labor includes this wage-bidding-up effect."
    ))


def slide_25(prs):
    """Salary comparisons (chart picture)."""
    def draw(slide):
        # The source had a single large chart image
        _add_source_image(slide, 25, "rId5",
                           left=Inches(1.0), top=Inches(1.85),
                           width=Inches(11.3))
        _add_text(slide, MARGIN, Inches(6.9), RULE_W, Inches(0.3),
                   "Bigger firms tend to pay more  —  consistent with the wage-search story",
                   size=14, italic=True, color=GRAY, font="Calibri",
                   align=PP_ALIGN.CENTER)

    s = make_diagram_slide(
        prs, page_num=25,
        section_tag=SECTION_TAG_WAGE,
        title="Salary Comparisons across Firm Size",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Real wage data across firms of different sizes. The pattern – "
        "larger firms tend to pay more – is consistent with the wage-search "
        "story, though many other things matter too (productivity, "
        "location, benefits)."
    ))


def slide_26(prs):
    """Example: poaching an AI researcher (Anthropic + DeepMind)."""
    def draw(slide):
        bullets = [
            ("Anthropic is trying to poach a star researcher from Google DeepMind", 0),
            ("She would join Anthropic for a $5M annual salary", 1),
            ("Anthropic already employs 2 star researchers, each earning $3.5M", 0),
            ("If the new researcher is hired, the existing two will demand the same salary", 1),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(9.0), height=Inches(4.0),
            items=bullets,
            size=22, sub_size=18, line_spacing_pts=10,
        )

        # Hassabis picture (the Wikimedia / Nobel 2024 photo) on the right
        _add_source_image(slide, 26, "rId4",
                           left=Inches(9.7), top=Inches(2.0),
                           width=Inches(3.3))
        _add_text(slide, Inches(9.7), Inches(5.65), Inches(3.3), Inches(0.25),
                   "Demis Hassabis  (CC BY, C. Michel via Wikimedia)",
                   size=11, italic=True, color=GRAY, font="Calibri",
                   align=PP_ALIGN.CENTER)

        # Gold takeaway bar – the question we'll vote on
        _add_takeaway_bar(slide,
                           "What is the marginal cost of the 3rd researcher?",
                           top=Inches(6.45), fill=GOLD, text_color=NAVY,
                           width=Inches(10.0))

    s = make_diagram_slide(
        prs, page_num=26,
        section_tag=SECTION_TAG_WAGE,
        title="Example:  The Full Cost of Poaching an AI Researcher",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "A torn-from-the-headlines wage-search example. Anthropic wants to "
        "poach a star researcher from Google DeepMind. She's rare; the third "
        "hire requires bumping up the existing senior researchers too. So "
        "the third hire is way more expensive than her salary alone. The "
        "2024-2026 AI talent wars are the textbook wage-searcher story."
    ))


def slide_27(prs):
    """Poll: what is the full marginal cost of poaching the researcher?"""
    def draw(slide):
        _add_source_image(slide, 27, "rId4",
                           left=Inches(3.2), top=Inches(1.85),
                           height=Inches(5.1))
        _add_text(slide, MARGIN, Inches(7.0), RULE_W, Inches(0.3),
                   "Respond at PollEv.com/nvoigtlaender",
                   size=14, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")

    s = make_diagram_slide(
        prs, page_num=27,
        section_tag=SECTION_TAG_WAGE,
        title="What Is the Full Marginal Cost of the New Researcher?",
        draw_diagram=draw,
    )
    _draw_poll_pill(s)
    _set_notes(s, (
        "PollEv – what's the full marginal cost of the new researcher? "
        "Watch for the common trap of just reporting her $5M salary; the "
        "real answer includes the raises paid to researchers 1 and 2."
    ))


def slide_28(prs):
    """Solution: marginal cost of the 3rd researcher = $8M."""
    def draw(slide):
        # Step-by-step calculation
        steps = [
            "1.  The star researcher herself is paid  $5M",
            "2.  The two existing researchers each get a raise of  ($5M − $3.5M) = $1.5M",
            "3.  Total extra wage bill:  $5M + 2 × $1.5M",
        ]
        for i, step in enumerate(steps):
            _add_text(slide, MARGIN + Inches(0.5), Inches(2.0 + i*0.7),
                       RULE_W - Inches(1.0), Inches(0.55),
                       step, size=22, bold=False, color=NAVY,
                       font="Calibri")

        # Boxed result
        _add_filled_box(slide,
                         left=Inches(2.5), top=Inches(4.4),
                         width=Inches(8.3), height=Inches(1.1),
                         label="Marginal cost of the 3rd researcher  =  $8M",
                         fill=NAVY, text_color=WHITE,
                         size=30, bold=True)

        # Bottom takeaway
        _add_takeaway_bar(slide,
                           "Big buyers of scarce talent move the market price  —  factor it in",
                           top=Inches(6.45), fill=GOLD, text_color=NAVY,
                           width=Inches(11.0))

    s = make_diagram_slide(
        prs, page_num=28,
        section_tag=SECTION_TAG_WAGE,
        title="Solution:  Marginal Cost of the 3rd Researcher = $8M",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Reveal: marginal cost of the third researcher is $8M, not her $5M "
        "salary. The lesson: when you're a big enough buyer of scarce talent "
        "(or anything), your hiring moves the market price. Factor that in. "
        "The same logic applied at Meta when they paid up to keep AI "
        "researchers from leaving in 2024."
    ))


def slide_29(prs):
    """Are real-world wages = MRPL? (UC wage-search tool)"""
    def draw(slide):
        # Source had two pictures stacked at the same position; use the
        # second (rId5) as the final state.
        _add_source_image(slide, 29, "rId5",
                           left=Inches(2.5), top=Inches(1.85),
                           width=Inches(8.3))

        # The UC search tool URL box at the bottom (source's Rectangle 7)
        _add_filled_box(slide,
                         left=Inches(2.5), top=Inches(6.4),
                         width=Inches(8.3), height=Inches(0.55),
                         label="Search tool:  https://ucannualwage.ucop.edu/wage/",
                         fill=NAVY, text_color=WHITE,
                         size=18, bold=True)

    s = make_diagram_slide(
        prs, page_num=29,
        section_tag=SECTION_TAG_WAGE,
        title="Are Real-World Wages = MRPL?",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The empirical question that closes this section. Do real-world "
        "wages roughly equal MRPL? The UC wage-comparison tool lets you "
        "check for yourself. Spoiler: yes, broadly, but with persistent "
        "gaps that economists still argue about."
    ))


def slide_30(prs):
    """Section divider – Part 1.2: Long Run.

    Same Layout 2 / agenda view as slide 7 (Part 1 navy, Part 2 faded), but
    with an action title signalling the sub-section transition.
    """
    s = make_section_agenda(
        prs, page_num=30,
        current_part_idx=0,
        section_tag=SECTION_TAG_DIV,
        title="Part 1.2:  Long Run – Choosing the Right Input Mix",
    )
    _set_notes(s, (
        "Switching gears now from short run to long run. In the long run, "
        "capacity is no longer fixed – we get to choose K AND L from "
        "scratch. New decision: what's the right MIX of capital and labor?"
    ))


def slide_31(prs):
    """Long-run context: Rivian builds a new Georgia plant."""
    def draw(slide):
        bullets = [
            ("Context: Rivian builds its new Georgia plant", 0),
            ("Both capital and labor are flexible inputs", 0),
            ("What is the optimal input mix?", 0),
            ("E.g., robots (K) and workers (L)", 1),
            ("We will:", 0),
            ("Use marginal analysis", 1),
            ("Learn a simple rule for the optimal combination of inputs", 1),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(8.0), height=Inches(4.6),
            items=bullets,
            size=22, sub_size=18, line_spacing_pts=10,
        )

        # Rivian R1T picture
        _add_source_image(slide, 31, "rId3",
                           left=Inches(8.6), top=Inches(2.6),
                           width=Inches(4.4))
        _add_text(slide, Inches(8.6), Inches(5.55), Inches(4.4), Inches(0.25),
                   "Rivian R1T  (CC BY-SA, Wikimedia)",
                   size=12, italic=True, color=GRAY, font="Calibri",
                   align=PP_ALIGN.CENTER)

        # Bottom takeaway
        _add_takeaway_bar(slide,
                           "Long run  ⇒  pick the right K-and-L mix from scratch",
                           top=Inches(6.5), fill=NAVY, width=Inches(10.0))

    s = make_diagram_slide(
        prs, page_num=31,
        section_tag=SECTION_TAG_LR,
        title="Now:  Long Run – Rivian Builds a New Georgia Plant",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Concrete setup: Rivian building its new plant in Stanton Springs, "
        "Georgia (recently revived after the VW partnership in 2024). They "
        "get to pick everything – plant size, machinery, workforce, "
        "layout. How should they choose?"
    ))


def slide_32(prs):
    """Optimal combination of inputs (concept introduction)."""
    def draw(slide):
        # Major-concept callout box centered upper half
        _add_filled_box(slide,
                         left=Inches(2.5), top=Inches(2.1),
                         width=Inches(8.3), height=Inches(0.85),
                         label="Optimal Input Mix  →  How much K and how much L?",
                         fill=GOLD, text_color=NAVY,
                         size=24, bold=True)

        _add_text(slide, MARGIN, Inches(3.4), RULE_W, Inches(0.5),
                   "Decision rule for the long run",
                   size=24, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")

        # Lead-in to the rule – the actual stacked-fraction formula
        _add_math_equation(
            slide,
            left=Inches(2.5), top=Inches(4.1),
            width=Inches(8.3), height=Inches(1.5),
            omml_content=_formula_bang_for_buck(),
            size_pt=44, color=NAVY,
        )
        _add_text(slide, MARGIN, Inches(5.85), RULE_W, Inches(0.4),
                   "—  i.e.,  the  'bang for the buck'  rule",
                   size=20, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")

        # External-document reference (Teaching Note)
        _add_teaching_note(slide,
                            "Bang-for-the-Buck Rule",
                            top=Inches(6.5), width=Inches(7.5))

    s = make_diagram_slide(
        prs, page_num=32,
        section_tag=SECTION_TAG_LR,
        title="Optimal Combination of Inputs",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The general framework: we need a decision rule for combining "
        "inputs when both are variable. The rule will look familiar – it's "
        "the same logic as the short-run hiring rule, generalized."
    ))


def slide_33(prs):
    """The "bang for the buck" rule: equalize MP per dollar."""
    def draw(slide):
        # Headline rule – proper stacked-fraction OMML equation
        _add_math_equation(
            slide,
            left=Inches(2.5), top=Inches(1.95),
            width=Inches(8.3), height=Inches(1.6),
            omml_content=_formula_bang_for_buck(),
            size_pt=54, color=NAVY,
        )
        _add_text(slide, MARGIN, Inches(3.65), RULE_W, Inches(0.4),
                   "—  equalize  'bang for the buck'  across all inputs",
                   size=20, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")

        # Conditions / fine print as bullets
        bullets = [
            ("Simple rule for the optimal use of inputs in the long run", 0),
            ("Holds when both L and K are flexible", 1),
            ("Refers to a given output quantity Q", 1),
            ("Assumes input prices w and pₖ are constant", 1),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN + Inches(0.5), top=Inches(4.2),
            width=RULE_W - Inches(1.0), height=Inches(2.2),
            items=bullets,
            size=20, sub_size=18, line_spacing_pts=8,
        )

        # Bottom takeaway
        _add_takeaway_bar(slide,
                           "Spend each extra $ on whichever input gives the most extra output per $",
                           top=Inches(6.5), fill=GOLD, text_color=NAVY,
                           width=Inches(11.5), size=18)

    s = make_diagram_slide(
        prs, page_num=33,
        section_tag=SECTION_TAG_LR,
        title="The 'Bang for the Buck' Rule:  Equalize MP per Dollar",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The 'bang for the buck' rule. Spend each additional dollar on "
        "whichever input gives you the most extra output per dollar. At the "
        "optimum, MP per dollar is the same across all inputs: "
        "MP_K / p_K = MP_L / w."
    ))


def slide_34(prs):
    """Applying the rule – recipe for exams."""
    def draw(slide):
        # Step-by-step recipe.  Steps that reference the formula get an
        # inline OMML equation rendered next to the step text.
        def step(slide, num, y, text_left, text_right=None,
                 inline_formula=None):
            _add_text(slide, MARGIN + Inches(0.3), y,
                       Inches(0.5), Inches(0.55),
                       num, size=22, bold=True, color=GOLD,
                       font="Calibri", anchor=MSO_ANCHOR.TOP)
            _add_text(slide, MARGIN + Inches(0.9), y,
                       Inches(text_left[1]), Inches(0.6),
                       text_left[0], size=20, bold=False, color=NAVY,
                       font="Calibri")
            if inline_formula is not None:
                x, w = inline_formula['x_w']
                _add_math_equation(
                    slide,
                    left=Inches(x), top=y - Inches(0.05),
                    width=Inches(w), height=Inches(0.85),
                    omml_content=inline_formula['omml'],
                    size_pt=22, color=NAVY,
                )
            if text_right is not None:
                x, w = text_right[1]
                _add_text(slide, Inches(x), y, Inches(w), Inches(0.6),
                           text_right[0], size=20, bold=False, color=NAVY,
                           font="Calibri")

        # Step 1
        y = Inches(2.0)
        _add_text(slide, MARGIN + Inches(0.3), y,
                   Inches(0.5), Inches(0.55),
                   "1.", size=22, bold=True, color=GOLD,
                   font="Calibri", anchor=MSO_ANCHOR.TOP)
        _add_text(slide, MARGIN + Inches(0.9), y,
                   Inches(2.2), Inches(0.6),
                   "Compute", size=20, color=NAVY, font="Calibri")
        _add_math_equation(slide,
                            left=Inches(3.2), top=y - Inches(0.05),
                            width=Inches(1.5), height=Inches(0.85),
                            omml_content=_formula_mp_ratio('K', 'p'),
                            size_pt=22, color=NAVY)
        _add_text(slide, Inches(4.7), y, Inches(0.6), Inches(0.6),
                   "and", size=20, color=NAVY, font="Calibri")
        _add_math_equation(slide,
                            left=Inches(5.3), top=y - Inches(0.05),
                            width=Inches(1.3), height=Inches(0.85),
                            omml_content=_formula_mp_ratio('L', 'w'),
                            size_pt=22, color=NAVY)
        _add_text(slide, Inches(6.7), y, Inches(6.0), Inches(0.6),
                   "at the current input mix",
                   size=20, color=NAVY, font="Calibri")

        # Step 2
        y = Inches(2.85)
        _add_text(slide, MARGIN + Inches(0.3), y,
                   Inches(0.5), Inches(0.55),
                   "2.", size=22, bold=True, color=GOLD,
                   font="Calibri", anchor=MSO_ANCHOR.TOP)
        _add_text(slide, MARGIN + Inches(0.9), y,
                   Inches(11.5), Inches(0.6),
                   "If the two ratios are equal:  the mix is optimal — stop.",
                   size=20, color=NAVY, font="Calibri")

        # Step 3
        y = Inches(3.7)
        _add_text(slide, MARGIN + Inches(0.3), y,
                   Inches(0.5), Inches(0.55),
                   "3.", size=22, bold=True, color=GOLD,
                   font="Calibri", anchor=MSO_ANCHOR.TOP)
        _add_text(slide, MARGIN + Inches(0.9), y,
                   Inches(0.4), Inches(0.6),
                   "If", size=20, color=NAVY, font="Calibri")
        _add_math_equation(slide,
                            left=Inches(1.5), top=y - Inches(0.05),
                            width=Inches(1.5), height=Inches(0.85),
                            omml_content=_formula_mp_ratio('K', 'p'),
                            size_pt=22, color=NAVY)
        _add_text(slide, Inches(3.0), y, Inches(0.5), Inches(0.6),
                   " > ", size=20, bold=True, color=NAVY,
                   font="Calibri", align=PP_ALIGN.CENTER)
        _add_math_equation(slide,
                            left=Inches(3.5), top=y - Inches(0.05),
                            width=Inches(1.3), height=Inches(0.85),
                            omml_content=_formula_mp_ratio('L', 'w'),
                            size_pt=22, color=NAVY)
        _add_text(slide, Inches(4.9), y, Inches(8.0), Inches(0.6),
                   ":  shift dollars toward K  (buy more capital)",
                   size=20, color=NAVY, font="Calibri")

        # Step 4
        y = Inches(4.55)
        _add_text(slide, MARGIN + Inches(0.3), y,
                   Inches(0.5), Inches(0.55),
                   "4.", size=22, bold=True, color=GOLD,
                   font="Calibri", anchor=MSO_ANCHOR.TOP)
        _add_text(slide, MARGIN + Inches(0.9), y,
                   Inches(0.4), Inches(0.6),
                   "If", size=20, color=NAVY, font="Calibri")
        _add_math_equation(slide,
                            left=Inches(1.5), top=y - Inches(0.05),
                            width=Inches(1.5), height=Inches(0.85),
                            omml_content=_formula_mp_ratio('K', 'p'),
                            size_pt=22, color=NAVY)
        _add_text(slide, Inches(3.0), y, Inches(0.5), Inches(0.6),
                   " < ", size=20, bold=True, color=NAVY,
                   font="Calibri", align=PP_ALIGN.CENTER)
        _add_math_equation(slide,
                            left=Inches(3.5), top=y - Inches(0.05),
                            width=Inches(1.3), height=Inches(0.85),
                            omml_content=_formula_mp_ratio('L', 'w'),
                            size_pt=22, color=NAVY)
        _add_text(slide, Inches(4.9), y, Inches(8.0), Inches(0.6),
                   ":  shift dollars toward L  (hire more labor)",
                   size=20, color=NAVY, font="Calibri")

        # Step 5
        y = Inches(5.4)
        _add_text(slide, MARGIN + Inches(0.3), y,
                   Inches(0.5), Inches(0.55),
                   "5.", size=22, bold=True, color=GOLD,
                   font="Calibri", anchor=MSO_ANCHOR.TOP)
        _add_text(slide, MARGIN + Inches(0.9), y,
                   Inches(11.5), Inches(0.6),
                   "Repeat until the two ratios are equal.",
                   size=20, color=NAVY, font="Calibri")

        # External-document reference
        _add_teaching_note(slide,
                            "Recipe for Exams",
                            top=Inches(6.5), width=Inches(6.5))

    s = make_diagram_slide(
        prs, page_num=34,
        section_tag=SECTION_TAG_LR,
        title="Applying the Rule – Recipe for Exams",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Step-by-step procedure for applying the rule on the exam, and in "
        "practice. Compute MP/$ for each input; if they're not equal, "
        "shift dollars toward the higher one."
    ))


def slide_35(prs):
    """Example: Rivian's New Georgia plant."""
    def draw(slide):
        bullets = [
            ("Rivian is building a new plant in Stanton Springs, Georgia", 0),
            ("They ask for your advice on the optimal mix of robots and workers", 0),
            ("You know:", 0),
            ("Planned output:  500 vehicles per week", 1),
            ("Weekly wage for suitable workers:  w = $1,200", 1),
            ("Cost of one robot (per week):  pₖ = $20,000", 1),
            ("Current plan:  200 robots and 5,000 workers", 1),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(8.0), height=Inches(4.4),
            items=bullets,
            size=20, sub_size=18, line_spacing_pts=8,
        )

        # Rivian R1T picture on the right (the source slide had a stale
        # Tesla picture — use the proper Rivian image instead)
        rivian = OUT_DIR / "_rivian.jpg"
        if rivian.exists():
            slide.shapes.add_picture(
                str(rivian),
                int(Inches(8.7)), int(Inches(2.4)),
                width=int(Inches(4.3)),
            )
        _add_text(slide, Inches(8.7), Inches(5.6), Inches(4.3), Inches(0.25),
                   "Rivian R1T  (CC BY-SA, Wikimedia)",
                   size=12, italic=True, color=GRAY, font="Calibri",
                   align=PP_ALIGN.CENTER)

        # Takeaway: the question we'll answer
        _add_takeaway_bar(slide,
                           "Is Rivian's 200 robots / 5,000 workers plan optimal?",
                           top=Inches(6.5), fill=GOLD, text_color=NAVY,
                           width=Inches(10.5))

    s = make_diagram_slide(
        prs, page_num=35,
        section_tag=SECTION_TAG_LR,
        title="Example:  Rivian's New Georgia Plant",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Real data from Rivian's Georgia plant project. We'll apply the "
        "bang-for-the-buck rule to actual numbers and see whether the "
        "current mix is optimal."
    ))


def slide_36(prs):
    """Is Rivian's current plan optimal? (production function)."""
    def draw(slide):
        bullets = [
            ("The production function at Rivian's new plant:", 0),
            ("Current mix to produce 500 vehicles per week:  200 robots, 5,000 workers", 0),
            ("Other input mixes are possible — which is best?", 0),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(6.5), height=Inches(3.0),
            items=bullets,
            size=20, sub_size=18, line_spacing_pts=10,
        )

        # Production function picture on the right
        _add_source_image(slide, 36, "rId4",
                           left=Inches(7.5), top=Inches(1.9),
                           width=Inches(5.5))

        # Bottom takeaway
        _add_takeaway_bar(slide,
                           "Compare MP per dollar across inputs at the current mix",
                           top=Inches(6.5), fill=NAVY, width=Inches(10.0))

    s = make_diagram_slide(
        prs, page_num=36,
        section_tag=SECTION_TAG_LR,
        title="Is Rivian's Current Plan Optimal?  (Production Function)",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The production function for Georgia in numbers. Setup question: is "
        "Rivian's current K/L mix optimal? Don't answer yet – the next "
        "slide does the math."
    ))


def slide_37(prs):
    """Is Rivian's current plan optimal? (analysis)."""
    def draw(slide):
        # The analysis is fundamentally visual – show the production-function
        # table with MP_K and MP_L highlighted, plus the per-$ calculation.
        bullets = [
            ("At the current mix:  read MPₖ and MPₗ from the table", 0),
            ("Compute MPₖ / pₖ  and  MPₗ / w", 0),
            ("If unequal → shift toward the input with the higher ratio", 0),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(6.5), height=Inches(2.6),
            items=bullets,
            size=20, sub_size=18, line_spacing_pts=10,
        )

        # Analysis picture (annotated production function) on the right
        _add_source_image(slide, 37, "rId5",
                           left=Inches(7.5), top=Inches(1.9),
                           width=Inches(5.5))

        # Major-concept callout with OMML inequality
        _add_text(slide, Inches(0.5), Inches(4.9), Inches(0.6), Inches(0.4),
                   "If", size=20, color=NAVY, bold=True, font="Calibri")
        _add_math_equation(
            slide,
            left=Inches(1.1), top=Inches(4.85),
            width=Inches(5.5), height=Inches(0.9),
            omml_content=_formula_bang_for_buck(op=' > '),
            size_pt=22, color=NAVY,
        )
        _add_text(slide, Inches(0.5), Inches(5.9), Inches(6.5), Inches(0.4),
                   "→  shift toward robots",
                   size=20, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")

        # Bottom takeaway
        _add_takeaway_bar(slide,
                           "Bang-for-the-buck tells you which way to re-allocate",
                           top=Inches(6.5), fill=NAVY, width=Inches(10.0))

    s = make_diagram_slide(
        prs, page_num=37,
        section_tag=SECTION_TAG_LR,
        title="Is Rivian's Current Plan Optimal?  (Analysis)",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Apply the rule. Compute MP/$ for K and MP/$ for L. If they're not "
        "equal, the mix isn't optimal and Rivian should shift toward the "
        "input with higher bang-for-the-buck."
    ))


def slide_38(prs):
    """Poll: Is Rivian's input mix optimal?"""
    def draw(slide):
        _add_source_image(slide, 38, "rId4",
                           left=Inches(3.2), top=Inches(1.85),
                           height=Inches(5.1))
        _add_text(slide, MARGIN, Inches(7.0), RULE_W, Inches(0.3),
                   "Respond at PollEv.com/nvoigtlaender",
                   size=14, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")

    s = make_diagram_slide(
        prs, page_num=38,
        section_tag=SECTION_TAG_LR,
        title="Is Rivian's Input Mix Optimal?",
        draw_diagram=draw,
    )
    _draw_poll_pill(s)
    _set_notes(s, (
        "PollEv – vote on whether Rivian's current mix is optimal. Some "
        "will say yes, some no. Reveal in the next slide."
    ))


def slide_39(prs):
    """Solution on optimal input mix."""
    def draw(slide):
        # Source had just a title — build the analysis from speaker notes.
        # Show MP_K and MP_L estimates, then ratios, then conclusion.
        # Use a table-like layout with rows for K and L.
        rows = [
            ("",            "MP",          "Price",  "MP per $"),
            ("Robots  (K)", "≈ 4 cars",    "$20,000", "0.0002 cars / $"),
            ("Workers (L)", "≈ 0.1 cars",  "$1,200",  "0.0001 cars / $"),
        ]
        col_w = [Inches(3.0), Inches(2.5), Inches(2.5), Inches(3.0)]
        x0 = (SLIDE_W - sum(col_w)) // 2
        y0 = Inches(2.0)
        row_h = Inches(0.7)
        for r, row in enumerate(rows):
            cx = x0
            for c, val in enumerate(row):
                fill = NAVY if r == 0 else (RGBColor(0xF4, 0xF1, 0xEA) if r % 2 else WHITE)
                txt_color = WHITE if r == 0 else NAVY
                bold = r == 0 or c == 0
                shp = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    int(cx), int(y0 + r * row_h),
                    int(col_w[c]), int(row_h),
                )
                shp.fill.solid()
                shp.fill.fore_color.rgb = fill
                shp.line.color.rgb = RULE
                shp.line.width = Pt(0.5)
                shp.shadow.inherit = False
                tf = shp.text_frame
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf.margin_left = Inches(0.1)
                tf.margin_right = Inches(0.1)
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                run = p.add_run()
                run.text = val
                run.font.name = "Calibri"
                run.font.size = Pt(18)
                run.font.bold = bold
                run.font.color.rgb = txt_color
                cx += col_w[c]

        # Conclusion: stacked-fraction OMML formula + arrow + advice
        _add_math_equation(
            slide,
            left=Inches(2.5), top=Inches(4.7),
            width=Inches(8.3), height=Inches(1.1),
            omml_content=_formula_bang_for_buck(op=' > '),
            size_pt=32, color=NAVY,
            fill=RGBColor(0xF4, 0xF1, 0xEA),
            line=NAVY,
        )
        _add_text(slide, MARGIN, Inches(5.85), RULE_W, Inches(0.4),
                   "→  Rivian should use more robots, fewer workers",
                   size=20, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")

        # Bottom takeaway
        _add_takeaway_bar(slide,
                           "Equalize MP per $  →  reach the optimal mix",
                           top=Inches(6.5), fill=GOLD, text_color=NAVY,
                           width=Inches(9.5))

    s = make_diagram_slide(
        prs, page_num=39,
        section_tag=SECTION_TAG_LR,
        title="Solution:  The Optimal Input Mix",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Reveal: the mix isn't optimal. Discuss which input was underused, "
        "and what Rivian should do to fix it – hire more L or buy more "
        "robots. Numbers shown are illustrative; replace with your "
        "classroom-version values."
    ))


def slide_40(prs):
    """When prices change, the input mix shifts: robot tax & union wages.

    The two body boxes use OMML for the math (p_K, MP_K / p_K with real
    subscripts + stacked fractions) and Calibri for the descriptive
    'shift toward …' line beneath each formula.
    """
    def draw(slide):
        _add_text(slide, MARGIN, Inches(1.9), RULE_W, Inches(0.55),
                   "What's the effect on the optimal input mix if…",
                   size=24, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")

        col_w = Inches(6.0)
        col_h = Inches(3.4)
        gap = Inches(0.4)
        left_x = (SLIDE_W - 2 * col_w - gap) // 2
        right_x = left_x + col_w + gap
        y = Inches(2.7)
        cream = RGBColor(0xF4, 0xF1, 0xEA)
        body_h = col_h - Inches(0.75)

        def _column(x, header_text, eq_xml, conclusion):
            # Header band
            _add_filled_box(slide, x, y, col_w, Inches(0.7),
                             header_text,
                             fill=NAVY, text_color=WHITE,
                             size=18, bold=True)
            # Cream body background (no text)
            _add_filled_box(slide, x, y + Inches(0.75),
                             col_w, body_h, "",
                             fill=cream, text_color=NAVY,
                             size=20, bold=False)
            # OMML equation centered in the upper part of the cream body
            _add_math_equation(
                slide, x, y + Inches(0.95),
                col_w, Inches(1.45),
                eq_xml, size_pt=24, color=NAVY,
            )
            # Conclusion line beneath the formula
            _add_text(slide, x, y + Inches(2.55),
                       col_w, Inches(0.5),
                       conclusion,
                       size=20, color=NAVY, font="Calibri",
                       align=PP_ALIGN.CENTER)

        # LEFT column: tax on robots → p_K rises → MP_K / p_K falls
        eq_left = (
            _omml_sub(_omml_run('p'), _omml_run('K')) +
            _omml_text('  ↑   ⇒   ') +
            _omml_frac(
                _omml_sub(_omml_run('MP'), _omml_run('K')),
                _omml_sub(_omml_run('p'), _omml_run('K'))
            ) +
            _omml_text('  ↓')
        )
        _column(
            left_x,
            "The government introduces a high tax on robots",
            eq_left,
            "→  shift toward more labor",
        )

        # RIGHT column: union wages → w rises → MP_L / w falls
        eq_right = (
            _omml_run('w') +
            _omml_text('  ↑   ⇒   ') +
            _omml_frac(
                _omml_sub(_omml_run('MP'), _omml_run('L')),
                _omml_run('w')
            ) +
            _omml_text('  ↓')
        )
        _column(
            right_x,
            "Labor unions demand significantly higher wages",
            eq_right,
            "→  shift toward more automation",
        )

        # Bottom takeaway
        _add_takeaway_bar(slide,
                           "When input prices change, the optimal mix shifts toward the cheaper input",
                           top=Inches(6.5), fill=GOLD, text_color=NAVY,
                           width=Inches(12.0), size=18)

    s = make_diagram_slide(
        prs, page_num=40,
        section_tag=SECTION_TAG_LR,
        title="When Prices Change, the Input Mix Shifts:  Robot Tax & Union Wages",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Comparative statics. When input prices change, the optimal mix "
        "shifts: a tax on robots pushes Rivian toward more labor; rising "
        "wages push them toward more automation. Real strategic "
        "implications for any firm facing input-price shocks – including AI "
        "labs deciding between GPU spend and engineer headcount."
    ))


def slide_41(prs):
    """'Bang for the buck' in grocery shopping (intuition reinforcer)."""
    def draw(slide):
        # Left text: the intuition
        bullets = [
            ("The bang-for-the-buck rule isn't just for factories", 0),
            ("You apply it every week at the grocery store", 0),
            ("Spend each $ on the item that gives the most extra utility per $", 1),
            ("Same logic, different decision", 0),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(6.5), height=Inches(4.0),
            items=bullets,
            size=22, sub_size=18, line_spacing_pts=12,
        )

        # Two product pictures on the right – mirroring the source
        _add_source_image(slide, 41, "rId6",
                           left=Inches(7.5), top=Inches(2.2),
                           height=Inches(2.4))
        _add_source_image(slide, 41, "rId5",
                           left=Inches(10.4), top=Inches(2.2),
                           height=Inches(2.4))

        # Bottom takeaway
        _add_takeaway_bar(slide,
                           "Bang-for-the-buck:  a universal decision rule",
                           top=Inches(6.5), fill=NAVY, width=Inches(9.0))

    s = make_diagram_slide(
        prs, page_num=41,
        section_tag=SECTION_TAG_LR,
        title="'Bang for the Buck' in Grocery Shopping",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The bang-for-the-buck rule isn't just for factories. You apply it "
        "every week at the grocery store – balancing what you spend on "
        "each item against the extra utility you get. Same logic, "
        "different decision."
    ))


def slide_42(prs):
    """Section divider – Part 2: Costs."""
    s = make_section_agenda(
        prs, page_num=42,
        current_part_idx=1,        # Part 2 now active
        section_tag=SECTION_TAG_DIV,
        title="Part 2:  Costs – Producing at the Lowest Price",
    )
    _set_notes(s, (
        "Part 2 – Costs. We've covered what to PRODUCE; now we cover what "
        "it COSTS, and crucially, which costs actually matter for "
        "decisions."
    ))


# --------------------------------------------------------------------------
# §2.1 Cost Concepts (slides 43-62)
# --------------------------------------------------------------------------

SECTION_TAG_P2 = "Module 3 · Costs · Cost Concepts"
SECTION_TAG_P2_LR = "Module 3 · Costs · Long-Run & Scale"


def slide_43(prs):
    """Cost types: fixed / sunk / variable, with takeaway 'sunk costs
    should never drive decisions'."""
    def draw(slide):
        # Three vertical bands, each with a category header and definition
        bands = [
            ("Fixed Costs", "Do not depend on quantity produced (Q)", NAVY, WHITE),
            ("Sunk Costs",
             "A fixed cost that cannot be recovered\n(may be partially sunk)",
             GOLD, NAVY),
            ("Variable Costs", "Depend on volume produced (Q)", NAVY, WHITE),
        ]
        band_w = Inches(3.95)
        band_h = Inches(2.2)
        gap = Inches(0.15)
        total_w = band_w * 3 + gap * 2
        start_x = (SLIDE_W - total_w) // 2
        for i, (label, body, fill, txt) in enumerate(bands):
            bx = start_x + (band_w + gap) * i
            # Header band
            _add_filled_box(
                slide, bx, Inches(2.05), band_w, Inches(0.7), label,
                fill=fill, text_color=txt, size=22, bold=True,
            )
            # Body description
            _add_outlined_box(
                slide, bx, Inches(2.75), band_w, Inches(1.5), body,
                fill=WHITE, line=fill, text_color=NAVY,
                size=18, bold=False, line_w=1.5,
            )

        # Decision rule in the middle (links sunk to action)
        _add_text(slide, MARGIN, Inches(4.6), RULE_W, Inches(0.5),
                  "→  Decision rule:  ignore sunk costs when choosing what to do next",
                  size=22, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)

        _add_takeaway_bar(
            slide,
            "Sunk costs should never drive decisions",
            top=Inches(6.45), fill=GOLD, text_color=NAVY,
            width=Inches(10.0),
        )

    s = make_diagram_slide(
        prs, page_num=43,
        section_tag=SECTION_TAG_P2,
        title="Cost Types:  Fixed,  Sunk,  Variable",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The single most important cost concept for executives, in five "
        "words: sunk costs are not costs. They've already been spent; they "
        "cannot be recovered. Any forward-looking decision should ignore "
        "them. Period."
    ))


def slide_44(prs):
    """Group work: choosing a car (own car vs. company car)."""
    def draw(slide):
        # Setup paragraph
        _add_text(slide, MARGIN, Inches(1.85), RULE_W, Inches(0.55),
                  "For a business trip, you can use:",
                  size=22, bold=True, color=NAVY, font="Calibri")

        # Two options side by side
        opt_w = Inches(5.5)
        opt_h = Inches(1.4)
        gap = Inches(0.6)
        start_x = (SLIDE_W - opt_w * 2 - gap) // 2
        _add_outlined_box(
            slide, start_x, Inches(2.45), opt_w, opt_h,
            "Your own car\n+  reimbursed 50¢ / mile",
            fill=WHITE, line=NAVY, text_color=NAVY,
            size=20, bold=True, line_w=1.5,
        )
        _add_outlined_box(
            slide, start_x + opt_w + gap, Inches(2.45), opt_w, opt_h,
            "Company car\n(full cost incl. gas covered)",
            fill=WHITE, line=NAVY, text_color=NAVY,
            size=20, bold=True, line_w=1.5,
        )

        # Cost breakdown for own car
        _add_text(slide, MARGIN, Inches(4.05), RULE_W, Inches(0.4),
                  "Costs associated with your car  (per mile driven):",
                  size=18, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)
        costs = [
            ("20¢", "insurance"),
            ("20¢", "maintenance"),
            ("15¢", "electricity"),
            ("45¢", "lease on the vehicle"),
        ]
        cost_w = Inches(2.85)
        cost_h = Inches(1.1)
        gap2 = Inches(0.1)
        total_cw = cost_w * 4 + gap2 * 3
        cx0 = (SLIDE_W - total_cw) // 2
        for i, (amt, lbl) in enumerate(costs):
            cx = cx0 + (cost_w + gap2) * i
            _add_filled_box(
                slide, cx, Inches(4.55), cost_w, cost_h,
                f"{amt}\n{lbl}",
                fill=NAVY, text_color=WHITE,
                size=18, bold=True,
            )

        # Question + Discussion break
        _add_text(slide, MARGIN, Inches(5.95), RULE_W, Inches(0.45),
                  "Should you use your own car or the company car?",
                  size=22, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)
        _add_discussion_break(slide, top=Inches(6.55), width=Inches(4.8))

    s = make_diagram_slide(
        prs, page_num=44,
        section_tag=SECTION_TAG_P2,
        title="Group Work:  Your Car or the Company Car?",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "A group-work exercise. Two cars, same daily operating cost, but "
        "different sunk amounts. Which should you drive today? Whatever "
        "your gut says, the answer is: ignore the sunk cost – it's the "
        "same either way (the lease is sunk; only the variable per-mile "
        "costs matter for the marginal trip decision)."
    ))


def slide_45(prs):
    """Why studios finish movies they know will flop: Waterworld (1995).

    Source has the iconic Waterworld poster as a large background image.
    """
    def draw(slide):
        # The Waterworld poster, large and centered
        _add_source_image(slide, 45, "rId3",
                          left=Inches(3.7), top=Inches(1.65),
                          height=Inches(4.85))

        # Caption
        _add_text(slide, MARGIN, Inches(6.55), RULE_W, Inches(0.4),
                  "Sunk cost in Waterworld (1995):  $175M spent before release",
                  size=18, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)

    s = make_diagram_slide(
        prs, page_num=45,
        section_tag=SECTION_TAG_P2,
        title="Why Studios Finish Movies They Know Will Flop:  Waterworld",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Hollywood case. Kevin Costner's Waterworld – they knew it would "
        "flop, why release it anyway? Answer: even a flop adds revenue "
        "net of marketing/release costs. The hundreds of millions already "
        "spent on production are sunk. Decision should be forward-looking "
        "only."
    ))


def slide_46(prs):
    """Sunk cost in Waterworld – decision tree across 3 scenarios.

    Source has a complex 3-column decision table with budget/sunk/raised
    figures and the conclusion 'Make the film!' for each.  Rebuild as a
    cleaner 3-column matrix.
    """
    def draw(slide):
        # Three scenarios: Optimistic / Neutral / Pessimistic
        scenarios = [
            ("Optimistic",  150, 16,  100, "+50", "Make the film!"),
            ("Neutral",     150, 100, 140, "+10", "Make the film!"),
            ("Pessimistic", 150, 140, 175, "−25", "Make the film!"),
        ]
        col_w = Inches(3.9)
        col_gap = Inches(0.2)
        col_x0 = (SLIDE_W - col_w * 3 - col_gap * 2) // 2

        # Row headers down the left
        labels = [
            "Scenario",
            "Sunk cost  (already spent, $M)",
            "Extra cost to release  ($M)",
            "Expected revenue if released  ($M)",
            "Net  (release vs. shelve)",
            "Decision",
        ]
        for j, sc in enumerate(scenarios):
            x = col_x0 + (col_w + col_gap) * j
            # Header row (band)
            _add_filled_box(
                slide, x, Inches(1.85), col_w, Inches(0.55),
                sc[0], fill=NAVY, text_color=WHITE,
                size=20, bold=True,
            )
            # 4 number rows
            vals = [
                f"{sc[1]}",           # sunk
                f"{sc[3]}",           # extra cost to release
                f"{sc[2]} (vs. {sc[1]})" if False else f"{sc[2]}",  # revenue
                sc[4],                # net
            ]
            for i, v in enumerate(vals):
                _add_outlined_box(
                    slide, x, Inches(2.4 + i * 0.7), col_w, Inches(0.6),
                    v, fill=WHITE, line=NAVY, text_color=NAVY,
                    size=20, bold=False, line_w=1.0,
                )
            # Decision band
            _add_filled_box(
                slide, x, Inches(5.25), col_w, Inches(0.65),
                sc[5], fill=GOLD, text_color=NAVY,
                size=20, bold=True,
            )

        # Bottom takeaway
        _add_takeaway_bar(
            slide,
            "Sunk costs are sunk  —  release whenever  revenue  >  marginal release cost",
            top=Inches(6.45), fill=NAVY, width=Inches(11.5),
        )

    s = make_diagram_slide(
        prs, page_num=46,
        section_tag=SECTION_TAG_P2,
        title="Waterworld:  Three Scenarios,  Same Decision",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The decision tree behind the Waterworld decision. Three scenarios "
        "– optimistic, neutral, pessimistic – differ in how much was sunk "
        "vs. how much more it costs to release. In ALL three the "
        "release-anyway revenue exceeds the marginal release cost. So "
        "they released. Sunk costs are sunk."
    ))


def slide_47(prs):
    """Modern sunk cost: Meta's Reality Labs has lost $50B+ since 2020.

    Layout: bullets left + Meta Quest 3 image right (mirrors revision-v1).
    """
    def draw(slide):
        bullets = [
            ("Meta has poured ~$50B into Reality Labs since 2020", 0),
            ("Metaverse / VR / AR – Quest headsets, Horizon Worlds", 1),
            ("Wall Street keeps asking when it pays off", 0),
            ("Zuckerberg keeps investing – past losses are sunk", 0),
            ("Right question: does the next $10B have positive expected value?", 0),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(8.7), height=Inches(4.5),
            items=bullets,
            size=22, sub_size=18, line_spacing_pts=12,
        )

        # Meta Quest 3 picture on the right
        _add_source_image(slide, 47, "rId2",
                          left=Inches(9.55), top=Inches(2.0),
                          width=Inches(3.3))
        _add_text(slide, Inches(9.55), Inches(6.05), Inches(3.3), Inches(0.3),
                  "Meta Quest 3  (CC BY-SA, Wikimedia)",
                  size=11, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)

        _add_takeaway_bar(
            slide,
            "Same lesson as Waterworld  —  classic sunk-cost discipline, 2020s edition",
            top=Inches(6.5), fill=GOLD, text_color=NAVY,
            width=Inches(11.0),
        )

    s = make_diagram_slide(
        prs, page_num=47,
        section_tag=SECTION_TAG_P2,
        title="Modern Sunk Cost:  Meta's Reality Labs Has Lost $50B+ Since 2020",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The same logic, in a current strategic context. Meta's Reality "
        "Labs has lost roughly $50B from 2020-2024 on Metaverse and VR "
        "investments. Wall Street keeps asking when it pays off. "
        "Zuckerberg keeps investing – correctly – because the past losses "
        "are sunk. The right question is forward-looking: does the next "
        "$10B have positive expected value? Same lesson as Waterworld, "
        "dressed in 2025 clothes."
    ))


def slide_48(prs):
    """Opportunity cost is a real cost: Apple's canceled Apple Car."""
    def draw(slide):
        bullets = [
            ("Apple killed Project Titan in 2024 after ~10 years and ~$10B spent", 0),
            ("Sunk costs ≠ a reason to keep going", 0),
            ("Real reason to stop: opportunity cost of capital + 2,000 engineers", 0),
            ("Reallocated → AI / Apple Intelligence (the higher-MPL use)", 1),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(7.0), height=Inches(4.5),
            items=bullets,
            size=22, sub_size=18, line_spacing_pts=12,
        )

        # Vanarama Apple Car concept render on the right
        _add_source_image(slide, 48, "rId3",
                          left=Inches(7.45), top=Inches(2.2),
                          width=Inches(5.55))
        _add_text(slide, Inches(7.45), Inches(5.25), Inches(5.55), Inches(0.35),
                  "Vanarama Apple Car concept  (fair use, © Vanarama)",
                  size=11, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)

        _add_takeaway_bar(
            slide,
            "Opportunity cost  =  the next-best use of the same dollars  &  people",
            top=Inches(6.5), fill=NAVY, width=Inches(10.5),
        )

    s = make_diagram_slide(
        prs, page_num=48,
        section_tag=SECTION_TAG_P2,
        title="Opportunity Cost Is a Real Cost:  Apple's Canceled Apple Car",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The flip side of sunk costs is opportunity cost. Apple killed "
        "Project Titan in 2024 after roughly a decade and $10B spent. "
        "The sunk costs were sunk. What killed the project was "
        "opportunity cost: those engineers and that capital had a "
        "higher-MPL use in Apple Intelligence and AI."
    ))


def slide_49(prs):
    """Dictionary of costs – the cheat-sheet taxonomy."""
    def draw(slide):
        # Use the source taxonomy image; it's a clean diagram of all 6 cost
        # types organized hierarchically.
        _add_source_image(slide, 49, "rId3",
                          left=Inches(2.0), top=Inches(1.85),
                          width=Inches(9.3))

        _add_text(slide, MARGIN, Inches(6.55), RULE_W, Inches(0.4),
                  "Cheat-sheet to refer back to for the rest of the module",
                  size=18, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)

    s = make_diagram_slide(
        prs, page_num=49,
        section_tag=SECTION_TAG_P2,
        title="Dictionary of Costs",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Quick reference: fixed, variable, sunk, marginal, average, "
        "opportunity. A cheat sheet you'll refer to for the rest of the "
        "module. Make sure you can give a one-sentence example of each."
    ))


def slide_50(prs):
    """Ross Stores annual report – grounding cost concepts in a real 10-K."""
    def draw(slide):
        # Main image: a page from Ross Stores' annual report
        _add_source_image(slide, 50, "rId3",
                          left=Inches(0.5), top=Inches(1.85),
                          height=Inches(4.6))

        # Labels on the right — categorize lines as FC / VC
        _add_text(slide, Inches(10.0), Inches(2.4), Inches(3.0), Inches(0.45),
                  "Mostly Fixed", size=20, bold=True, color=NAVY,
                  font="Calibri")
        _add_filled_box(
            slide, Inches(10.0), Inches(2.85), Inches(3.0), Inches(0.7),
            "Rent, depreciation,\noccupancy",
            fill=NAVY, text_color=WHITE, size=14, bold=False,
        )
        _add_text(slide, Inches(10.0), Inches(3.85), Inches(3.0), Inches(0.45),
                  "Mostly Variable", size=20, bold=True, color=NAVY,
                  font="Calibri")
        _add_filled_box(
            slide, Inches(10.0), Inches(4.3), Inches(3.0), Inches(0.7),
            "Cost of goods sold,\nfreight, payment fees",
            fill=GOLD, text_color=NAVY, size=14, bold=False,
        )
        _add_text(slide, Inches(10.0), Inches(5.3), Inches(3.0), Inches(0.45),
                  "Mix", size=20, bold=True, color=NAVY,
                  font="Calibri")
        _add_outlined_box(
            slide, Inches(10.0), Inches(5.75), Inches(3.0), Inches(0.7),
            "Wages,\nstore operations",
            fill=WHITE, line=NAVY, text_color=NAVY, size=14, bold=False,
            line_w=1.5,
        )

        _add_takeaway_bar(
            slide,
            "Every 10-K can be read as a fixed-vs-variable split",
            top=Inches(6.55), fill=GOLD, text_color=NAVY,
            width=Inches(9.5),
        )

    s = make_diagram_slide(
        prs, page_num=50,
        section_tag=SECTION_TAG_P2,
        title="Cost Concepts in the Real World:  Ross Stores Annual Report",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Cost concepts in the real world – a page from Ross Stores' "
        "annual report. Have students classify each line as fixed or "
        "variable. The point is to ground the abstract concepts in "
        "something they'll see in a 10-K."
    ))


def slide_51(prs):
    """ChatGPT subscription tiers — Plus vs Team marginal cost question."""
    def draw(slide):
        # Two tier panels, side by side
        tier_w = Inches(4.0)
        tier_h = Inches(2.6)
        gap = Inches(0.4)
        x0 = MARGIN + Inches(0.3)

        # Plus tier
        _add_filled_box(
            slide, x0, Inches(2.0), tier_w, Inches(0.6),
            "ChatGPT Plus", fill=NAVY, text_color=WHITE,
            size=22, bold=True,
        )
        _add_outlined_box(
            slide, x0, Inches(2.6), tier_w, Inches(2.0),
            "$20 / user / month\n\n1 user minimum",
            fill=WHITE, line=NAVY, text_color=NAVY,
            size=22, bold=False, line_w=1.5,
        )

        # Team tier
        _add_filled_box(
            slide, x0 + tier_w + gap, Inches(2.0), tier_w, Inches(0.6),
            "ChatGPT Team", fill=GOLD, text_color=NAVY,
            size=22, bold=True,
        )
        _add_outlined_box(
            slide, x0 + tier_w + gap, Inches(2.6), tier_w, Inches(2.0),
            "$25 / user / month\n\n2 users minimum",
            fill=WHITE, line=GOLD, text_color=NAVY,
            size=22, bold=False, line_w=1.5,
        )

        # ChatGPT phone image on the right
        _add_source_image(slide, 51, "rId5",
                          left=Inches(9.6), top=Inches(2.0),
                          height=Inches(3.5))
        _add_text(slide, Inches(9.6), Inches(5.55), Inches(3.2), Inches(0.3),
                  "ChatGPT on iPhone  (CC BY-SA, Wikimedia)",
                  size=11, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)

        # Question
        _add_text(slide, MARGIN, Inches(5.0), Inches(8.6), Inches(0.55),
                  "You're on Plus and want to add a 2nd user.",
                  size=22, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)
        _add_text(slide, MARGIN, Inches(5.5), Inches(8.6), Inches(0.55),
                  "What is the marginal cost of that 2nd user?",
                  size=22, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)

        _add_takeaway_bar(
            slide,
            "Hint:  it's not just $25",
            top=Inches(6.5), fill=GOLD, text_color=NAVY,
            width=Inches(5.0),
        )

    s = make_diagram_slide(
        prs, page_num=51,
        section_tag=SECTION_TAG_P2,
        title="Marginal Cost ≠ Average Cost:  ChatGPT Subscription Tiers",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "A pricing case students will recognize from their own "
        "subscriptions. ChatGPT Plus costs $20/user/month; ChatGPT Team "
        "costs $25/user/month but requires a 2-user minimum. If you have "
        "one user on Plus and want to add a second, what's the marginal "
        "cost of that second user? Hint: it's not just $25."
    ))


def slide_52(prs):
    """Poll: MC of the 2nd ChatGPT user?

    Source slide is a full-bleed PollEv screenshot.
    """
    def draw(slide):
        _add_source_image(slide, 52, "rId4",
                          left=Inches(3.2), top=Inches(1.85),
                          height=Inches(5.1))
        _add_text(slide, MARGIN, Inches(7.0), RULE_W, Inches(0.3),
                  "Respond at PollEv.com/nvoigtlaender",
                  size=14, italic=True, color=GRAY,
                  align=PP_ALIGN.CENTER, font="Calibri")

    s = make_diagram_slide(
        prs, page_num=52,
        section_tag=SECTION_TAG_P2,
        title="What's the MC of Adding the 2nd User?",
        draw_diagram=draw,
    )
    _draw_poll_pill(s)
    _set_notes(s, (
        "PollEv – compute MC of adding the second user. Watch for them "
        "assuming MC = the Team rate of $25."
    ))


def slide_53(prs):
    """Solution: MC = $30 / user-month."""
    def draw(slide):
        # Two side-by-side cost lines
        col_w = Inches(5.6)
        col_gap = Inches(0.5)
        x0 = (SLIDE_W - col_w * 2 - col_gap) // 2

        _add_filled_box(
            slide, x0, Inches(1.95), col_w, Inches(0.55),
            "1 user on Plus", fill=NAVY, text_color=WHITE,
            size=20, bold=True,
        )
        _add_outlined_box(
            slide, x0, Inches(2.5), col_w, Inches(0.9),
            "1 × $20  =  $20 / month",
            fill=WHITE, line=NAVY, text_color=NAVY,
            size=22, bold=True, line_w=1.5,
        )

        _add_filled_box(
            slide, x0 + col_w + col_gap, Inches(1.95), col_w, Inches(0.55),
            "2 users on Team", fill=GOLD, text_color=NAVY,
            size=20, bold=True,
        )
        _add_outlined_box(
            slide, x0 + col_w + col_gap, Inches(2.5), col_w, Inches(0.9),
            "2 × $25  =  $50 / month",
            fill=WHITE, line=GOLD, text_color=NAVY,
            size=22, bold=True, line_w=1.5,
        )

        # MC calculation row
        _add_text(slide, MARGIN, Inches(3.7), RULE_W, Inches(0.5),
                  "Marginal cost of the 2nd user:",
                  size=22, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)
        _add_filled_box(
            slide, (SLIDE_W - Inches(7.5)) // 2, Inches(4.25),
            Inches(7.5), Inches(1.0),
            "MC  =  $50  −  $20  =  $30 / month",
            fill=NAVY, text_color=WHITE, size=28, bold=True,
        )

        # Intuition
        _add_text(slide, MARGIN, Inches(5.5), RULE_W, Inches(0.5),
                  "Re-price the existing user ($20 → $25)  +  add a new one ($25)  =  $30",
                  size=18, italic=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)

        _add_takeaway_bar(
            slide,
            "MC  >  AC :  subscription tiers can hide a higher marginal cost",
            top=Inches(6.5), fill=GOLD, text_color=NAVY,
            width=Inches(11.0),
        )

    s = make_diagram_slide(
        prs, page_num=53,
        section_tag=SECTION_TAG_P2,
        title="Solution:  MC = $30 / user · month",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Reveal: MC = $30/user-month for the second user (you go from "
        "$20 to $50 total). The lesson: tiered subscription pricing can "
        "hide a marginal cost that's HIGHER than the average rate. The "
        "opposite of the classic 'volume discount' story – and "
        "increasingly common in SaaS."
    ))


def slide_54(prs):
    """Marginal cost in finance: the true rate on a bigger loan."""
    def draw(slide):
        # Setup
        _add_text(slide, MARGIN, Inches(1.85), RULE_W, Inches(0.5),
                  "Two options for a personal loan:",
                  size=22, bold=True, color=NAVY, font="Calibri")

        # Two options side by side
        opt_w = Inches(5.6)
        opt_h = Inches(1.0)
        gap = Inches(0.4)
        x0 = (SLIDE_W - opt_w * 2 - gap) // 2

        _add_filled_box(
            slide, x0, Inches(2.55), opt_w, opt_h,
            "Loan A:  $100K  @  5% annual",
            fill=NAVY, text_color=WHITE, size=22, bold=True,
        )
        _add_filled_box(
            slide, x0 + opt_w + gap, Inches(2.55), opt_w, opt_h,
            "Loan B:  $110K  @  6% annual",
            fill=NAVY, text_color=WHITE, size=22, bold=True,
        )

        # Question
        _add_text(slide, MARGIN, Inches(3.7), RULE_W, Inches(0.5),
                  "What's the marginal interest rate on the extra $10K?",
                  size=22, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)

        # Calculation
        steps = [
            "Extra interest:   $6,600  −  $5,000   =   $1,600",
            "Extra loan amount:   $10,000",
            "Marginal rate:   $1,600  /  $10,000   =   16%",
        ]
        for i, step in enumerate(steps):
            _add_text(slide, MARGIN, Inches(4.4 + i * 0.45),
                      RULE_W, Inches(0.4),
                      step, size=20, color=NAVY, font="Calibri",
                      align=PP_ALIGN.CENTER)

        _add_takeaway_bar(
            slide,
            "Marginal cost (16%) is much higher than average rate (6%)",
            top=Inches(6.5), fill=GOLD, text_color=NAVY,
            width=Inches(11.0),
        )

    s = make_diagram_slide(
        prs, page_num=54,
        section_tag=SECTION_TAG_P2,
        title="Marginal Cost in Finance:  The True Rate on a Bigger Loan",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Same concept applied to finance: when a bigger loan comes with "
        "a worse rate, the marginal cost of the extra dollars is much "
        "higher than the average rate the loan was quoted at. This is a "
        "very common executive trap when comparing financing options."
    ))


def slide_55(prs):
    """Rivian's Georgia plant – weekly Total Cost function.

    Layout: bullet intro on top, cost function box, source cost-curve image.
    """
    def draw(slide):
        # Intro lines
        _add_text(slide, MARGIN, Inches(1.85), RULE_W, Inches(0.5),
                  "Collect data on total cost at different output levels (Q)",
                  size=22, color=NAVY, font="Calibri")
        _add_text(slide, MARGIN, Inches(2.3), RULE_W, Inches(0.5),
                  "→  Estimate the cost function",
                  size=22, italic=True, color=GRAY, font="Calibri")

        # Cost function (OMML / Cambria Math, with real Q² superscript).
        # TC = 10,000,000 + 30,000 · Q + 40 · Q²
        eq_xml = (
            _omml_text('TC') +
            _omml_text(' = ') +
            _omml_text('10,000,000') +
            _omml_text(' + ') +
            _omml_text('30,000') +
            _omml_text(' · ') +
            _omml_run('Q') +
            _omml_text(' + ') +
            _omml_text('40') +
            _omml_text(' · ') +
            _omml_sup(_omml_run('Q'), _omml_text('2'))
        )
        _add_math_equation(
            slide, (SLIDE_W - Inches(9.0)) // 2, Inches(2.95),
            Inches(9.0), Inches(0.85),
            eq_xml, size_pt=26, color=WHITE, fill=NAVY,
        )

        # Cost-curve image from source
        _add_source_image(slide, 55, "rId4",
                          left=Inches(3.0), top=Inches(4.0),
                          height=Inches(2.5))

        _add_takeaway_bar(
            slide,
            "Fixed plus variable, with a convex quadratic term as scale rises",
            top=Inches(6.55), fill=GOLD, text_color=NAVY,
            width=Inches(10.5),
        )

    s = make_diagram_slide(
        prs, page_num=55,
        section_tag=SECTION_TAG_P2,
        title="Rivian's Georgia Plant —  Weekly Cost",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Now back to Rivian. The Georgia plant's weekly total cost "
        "function – how total cost varies with output Q. Fixed cost of "
        "$10M, variable cost of $30K per vehicle, plus a quadratic term "
        "($40 × Q²) that captures the rising marginal cost at high "
        "volumes."
    ))


def slide_56(prs):
    """Rivian's Georgia plant – Cost Components (the decomposition chart)."""
    def draw(slide):
        # The cost-components stacked-area chart from the source
        _add_source_image(slide, 56, "rId3",
                          left=Inches(2.5), top=Inches(1.85),
                          height=Inches(4.4))

        _add_takeaway_bar(
            slide,
            "Fixed costs dominate at low Q;  variable rises linearly,  quadratic kicks in late",
            top=Inches(6.5), fill=NAVY, width=Inches(11.5),
        )

    s = make_diagram_slide(
        prs, page_num=56,
        section_tag=SECTION_TAG_P2,
        title="Rivian's Georgia Plant —  Cost Components",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Decompose total cost into its components: fixed plus variable, "
        "broken into sub-categories. The visual lets students see how "
        "the cost structure shifts as output grows: at low Q, the fixed "
        "cost ($10M) dominates; at high Q, the quadratic term takes over."
    ))


def slide_57(prs):
    """Rivian's Georgia plant – Per-Unit Costs (AC, AVC, MC)."""
    def draw(slide):
        # Per-unit cost chart from the source
        _add_source_image(slide, 57, "rId3",
                          left=Inches(2.3), top=Inches(1.85),
                          height=Inches(4.5))

        # Legend labels (3 curves)
        legend_w = Inches(2.0)
        ly = Inches(2.0)
        _add_filled_box(slide, Inches(10.5), ly, legend_w, Inches(0.5),
                         "AC", fill=NAVY, text_color=WHITE,
                         size=20, bold=True)
        _add_filled_box(slide, Inches(10.5), ly + Inches(0.6),
                         legend_w, Inches(0.5),
                         "AVC", fill=GOLD, text_color=NAVY,
                         size=20, bold=True)
        _add_filled_box(slide, Inches(10.5), ly + Inches(1.2),
                         legend_w, Inches(0.5),
                         "MC", fill=GRAY, text_color=WHITE,
                         size=20, bold=True)

        _add_takeaway_bar(
            slide,
            "MC crosses AVC and AC at their minima  —  the textbook U-shape",
            top=Inches(6.5), fill=GOLD, text_color=NAVY,
            width=Inches(11.0),
        )

    s = make_diagram_slide(
        prs, page_num=57,
        section_tag=SECTION_TAG_P2,
        title="Rivian's Georgia Plant —  Per-Unit Costs",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Same Rivian data, but expressed per vehicle: average cost (AC), "
        "average variable cost (AVC), and marginal cost (MC). Three "
        "curves on one chart. Note: MC crosses AVC and AC at their "
        "respective minima – this is the textbook U-shape and a "
        "diagnostic students should be able to read."
    ))


def slide_58(prs):
    """Cost Estimation – iPhone 17 teardown (discussion setup, no internet)."""
    def draw(slide):
        # Setup prompt
        _add_text(slide, MARGIN, Inches(1.85), RULE_W, Inches(0.55),
                  "Without Internet Access!",
                  size=24, bold=True, color=GOLD, font="Calibri",
                  align=PP_ALIGN.CENTER)

        _add_text(slide, MARGIN, Inches(2.5), RULE_W, Inches(0.55),
                  "What is the average variable cost of producing a current-generation iPhone 17?",
                  size=22, bold=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)

        # Components grid (3 boxes)
        _add_text(slide, MARGIN, Inches(3.4), RULE_W, Inches(0.4),
                  "Provide a numerical guess  +  rough split:",
                  size=18, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)

        comps = ["Processor", "Other material inputs", "Labor"]
        cw = Inches(3.3)
        ch = Inches(1.1)
        gap = Inches(0.2)
        x0 = (SLIDE_W - cw * 3 - gap * 2) // 2
        for i, name in enumerate(comps):
            _add_outlined_box(
                slide, x0 + (cw + gap) * i, Inches(3.9),
                cw, ch, f"{name}\n\n$ ?",
                fill=WHITE, line=NAVY, text_color=NAVY,
                size=20, bold=True, line_w=1.5,
            )

        # Hint as gold callout
        _add_text(slide, MARGIN, Inches(5.4), RULE_W, Inches(0.4),
                  "Total retail price of an iPhone 17 ≈  $1,199",
                  size=20, italic=True, color=NAVY, font="Calibri",
                  align=PP_ALIGN.CENTER)

        _add_discussion_break(slide, top=Inches(6.45), width=Inches(5.0))

    s = make_diagram_slide(
        prs, page_num=58,
        section_tag=SECTION_TAG_P2,
        title="Cost Estimation:  What Does an iPhone Cost to Make?",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "iPhone manufacturing teardowns – what do the parts and assembly "
        "actually cost for the current-gen iPhone 17? Have students "
        "guess BEFORE you reveal in two slides. Discuss what's included "
        "(components, labor, assembly) and what's not (R&D, marketing, "
        "retail). Most students wildly overestimate the build cost."
    ))


def slide_59(prs):
    """Poll: AVC of an iPhone 17?

    Source slide is a full-bleed PollEv screenshot.
    """
    def draw(slide):
        _add_source_image(slide, 59, "rId4",
                          left=Inches(3.2), top=Inches(1.85),
                          height=Inches(5.1))
        _add_text(slide, MARGIN, Inches(7.0), RULE_W, Inches(0.3),
                  "Respond at PollEv.com/nvoigtlaender",
                  size=14, italic=True, color=GRAY,
                  align=PP_ALIGN.CENTER, font="Calibri")

    s = make_diagram_slide(
        prs, page_num=59,
        section_tag=SECTION_TAG_P2,
        title="What's the AVC of an iPhone 17?",
        draw_diagram=draw,
    )
    _draw_poll_pill(s)
    _set_notes(s, (
        "PollEv – estimate AVC of an iPhone 17 given the teardown data. "
        "Most students will overestimate."
    ))


def slide_60(prs):
    """Solution: AVC of iPhone 17 ≈ $580 (vs. $1,199 retail)."""
    def draw(slide):
        # Teardown image on the left
        _add_source_image(slide, 60, "rId4",
                          left=Inches(0.5), top=Inches(1.85),
                          height=Inches(4.6))

        # Numbers on the right
        _add_text(slide, Inches(7.5), Inches(2.0), Inches(5.4), Inches(0.55),
                  "Retail price:",
                  size=22, color=GRAY, font="Calibri")
        _add_text(slide, Inches(7.5), Inches(2.55), Inches(5.4), Inches(0.6),
                  "$1,199",
                  size=36, bold=True, color=NAVY, font="Calibri")
        _add_text(slide, Inches(7.5), Inches(3.4), Inches(5.4), Inches(0.55),
                  "Total variable cost  (TVC):",
                  size=22, color=GRAY, font="Calibri")
        _add_text(slide, Inches(7.5), Inches(3.95), Inches(5.4), Inches(0.6),
                  "≈  $580",
                  size=36, bold=True, color=GOLD, font="Calibri")

        # Missing components (small list)
        _add_text(slide, Inches(7.5), Inches(4.85), Inches(5.4), Inches(0.4),
                  "Plus missing components:",
                  size=14, italic=True, color=GRAY, font="Calibri")
        miss = ["Shipping & handling", "Customer service", "Warranty costs"]
        for i, m in enumerate(miss):
            _add_text(slide, Inches(7.7), Inches(5.2 + i * 0.32),
                      Inches(5.2), Inches(0.3),
                      f"–  {m}",
                      size=14, color=GRAY, font="Calibri")

        _add_takeaway_bar(
            slide,
            "About half of retail goes to fixed-cost recovery, R&D, retail margin",
            top=Inches(6.5), fill=NAVY, width=Inches(11.0),
        )

    s = make_diagram_slide(
        prs, page_num=60,
        section_tag=SECTION_TAG_P2,
        title="AVC of iPhone 17  ≈  $580",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Reveal: AVC ≈ $580 for iPhone 17 (vs. ~$1,199 retail). About "
        "half of retail. The rest – fixed-cost recovery, gross margin – "
        "funds Apple's R&D, retail network, and ecosystem. Students "
        "consistently overestimate this number."
    ))


def slide_61(prs):
    """iPhone naïve cost function – total cost (line chart, hand-drawn)."""
    def draw(slide):
        # Axes
        AX_L = Inches(1.6)
        AX_T = Inches(2.0)
        AX_W = Inches(7.5)
        AX_H = Inches(4.2)

        # Y-axis line
        _add_arrow(slide,
                   start_xy=(int(AX_L), int(AX_T + AX_H)),
                   end_xy=(int(AX_L), int(AX_T)),
                   color=NAVY, weight_pt=1.5, head=True)
        # X-axis line
        _add_arrow(slide,
                   start_xy=(int(AX_L), int(AX_T + AX_H)),
                   end_xy=(int(AX_L + AX_W), int(AX_T + AX_H)),
                   color=NAVY, weight_pt=1.5, head=True)

        # Axis labels
        _add_text(slide, AX_L - Inches(1.0), AX_T - Inches(0.45),
                   Inches(2.0), Inches(0.4),
                   "$  (Total cost)", size=18, bold=True, color=NAVY,
                   font="Calibri")
        _add_text(slide, AX_L + AX_W - Inches(2.5),
                   AX_T + AX_H + Inches(0.15),
                   Inches(3.5), Inches(0.4),
                   "Q  (quantity produced)", size=18, bold=True, color=NAVY,
                   font="Calibri")

        # Fixed-cost intercept (horizontal mark) at ~ TFC level
        tfc_y = AX_T + AX_H - Inches(0.9)   # TFC ~ 20% of vertical
        _add_text(slide, AX_L - Inches(0.7), tfc_y - Inches(0.2),
                   Inches(0.7), Inches(0.4),
                   "TFC", size=18, bold=True, color=GOLD, font="Calibri")
        # Short tick from y-axis
        _add_arrow(slide,
                   start_xy=(int(AX_L - Inches(0.05)), int(tfc_y)),
                   end_xy=(int(AX_L + Inches(0.05)), int(tfc_y)),
                   color=GOLD, weight_pt=2.0, head=False)

        # The TC line: starts at (AX_L, tfc_y), goes up-right
        line_end_x = AX_L + AX_W - Inches(0.5)
        line_end_y = AX_T + Inches(0.4)
        _add_arrow(slide,
                   start_xy=(int(AX_L), int(tfc_y)),
                   end_xy=(int(line_end_x), int(line_end_y)),
                   color=NAVY, weight_pt=3.0, head=False)

        # Equation callout on the right (OMML – upright acronyms + italic Q)
        eq_xml = (
            _omml_text('TC') +
            _omml_text(' = ') +
            _omml_text('TFC') +
            _omml_text(' + ') +
            _omml_text('500') +
            _omml_text(' · ') +
            _omml_run('Q')
        )
        _add_math_equation(
            slide, Inches(9.5), Inches(3.0), Inches(3.5), Inches(0.95),
            eq_xml, size_pt=24, color=NAVY, fill=GOLD,
        )

        # Slope annotation
        _add_callout_box(
            slide, Inches(5.8), Inches(4.3), Inches(2.4), Inches(0.45),
            "slope  =  $500",
            fill=NAVY, text_color=WHITE, size=14, bold=True,
        )

        _add_takeaway_bar(
            slide,
            "Linear naïve TC :  fixed cost  +  constant marginal cost  ($500 / unit)",
            top=Inches(6.55), fill=NAVY, width=Inches(11.5),
        )

    s = make_diagram_slide(
        prs, page_num=61,
        section_tag=SECTION_TAG_P2,
        title="iPhone:  Naïve Cost Function (Total Cost)",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "iPhone naïve cost function – total cost as output rises. "
        "Linear: a fixed cost plus a constant marginal cost. This is "
        "the simplest model of a cost function; we'll see more complex "
        "shapes later (capacity constraints, increasing returns)."
    ))


def slide_62(prs):
    """iPhone naïve cost function – per-unit (constant MC line)."""
    def draw(slide):
        AX_L = Inches(1.6)
        AX_T = Inches(2.0)
        AX_W = Inches(7.5)
        AX_H = Inches(4.2)

        # Y-axis
        _add_arrow(slide,
                   start_xy=(int(AX_L), int(AX_T + AX_H)),
                   end_xy=(int(AX_L), int(AX_T)),
                   color=NAVY, weight_pt=1.5, head=True)
        # X-axis
        _add_arrow(slide,
                   start_xy=(int(AX_L), int(AX_T + AX_H)),
                   end_xy=(int(AX_L + AX_W), int(AX_T + AX_H)),
                   color=NAVY, weight_pt=1.5, head=True)

        _add_text(slide, AX_L - Inches(1.0), AX_T - Inches(0.45),
                   Inches(2.0), Inches(0.4),
                   "$ per unit", size=18, bold=True, color=NAVY,
                   font="Calibri")
        _add_text(slide, AX_L + AX_W - Inches(2.5),
                   AX_T + AX_H + Inches(0.15),
                   Inches(3.5), Inches(0.4),
                   "Q  (quantity produced)", size=18, bold=True, color=NAVY,
                   font="Calibri")

        # Constant horizontal line at $500 mark
        mc_y = AX_T + AX_H - Inches(2.0)   # half-way up
        _add_arrow(slide,
                   start_xy=(int(AX_L + Inches(0.2)), int(mc_y)),
                   end_xy=(int(AX_L + AX_W - Inches(0.3)), int(mc_y)),
                   color=NAVY, weight_pt=3.0, head=False)
        # Y-axis tick label "500"
        _add_text(slide, AX_L - Inches(0.7), mc_y - Inches(0.2),
                   Inches(0.7), Inches(0.4),
                   "500", size=18, bold=True, color=GOLD, font="Calibri",
                   align=PP_ALIGN.RIGHT)

        # Line label (OMML)
        eq_label = (
            _omml_text('MC') +
            _omml_text(' = ') +
            _omml_text('AVC') +
            _omml_text(' = ') +
            _omml_text('$500')
        )
        _add_math_equation(
            slide,
            AX_L + AX_W // 2 - Inches(1.8),
            mc_y - Inches(0.65),
            Inches(3.6), Inches(0.55),
            eq_label, size_pt=22, color=NAVY,
        )

        # Annotation callout (OMML inside)
        callout_xml = (
            _omml_text('When ') +
            _omml_text('MC') +
            _omml_text(' is constant, it equals ') +
            _omml_text('AVC')
        )
        _add_math_equation(
            slide,
            Inches(9.5), Inches(3.0), Inches(3.5), Inches(1.5),
            callout_xml, size_pt=18, color=NAVY, fill=GOLD,
        )

        _add_takeaway_bar(
            slide,
            "Per-unit view :  constant marginal cost,  AC falls as fixed cost spreads",
            top=Inches(6.55), fill=GOLD, text_color=NAVY,
            width=Inches(11.5),
        )

    s = make_diagram_slide(
        prs, page_num=62,
        section_tag=SECTION_TAG_P2,
        title="iPhone:  Naïve Cost Function (Per-Unit)",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Same data, per unit. Note the constant MC and (implicit) "
        "declining AC – the classic shape when fixed costs spread over "
        "more units. With constant MC, AVC = MC at every Q. "
        "AC = AVC + AFC, so AC starts above MC and falls toward it as "
        "Q → ∞."
    ))


# --------------------------------------------------------------------------
# Batch 5 – §2.2 Long-Run Costs & Economies of Scale  (slides 63-74)
# --------------------------------------------------------------------------


def slide_63(prs):
    """Bridge slide: real-world cost functions can be non-linear (U-shape
    MC). We will keep MC linear where possible – the iPhone toy model is
    fine for most decisions; the Rivian quadratic was the exception."""
    def draw(slide):
        # Two source images side by side: TC/VC/FC curves + U-shape MC
        _add_source_image(slide, 63, "rId1",
                          left=Inches(0.5), top=Inches(2.0),
                          height=Inches(3.6))
        _add_source_image(slide, 63, "rId2",
                          left=Inches(7.0), top=Inches(2.0),
                          height=Inches(3.6))
        # Captions
        _add_text(slide, Inches(0.5), Inches(5.65),
                  Inches(6.0), Inches(0.3),
                  "Total / variable / fixed cost — convex at high Q",
                  size=14, italic=True, color=GRAY,
                  font="Calibri", align=PP_ALIGN.CENTER)
        _add_text(slide, Inches(7.0), Inches(5.65),
                  Inches(6.0), Inches(0.3),
                  "Marginal cost — U-shape (high at low + high Q)",
                  size=14, italic=True, color=GRAY,
                  font="Calibri", align=PP_ALIGN.CENTER)

        _add_takeaway_bar(
            slide,
            "Keep MC linear when possible — use the U-shape only when scale really matters",
            top=Inches(6.5), fill=GOLD, text_color=NAVY,
            width=Inches(12.0),
        )

    s = make_diagram_slide(
        prs, page_num=63,
        section_tag=SECTION_TAG_P2,
        title="More Complex Cost Functions:  When MC Isn't Linear",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Reality check: cost functions can be more complex than a "
        "straight line. The classic textbook shape is a U for marginal "
        "cost – MC high at low Q (under-utilised plant), falls as you "
        "hit the sweet spot, then rises again at high Q (over-stretched "
        "capacity, overtime, congestion). We saw the convex piece in the "
        "Rivian quadratic. For most decisions in this course we will "
        "still use linear MC – it captures the right intuition without "
        "the algebra. Pull out the U-shape only when scale effects are "
        "the whole point."
    ))


def slide_64(prs):
    """Section divider – Part 2.2: Long-Run Costs & Economies of Scale.
    Mirror of slide_30 (Part 1.2) using the Part-2 highlight."""
    s = make_section_agenda(
        prs, page_num=64,
        current_part_idx=1,
        section_tag=SECTION_TAG_DIV,
        title="Part 2.2:  Long-Run Costs & Economies of Scale",
    )
    _set_notes(s, (
        "Transitioning from the cost concepts (fixed/variable/marginal/"
        "average) to what happens in the LONG run, when the plant size "
        "itself can change. Three big ideas next: long-run vs short-run "
        "cost curves, the LR-AC envelope, and economies of scale."
    ))


def slide_65(prs):
    """Short-run v. long-run costs – two-column comparison."""
    def draw(slide):
        col_w = Inches(6.0)
        col_h = Inches(0.85)
        gap = Inches(0.2)
        x_l = MARGIN
        x_r = MARGIN + col_w + gap
        y0 = Inches(2.0)

        # Headers
        _add_filled_box(slide, x_l, y0, col_w, col_h,
                         "Short Run",
                         fill=NAVY, text_color=WHITE, size=26, bold=True)
        _add_filled_box(slide, x_r, y0, col_w, col_h,
                         "Long Run",
                         fill=NAVY, text_color=WHITE, size=26, bold=True)

        # Body bullets
        sr_items = [
            "Capital  (K)  is FIXED",
            "Labor  (L)  is flexible",
            "Plant size already chosen",
            "→  Cost of changing Q given the plant you have",
        ]
        lr_items = [
            "BOTH K and L flexible",
            "Choose plant size from scratch",
            "Pick the optimal input mix for each Q",
            "→  Lowest cost of producing any Q",
        ]

        for i, (l, r) in enumerate(zip(sr_items, lr_items)):
            yi = y0 + col_h + Inches(0.25) + Inches(0.7) * i
            _add_text(slide, x_l + Inches(0.1), yi, col_w - Inches(0.2),
                      Inches(0.6), l,
                      size=18, color=NAVY, font="Calibri")
            _add_text(slide, x_r + Inches(0.1), yi, col_w - Inches(0.2),
                      Inches(0.6), r,
                      size=18, color=NAVY, font="Calibri")

        # Inequality at the bottom of the body
        eq_xml = (
            _omml_text('TC') +
            _omml_sub(_omml_text(''), _omml_text('SR')) +
            _omml_text('  ≥  ') +
            _omml_text('TC') +
            _omml_sub(_omml_text(''), _omml_text('LR'))
        )
        _add_math_equation(
            slide,
            (SLIDE_W - Inches(7.0)) // 2, Inches(5.55),
            Inches(7.0), Inches(0.7),
            eq_xml, size_pt=26, color=WHITE, fill=NAVY,
        )

        _add_takeaway_bar(
            slide,
            "Long-run costs are the lower envelope:  more freedom  ⇒  weakly cheaper",
            top=Inches(6.5), fill=GOLD, text_color=NAVY,
            width=Inches(11.0),
        )

    s = make_diagram_slide(
        prs, page_num=65,
        section_tag=SECTION_TAG_P2_LR,
        title="Short-Run vs. Long-Run Costs",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Key distinction. Short run: you're stuck with the plant you "
        "have, so you're paying for capacity you may not be using (or "
        "scrambling to get more). Long run: you get to build the right "
        "plant for the output you want. With more freedom comes weakly "
        "lower cost – long-run TC is never higher than short-run TC at "
        "any given Q, because the long run is the option to pick the "
        "best short-run plant for that Q."
    ))


def slide_66(prs):
    """LR-AC envelope schematic.

    A clean schematic: three plant-size SAC bands across a quantity axis,
    with the LR-AC line traced as their lower envelope. Drawn with shapes
    rather than curves (which python-pptx doesn't render smoothly).
    """
    def draw(slide):
        # Plot bounding box
        AX_L = Inches(1.2)
        AX_R = Inches(12.5)
        AX_T = Inches(2.1)
        AX_B = Inches(5.8)
        AX_W = AX_R - AX_L
        AX_H = AX_B - AX_T

        # Axes
        _add_arrow(slide, (AX_L, AX_B), (AX_R, AX_B),
                   color=NAVY, weight_pt=1.5, head=True)
        _add_arrow(slide, (AX_L, AX_B), (AX_L, AX_T),
                   color=NAVY, weight_pt=1.5, head=True)
        _add_text(slide, AX_R - Inches(0.9), AX_B + Inches(0.05),
                  Inches(1.0), Inches(0.3),
                  "Q  (output)", size=14, italic=True, color=GRAY,
                  font="Calibri")
        _add_text(slide, AX_L - Inches(0.9), AX_T - Inches(0.05),
                  Inches(1.0), Inches(0.3),
                  "$ / unit", size=14, italic=True, color=GRAY,
                  font="Calibri", align=PP_ALIGN.RIGHT)

        # Three plant-size U-shape bands (drawn as inverted parabolas via
        # MSO_SHAPE.ARC won't be ideal — use 3 mini schematic U's drawn
        # with 3 short navy lines per plant: descending-flat-ascending).
        def draw_sac(label, x_min, x_max, y_min, y_left, y_right,
                      color=GRAY):
            # left descending segment
            _add_arrow(slide, (x_min, y_left), (x_min + (x_max - x_min) // 3,
                       y_min), color=color, weight_pt=1.5, head=False)
            # right ascending segment
            _add_arrow(slide, (x_max - (x_max - x_min) // 3, y_min),
                       (x_max, y_right),
                       color=color, weight_pt=1.5, head=False)
            # flat bottom
            _add_arrow(slide,
                        (x_min + (x_max - x_min) // 3, y_min),
                        (x_max - (x_max - x_min) // 3, y_min),
                        color=color, weight_pt=1.5, head=False)
            # label above min
            _add_text(slide,
                       (x_min + x_max) // 2 - Inches(0.9),
                       y_min - Inches(0.45),
                       Inches(1.8), Inches(0.3),
                       label, size=12, italic=True, color=color,
                       font="Calibri", align=PP_ALIGN.CENTER)

        # Three SAC bands placed along x — minima fall as you go right
        # (economies of scale).
        w = AX_W
        b1_x_min = AX_L + int(0.05 * w); b1_x_max = AX_L + int(0.32 * w)
        b1_y_min = AX_T + int(0.55 * AX_H)
        b1_y_left  = AX_T + int(0.15 * AX_H)
        b1_y_right = AX_T + int(0.20 * AX_H)
        draw_sac("SAC  (small plant)", b1_x_min, b1_x_max,
                  b1_y_min, b1_y_left, b1_y_right)

        b2_x_min = AX_L + int(0.30 * w); b2_x_max = AX_L + int(0.62 * w)
        b2_y_min = AX_T + int(0.65 * AX_H)
        b2_y_left  = AX_T + int(0.25 * AX_H)
        b2_y_right = AX_T + int(0.30 * AX_H)
        draw_sac("SAC  (medium plant)", b2_x_min, b2_x_max,
                  b2_y_min, b2_y_left, b2_y_right)

        b3_x_min = AX_L + int(0.60 * w); b3_x_max = AX_L + int(0.95 * w)
        b3_y_min = AX_T + int(0.75 * AX_H)
        b3_y_left  = AX_T + int(0.35 * AX_H)
        b3_y_right = AX_T + int(0.40 * AX_H)
        draw_sac("SAC  (large plant)", b3_x_min, b3_x_max,
                  b3_y_min, b3_y_left, b3_y_right)

        # Lower-envelope LAC – piecewise across the three minima
        _add_arrow(slide,
                    (b1_x_min, b1_y_min - Inches(0.05)),
                    (b3_x_max, b3_y_min - Inches(0.05)),
                    color=GOLD, weight_pt=3.0, head=False)
        _add_text(slide, AX_R - Inches(2.6), b3_y_min - Inches(0.55),
                  Inches(2.5), Inches(0.3),
                  "LAC  =  lower envelope",
                  size=14, bold=True, italic=True, color=GOLD,
                  font="Calibri")

        _add_takeaway_bar(
            slide,
            "LR-AC is the lower envelope of all SAC curves:  pick the best plant for each Q",
            top=Inches(6.4), fill=GOLD, text_color=NAVY,
            width=Inches(12.0),
        )

    s = make_diagram_slide(
        prs, page_num=66,
        section_tag=SECTION_TAG_P2_LR,
        title="LR Average Cost is the Lower Envelope of SR Curves",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Classic envelope diagram. Each plant size has its own short-run "
        "AC curve – U-shaped, with a minimum at the output it was "
        "designed for. The long-run AC is the lower envelope of all "
        "these short-run curves: for each output level Q, you pick the "
        "plant size that produces it at the lowest AC. In this drawing, "
        "the LAC slopes down – meaning economies of scale: bigger plant "
        "→ lower AC at its respective optimum."
    ))


def slide_67(prs):
    """Economies of scale: definition + the three cases."""
    def draw(slide):
        # Question header
        _add_text(slide, MARGIN, Inches(1.85), RULE_W, Inches(0.55),
                  "What happens to long-run AC as output grows?",
                  size=24, italic=True, color=GRAY,
                  font="Calibri", align=PP_ALIGN.CENTER)

        # Three cases as horizontal cards
        case_w = Inches(4.0)
        case_h = Inches(2.0)
        gap = Inches(0.25)
        x0 = (SLIDE_W - case_w * 3 - gap * 2) // 2
        y0 = Inches(2.8)

        _add_filled_box(slide, x0, y0, case_w, case_h,
                         "Economies of Scale\n\nLAC FALLS with Q\n\n(bigger ⇒ cheaper / unit)",
                         fill=NAVY, text_color=WHITE,
                         size=18, bold=True)
        _add_filled_box(slide, x0 + (case_w + gap), y0, case_w, case_h,
                         "Constant Returns\n\nLAC is FLAT in Q\n\n(size doesn't matter)",
                         fill=GRAY, text_color=WHITE,
                         size=18, bold=True)
        _add_filled_box(slide, x0 + 2 * (case_w + gap), y0, case_w, case_h,
                         "Diseconomies of Scale\n\nLAC RISES with Q\n\n(too big to manage)",
                         fill=NAVY, text_color=WHITE,
                         size=18, bold=True)

        # Why bullet header
        _add_text(slide, MARGIN, Inches(5.05), RULE_W, Inches(0.4),
                  "Why?  Two big drivers of economies of scale:",
                  size=18, bold=True, color=NAVY,
                  font="Calibri", align=PP_ALIGN.CENTER)
        _add_text(slide, MARGIN, Inches(5.50), RULE_W, Inches(0.4),
                  "(1) input prices fall as you grow      (2) technology favours larger scale (increasing returns)",
                  size=16, italic=True, color=GRAY,
                  font="Calibri", align=PP_ALIGN.CENTER)

        _add_takeaway_bar(
            slide,
            "Economies of scale is a COST concept;  returns to scale is a TECHNOLOGY concept",
            top=Inches(6.4), fill=GOLD, text_color=NAVY,
            width=Inches(12.0),
        )

    s = make_diagram_slide(
        prs, page_num=67,
        section_tag=SECTION_TAG_P2_LR,
        title="Economies of Scale:  Three Possible Patterns",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Economies of scale describes what happens to long-run average "
        "cost as output rises. Three cases: falling (EoS), flat "
        "(constant), rising (DEoS). Drivers are either input-price "
        "effects – you get bulk discounts, or you bid up prices when "
        "you're huge – or pure technology effects. Important: "
        "'economies of scale' is a COST-side concept, while 'returns to "
        "scale' (a textbook term) is a TECHNOLOGY concept. The two "
        "overlap but aren't the same. A firm can have constant returns "
        "to scale but enjoy economies of scale from bulk pricing."
    ))


def slide_68(prs):
    """Technological reasons for economies of scale."""
    bullets = [
        ("Specialisation and division of labor", 0),
        ("E.g., Ford assembly line vs. one craftsman building the whole car", 1),
        ("Lumpiness / indivisibilities of inputs", 0),
        ("R&D, network infrastructure, brand investment", 1),
        ("Same fixed asset serves more customers as scale grows", 1),
        ("Geometry — volume scales faster than surface", 0),
        ("Cargo ship, aircraft fuselage, pipeline:  capacity grows with r²", 1),
        ("Doubling materials more than doubles useful capacity", 1),
        ("Appropriate technology shifts with scale", 0),
        ("Big firms can run dedicated lines, automation, AI/data infra", 1),
    ]
    s = make_content_bulleted(
        prs, page_num=68,
        section_tag=SECTION_TAG_P2_LR,
        title="Technological Reasons for Economies of Scale",
        bullets=bullets,
        size=22, sub_size=17, line_spacing_pts=10,
    )
    _set_notes(s, (
        "Why does technology often favour larger scale? Four classic "
        "drivers. Specialisation – Adam Smith's pin factory – workers "
        "get better at narrow tasks. Lumpiness – you can't build half a "
        "network or half an R&D lab, so the fixed cost is the same "
        "whether you have 1M or 100M users. Geometry – for cylinders "
        "and tanks, volume grows faster than the surface area you have "
        "to build, so cost per unit of capacity drops with size. And "
        "scale unlocks DIFFERENT technologies entirely: only Amazon-"
        "scale firms can justify their own AI infrastructure or "
        "fulfilment robotics."
    ))


def slide_69(prs):
    """Embraer ERJ-145 vs. Boeing 787 – scale economies in aviation."""
    def draw(slide):
        # Two airplane cards side by side
        card_w = Inches(6.0)
        card_h = Inches(4.4)
        gap = Inches(0.3)
        x_l = (SLIDE_W - card_w * 2 - gap) // 2
        x_r = x_l + card_w + gap
        y0 = Inches(1.95)

        # Left card – Embraer
        _add_outlined_box(slide, x_l, y0, card_w, card_h,
                          "", fill=WHITE, line=NAVY, line_w=1.5)
        _add_text(slide, x_l, y0 + Inches(0.1), card_w, Inches(0.4),
                  "Embraer ERJ-145  ·  Regional Jet",
                  size=18, bold=True, color=NAVY,
                  font="Calibri", align=PP_ALIGN.CENTER)
        _add_source_image(slide, 69, "rId1",
                          left=x_l + Inches(0.3), top=y0 + Inches(0.55),
                          width=card_w - Inches(0.6))
        # Stats
        stats_y = y0 + Inches(2.4)
        stat_lines = [
            "List price:    ~ $25 M",
            "Seats:           50 passengers",
            "Cost / flight-hour:    ~ $1,400",
            "Cost / passenger-hour:   ≈ $28",
        ]
        for i, line in enumerate(stat_lines):
            _add_text(slide, x_l + Inches(0.4),
                      stats_y + Inches(0.4) * i,
                      card_w - Inches(0.8), Inches(0.35),
                      line, size=16, color=NAVY, font="Calibri")

        # Right card – Boeing 787
        _add_outlined_box(slide, x_r, y0, card_w, card_h,
                          "", fill=WHITE, line=NAVY, line_w=1.5)
        _add_text(slide, x_r, y0 + Inches(0.1), card_w, Inches(0.4),
                  "Boeing 787-9  ·  Wide-Body",
                  size=18, bold=True, color=NAVY,
                  font="Calibri", align=PP_ALIGN.CENTER)
        _add_source_image(slide, 69, "rId2",
                          left=x_r + Inches(0.3), top=y0 + Inches(0.55),
                          width=card_w - Inches(0.6))
        stat_lines_r = [
            "List price:    ~ $290 M",
            "Seats:           ~ 290 passengers",
            "Cost / flight-hour:    ~ $9,000",
            "Cost / passenger-hour:   ≈ $31",
        ]
        for i, line in enumerate(stat_lines_r):
            _add_text(slide, x_r + Inches(0.4),
                      stats_y + Inches(0.4) * i,
                      card_w - Inches(0.8), Inches(0.35),
                      line, size=16, color=NAVY, font="Calibri")

        _add_takeaway_bar(
            slide,
            "Bigger plane  →  similar (or lower) cost per passenger-hour:  geometry + load",
            top=Inches(6.45), fill=GOLD, text_color=NAVY,
            width=Inches(12.0),
        )

    s = make_diagram_slide(
        prs, page_num=69,
        section_tag=SECTION_TAG_P2_LR,
        title="Economies of Scale in Aviation:  ERJ-145 vs. 787",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Concrete example. Both aircraft burn fuel, pay pilots, pay "
        "landing fees – many of those costs scale with the airframe, "
        "not with passenger count. Spread across far more seats, the "
        "big plane achieves a similar or lower cost per passenger-hour. "
        "(Numbers are illustrative, based on FAA operating-cost data.) "
        "This is why long-haul routes use wide-body and short regional "
        "hops use small jets: the economics of scale depend on the load."
    ))


def slide_70(prs):
    """Reasons for diseconomies of scale."""
    bullets = [
        ("Coordination, communication, control  get harder", 0),
        ("More layers of management between strategy and the line", 1),
        ("Information has to travel up and down a longer chain", 1),
        ("Monitoring a bigger workforce is disproportionately costly", 0),
        ("Misaligned incentives, free-riding, principal-agent problems", 1),
        ("Bureaucracy and staff functions grow super-linearly", 0),
        ("HR, accounting, legal, compliance hire to support scale", 1),
        ("Real-world signals", 0),
        ("Boeing's recent quality issues at scale  (2024-25)", 1),
        ("Big-tech reorgs to break ranks into smaller, accountable units", 1),
    ]
    s = make_content_bulleted(
        prs, page_num=70,
        section_tag=SECTION_TAG_P2_LR,
        title="Reasons for Diseconomies of Scale",
        bullets=bullets,
        size=22, sub_size=17, line_spacing_pts=10,
    )
    _set_notes(s, (
        "Why bigger isn't always cheaper. Coordination costs explode "
        "with size: longer reporting chains, more meetings, more "
        "alignment overhead. Monitoring a 50-person team is one thing; "
        "monitoring 5,000 is qualitatively different. Bureaucracy "
        "scales super-linearly – legal, HR, compliance all expand "
        "faster than headcount. Recent business-press examples drive "
        "this home: Boeing's quality issues at scale, and the wave of "
        "big-tech reorganisations explicitly framed around 'getting "
        "back to small-team velocity'."
    ))


def slide_71(prs):
    """Economies of scope – producing 2+ products together."""
    def draw(slide):
        # Definition box at top
        _add_filled_box(slide, MARGIN, Inches(1.95),
                         RULE_W, Inches(0.9),
                         "Economies of scope:   producing 2+ related products together is cheaper than separately",
                         fill=NAVY, text_color=WHITE,
                         size=18, bold=True)

        # Sub-bullet drivers (left side)
        bullets = [
            ("Shared input production", 0),
            ("Shared engineering know-how & R&D", 0),
            ("Shared brand, sales channel, marketing", 0),
            ("Shared supply chain & customer data", 0),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(3.15),
            width=Inches(6.5), height=Inches(2.8),
            items=bullets,
            size=20, line_spacing_pts=10,
        )

        # Example image on right – Airbus A380 + A318
        _add_source_image(slide, 71, "rId1",
                          left=Inches(7.5), top=Inches(3.15),
                          width=Inches(5.5))
        _add_text(slide, Inches(7.5), Inches(5.85),
                  Inches(5.5), Inches(0.25),
                  "British Airways A380 + A318  (CC BY-SA, Wikimedia)",
                  size=11, italic=True, color=GRAY,
                  font="Calibri", align=PP_ALIGN.CENTER)
        _add_text(slide, Inches(7.5), Inches(3.0),
                  Inches(5.5), Inches(0.25),
                  "Example: Airbus's A380 and A318 share engineering, supply chain, brand",
                  size=12, italic=True, color=NAVY,
                  font="Calibri", align=PP_ALIGN.CENTER)

        _add_takeaway_bar(
            slide,
            "Scope ≠ scale:  one firm, MANY products  →  cheaper than splitting them up",
            top=Inches(6.5), fill=GOLD, text_color=NAVY,
            width=Inches(11.5),
        )

    s = make_diagram_slide(
        prs, page_num=71,
        section_tag=SECTION_TAG_P2_LR,
        title="Economies of Scope:  Cheaper Together than Apart",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Scope is the cousin of scale – but about the BREADTH of "
        "products a firm makes, not the depth of any one. Same firm, "
        "many products, sharing capabilities that are expensive to "
        "build: engineering knowledge, supply chain, brand. Airbus "
        "designed the A380 and A318 with shared engineering DNA – "
        "cockpit commonalities, pilot training, supplier base. Apple "
        "does it across iPhone/iPad/Mac (shared silicon, OS, retail). "
        "Recognising scope is what separates a firm that 'diversifies' "
        "into unrelated junk from one that genuinely lowers costs."
    ))


def slide_72(prs):
    """Amazon – scale, scope, or both? Discussion."""
    def draw(slide):
        _add_text(slide, MARGIN, Inches(1.95), RULE_W, Inches(0.55),
                  "Amazon is the textbook case for BOTH at once  —  where exactly?",
                  size=22, italic=True, bold=True, color=NAVY,
                  font="Calibri", align=PP_ALIGN.CENTER)

        # Two columns: Scale | Scope
        col_w = Inches(6.0)
        col_h = Inches(0.75)
        gap = Inches(0.3)
        x_l = (SLIDE_W - col_w * 2 - gap) // 2
        x_r = x_l + col_w + gap
        y_hdr = Inches(2.7)

        _add_filled_box(slide, x_l, y_hdr, col_w, col_h,
                         "Economies of SCALE",
                         fill=NAVY, text_color=WHITE,
                         size=22, bold=True)
        _add_filled_box(slide, x_r, y_hdr, col_w, col_h,
                         "Economies of SCOPE",
                         fill=NAVY, text_color=WHITE,
                         size=22, bold=True)

        scale_items = [
            "AWS — massive data-center fixed costs spread",
            "Fulfilment network density (FBA)",
            "Bargaining power with suppliers",
        ]
        scope_items = [
            "Prime  =  shipping + video + music + grocery",
            "Customer data shared across retail/ads/AWS",
            "Devices (Alexa, Kindle) lever the brand",
        ]
        y_items = y_hdr + col_h + Inches(0.25)
        for i, (l, r) in enumerate(zip(scale_items, scope_items)):
            _add_text(slide, x_l + Inches(0.2),
                      y_items + Inches(0.55) * i,
                      col_w - Inches(0.4), Inches(0.5),
                      "•  " + l,
                      size=17, color=NAVY, font="Calibri")
            _add_text(slide, x_r + Inches(0.2),
                      y_items + Inches(0.55) * i,
                      col_w - Inches(0.4), Inches(0.5),
                      "•  " + r,
                      size=17, color=NAVY, font="Calibri")

        _add_discussion_break(slide, top=Inches(6.55), width=Inches(5.0))

    s = make_diagram_slide(
        prs, page_num=72,
        section_tag=SECTION_TAG_P2_LR,
        title="Amazon:  Economies of Scale,  Scope,  or Both?",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Discussion prompt. Amazon is the canonical 'both' case. "
        "Scale: AWS is a fixed-cost-heavy business where capacity is "
        "shared across millions of customers. Fulfilment is denser the "
        "more orders flow through. Suppliers offer better terms at "
        "scale. Scope: Prime bundles shipping + video + music + "
        "grocery; advertising, devices, and AWS all benefit from "
        "customer data created in retail. The strategic question for "
        "students: which of these is most defensible, and which is "
        "just incremental?"
    ))


def slide_73(prs):
    """Shark Tank mini-case setup (video + group discussion)."""
    def draw(slide):
        # Setup line
        _add_text(slide, MARGIN, Inches(1.95), RULE_W, Inches(0.5),
                  "Watch first:   vimeo.com/236977187   (focus 4:50 – 5:40)",
                  size=20, italic=True, bold=True, color=NAVY,
                  font="Calibri", align=PP_ALIGN.CENTER)

        # Two PollEV question cards
        card_w = Inches(6.0)
        card_h = Inches(2.0)
        gap = Inches(0.3)
        x_l = (SLIDE_W - card_w * 2 - gap) // 2
        x_r = x_l + card_w + gap
        y0 = Inches(2.8)

        _add_outlined_box(slide, x_l, y0, card_w, card_h,
                          "PollEV  ·  Q1\n\nAre there economies of scale\nin this business?",
                          fill=WHITE, line=NAVY, text_color=NAVY,
                          size=20, bold=True, line_w=2.0)
        _add_outlined_box(slide, x_r, y0, card_w, card_h,
                          "PollEV  ·  Q2\n\nWhich deal would you choose?\n"
                          "$100K, royalty 25¢/can   vs.   $75K, 15% equity",
                          fill=WHITE, line=NAVY, text_color=NAVY,
                          size=18, bold=True, line_w=2.0)

        # Cue bullet under the cards
        _add_text(slide, MARGIN, Inches(5.1), RULE_W, Inches(0.5),
                  "Look for:  volume last year vs. this year,  and average cost per can",
                  size=18, italic=True, color=GRAY,
                  font="Calibri", align=PP_ALIGN.CENTER)
        _add_text(slide, MARGIN, Inches(5.55), RULE_W, Inches(0.5),
                  "Then estimate profit per can to compare the two deals on an apples-to-apples basis",
                  size=18, italic=True, color=GRAY,
                  font="Calibri", align=PP_ALIGN.CENTER)

        _add_discussion_break(slide, top=Inches(6.55), width=Inches(5.0))

    s = make_diagram_slide(
        prs, page_num=73,
        section_tag=SECTION_TAG_P2_LR,
        title="Mini-Case:  Shark Tank Pitch — Group Discussion",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Group discussion. Watch the pitch (Vimeo link), focus 4:50 – "
        "5:40 where they discuss sales volumes and costs. Two questions "
        "on PollEV. First: is this a business with economies of scale? "
        "Compare last year's volume and unit cost with this year's "
        "projection. Second: of the two deals the sharks offered, which "
        "is more founder-friendly? Comparison requires estimating "
        "profit per can."
    ))


def slide_74(prs):
    """Shark Tank mini-case – the numbers / solution."""
    def draw(slide):
        # Two columns: Volume/Cost evidence  |  Deal comparison
        col_w = Inches(6.0)
        gap = Inches(0.3)
        x_l = (SLIDE_W - col_w * 2 - gap) // 2
        x_r = x_l + col_w + gap
        y0 = Inches(1.95)

        # Left – volume + AC
        _add_filled_box(slide, x_l, y0, col_w, Inches(0.7),
                         "Volume & average cost",
                         fill=NAVY, text_color=WHITE,
                         size=20, bold=True)
        evid = [
            "Sales last year:     135K cans",
            "Sales this year:    300K cans   ↑",
            "Avg. cost last yr:   $1.30 / can",
            "Avg. cost this yr:   $1.10 / can   ↓",
            "⇒  AC falls as volume rises  =  Economies of Scale",
        ]
        for i, line in enumerate(evid):
            _add_text(slide, x_l + Inches(0.2),
                      y0 + Inches(0.85) + Inches(0.4) * i,
                      col_w - Inches(0.4), Inches(0.35),
                      line, size=16,
                      color=GOLD if line.startswith("⇒") else NAVY,
                      bold=line.startswith("⇒"),
                      font="Calibri")

        # Right – deal comparison
        _add_filled_box(slide, x_r, y0, col_w, Inches(0.7),
                         "Deals on a per-can basis",
                         fill=NAVY, text_color=WHITE,
                         size=20, bold=True)
        deals = [
            "Wholesale price:    $2.69 / can",
            "Profit per can:     $2.69 − $1.10  ≈  $1.50",
            "Deal A:  25¢ royalty / can  =  16.6% of profit",
            "Deal B:  15% equity",
            "⇒  Very similar in expected NPV terms",
        ]
        for i, line in enumerate(deals):
            _add_text(slide, x_r + Inches(0.2),
                      y0 + Inches(0.85) + Inches(0.4) * i,
                      col_w - Inches(0.4), Inches(0.35),
                      line, size=16,
                      color=GOLD if line.startswith("⇒") else NAVY,
                      bold=line.startswith("⇒"),
                      font="Calibri")

        _add_takeaway_bar(
            slide,
            "EoS is real here.  Royalty and equity deals look similar  —  pick based on control & risk",
            top=Inches(6.45), fill=GOLD, text_color=NAVY,
            width=Inches(12.5),
        )

    s = make_diagram_slide(
        prs, page_num=74,
        section_tag=SECTION_TAG_P2_LR,
        title="Shark Tank Solution:  Scale + Deal Comparison",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Reveal. The numbers show a clear AC drop – $1.30 → $1.10 per "
        "can – as volume more than doubles. That's economies of scale "
        "in the can-production business. Wholesale price $2.69 leaves "
        "about $1.50 profit per can. Deal A (25¢ royalty) is roughly "
        "16.6% of profit per can; Deal B (15% equity) is structurally "
        "similar in expected value. The real decision then turns on "
        "non-cash factors: control, signalling, dilution path, and the "
        "founder's view on probability of upside."
    ))


# --------------------------------------------------------------------------
# Layout-stripping surgery (kept from previous version)
# --------------------------------------------------------------------------

KEEP_LAYOUT = 'slideLayout7.xml'
LAYOUT_DISPLAY_NAME = '405 Slides Layout'


def strip_unused_layouts(pptx_path: Path):
    src = pptx_path
    tmp = pptx_path.with_suffix(pptx_path.suffix + '.tmp')

    with zipfile.ZipFile(src, 'r') as zin:
        names = zin.namelist()

        layouts_to_drop = []
        for n in names:
            if n.startswith('ppt/slideLayouts/') and n.endswith('.xml'):
                fname = n.rsplit('/', 1)[-1]
                if fname != KEEP_LAYOUT:
                    layouts_to_drop.append(n)
            elif n.startswith('ppt/slideLayouts/_rels/') and n.endswith('.xml.rels'):
                fname = n.rsplit('/', 1)[-1].replace('.rels', '')
                if fname != KEEP_LAYOUT:
                    layouts_to_drop.append(n)

        drop_set = set(layouts_to_drop)

        master_rels_xml = zin.read('ppt/slideMasters/_rels/slideMaster1.xml.rels').decode('utf-8')
        rels_root = ET.fromstring(master_rels_xml.encode('utf-8'))
        REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

        dropped_rids = []
        for rel in rels_root.findall(f'{{{REL_NS}}}Relationship'):
            target = rel.get('Target', '')
            if 'slideLayouts/' in target:
                if not target.endswith(KEEP_LAYOUT):
                    dropped_rids.append(rel.get('Id'))
                    rels_root.remove(rel)

        new_master_rels = ET.tostring(
            rels_root, xml_declaration=True, encoding='UTF-8',
            standalone=True,
        ).decode('utf-8')

        master_xml = zin.read('ppt/slideMasters/slideMaster1.xml').decode('utf-8')
        for rid in dropped_rids:
            master_xml = re.sub(
                rf'<p:sldLayoutId\s+id="\d+"\s+r:id="{rid}"\s*/>',
                '',
                master_xml,
            )

        ct_xml = zin.read('[Content_Types].xml').decode('utf-8')
        for n in layouts_to_drop:
            if n.endswith('.xml'):
                part_name = '/' + n
                ct_xml = re.sub(
                    rf'<Override\s+PartName="{re.escape(part_name)}"\s+ContentType="[^"]*"\s*/>',
                    '',
                    ct_xml,
                )

        kept_layout_xml = zin.read(f'ppt/slideLayouts/{KEEP_LAYOUT}').decode('utf-8')
        kept_layout_xml = re.sub(
            r'<p:cSld\s+name="[^"]*"',
            f'<p:cSld name="{LAYOUT_DISPLAY_NAME}"',
            kept_layout_xml,
            count=1,
        )
        if f'name="{LAYOUT_DISPLAY_NAME}"' not in kept_layout_xml:
            kept_layout_xml = kept_layout_xml.replace(
                '<p:cSld>', f'<p:cSld name="{LAYOUT_DISPLAY_NAME}">', 1,
            )

        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                if n in drop_set:
                    continue
                if n == 'ppt/slideMasters/_rels/slideMaster1.xml.rels':
                    zout.writestr(n, new_master_rels)
                elif n == 'ppt/slideMasters/slideMaster1.xml':
                    zout.writestr(n, master_xml)
                elif n == '[Content_Types].xml':
                    zout.writestr(n, ct_xml)
                elif n == f'ppt/slideLayouts/{KEEP_LAYOUT}':
                    zout.writestr(n, kept_layout_xml)
                else:
                    zout.writestr(n, zin.read(n))

    shutil.move(str(tmp), str(src))


def build_deck():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # Front matter
    slide_1(prs)
    slide_2(prs)
    slide_3(prs)
    slide_4(prs)
    slide_5(prs)
    # Big-picture concept map – at page 6 (replaces old textual outline)
    slide_concept_map(prs)

    # Part 1 §1.1 Short Run
    slide_7(prs)
    slide_8(prs)
    slide_9(prs)
    slide_10(prs)
    slide_11(prs)
    slide_12(prs)
    slide_13(prs)
    slide_14(prs)
    slide_15(prs)
    slide_16(prs)
    slide_17(prs)
    slide_18(prs)
    slide_19(prs)
    slide_20(prs)
    slide_21(prs)
    slide_22(prs)

    # Part 1 §1.1b Wage Searchers
    slide_23(prs)
    slide_24(prs)
    slide_25(prs)
    slide_26(prs)
    slide_27(prs)
    slide_28(prs)
    slide_29(prs)

    # Part 1.2 Long Run (with section divider)
    slide_30(prs)
    slide_31(prs)
    slide_32(prs)
    slide_33(prs)
    slide_34(prs)
    slide_35(prs)
    slide_36(prs)
    slide_37(prs)
    slide_38(prs)
    slide_39(prs)
    slide_40(prs)
    slide_41(prs)

    # Part 2 section divider
    slide_42(prs)

    # Part 2 §2.1 Cost Concepts
    slide_43(prs)
    slide_44(prs)
    slide_45(prs)
    slide_46(prs)
    slide_47(prs)
    slide_48(prs)
    slide_49(prs)
    slide_50(prs)
    slide_51(prs)
    slide_52(prs)
    slide_53(prs)
    slide_54(prs)
    slide_55(prs)
    slide_56(prs)
    slide_57(prs)
    slide_58(prs)
    slide_59(prs)
    slide_60(prs)
    slide_61(prs)
    slide_62(prs)

    # Part 2 §2.2 Long-Run Costs & Economies of Scale
    slide_63(prs)
    slide_64(prs)
    slide_65(prs)
    slide_66(prs)
    slide_67(prs)
    slide_68(prs)
    slide_69(prs)
    slide_70(prs)
    slide_71(prs)
    slide_72(prs)
    slide_73(prs)
    slide_74(prs)

    out = OUT_DIR / "Module 3_clean.pptx"
    prs.save(out)
    strip_unused_layouts(out)
    return out


if __name__ == "__main__":
    out = build_deck()
    print(f"Wrote {out}")
