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
from pptx.dml.color import RGBColor
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
                        section_tag="Module 3 · Section Divider",
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
        section_tag=SECTION_TAG_FRONT,
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
        section_tag=SECTION_TAG_FRONT,
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
        section_tag="Module 3 · Big Picture",
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


def slide_6(prs):
    """Outline of Module 3 — single master outline (was old #6).

    Restructured per 2026-05-12 feedback: 4 main points (Economic Models of
    Production / Short-run / Long-run / Costs), each with sub-bullets.
    """
    bullets = [
        ("Economic Models of Production", 0),
        ("Production functions", 1),
        ("Short run vs. long run", 1),
        ("Short-run production decisions", 0),
        ("Diminishing marginal returns", 1),
        ("Optimal hiring decisions", 1),
        ("Wage searchers", 1),
        ("Long-run production decisions", 0),
        ("Optimal mix of inputs", 1),
        ("Costs", 0),
        ("Fixed, sunk, variable costs", 1),
        ("Marginal costs", 1),
        ("Long-run costs and economies of scale", 1),
    ]
    s = make_content_bulleted(
        prs, page_num=6,
        section_tag=SECTION_TAG_FRONT,
        title="Outline of Module 3",
        bullets=bullets,
        # 13 lines total – use 24/20 with tight spacing to fit while
        # staying readable on 2-up handouts.
        size=24, sub_size=20,
        line_spacing_pts=8,
    )
    _set_notes(s, (
        "One look at where the night is going. Part 1 – Production – starts "
        "with how output depends on inputs, then takes us through the short-"
        "run hiring rule and the long-run input-mix decision. Part 2 – Costs "
        "– separates the cost types that actually drive decisions, builds the "
        "marginal-cost concept, and finishes with economies of scale. We will "
        "come back to this slide briefly at each transition."
    ))


# --------------------------------------------------------------------------
# Batch 2 – Part 1 (Production) §1.1 Short Run: slides 7-22
# --------------------------------------------------------------------------

SECTION_TAG_P1 = "Module 3 · Part 1 · Production"
SECTION_TAG_P1_DIV = "Module 3 · Section Divider"


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


def slide_10(prs):
    """Tesla's production function: output of the Gigafactory per week.

    This is a 2D table showing output as a function of (#workers, #robots).
    We rebuild it as a clean python-pptx table.
    """
    def draw(slide):
        # Clean rebuild of the production-function table.
        # Data: rows = workers (1000s), cols = robots (count)
        data = [
            ["",        "0",   "20",  "40",  "60",  "80",  "100", "120", "140"],
            ["0",       0,     0,     0,     0,     0,     0,     0,     0],
            ["1,000",   100,   200,   300,   400,   480,   560,   600,   620],
            ["2,000",   180,   360,   540,   720,   840,   960,  1020,  1040],
            ["3,000",   240,   480,   720,   960,  1140,  1300,  1380,  1400],
            ["4,000",   280,   560,   840,  1120,  1340,  1540,  1620,  1640],
            ["5,000",   300,   600,   900,  1200,  1440,  1660,  1740,  1760],
            ["6,000",   310,   620,   930,  1240,  1490,  1720,  1800,  1820],
            ["7,000",   315,   630,   945,  1260,  1515,  1750,  1830,  1850],
        ]
        rows = len(data)
        cols = len(data[0])
        tbl_left = Inches(2.4)
        tbl_top = Inches(2.4)
        tbl_w = Inches(9.0)
        tbl_h = Inches(4.4)
        table_shape = slide.shapes.add_table(rows, cols, tbl_left, tbl_top,
                                              tbl_w, tbl_h)
        tbl = table_shape.table
        for r, row in enumerate(data):
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
                   "Output = cars per week.  Note: marginal product diminishes as L grows for any fixed K.",
                   size=13, italic=True, color=GRAY, font="Calibri")

    s = make_diagram_slide(
        prs, page_num=10,
        section_tag=SECTION_TAG_P1,
        title="Tesla's Production Function: Output of the Gigafactory per Week",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Concrete example: Tesla's Gigafactory weekly output as a function "
        "of workers, given a fixed number of robots. Look at the pattern: "
        "more workers, more output, but the gains per worker shrink as you "
        "move down any column. That's diminishing marginal product in action. "
        "Note: numbers are approximate – the point is the shape, not precision."
    ))


def slide_11(prs):
    """Plotting total output (image of the production-function curve)."""
    def draw(slide):
        # The source slide had four animation frames; use the final one (rId6)
        # as a single static picture, centred in the body region.
        _add_source_image(slide, 11, "rId6",
                           left=Inches(2.5), top=Inches(1.95),
                           width=Inches(8.3))
        _add_text(slide, MARGIN, Inches(6.65), RULE_W, Inches(0.3),
                   "Output rises but flattens as L grows — the visual signature of diminishing returns.",
                   size=14, italic=True, color=GRAY, font="Calibri",
                   align=PP_ALIGN.CENTER)

    s = make_diagram_slide(
        prs, page_num=11,
        section_tag=SECTION_TAG_P1,
        title="Plotting Total Output",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Same data as the previous slide, now plotted. The shape – rising "
        "but flattening – is the visual signature of diminishing returns. "
        "Memorize this curve shape; you'll see it everywhere."
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

        # Bottom takeaway: the rule
        _add_takeaway_bar(slide,
                           "Optimum:  L*  where  MRPL  =  w",
                           top=Inches(6.45), fill=NAVY, text_color=WHITE,
                           width=Inches(10.0))

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

SECTION_TAG_WAGE = "Module 3 · Part 1 · Wage Searchers"
SECTION_TAG_LR   = "Module 3 · Part 1 · Long Run"
SECTION_TAG_DIV  = "Module 3 · Section Divider"


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
            ("Assumes input prices w and p_K are constant", 1),
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
            ("Cost of one robot (per week):  p_K = $20,000", 1),
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
            ("At the current mix:  read MP_K and MP_L from the table", 0),
            ("Compute MP_K / p_K  and  MP_L / w", 0),
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
    """When prices change, the input mix shifts: robot tax & union wages."""
    def draw(slide):
        # Two-column comparison: tax on robots | union wages up
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

        # Left column header
        _add_filled_box(slide, left_x, y, col_w, Inches(0.7),
                         "The government introduces a high tax on robots",
                         fill=NAVY, text_color=WHITE, size=18, bold=True)
        # Left column body — outcome
        _add_filled_box(slide, left_x, y + Inches(0.75),
                         col_w, col_h - Inches(0.75),
                         "p_K rises  →  MP_K / p_K falls\n→  shift toward more labor",
                         fill=RGBColor(0xF4, 0xF1, 0xEA),
                         text_color=NAVY, size=20, bold=False)

        # Right column header
        _add_filled_box(slide, right_x, y, col_w, Inches(0.7),
                         "Labor unions demand significantly higher wages",
                         fill=NAVY, text_color=WHITE, size=18, bold=True)
        _add_filled_box(slide, right_x, y + Inches(0.75),
                         col_w, col_h - Inches(0.75),
                         "w rises  →  MP_L / w falls\n→  shift toward more automation",
                         fill=RGBColor(0xF4, 0xF1, 0xEA),
                         text_color=NAVY, size=20, bold=False)

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
    slide_6(prs)

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

    out = OUT_DIR / "Module 3_clean.pptx"
    prs.save(out)
    strip_unused_layouts(out)
    return out


if __name__ == "__main__":
    out = build_deck()
    print(f"Wrote {out}")
