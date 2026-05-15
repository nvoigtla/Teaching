"""
Build a clean Module 3 deck from scratch, using ONLY the six template layouts
defined in `_build_template_samples.py`.

Goal: every slide in this deck uses one of six layout types (title, section
header, content bulleted, content two-column, poll, closing synthesis), all
on the Blank layout, so PowerPoint's Layout dropdown stays clean.

Build is by batches – front matter (1-6), then §1.1 Short Run (7-22), etc.

Output: `Module 3_clean.pptx`
"""

import math
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


def _add_convention_box(slide, left, top, width, height, *,
                          prefix=None, body=None, runs=None,
                          fill_rgb=None, border=None, line_w=1.0,
                          corner_pct=0.12, size=15, align=PP_ALIGN.LEFT,
                          font="Calibri", pad_h=None, pad_v=None):
    """Cream-fill / navy-border rounded-rect explanation callout.

    The "Convention" textbox pattern from slide 14 generalised — use it
    anywhere a slide needs a compact, visually-distinct box for a
    short conceptual explanation or notational convention.  Sits well
    below a table, beside a hero formula, or as a slide-wide footer.

    Two ways to populate the text:
      • ``prefix`` (bold) + ``body`` (regular) — simplest path; matches
        slide 14's "Convention:  <text>" pattern.
      • ``runs`` — a list of ``(text, {"bold": .., "italic": .., ...})``
        tuples for finer-grained styling (multi-line, mixed formatting).

    Style defaults follow the course-layer CLAUDE.md "Convention callout
    box" spec — cream fill, thin primary-color border, slight rounding,
    primary-color text.  Override ``fill_rgb`` / ``border`` only when you
    need a different accent.
    """
    fill = fill_rgb if fill_rgb is not None else RGBColor(0xFD, 0xF6, 0xE6)
    border = border if border is not None else NAVY

    left, top, width, height = int(left), int(top), int(width), int(height)
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height,
    )
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.color.rgb = border
    shp.line.width = Pt(line_w)
    shp.shadow.inherit = False
    try:
        shp.adjustments[0] = corner_pct
    except Exception:
        pass

    # Inset text box so the rounded corners breathe — matches slide 14.
    pad_h = Inches(0.20) if pad_h is None else pad_h
    pad_v = Inches(0.12) if pad_v is None else pad_v
    tb = slide.shapes.add_textbox(
        left + int(pad_h), top + int(pad_v),
        width - 2 * int(pad_h), height - 2 * int(pad_v),
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0); tf.margin_bottom = Inches(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    def _style_run(r, opts):
        r.font.name = opts.get('font', font)
        r.font.size = Pt(opts.get('size', size))
        r.font.bold = opts.get('bold', False)
        r.font.italic = opts.get('italic', False)
        r.font.color.rgb = opts.get('color', NAVY)

    if runs is not None:
        first = True
        for entry in runs:
            text, opts = entry if isinstance(entry, tuple) else (entry, {})
            if opts.get('newline') and not first:
                p = tf.add_paragraph()
            elif first:
                p = tf.paragraphs[0]
            else:
                # Same paragraph — append run to the most-recent paragraph.
                p = tf.paragraphs[-1]
            p.alignment = align
            r = p.add_run()
            r.text = text
            _style_run(r, opts)
            first = False
    else:
        p = tf.paragraphs[0]
        p.alignment = align
        if prefix:
            r1 = p.add_run(); r1.text = prefix
            _style_run(r1, {'bold': True, 'color': NAVY, 'size': size})
        if body:
            r2 = p.add_run(); r2.text = body
            _style_run(r2, {'color': NAVY, 'size': size})
    return shp


def _add_rounded_filled_box(slide, left, top, width, height, label, *,
                             fill=NAVY, text_color=WHITE, line=None,
                             size=18, bold=True, font="Calibri",
                             corner_pct=0.06, shadow=True):
    """Rounded-corner filled rectangle with centered text and soft drop shadow.

    Mirrors :func:`_add_filled_box` but renders ``MSO_SHAPE.ROUNDED_RECTANGLE``
    with the corner-adjust set to ``corner_pct`` (6 % per course CLAUDE.md
    "slight rounding") and a soft drop shadow via :func:`_add_drop_shadow`.
    """
    left, top, width, height = int(left), int(top), int(width), int(height)
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height,
    )
    try:
        shp.adjustments[0] = corner_pct
    except Exception:
        pass
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    if shadow:
        _add_drop_shadow(shp)
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
               head=True, dash=None):
    """Draw a line/arrow from start to end (in EMU/Inches values).

    EMU coordinates MUST be integers — PowerPoint rejects decimal values
    in <a:off>/<a:ext> and refuses to open the file. Cast to int defensively.

    ``dash`` accepts any OOXML preset-dash name (e.g., ``"dash"``,
    ``"dashDot"``, ``"sysDash"``).  Default ``None`` = solid line.
    """
    sx, sy = int(start_xy[0]), int(start_xy[1])
    ex, ey = int(end_xy[0]), int(end_xy[1])
    line = slide.shapes.add_connector(1, sx, sy, ex, ey)  # 1 = STRAIGHT
    line.line.color.rgb = color
    line.line.width = Pt(weight_pt)
    ln = line.line._get_or_add_ln()
    if dash is not None:
        for old in ln.findall(qn('a:prstDash')):
            ln.remove(old)
        prst = ET.SubElement(ln, qn('a:prstDash'))
        prst.set('val', dash)
    if head:
        tailEnd = ET.SubElement(ln, qn('a:tailEnd'))
        tailEnd.set('type', 'triangle')
        tailEnd.set('w', 'med')
        tailEnd.set('h', 'med')
    return line


def _add_wavy_line(slide, x_start, x_end, y_center, *,
                    amplitude=None, cycles=1.75, segments=36,
                    color=NAVY, weight_pt=1.5):
    """Horizontal sinusoidal line from x_start to x_end at y_center.

    Renders a polyline approximation of ``sin(2π · t · cycles)`` inside a
    custGeom shape so the line reads as a gentle wave rather than a
    straight connector.  ``amplitude`` is the peak-to-baseline height in
    EMU; defaults to ~0.04".  ``segments`` controls how finely the wave
    is discretised — 30+ is smooth enough that the polyline reads as a
    curve.  No arrowhead.
    """
    if amplitude is None:
        amplitude = Inches(0.04)
    L = int(x_end - x_start)
    A = int(amplitude)
    if L == 0:
        return None
    flip = "1" if L < 0 else "0"
    bbox_left = int(min(x_start, x_end))
    bbox_top = int(y_center - A)
    bbox_w = abs(L)
    bbox_h = 2 * A

    pts = []
    for i in range(segments + 1):
        t = i / segments
        lx = int(round(t * 100000))
        sin_val = math.sin(2 * math.pi * t * cycles)
        # Path coords: y=0 is top.  Centre at 50000; +1·amplitude → top
        # (0), −1·amplitude → bottom (100000).
        ly = int(round(50000 - sin_val * 50000))
        pts.append((lx, ly))

    path_segs = [f'<a:moveTo><a:pt x="{pts[0][0]}" y="{pts[0][1]}"/></a:moveTo>']
    for lx, ly in pts[1:]:
        path_segs.append(f'<a:lnTo><a:pt x="{lx}" y="{ly}"/></a:lnTo>')
    path_inner = ''.join(path_segs)

    color_hex = f'{color[0]:02X}{color[1]:02X}{color[2]:02X}'
    width_emu = int(weight_pt * 12700)

    P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    A_NS_LOCAL = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    sp_xml = (
        f'<p:sp xmlns:p="{P_NS}" xmlns:a="{A_NS_LOCAL}">'
        f'<p:nvSpPr><p:cNvPr id="0" name="WavyLine"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm flipH="{flip}">'
        f'<a:off x="{bbox_left}" y="{bbox_top}"/>'
        f'<a:ext cx="{bbox_w}" cy="{bbox_h}"/>'
        f'</a:xfrm>'
        f'<a:custGeom>'
        f'<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        f'<a:rect l="0" t="0" r="0" b="0"/>'
        f'<a:pathLst>'
        f'<a:path w="100000" h="100000" fill="none">'
        f'{path_inner}'
        f'</a:path>'
        f'</a:pathLst>'
        f'</a:custGeom>'
        f'<a:noFill/>'
        f'<a:ln w="{width_emu}" cap="rnd">'
        f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>'
        f'</a:ln>'
        f'</p:spPr>'
        f'</p:sp>'
    )
    elem = ET.fromstring(sp_xml)
    slide.shapes._spTree.append(elem)
    return elem


def _add_arrow_shape(slide, left, top, width, height, *,
                     direction="right", fill=GOLD, line=None):
    """Block arrow shape (the 'we-are-here' indicator).

    direction: "right" (default), "left", "up", or "down".
    """
    left, top, width, height = int(left), int(top), int(width), int(height)
    geom_map = {
        "left": MSO_SHAPE.LEFT_ARROW,
        "right": MSO_SHAPE.RIGHT_ARROW,
        "up": MSO_SHAPE.UP_ARROW,
        "down": MSO_SHAPE.DOWN_ARROW,
    }
    geom = geom_map.get(direction, MSO_SHAPE.RIGHT_ARROW)
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
                          sub_line_spacing_pts=None,
                          extras=None, bullets_top=None):
    """bullets: list of (text, level) tuples OR plain strings (level=0).

    ``bullets_top`` overrides the default body-region start (Inches(1.85))
    — useful when the slide also hosts large diagrams below the bullets
    and the bullets need to lift up to avoid overlap.

    ``sub_line_spacing_pts`` overrides the legacy
    ``max(6, line_spacing_pts - 8)`` formula for sub-bullet space-before.
    """
    slide = _blank_slide(prs)
    _draw_top_bar_tc(slide, section_tag)
    _draw_action_title(slide, title)

    normalized = [(b, 0) if isinstance(b, str) else b for b in bullets]

    if bullets_top is None:
        bullets_top = Inches(1.85)

    _add_hierarchical_bullets(
        slide,
        left=MARGIN,
        top=bullets_top,
        width=RULE_W,
        height=Inches(5.0),
        items=normalized,
        size=size,
        sub_size=sub_size,
        line_spacing_pts=line_spacing_pts,
        sub_line_spacing_pts=sub_line_spacing_pts,
    )

    if extras is not None:
        extras(slide)

    _draw_footer(slide, FOOTER_TEXT, page_num)
    return slide


def _add_hierarchical_bullets(slide, left, top, width, height, items,
                              *, size=24, sub_size=None, line_spacing_pts=18,
                              sub_line_spacing_pts=None):
    """Render bullets with indent levels (0 = top-level navy ▪,
    1+ = sub-bullets smaller and grey with – marker).

    size: font size for level-0 bullets.
    sub_size: font size for level-1+ bullets (default: size - 4).
    sub_line_spacing_pts: ``spcBef`` (points) applied to sub-bullets.
        Default ``None`` → ``max(6, line_spacing_pts - 8)`` (legacy
        formula).  Pass 0 (or any non-negative int) to override.
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
            if level == 0:
                sp = line_spacing_pts
            elif sub_line_spacing_pts is not None:
                sp = sub_line_spacing_pts
            else:
                sp = max(6, line_spacing_pts - 8)
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


def _add_drop_shadow(shape, *, blur="50800", dist="38100",
                      direction="2700000", alpha="45000"):
    """Add a soft drop shadow to any shape that exposes spPr (pictures,
    rectangles, rounded rects).  Default: 4pt blur, 3pt offset, 45° down-
    right, 45% opacity black.  Used deck-wide for figures/boxes."""
    try:
        spPr = shape._element.spPr
    except AttributeError:
        # Fallback for shapes whose XML element exposes spPr only via find()
        spPr = shape._element.find(qn('p:spPr'))
    if spPr is None:
        return shape
    for old in spPr.findall(qn('a:effectLst')):
        spPr.remove(old)
    effLst = ET.SubElement(spPr, qn('a:effectLst'))
    outerShdw = ET.SubElement(effLst, qn('a:outerShdw'))
    outerShdw.set('blurRad', str(blur))
    outerShdw.set('dist', str(dist))
    outerShdw.set('dir', str(direction))
    outerShdw.set('algn', 'tl')
    outerShdw.set('rotWithShape', '0')
    rgb = ET.SubElement(outerShdw, qn('a:srgbClr'))
    rgb.set('val', '000000')
    a = ET.SubElement(rgb, qn('a:alpha'))
    a.set('val', str(alpha))
    return shape


def _add_graphicframe_shadow(slide, left, top, width, height, *,
                              shadow_alpha=45000):
    """White backing rectangle with an outerShdw effect, behind a
    graphicFrame (table or chart).  graphicFrames can't host
    a:effectLst directly; this rect supplies the shadow projected
    OUTSIDE its bounds.

    Charts have transparent plot areas by default, so a coloured backing
    bleeds through and tints the figure.  Use white instead — the chart
    sees white through its own transparent areas (clean background) and
    the shadow renders only at the visible edges.  Call BEFORE adding
    the table/chart so z-order is correct.
    """
    shdw = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        int(left), int(top), int(width), int(height),
    )
    shdw.fill.solid()
    shdw.fill.fore_color.rgb = WHITE
    shdw.line.fill.background()
    shdw.shadow.inherit = False
    sp_pr = shdw._element.spPr
    # Strip any default <a:effectLst/> python-pptx may have inserted
    # (duplicate effectLst makes PowerPoint refuse to open the file).
    for old in sp_pr.findall(qn('a:effectLst')):
        sp_pr.remove(old)
    effLst = ET.SubElement(sp_pr, qn('a:effectLst'))
    outerShdw = ET.SubElement(effLst, qn('a:outerShdw'))
    outerShdw.set('blurRad', '50800')
    outerShdw.set('dist', '38100')
    outerShdw.set('dir', '2700000')
    outerShdw.set('algn', 'tl')
    outerShdw.set('rotWithShape', '0')
    rgb = ET.SubElement(outerShdw, qn('a:srgbClr'))
    rgb.set('val', '000000')
    a = ET.SubElement(rgb, qn('a:alpha'))
    a.set('val', str(int(shadow_alpha)))
    return shdw


def _add_source_image(slide, src_slide_no, rid, *, left, top, width=None,
                      height=None, shadow=True):
    """Place a source-deck image on the new slide.

    `shadow=True` (default) adds a soft drop shadow so figures pop off the
    background — applied deck-wide per the latest visual direction.  Set
    `shadow=False` for niche cases (transparent PNGs, screenshots that
    already include a shadow, etc.).
    """
    candidates = list(SRC_IMG_DIR.glob(f"slide{src_slide_no}_{rid}.*"))
    if not candidates:
        return None
    img = candidates[0]
    kwargs = {"left": int(left), "top": int(top)}
    if width is not None:
        kwargs["width"] = int(width)
    if height is not None:
        kwargs["height"] = int(height)
    pic = slide.shapes.add_picture(str(img), **kwargs)
    if shadow:
        _add_drop_shadow(pic)
    return pic


def _apply_picture_style(pic, *, corner_pct=8,
                          shadow_blur=50800, shadow_dist=38100,
                          shadow_dir=2700000, shadow_alpha=50000):
    """Apply rounded corners + drop shadow to a picture shape.

    corner_pct: rounded-corner radius as percent of shorter side (8 ≈ subtle).
    shadow_blur/dist: EMU; defaults give a soft 4pt blur, 3pt offset.
    shadow_dir: 2700000 = 45° down-right (standard).
    shadow_alpha: 50000 = 50% opacity black shadow.
    """
    spPr = pic._element.find(qn('p:spPr'))
    if spPr is None:
        return pic
    # Replace any existing prstGeom with roundRect at corner_pct
    for old in spPr.findall(qn('a:prstGeom')):
        spPr.remove(old)
    prstGeom = ET.Element(qn('a:prstGeom'))
    prstGeom.set('prst', 'roundRect')
    avLst = ET.SubElement(prstGeom, qn('a:avLst'))
    gd = ET.SubElement(avLst, qn('a:gd'))
    gd.set('name', 'adj')
    gd.set('fmla', f'val {int(corner_pct * 1000)}')
    # prstGeom must come after a:xfrm
    xfrm = spPr.find(qn('a:xfrm'))
    if xfrm is not None:
        xfrm.addnext(prstGeom)
    else:
        spPr.insert(0, prstGeom)
    # Replace any existing effectLst with a single outer shadow
    for old in spPr.findall(qn('a:effectLst')):
        spPr.remove(old)
    effectLst = ET.SubElement(spPr, qn('a:effectLst'))
    outerShdw = ET.SubElement(effectLst, qn('a:outerShdw'))
    outerShdw.set('blurRad', str(int(shadow_blur)))
    outerShdw.set('dist', str(int(shadow_dist)))
    outerShdw.set('dir', str(int(shadow_dir)))
    outerShdw.set('algn', 'tl')
    outerShdw.set('rotWithShape', '0')
    rgb = ET.SubElement(outerShdw, qn('a:srgbClr'))
    rgb.set('val', '000000')
    alpha = ET.SubElement(rgb, qn('a:alpha'))
    alpha.set('val', str(int(shadow_alpha)))
    return pic


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
                           text="Discussion Break"):
    """Rounded-parallelogram 'discussion break' badge (bottom-right).

    Custom-geometry shape: top and bottom edges are horizontal; the left
    and right edges slant at 45° in real space (skew = height of the
    shape).  All four corners are slightly rounded.  Gold fill, navy
    bold text, soft drop shadow.
    """
    height = Inches(0.72)
    left = SLIDE_W - MARGIN - width
    left, top, width, height = int(left), int(top), int(width), int(height)
    # Compute skew in path-coordinate units so that left/right sides slant
    # at 45° in REAL space:  horizontal offset of top edge == shape height.
    skew = int(100000 * height / width) if width else 15000
    skew = min(max(skew, 6000), 35000)
    r = 5000          # corner radius in path units — gentle rounding

    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = GOLD
    shp.line.fill.background()
    shp.shadow.inherit = False
    # Strip the default prstGeom – we'll inject custGeom instead.
    spPr = shp._element.spPr
    for old in spPr.findall(qn('a:prstGeom')):
        spPr.remove(old)
    rs = int(r * skew / 100000)
    # IMPORTANT: <a:rect> defines the TEXT bounding rectangle inside the
    # custom geometry.  Set it to the parallelogram's inscribed rectangle
    # (from TL vertex to BR vertex) so PowerPoint won't render text past
    # the slanted edges, regardless of the text frame's own lIns/rIns.
    custgeom_xml = (
        f'<a:custGeom xmlns:a="{A_NS}">'
        f'<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        f'<a:rect l="{skew}" t="0" r="{100000-skew}" b="100000"/>'
        f'<a:pathLst><a:path w="100000" h="100000">'
        # Top-left rounded corner (vertex at (skew, 0))
        f'<a:moveTo><a:pt x="{skew+r}" y="0"/></a:moveTo>'
        # Top edge
        f'<a:lnTo><a:pt x="{100000-r}" y="0"/></a:lnTo>'
        # Top-right corner round
        f'<a:cubicBezTo>'
        f'<a:pt x="100000" y="0"/><a:pt x="100000" y="0"/>'
        f'<a:pt x="{100000-rs}" y="{r}"/>'
        f'</a:cubicBezTo>'
        # Right slanted side (down-left)
        f'<a:lnTo><a:pt x="{100000-skew+rs}" y="{100000-r}"/></a:lnTo>'
        # Bottom-right corner round (vertex at (100000-skew, 100000))
        f'<a:cubicBezTo>'
        f'<a:pt x="{100000-skew}" y="100000"/>'
        f'<a:pt x="{100000-skew}" y="100000"/>'
        f'<a:pt x="{100000-skew-r}" y="100000"/>'
        f'</a:cubicBezTo>'
        # Bottom edge
        f'<a:lnTo><a:pt x="{r}" y="100000"/></a:lnTo>'
        # Bottom-left corner round (vertex at (0, 100000))
        f'<a:cubicBezTo>'
        f'<a:pt x="0" y="100000"/><a:pt x="0" y="100000"/>'
        f'<a:pt x="{rs}" y="{100000-r}"/>'
        f'</a:cubicBezTo>'
        # Left slanted side (up-right)
        f'<a:lnTo><a:pt x="{skew-rs}" y="{r}"/></a:lnTo>'
        # Top-left corner round (close the path)
        f'<a:cubicBezTo>'
        f'<a:pt x="{skew}" y="0"/><a:pt x="{skew}" y="0"/>'
        f'<a:pt x="{skew+r}" y="0"/>'
        f'</a:cubicBezTo>'
        f'<a:close/>'
        f'</a:path></a:pathLst>'
        f'</a:custGeom>'
    )
    custgeom = ET.fromstring(custgeom_xml)
    # Insert custGeom right after a:xfrm (schema order)
    xfrm = spPr.find(qn('a:xfrm'))
    if xfrm is not None:
        xfrm.addnext(custgeom)
    else:
        spPr.insert(0, custgeom)

    # Drop shadow (45° down-right, 50% opacity).
    for old in spPr.findall(qn('a:effectLst')):
        spPr.remove(old)
    effectLst = ET.SubElement(spPr, qn('a:effectLst'))
    outerShdw = ET.SubElement(effectLst, qn('a:outerShdw'))
    outerShdw.set('blurRad', '50800')
    outerShdw.set('dist', '38100')
    outerShdw.set('dir', '2700000')
    outerShdw.set('algn', 'tl')
    outerShdw.set('rotWithShape', '0')
    rgb = ET.SubElement(outerShdw, qn('a:srgbClr'))
    rgb.set('val', '000000')
    alpha = ET.SubElement(rgb, qn('a:alpha'))
    alpha.set('val', '50000')

    # Text — render as a SEPARATE textbox overlaid on top of the
    # parallelogram, positioned exactly inside the inscribed rectangle.
    # This decouples text placement from the shape geometry and avoids
    # PowerPoint rendering the run past the slanted edges (which can
    # happen with the in-shape text frame on some PowerPoint versions
    # even with <a:rect> set inside custGeom).
    # In real EMU, the skew equals the shape height (45° slant), so the
    # inscribed rectangle spans (left + height) → (left + width - height).
    skew_emu = height
    ins_left = left + skew_emu
    ins_top = top
    ins_w = width - 2 * skew_emu
    ins_h = height
    txt = slide.shapes.add_textbox(int(ins_left), int(ins_top),
                                     int(ins_w), int(ins_h))
    tf = txt.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.name = "Calibri"
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = NAVY
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
    # Soft drop shadow — added 2026-05-15 per user request, so the MB=MC
    # star reads as lifted off the slide like every other content shape.
    _add_drop_shadow(shp)
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


def _omml_fill(color):
    """Build the inner ``<a:solidFill>`` clause for an OMML run's rPr.

    ``color`` may be an ``RGBColor`` instance (or a 3-tuple of ints).
    Returns an empty string when ``color`` is None so callers can splice
    the result into rPr unconditionally.
    """
    if color is None:
        return ''
    return (
        f'<a:solidFill><a:srgbClr val="'
        f'{color[0]:02X}{color[1]:02X}{color[2]:02X}'
        f'"/></a:solidFill>'
    )


def _omml_run(text, *, color=None):
    """OMML run for an italic variable (default math style).

    Inside an oMath, italic style is the math default for Latin letters;
    we leave m:rPr out entirely so the Cambria Math italic comes through.
    The a:rPr applies drawing-level font sizing/coloring.  Pass ``color``
    to tint the run (e.g., green ΔL / ΔQ in the slide-14 Convention box).
    """
    return (
        f'<m:r xmlns:m="{M_NS}">'
        f'<a:rPr xmlns:a="{A_NS}" lang="en-US" b="0" i="1">'
        f'{_omml_fill(color)}'
        f'<a:latin typeface="Cambria Math"/>'
        f'<a:ea typeface="Cambria Math"/>'
        f'</a:rPr>'
        f'<m:t>{text}</m:t>'
        f'</m:r>'
    )


def _omml_text(text, *, color=None):
    """Upright-style OMML run (for operators, numbers, acronyms).

    Force plain (upright) style via <m:rPr><m:sty m:val="p"/></m:rPr> – this
    is the documented way to disable the math-default italics for the
    enclosed run.  Pass ``color`` to tint the run.
    """
    return (
        f'<m:r xmlns:m="{M_NS}">'
        f'<m:rPr><m:sty m:val="p"/></m:rPr>'
        f'<a:rPr xmlns:a="{A_NS}" lang="en-US" b="0" i="0">'
        f'{_omml_fill(color)}'
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


def _add_dashed_gridlines(axis_el):
    """Dashed light-grey major gridlines on an axis (Cobb-Douglas style)."""
    for old in axis_el.findall(qn('c:majorGridlines')):
        axis_el.remove(old)
    gl = ET.Element(qn('c:majorGridlines'))
    sp = ET.SubElement(gl, qn('c:spPr'))
    ln = ET.SubElement(sp, qn('a:ln'))
    ln.set('w', '9525')
    ln.set('cap', 'flat'); ln.set('cmpd', 'sng'); ln.set('algn', 'ctr')
    fill = ET.SubElement(ln, qn('a:solidFill'))
    clr = ET.SubElement(fill, qn('a:srgbClr')); clr.set('val', 'C8CDD3')
    dash = ET.SubElement(ln, qn('a:prstDash')); dash.set('val', 'dash')
    axpos = axis_el.find(qn('c:axPos'))
    if axpos is not None:
        axpos.addnext(gl)


def _make_simple_line_chart(slide, x, y, w, h, categories, values, *,
                              line_color, x_title, y_title,
                              y_min=0, y_max=None, y_unit=None,
                              marker='circle'):
    """Single-series line+markers chart with dashed light-grey gridlines.

    Same visual conventions as slide 11 (Calibri navy labels, dashed C8CDD3
    major gridlines, marker size 7) but no legend/title.
    """
    cd = CategoryChartData()
    cd.categories = list(categories)
    cd.add_series("Y", values)
    # Drop-shadow rectangle behind the chart (graphicFrames can't host shadow).
    _add_graphicframe_shadow(slide, x, y, w, h)
    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        int(x), int(y), int(w), int(h), cd,
    )
    chart = chart_shape.chart
    chart.has_title = False
    chart.has_legend = False

    clr_hex = f'{line_color[0]:02X}{line_color[1]:02X}{line_color[2]:02X}'

    for series in chart.series:
        line = series.format.line
        line.color.rgb = line_color
        line.width = Pt(2.5)
        ser_xml = series._element
        for old in ser_xml.findall(qn('c:marker')):
            ser_xml.remove(old)
        m = ET.SubElement(ser_xml, qn('c:marker'))
        sym = ET.SubElement(m, qn('c:symbol')); sym.set('val', marker)
        sz_el = ET.SubElement(m, qn('c:size')); sz_el.set('val', '7')
        sp = ET.SubElement(m, qn('c:spPr'))
        fl = ET.SubElement(sp, qn('a:solidFill'))
        rg = ET.SubElement(fl, qn('a:srgbClr')); rg.set('val', clr_hex)
        ln = ET.SubElement(sp, qn('a:ln'))
        lf = ET.SubElement(ln, qn('a:solidFill'))
        lr = ET.SubElement(lf, qn('a:srgbClr')); lr.set('val', clr_hex)
        # disable smoothing (straight segments between points)
        for sm in ser_xml.findall(qn('c:smooth')):
            ser_xml.remove(sm)
        smooth = ET.SubElement(ser_xml, qn('c:smooth'))
        smooth.set('val', '0')

    # Axes – axis titles in BOLD ITALIC navy (per course CLAUDE.md);
    # tick labels in regular Calibri navy.
    cat = chart.category_axis
    cat.tick_labels.font.name = "Calibri"
    cat.tick_labels.font.size = Pt(10)
    cat.tick_labels.font.color.rgb = NAVY
    cat.has_title = True
    cat.axis_title.text_frame.text = x_title
    ar = cat.axis_title.text_frame.paragraphs[0].runs[0]
    ar.font.name = "Calibri"; ar.font.size = Pt(12)
    ar.font.bold = True; ar.font.italic = True
    ar.font.color.rgb = NAVY

    val = chart.value_axis
    val.tick_labels.font.name = "Calibri"
    val.tick_labels.font.size = Pt(10)
    val.tick_labels.font.color.rgb = NAVY
    val.has_title = True
    val.axis_title.text_frame.text = y_title
    ar = val.axis_title.text_frame.paragraphs[0].runs[0]
    ar.font.name = "Calibri"; ar.font.size = Pt(12)
    ar.font.bold = True; ar.font.italic = True
    ar.font.color.rgb = NAVY
    if y_max is not None:
        val.minimum_scale = y_min
        val.maximum_scale = y_max
    if y_unit is not None:
        val.major_unit = y_unit

    _add_dashed_gridlines(cat._element)
    _add_dashed_gridlines(val._element)
    return chart_shape


def _make_multi_line_chart(slide, x, y, w, h, categories, series, *,
                             x_title, y_title,
                             y_min=0, y_max=None, y_unit=None,
                             legend=True, legend_pos=('0.08', '0.10', '0.20', '0.20')):
    """Multi-series line+markers chart with the deck's standard styling.

    series: list of (name, values, color: RGBColor, marker: str) tuples.
    legend_pos: (x, y, w, h) in chart-fraction units (str) – top-left default.
    """
    _add_graphicframe_shadow(slide, x, y, w, h)
    cd = CategoryChartData()
    cd.categories = list(categories)
    for name, values, _color, _marker in series:
        cd.add_series(name, values)
    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        int(x), int(y), int(w), int(h), cd,
    )
    chart = chart_shape.chart
    chart.has_title = False
    chart.has_legend = legend

    # Per-series styling: line color + marker
    for idx, ser in enumerate(chart.series):
        name, values, color, marker = series[idx]
        line = ser.format.line
        line.color.rgb = color
        line.width = Pt(2.5)
        ser_xml = ser._element
        clr_hex = f'{color[0]:02X}{color[1]:02X}{color[2]:02X}'
        for old in ser_xml.findall(qn('c:marker')):
            ser_xml.remove(old)
        m = ET.SubElement(ser_xml, qn('c:marker'))
        sym = ET.SubElement(m, qn('c:symbol')); sym.set('val', marker)
        sz_el = ET.SubElement(m, qn('c:size')); sz_el.set('val', '7')
        sp = ET.SubElement(m, qn('c:spPr'))
        fl = ET.SubElement(sp, qn('a:solidFill'))
        rg = ET.SubElement(fl, qn('a:srgbClr')); rg.set('val', clr_hex)
        ln = ET.SubElement(sp, qn('a:ln'))
        lf = ET.SubElement(ln, qn('a:solidFill'))
        lr = ET.SubElement(lf, qn('a:srgbClr')); lr.set('val', clr_hex)
        for sm in ser_xml.findall(qn('c:smooth')):
            ser_xml.remove(sm)
        smooth = ET.SubElement(ser_xml, qn('c:smooth'))
        smooth.set('val', '0')

    # Axes – bold italic navy titles per course style
    cat = chart.category_axis
    cat.tick_labels.font.name = "Calibri"
    cat.tick_labels.font.size = Pt(10)
    cat.tick_labels.font.color.rgb = NAVY
    cat.has_title = True
    cat.axis_title.text_frame.text = x_title
    ar = cat.axis_title.text_frame.paragraphs[0].runs[0]
    ar.font.name = "Calibri"; ar.font.size = Pt(12)
    ar.font.bold = True; ar.font.italic = True; ar.font.color.rgb = NAVY

    val = chart.value_axis
    val.tick_labels.font.name = "Calibri"
    val.tick_labels.font.size = Pt(10)
    val.tick_labels.font.color.rgb = NAVY
    val.has_title = True
    val.axis_title.text_frame.text = y_title
    ar = val.axis_title.text_frame.paragraphs[0].runs[0]
    ar.font.name = "Calibri"; ar.font.size = Pt(12)
    ar.font.bold = True; ar.font.italic = True; ar.font.color.rgb = NAVY
    if y_max is not None:
        val.minimum_scale = y_min
        val.maximum_scale = y_max
    if y_unit is not None:
        val.major_unit = y_unit

    _add_dashed_gridlines(cat._element)
    _add_dashed_gridlines(val._element)

    # Legend top-left inside plot with white fill
    if legend:
        leg_el = chart.legend._element
        chart.legend.font.name = "Calibri"
        chart.legend.font.size = Pt(11)
        chart.legend.font.color.rgb = NAVY
        chart.legend.include_in_layout = False
        # Strip default legendPos / layout, replace
        for old in leg_el.findall(qn('c:layout')):
            leg_el.remove(old)
        for old in leg_el.findall(qn('c:legendPos')):
            leg_el.remove(old)
        pos = ET.SubElement(leg_el, qn('c:legendPos')); pos.set('val', 'tr')
        leg_el.remove(pos); leg_el.insert(0, pos)
        # manualLayout positions in chart-fraction units
        layout = ET.Element(qn('c:layout'))
        ml = ET.SubElement(layout, qn('c:manualLayout'))
        ET.SubElement(ml, qn('c:xMode')).set('val', 'edge')
        ET.SubElement(ml, qn('c:yMode')).set('val', 'edge')
        ET.SubElement(ml, qn('c:x')).set('val', legend_pos[0])
        ET.SubElement(ml, qn('c:y')).set('val', legend_pos[1])
        ET.SubElement(ml, qn('c:w')).set('val', legend_pos[2])
        ET.SubElement(ml, qn('c:h')).set('val', legend_pos[3])
        pos.addnext(layout)
        # White fill behind legend
        for old in leg_el.findall(qn('c:spPr')):
            leg_el.remove(old)
        leg_spPr = ET.Element(qn('c:spPr'))
        sf = ET.SubElement(leg_spPr, qn('a:solidFill'))
        clr = ET.SubElement(sf, qn('a:srgbClr')); clr.set('val', 'FFFFFF')
        ln = ET.SubElement(leg_spPr, qn('a:ln')); ln.set('w', '6350')
        lf = ET.SubElement(ln, qn('a:solidFill'))
        lc = ET.SubElement(lf, qn('a:srgbClr')); lc.set('val', '0B2B4E')
        layout.addnext(leg_spPr)
    return chart_shape


def _omml_acc_overline(symbol):
    """Inline OMML accent (overline / bar) on a math symbol.

    symbol: e.g. 'K' or 'L' – the variable to wear the bar.
    Returns an OMML fragment intended to be embedded inside an <m:oMath>.
    """
    return (
        '<m:acc>'
          '<m:accPr><m:chr m:val="̅"/></m:accPr>'
          '<m:e>'
            '<m:r>'
              '<a:rPr lang="en-US" b="0" i="1">'
                '<a:latin typeface="Cambria Math"/>'
              '</a:rPr>'
              f'<m:t>{symbol}</m:t>'
            '</m:r>'
          '</m:e>'
        '</m:acc>'
    )


def _add_mixed_textbox(slide, left, top, width, height, segments, *,
                        align=PP_ALIGN.LEFT, default_color=NAVY,
                        default_size=24,
                        margin_left=None, margin_right=None,
                        margin_top=None, margin_bottom=None):
    """Build a textbox whose paragraphs mix plain text runs and inline OMML.

    segments: list of (kind, content, opts) tuples, with kind ∈ {"text",
    "omml", "break"}.  "break" inserts a new paragraph.  Opts may set
    `size`, `bold`, `italic`, `color`, `font` per run.
    """
    box = slide.shapes.add_textbox(int(left), int(top),
                                     int(width), int(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05) if margin_left is None else margin_left
    tf.margin_right = Inches(0.05) if margin_right is None else margin_right
    tf.margin_top = Inches(0.0) if margin_top is None else margin_top
    tf.margin_bottom = Inches(0.0) if margin_bottom is None else margin_bottom

    align_attr = ''
    if align == PP_ALIGN.CENTER: align_attr = ' algn="ctr"'
    elif align == PP_ALIGN.RIGHT: align_attr = ' algn="r"'

    def _start_para():
        return [f'<a:p xmlns:a="{A_NS}" xmlns:m="{M_NS}" xmlns:a14="{A14_NS}">',
                f'<a:pPr{align_attr}/>' if align_attr else '']

    paragraphs = [_start_para()]
    for kind, content, opts in segments:
        if kind == 'break':
            paragraphs[-1].append('<a:endParaRPr lang="en-US"/></a:p>')
            paragraphs.append(_start_para())
            continue
        size_pt = int(opts.get('size', default_size) * 100)
        color = opts.get('color', default_color)
        clr_hex = f'{color[0]:02X}{color[1]:02X}{color[2]:02X}'
        bold_attr = ' b="1"' if opts.get('bold') else ''
        italic_attr = ' i="1"' if opts.get('italic') else ''
        font = opts.get('font', 'Calibri')
        if kind == 'text':
            paragraphs[-1].append(
                f'<a:r><a:rPr lang="en-US" sz="{size_pt}"{bold_attr}{italic_attr}>'
                f'<a:solidFill><a:srgbClr val="{clr_hex}"/></a:solidFill>'
                f'<a:latin typeface="{font}"/>'
                f'</a:rPr><a:t>{content}</a:t></a:r>'
            )
        elif kind == 'omml':
            paragraphs[-1].append(
                f'<a14:m><m:oMath>{content}</m:oMath></a14:m>'
            )
    paragraphs[-1].append('<a:endParaRPr lang="en-US"/></a:p>')
    full_xml = ''.join(''.join(p) for p in paragraphs)

    txBody = tf._txBody
    for old in list(txBody.findall(qn('a:p'))):
        txBody.remove(old)
    # We have to parse each <a:p> separately so the namespaces resolve
    for p_str in full_xml.split('</a:p>'):
        if not p_str.strip(): continue
        p_xml = p_str + '</a:p>'
        new_p = ET.fromstring(p_xml)
        txBody.append(new_p)
        # Apply size+color to any OMML m:r elements inside this paragraph.
        # Color is set to ``default_color`` ONLY when no per-run solidFill
        # is already present — this lets callers tint individual OMML runs
        # via the optional ``color=`` argument on ``_omml_run`` / ``_omml_text``
        # (e.g., the green ΔL / ΔQ in the slide-14 Convention box) without
        # being silently overridden here.
        clr_hex = f'{default_color[0]:02X}{default_color[1]:02X}{default_color[2]:02X}'
        for r in new_p.iter(qn('m:r')):
            arPr = r.find(qn('a:rPr'))
            if arPr is None:
                arPr = ET.Element(qn('a:rPr'))
                arPr.set('lang', 'en-US')
                r.insert(0, arPr)
            arPr.set('sz', str(int(default_size * 100)))
            if arPr.find(qn('a:solidFill')) is None:
                sf = ET.SubElement(arPr, qn('a:solidFill'))
                srgb = ET.SubElement(sf, qn('a:srgbClr'))
                srgb.set('val', clr_hex)
    return box


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


def slide_announcements(prs):
    """Slide 3 – Announcements (midterm logistics).

    Reintroduced 2026-05-15 to mirror the original deck's slide 3.
    Dates left as ``{{MIDTERM_WINDOW}}`` / ``{{TA_WINDOW}}`` placeholders
    so the actual 2026 dates can be filled in later in PowerPoint.
    """
    bullets = [
        ("Midterm logistics", 0),
        ("3.5-hour window at home, any time during {{MIDTERM_WINDOW}}", 1),
        ("Guaranteed TA availability: {{TA_WINDOW}}", 1),
        ("Material covered", 0),
        ("All material from Modules 1 and 2 (includes PS 1 + 2)", 1),
        ("Problem-solving exercises similar to PS 1 + 2", 1),
        ("Review sessions during the midterm week", 0),
    ]
    s = make_content_bulleted(
        prs,
        page_num=3,
        section_tag="Module 3 · Announcements",
        title="Announcements",
        bullets=bullets,
        size=30, sub_size=26,
        line_spacing_pts=12,
    )
    _set_notes(s, (
        "Before we dive into Module 3 – two quick announcements about the "
        "midterm. It's a 6-hour at-home window over the dates shown; pick "
        "any contiguous block within that window. We'll have a TA on call "
        "for one guaranteed support window – use it if you'd like to ask "
        "live questions. Material covers Modules 1 and 2, including both "
        "problem sets. Review sessions will run during the week of the "
        "midterm; details by email."
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
        page_num=4,
        section_tag="Module 3 · Recap",
        title="Recap of Module 2",
        bullets=bullets,
        size=32, sub_size=26,
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
        _add_rounded_filled_box(slide, top_x, top_y, wide_w, box_h,
                                  "1. Basic Principles and Economic Way of Thinking",
                                  fill=FADED, text_color=WHITE, size=24, bold=True)

        # Row 2 – two parallel modules
        row2_y = Inches(3.65)
        left_x = slide_mid - gap // 2 - narrow_w
        right_x = slide_mid + gap // 2
        _add_rounded_filled_box(slide, left_x, row2_y, narrow_w, box_h,
                                  "2. Value and Demand",
                                  fill=FADED, text_color=WHITE, size=26, bold=True)
        # Current module (navy)
        _add_rounded_filled_box(slide, right_x, row2_y, narrow_w, box_h,
                                  "3. Supply and Cost",
                                  fill=NAVY, text_color=WHITE, size=26, bold=True)

        # Row 3 – bottom module (faded)
        bot_x = slide_mid - wide_w // 2
        bot_y = Inches(5.5)
        _add_rounded_filled_box(slide, bot_x, bot_y, wide_w, box_h,
                                  "4. Markets, Pricing, and Strategy",
                                  fill=FADED, text_color=WHITE, size=24, bold=True)

        # Connectors — top down to row 2 (faded grey lines).  Thicker
        # than before per user request 2026-05-15.
        top_bottom_y = top_y + box_h
        row2_top_y = row2_y
        _add_arrow(slide,
                    (top_x + wide_w // 2, top_bottom_y),
                    (left_x + narrow_w // 2, row2_top_y),
                    color=FADED, weight_pt=3.0, head=True)
        _add_arrow(slide,
                    (top_x + wide_w // 2, top_bottom_y),
                    (right_x + narrow_w // 2, row2_top_y),
                    color=NAVY, weight_pt=3.5, head=True)

        # Row 2 down to row 3
        row2_bottom_y = row2_y + box_h
        row3_top_y = bot_y
        _add_arrow(slide,
                    (left_x + narrow_w // 2, row2_bottom_y),
                    (bot_x + wide_w // 2, row3_top_y),
                    color=FADED, weight_pt=3.0, head=True)
        _add_arrow(slide,
                    (right_x + narrow_w // 2, row2_bottom_y),
                    (bot_x + wide_w // 2, row3_top_y),
                    color=FADED, weight_pt=3.0, head=True)

        # "We are here" — gold UP-pointing arrow positioned BELOW box 3,
        # head pointing INTO the bottom edge of the box.  Vertical budget
        # is tight: ~1.0" of clearance between row 2 bottom (4.50") and
        # row 3 top (5.50"), so the arrow + label must fit inside that.
        # Shifted right by ~0.55" per user request 2026-05-15.
        right_shift = Inches(0.55)
        arrow_w = Inches(0.55)
        arrow_h = Inches(0.55)
        arrow_left = right_x + (narrow_w - arrow_w) // 2 + right_shift
        arrow_top = row2_y + box_h + Inches(0.05)
        _add_arrow_shape(slide, arrow_left, arrow_top, arrow_w, arrow_h,
                          direction="up", fill=GOLD)
        # Label directly below the arrow, shifted with the arrow
        _add_text(slide, right_x + right_shift,
                   arrow_top + arrow_h + Inches(0.02),
                   narrow_w, Inches(0.32),
                   "we are here", size=16, italic=True, bold=True,
                   color=GOLD, font="Calibri", align=PP_ALIGN.CENTER)

    s = make_diagram_slide(
        prs, page_num=5,
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
        # Box wording restored to the original deck's slide 8 (single-line:
        # no parenthetical sub-text), so heights are shorter than before.
        box_h = Inches(1.05)
        small_h = Inches(1.05)
        m3_w = Inches(4.8)
        m2_w = Inches(4.2)
        m4_w = Inches(7.0)

        # Module 3 left stack
        m3_x = Inches(0.5)
        prod_y = Inches(2.05)
        costs_y = Inches(3.85)
        _add_rounded_filled_box(slide, m3_x, prod_y, m3_w, box_h,
                                  "Production Functions",
                                  fill=NAVY, text_color=WHITE,
                                  size=36, bold=True)
        _add_rounded_filled_box(slide, m3_x, costs_y, m3_w, box_h,
                                  "Costs",
                                  fill=NAVY, text_color=WHITE,
                                  size=36, bold=True)

        # Module 2 right (faded — already covered)
        m2_x = Inches(8.6)
        m2_y = Inches(2.05)
        _add_rounded_filled_box(slide, m2_x, m2_y, m2_w, small_h,
                                  "Demand",
                                  fill=FADED, text_color=WHITE,
                                  size=36, bold=True)

        # Module 4 bottom-center (gold — coming up)
        m4_x = (SLIDE_W - m4_w) // 2
        m4_y = Inches(5.8)
        _add_rounded_filled_box(slide, m4_x, m4_y, m4_w, box_h,
                                  "Output Decisions",
                                  fill=GOLD, text_color=WHITE,
                                  size=36, bold=True)

        # Arrows — thicker than before per user request 2026-05-15.
        # Production → Costs (vertical, inside the M3 stack)
        _add_arrow(slide,
                    (m3_x + m3_w // 2, prod_y + box_h),
                    (m3_x + m3_w // 2, costs_y),
                    color=NAVY, weight_pt=3.5, head=True)
        # Costs → Output Decisions (diagonal down to centre)
        _add_arrow(slide,
                    (m3_x + m3_w // 2, costs_y + box_h),
                    (int(m4_x + m4_w * 0.30), m4_y),
                    color=NAVY, weight_pt=3.5, head=True)
        # Demand → Output Decisions (diagonal down to centre-right)
        _add_arrow(slide,
                    (m2_x + m2_w // 2, m2_y + small_h),
                    (int(m4_x + m4_w * 0.70), m4_y),
                    color=FADED, weight_pt=3.0, head=True)

    s = make_diagram_slide(
        prs, page_num=6,
        section_tag="Module 3 · Production · Big Picture",
        title="Big Picture of Module 3",
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
                    color=NAVY, weight_pt=2.5, head=True)
        _add_arrow(slide,
                    (n1_x + n1_w // 2, n1_y + n1_h),
                    (lr_x + col_w // 2, sr_y),
                    color=NAVY, weight_pt=2.5, head=True)
        _add_arrow(slide,
                    (sr_x + col_w // 2, sr_y + col_h),
                    (sr_x + col_w // 2, rule_y),
                    color=NAVY, weight_pt=2.5, head=True)
        _add_arrow(slide,
                    (lr_x + col_w // 2, sr_y + col_h),
                    (lr_x + col_w // 2, rule_y),
                    color=NAVY, weight_pt=2.5, head=True)

        # Costs – C1 fans out into the three parallel children
        for cx in (c2_x, c3_x, c4_x):
            _add_arrow(slide,
                        (c1_x + c1_w // 2, c1_y + c1_h),
                        (cx + cw // 2, child_y),
                        color=NAVY, weight_pt=2.5, head=True)

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
                    color=GOLD, weight_pt=2.5, head=True)

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
                    color=GOLD, weight_pt=3.0, head=True)

        # ---- Outflow: bridge → cost cluster ----------------------------

        # Arrow departs from the "Minimum cost for any Q" bridge and
        # points up-right at the cost cluster (lands inside the cluster's
        # bottom-left area, NOT into any specific cost child).
        arrow_start = (bridge_x + bridge_w - Inches(0.3),
                        bridge_y + Inches(0.15))
        arrow_end = (c1_x + Inches(0.4), child_y + child_h)
        _add_arrow(slide, arrow_start, arrow_end,
                    color=GOLD, weight_pt=3.0, head=True)

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
                    color=GOLD, weight_pt=2.5, head=True)

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
                    color=NAVY, weight_pt=2.5, head=True)

    s = make_diagram_slide(
        prs, page_num=7,
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
        prs, page_num=8,
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
    """The Production Function: Q = f(K, L, etc)."""
    def draw(slide):
        # Big equation, centred near the top of the body region
        _add_text(slide, MARGIN, Inches(2.0), RULE_W, Inches(1.0),
                   "Q = f (K, L, etc)",
                   size=54, bold=True, color=NAVY, font="Calibri",
                   align=PP_ALIGN.CENTER)
        # Variable legend on the LEFT
        legend = [
            ("Q  =  Output", 0),
            ("f   =  a function of inputs:", 0),
            ("K  =  Capital  (physical: factories, machinery, software, IP)", 1),
            ("L  =  Labor", 1),
            ('"etc" can be raw materials, energy...', 1),
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
        # User decision: drop the drop-shadow here (book cover is a special
        # case — looks weird with shadow against its existing background).
        _add_source_image(slide, 8, "rId4",
                           left=Inches(10.3), top=Inches(3.3),
                           width=Inches(2.4),
                           shadow=False)
        _add_text(slide, Inches(10.3), Inches(5.85), Inches(2.4), Inches(0.25),
                   "Marx, Das Kapital  (1867)",
                   size=12, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        # Bottom explanation callout — Convention-style box (cream-fill
        # rounded rect with navy border).  Pattern documented in the
        # Teaching CLAUDE.md as the preferred format for concept-
        # explanation textboxes; mirrors the "Convention" box on slide 14.
        box_w = Inches(11.0)
        box_h = Inches(1.00)
        box_x = (SLIDE_W - box_w) // 2
        box_y = Inches(6.10)
        _add_convention_box(
            slide, box_x, box_y, box_w, box_h,
            runs=[
                ("A production function transforms inputs into outputs.",
                 {'size': 20, 'bold': True, 'color': NAVY}),
                ("The more efficient this process, the higher is productivity",
                 {'size': 20, 'color': NAVY, 'newline': True}),
            ],
            size=20, align=PP_ALIGN.CENTER,
        )

    s = make_diagram_slide(
        prs, page_num=9,
        section_tag=SECTION_TAG_P1,
        title="The Production Function",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "The production function in its most basic form: output Q is a "
        "function of capital K, labor L, and any other inputs the firm "
        "uses – raw materials, energy, and so on. Capital is more than "
        "buildings – it includes machinery, software, IP, AI systems – "
        "anything you've already paid for that keeps producing. The "
        "callout at the bottom captures the big idea: a production "
        "function maps inputs into outputs, and the more efficiently it "
        "does so, the higher is productivity. Everything else in the "
        "module is built on this expression."
    ))


def slide_9(prs):
    """In the short run, you're stuck with your capacity."""
    bullets = [
        "In the short run, some inputs (fixed factors) cannot be increased or decreased",
        "The long run is a period long enough for all inputs to be variable",
    ]

    def draw_pictures(slide):
        # Picture captions sit ABOVE the pictures (matches the original
        # slide 11's layout); attribution under the LEFT image stays
        # below the picture.
        CAP_TOP = Inches(3.30)
        PIC_TOP = Inches(3.65)
        PIC_W = Inches(5.0)
        PIC_H = Inches(3.0)
        # Captions — wording from the original deck's slide 11
        _add_text(slide, Inches(1.0), int(CAP_TOP), Inches(5.0), Inches(0.32),
                   "Capital fixed in short run",
                   size=14, italic=True, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        _add_text(slide, Inches(7.3), int(CAP_TOP), Inches(5.0), Inches(0.32),
                   "Labor (ophthalmologists) fixed in short run",
                   size=14, italic=True, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        # LEFT image: Rivian Normal IL assembly-plant floor (CC BY-SA,
        # Wikimedia).  Replaces the stale Tesla factory floor.
        rivian_plant = OUT_DIR / "_rivian_plant.jpg"
        if rivian_plant.exists():
            pic_left = slide.shapes.add_picture(
                str(rivian_plant),
                int(Inches(1.0)), int(PIC_TOP),
                width=int(PIC_W), height=int(PIC_H),
            )
            _apply_picture_style(pic_left)
        # RIGHT image: ophthalmologist photo from the original deck —
        # illustrates skilled labor as a short-run fixed factor (years of
        # specialised training mean the supply can't be ramped on demand).
        pic_right = _add_source_image(slide, 9, "rId5",
                           left=Inches(7.3), top=PIC_TOP,
                           width=PIC_W, height=PIC_H)
        if pic_right is not None:
            _apply_picture_style(pic_right)
        # Tiny attribution under the LEFT image (CC BY-SA author + license).
        attr_top = PIC_TOP + PIC_H + Inches(0.08)
        _add_text(slide, Inches(1.0), int(attr_top),
                   Inches(5.0), Inches(0.18),
                   "Rivian Normal, IL plant  (CC BY-SA, Wikimedia)",
                   size=9, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")

    s = make_content_bulleted(
        prs, page_num=10,
        section_tag=SECTION_TAG_P1,
        title="Short vs. Long Run:  A Critical Distinction",
        bullets=bullets,
        size=28, sub_size=22, line_spacing_pts=14,
        extras=draw_pictures,
    )
    # Shrink the bullet box so it doesn't overlap the images
    _set_notes(s, (
        "The single most important time-scale distinction in this course. "
        "Short run = some inputs are fixed factors — you cannot vary them "
        "within the planning horizon. Long run = a horizon long enough "
        "that EVERY input becomes variable. The two photos give parallel "
        "examples of short-run fixed factors. Left: Rivian's Normal, "
        "Illinois assembly plant — capital (the building, the robot fleet) "
        "is fixed in any given quarter; they can ramp shifts and headcount "
        "but not the four walls. Right: an ophthalmologist — skilled labor "
        "with many years of training is also a fixed factor in the short "
        "run; a hospital cannot conjure up another ophthalmologist on a "
        "month's notice. Both illustrate the same lesson: in the short "
        "run, you optimise around the factors you cannot change."
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

# --------------------------------------------------------------------------
# Rivian Georgia plant cost data (from Background Material/Module 3 - Make
# vs Buy.xlsx).  Quadratic cost function:  TC = TFC + 200 · Q²,  TFC = $800K.
# Q grid:  10, 20, …, 110 vehicles per week.  Drives the native cost charts
# on slides 55, 56, 57 so all three are consistent and locked to one source.
# --------------------------------------------------------------------------

COST_TFC = 800_000
COST_VAR_COEF = 200
COST_Q_VALS = list(range(10, 111, 10))

def _cost_tc(Q):  return COST_TFC + COST_VAR_COEF * Q * Q
def _cost_tvc(Q): return COST_VAR_COEF * Q * Q
def _cost_avc(Q): return COST_VAR_COEF * Q
def _cost_atc(Q): return _cost_tc(Q) / Q
def _cost_mc(Q, dQ=10):
    return (_cost_tc(Q + dQ) - _cost_tc(Q)) / dQ


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


def _add_compact_pf_table(slide, *, tbl_left, tbl_top, col_w_label=Inches(0.72),
                            col_w_data=Inches(0.55),
                            tbl_h=Inches(3.70),
                            font_size=11,
                            caption="Production-function table  (slide 10)",
                            with_axes=True):
    """Insert the compact production-function table (same data as slide 10),
    with a drop-shadow rect behind and optional K/L axis labels + caption.

    Returns the table_shape so callers can position related elements.
    """
    Q_t = _pf_table()
    header_row = [""] + [str(K) for K in PF_K_VALS]
    rows_data = [header_row]
    for ri, L in enumerate(PF_L_VALS):
        rows_data.append([f"{L:,}"] + [str(v) for v in Q_t[ri]])

    rows = len(rows_data); cols = len(rows_data[0])
    tbl_w = col_w_label + col_w_data * 4

    _add_graphicframe_shadow(slide, tbl_left, tbl_top, tbl_w, tbl_h)
    tshape = slide.shapes.add_table(rows, cols, int(tbl_left), int(tbl_top),
                                      int(tbl_w), int(tbl_h))
    tbl = tshape.table
    tbl.columns[0].width = col_w_label
    for c in range(1, cols):
        tbl.columns[c].width = col_w_data

    for r, row in enumerate(rows_data):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.margin_left = Inches(0.06)
            cell.margin_right = Inches(0.06)
            cell.margin_top = Inches(0.01)
            cell.margin_bottom = Inches(0.01)
            cell.text = str(val)
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(font_size)
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

    if caption:
        _add_text(slide,
                   tbl_left, tbl_top + tbl_h + Inches(0.08),
                   tbl_w, Inches(0.25),
                   caption,
                   size=11, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")

    if with_axes:
        _add_text(slide,
                   tbl_left + col_w_label, tbl_top - Inches(0.30),
                   col_w_data * 4, Inches(0.25),
                   "K  (robots)",
                   size=10, italic=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        _add_text(slide,
                   tbl_left - Inches(0.75), tbl_top + tbl_h / 2 - Inches(0.15),
                   Inches(0.70), Inches(0.30),
                   "L  (workers)",
                   size=10, italic=True, color=NAVY,
                   align=PP_ALIGN.RIGHT,
                   anchor=MSO_ANCHOR.MIDDLE, font="Calibri")
    return tshape


# --------------------------------------------------------------------------
# Slide-10 user-added "Number of cars" callout group.
#
# This XML was hand-built in PowerPoint (oval circling a table cell,
# rectangle label "Number of / cars" on the right, slanted line with
# arrowhead from label to oval) and copied verbatim so that running
# `_build_clean_deck.py` reproduces it identically. Do NOT regenerate
# from python-pptx primitives – the styling (Whitney-Book font,
# stealth arrowhead, drop shadow on the second-line text) is hard to
# recreate via the python-pptx API and is preserved here as-is.
# --------------------------------------------------------------------------

GROUP_XML_SLIDE10 = '''<p:grpSp><p:nvGrpSpPr><p:cNvPr id="101" name="Group 16"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{CCFFE88A-C1F2-920E-75A7-DC37624177EC}"/></a:ext></a:extLst></p:cNvPr><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="5669321" y="3150554"/><a:ext cx="5054639" cy="1236280"/><a:chOff x="3902753" y="3429000"/><a:chExt cx="5054639" cy="1236280"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="102" name="Oval 17"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{149BAB88-5F44-F462-B3EF-E90763362AA2}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="3902753" y="4279192"/><a:ext cx="969223" cy="386088"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="28575"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="103" name="Rectangle 8"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{64C69EA1-DAB2-A1C9-80EF-8CB4813BCBB1}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr><a:spLocks noChangeArrowheads="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr bwMode="auto"><a:xfrm><a:off x="7532631" y="3429000"/><a:ext cx="1424761" cy="422371"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="9525"><a:noFill/><a:miter lim="800000"/><a:headEnd/><a:tailEnd/></a:ln><a:effectLst/></p:spPr><p:txBody><a:bodyPr wrap="none" lIns="20048" tIns="28402" rIns="20048" bIns="28402"/><a:lstStyle/><a:p><a:pPr algn="ctr"><a:lnSpc><a:spcPts val="1700"/></a:lnSpc><a:tabLst><a:tab pos="481157" algn="l"/><a:tab pos="962315" algn="l"/><a:tab pos="1443472" algn="l"/></a:tabLst></a:pPr><a:r><a:rPr lang="en-US" b="0" dirty="0"><a:latin typeface="Whitney-Book" pitchFamily="50" charset="0"/><a:cs typeface="Whitney-Book" pitchFamily="50" charset="0"/></a:rPr><a:t>Number of</a:t></a:r></a:p><a:p><a:pPr algn="ctr"><a:lnSpc><a:spcPts val="1700"/></a:lnSpc><a:tabLst><a:tab pos="481157" algn="l"/><a:tab pos="962315" algn="l"/><a:tab pos="1443472" algn="l"/></a:tabLst></a:pPr><a:r><a:rPr lang="en-US" dirty="0"><a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="C0C0C0"/></a:outerShdw></a:effectLst><a:latin typeface="Whitney-Book" pitchFamily="50" charset="0"/><a:cs typeface="Whitney-Book" pitchFamily="50" charset="0"/></a:rPr><a:t>cars</a:t></a:r><a:endParaRPr lang="en-US" b="0" dirty="0"><a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="C0C0C0"/></a:outerShdw></a:effectLst><a:latin typeface="Whitney-Book" pitchFamily="50" charset="0"/><a:cs typeface="Whitney-Book" pitchFamily="50" charset="0"/></a:endParaRPr></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="104" name="Line 11"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{2D13A4FD-F41A-D960-86D4-A41F1454BC30}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr><a:spLocks noChangeShapeType="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr bwMode="auto"><a:xfrm flipH="1"><a:off x="4871976" y="3727766"/><a:ext cx="2885453" cy="633968"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="34925"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:round/><a:headEnd type="none" w="sm" len="sm"/><a:tailEnd type="stealth" w="med" len="lg"/></a:ln><a:effectLst/></p:spPr><p:txBody><a:bodyPr wrap="none" lIns="96231" tIns="48116" rIns="96231" bIns="48116" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" sz="3200" b="0"><a:latin typeface="Whitney-Book" pitchFamily="50" charset="0"/><a:cs typeface="Whitney-Book" pitchFamily="50" charset="0"/></a:endParaRPr></a:p></p:txBody></p:sp></p:grpSp>'''


def _inject_raw_xml(slide, xml_str):
    """Append a raw XML element (e.g. a <p:grpSp>) to a slide's shape tree.

    The XML must be a single root element. We inject xmlns:p and xmlns:a
    declarations onto the root so it parses standalone — when re-serialised
    as part of the slide, lxml strips redundant declarations.
    """
    NS = (
        ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    )
    # Insert namespaces right after the opening tag name
    tag_end = xml_str.find('>')
    space_in_tag = xml_str.find(' ', 0, tag_end)
    insert_at = space_in_tag if (space_in_tag != -1 and space_in_tag < tag_end) else tag_end
    wrapped = xml_str[:insert_at] + NS + xml_str[insert_at:]
    elem = ET.fromstring(wrapped)
    slide.shapes._spTree.append(elem)


def slide_10(prs):
    """Rivian's production function: weekly output of the Normal, IL plant.

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

        # Columns sized to leave ~0.5 cm padding either side of the longest
        # number at 14pt Calibri ("10,000" in col 0; "1000" in data cols).
        col_w_label = Inches(1.00)
        col_w_data = Inches(0.80)
        data_cols_w = col_w_data * 4
        tbl_w = col_w_label + data_cols_w
        tbl_h = Inches(4.00)           # ~0.33" per row × 12 rows
        tbl_top = Inches(2.20)         # shifted up from 2.45 on 2026-05-15
                                        # to make room for the Concept-
                                        # explanation callout at the bottom
        tbl_left = int((SLIDE_W - tbl_w) / 2)   # centre horizontally

        # Soft drop shadow rectangle BEHIND the table (graphicFrames can't host shadow).
        _add_graphicframe_shadow(slide, tbl_left, tbl_top, tbl_w, tbl_h)

        table_shape = slide.shapes.add_table(rows, cols, tbl_left, tbl_top,
                                              tbl_w, tbl_h)
        tbl = table_shape.table
        tbl.columns[0].width = col_w_label
        for c in range(1, cols):
            tbl.columns[c].width = col_w_data

        cell_pad = Inches(0.20)        # ≈ 0.5 cm horizontal
        cell_pad_v = Inches(0.02)
        for r, row in enumerate(rows_data):
            for c, val in enumerate(row):
                cell = tbl.cell(r, c)
                cell.margin_left = cell_pad
                cell.margin_right = cell_pad
                cell.margin_top = cell_pad_v
                cell.margin_bottom = cell_pad_v
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

        # --- Direction-of-increase arrows (gold block arrows) ---
        # Right-arrow above the data columns – more robots →
        data_cols_left = tbl_left + col_w_label
        top_arrow_h = Inches(0.30)
        top_arrow_top = tbl_top - top_arrow_h - Inches(0.10)
        top_arrow = slide.shapes.add_shape(
            MSO_SHAPE.RIGHT_ARROW,
            int(data_cols_left), int(top_arrow_top),
            int(data_cols_w), int(top_arrow_h),
        )
        top_arrow.fill.solid()
        top_arrow.fill.fore_color.rgb = GOLD
        top_arrow.line.fill.background()
        top_arrow.shadow.inherit = False

        # Down-arrow to the left of the data rows – more workers ↓
        # Aligned with rows 1..N (skip the header row).
        row_h_est = tbl_h / rows
        data_rows_top = tbl_top + row_h_est
        data_rows_h = tbl_h - row_h_est
        left_arrow_w = Inches(0.30)
        left_arrow_left = tbl_left - left_arrow_w - Inches(0.12)
        left_arrow = slide.shapes.add_shape(
            MSO_SHAPE.DOWN_ARROW,
            int(left_arrow_left), int(data_rows_top),
            int(left_arrow_w), int(data_rows_h),
        )
        left_arrow.fill.solid()
        left_arrow.fill.fore_color.rgb = GOLD
        left_arrow.line.fill.background()
        left_arrow.shadow.inherit = False

        # --- Axis labels (above top arrow / left of down arrow) ---
        top_label_h = Inches(0.40)
        top_label_y = top_arrow_top - top_label_h - Inches(0.10)
        _add_text(slide, int(data_cols_left), int(top_label_y),
                   int(data_cols_w), int(top_label_h),
                   "Number of robots (K)",
                   size=18, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        label_left_w = Inches(2.0)
        label_left_h = Inches(0.8)
        label_left_x = int(left_arrow_left - label_left_w - Inches(0.10))
        # Hand-nudge: centre with a -0.10" vertical offset to match user edit.
        label_left_y = int(data_rows_top + data_rows_h / 2 - label_left_h / 2
                            - Inches(0.10))
        _add_text(slide, label_left_x, label_left_y,
                   int(label_left_w), int(label_left_h),
                   "Number of\nworkers (L)",
                   size=18, bold=True, color=NAVY,
                   align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE,
                   font="Calibri")
        # Concept-explanation callout below the table — cream-fill rounded
        # rect, replacing the small italic gray caption that previously sat
        # here.  Two short lines, 17 pt navy bold.
        cap_w = Inches(11.0)
        cap_h = Inches(0.80)
        cap_x = (SLIDE_W - cap_w) // 2
        cap_y = Inches(6.30)
        _add_convention_box(
            slide, cap_x, cap_y, cap_w, cap_h,
            runs=[
                ("Output = cars per week",
                 {'size': 17, 'bold': True, 'color': NAVY}),
                ("MPL falls down each column;  MPK falls along each row",
                 {'size': 17, 'bold': True, 'color': NAVY, 'newline': True}),
            ],
            size=17, align=PP_ALIGN.CENTER,
        )
        # --- Inject the user-added "Number of cars" callout group ---
        _inject_raw_xml(slide, GROUP_XML_SLIDE10)

    s = make_diagram_slide(
        prs, page_num=11,
        section_tag=SECTION_TAG_P1,
        title="Rivian's Production Function:  Weekly Plant Output",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Rivian's weekly plant output as a function of workers (L) and "
        "robots (K).  Built from the Cobb-Douglas Q = 0.5 · √(K · L) – "
        "constant returns to scale overall, but strictly diminishing in "
        "each input individually.  Look down a column: MPL falls as you "
        "add labor with capital fixed.  Look along a row: MPK falls as "
        "you add robots with labor fixed.  Strict diminishing returns in "
        "both directions – the textbook story."
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

        # Chart frame: 8.4" wide, top y=1.30 (just below the title divider).
        # Height trimmed to 4.80" on 2026-05-15 to make room for the new
        # cream Concept-explanation callout below (which is larger than the
        # earlier thin banner).  Bottom at y≈6.10.
        chart_w = Inches(8.4)
        chart_h = Inches(4.80)
        chart_top = Inches(1.30)
        chart_left = Inches(2.636)
        # Drop-shadow rectangle behind the chart.
        _add_graphicframe_shadow(slide, chart_left, chart_top, chart_w, chart_h)
        chart_shape = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE,  # markers added per series below
            chart_left, chart_top, chart_w, chart_h,
            chart_data,
        )
        chart = chart_shape.chart

        # No internal "Output per Week" title – redundant with the slide
        # action title; dropping it reclaims ~0.4" of plot-area height.
        chart.has_title = False

        # --- Tighten the white margin around the plot ---------------------
        # By default PowerPoint reserves a generous border between the chart
        # frame and the plot region.  Force a manualLayout (layoutTarget=
        # "inner") so the inner plot fills ~88 % × 82 % of the chart frame,
        # leaving room only for the y-axis labels (left), x-axis labels +
        # title (bottom), and a thin top margin.  Added 2026-05-15.
        # chart._element is the <c:chartSpace>; plotArea lives at
        # chartSpace/chart/plotArea — navigate two levels down.
        chart_el = chart._element.find(qn('c:chart'))
        plot_el = chart_el.find(qn('c:plotArea')) if chart_el is not None else None
        if plot_el is not None:
            for old in plot_el.findall(qn('c:layout')):
                plot_el.remove(old)
            pl_layout = ET.Element(qn('c:layout'))
            pl_ml = ET.SubElement(pl_layout, qn('c:manualLayout'))
            ltgt = ET.SubElement(pl_ml, qn('c:layoutTarget')); ltgt.set('val', 'inner')
            xM = ET.SubElement(pl_ml, qn('c:xMode')); xM.set('val', 'edge')
            yM = ET.SubElement(pl_ml, qn('c:yMode')); yM.set('val', 'edge')
            xv = ET.SubElement(pl_ml, qn('c:x')); xv.set('val', '0.10')
            yv = ET.SubElement(pl_ml, qn('c:y')); yv.set('val', '0.03')
            wv = ET.SubElement(pl_ml, qn('c:w')); wv.set('val', '0.88')
            hv = ET.SubElement(pl_ml, qn('c:h')); hv.set('val', '0.82')
            plot_el.insert(0, pl_layout)

        # Native legend, positioned inside the plot area (top-left).
        # Font bumped 12 → 13 pt and box dims bumped on 2026-05-15 per
        # user request — agenda reads a touch larger.
        chart.has_legend = True
        chart.legend.font.name = "Calibri"
        chart.legend.font.size = Pt(13)
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
        # x bumped from 0.08 → 0.18 on 2026-05-15 so the legend clears the
        # y-axis area; h shrunk from 0.32 to 0.18 to keep the white-fill box
        # tight around the four legend entries (4 lines × ~12pt at 70 % LS).
        layout = ET.SubElement(leg_el, qn('c:layout'))
        ml = ET.SubElement(layout, qn('c:manualLayout'))
        xMode = ET.SubElement(ml, qn('c:xMode')); xMode.set('val', 'edge')
        yMode = ET.SubElement(ml, qn('c:yMode')); yMode.set('val', 'edge')
        x_el = ET.SubElement(ml, qn('c:x')); x_el.set('val', '0.18')
        y_el = ET.SubElement(ml, qn('c:y')); y_el.set('val', '0.05')
        w_el = ET.SubElement(ml, qn('c:w')); w_el.set('val', '0.17')
        h_el = ET.SubElement(ml, qn('c:h')); h_el.set('val', '0.24')
        # Re-order: legendPos must precede layout (already done by insert(0)).
        # Move <c:layout> right after <c:legendPos>.
        leg_el.remove(layout)
        leg_el.insert(list(leg_el).index(pos_el) + 1, layout)

        # --- Solid white fill + thin navy border on the legend box ---
        # Punches out the dashed gridlines underneath so the four series
        # labels read cleanly. Schema order: legendPos, legendEntry, layout,
        # overlay, spPr, txPr.  We ensure spPr sits after layout and before
        # any txPr that python-pptx may have created.
        for old in leg_el.findall(qn('c:spPr')):
            leg_el.remove(old)
        leg_spPr = ET.Element(qn('c:spPr'))
        sp_fill = ET.SubElement(leg_spPr, qn('a:solidFill'))
        sp_clr = ET.SubElement(sp_fill, qn('a:srgbClr'))
        sp_clr.set('val', 'FFFFFF')
        sp_ln = ET.SubElement(leg_spPr, qn('a:ln'))
        sp_ln.set('w', '6350')                 # 0.5 pt
        sp_lf = ET.SubElement(sp_ln, qn('a:solidFill'))
        sp_lc = ET.SubElement(sp_lf, qn('a:srgbClr'))
        sp_lc.set('val', '0B2B4E')             # NAVY
        # Insert immediately after c:layout (and any c:overlay that may exist)
        anchor = layout
        ovr = leg_el.find(qn('c:overlay'))
        if ovr is not None and list(leg_el).index(ovr) > list(leg_el).index(layout):
            anchor = ovr
        anchor.addnext(leg_spPr)

        # --- Tighter line spacing between entries (70%) ---
        # python-pptx already creates a c:txPr when font properties are set
        # above. We poke a:lnSpc into its first a:pPr; if no txPr exists,
        # build a minimal one.
        txPr = leg_el.find(qn('c:txPr'))
        if txPr is None:
            txPr = ET.Element(qn('c:txPr'))
            ET.SubElement(txPr, qn('a:bodyPr'))
            ET.SubElement(txPr, qn('a:lstStyle'))
            p_el = ET.SubElement(txPr, qn('a:p'))
            ET.SubElement(p_el, qn('a:pPr'))
            ET.SubElement(p_el, qn('a:endParaRPr'))
            leg_spPr.addnext(txPr)
        else:
            # Ensure txPr is after spPr (schema)
            leg_el.remove(txPr)
            leg_spPr.addnext(txPr)
        p_el = txPr.find(qn('a:p'))
        if p_el is None:
            p_el = ET.SubElement(txPr, qn('a:p'))
        pPr_el = p_el.find(qn('a:pPr'))
        if pPr_el is None:
            pPr_el = ET.Element(qn('a:pPr'))
            p_el.insert(0, pPr_el)
        for old in pPr_el.findall(qn('a:lnSpc')):
            pPr_el.remove(old)
        lnSpc = ET.Element(qn('a:lnSpc'))
        spcPct = ET.SubElement(lnSpc, qn('a:spcPct'))
        spcPct.set('val', '70000')             # 70%
        pPr_el.insert(0, lnSpc)

        # Axes
        cat = chart.category_axis
        cat.tick_labels.font.name = "Calibri"
        cat.tick_labels.font.size = Pt(11)
        cat.tick_labels.font.color.rgb = NAVY
        cat.has_title = True
        cat.axis_title.text_frame.text = "Number of Workers"
        ar = cat.axis_title.text_frame.paragraphs[0].runs[0]
        ar.font.name = "Calibri"; ar.font.size = Pt(14)
        ar.font.bold = True; ar.font.italic = True; ar.font.color.rgb = NAVY

        # Light-grey dashed vertical gridlines at each category tick
        # (1,000 / 2,000 / … / 10,000 workers). Schema order: majorGridlines
        # sits between c:axPos and c:title, so insert right after c:axPos.
        cat_el = cat._element
        for old in cat_el.findall(qn('c:majorGridlines')):
            cat_el.remove(old)
        gridlines = ET.Element(qn('c:majorGridlines'))
        sp = ET.SubElement(gridlines, qn('c:spPr'))
        ln = ET.SubElement(sp, qn('a:ln'))
        ln.set('w', '9525')                        # 0.75 pt
        ln.set('cap', 'flat'); ln.set('cmpd', 'sng'); ln.set('algn', 'ctr')
        fill = ET.SubElement(ln, qn('a:solidFill'))
        clr = ET.SubElement(fill, qn('a:srgbClr'))
        clr.set('val', 'C8CDD3')                   # RULE light grey
        dash = ET.SubElement(ln, qn('a:prstDash'))
        dash.set('val', 'dash')
        axpos = cat_el.find(qn('c:axPos'))
        axpos.addnext(gridlines)

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
        ar.font.bold = True; ar.font.italic = True; ar.font.color.rgb = NAVY

        # Light-grey dashed horizontal gridlines at each value tick
        # (100, 200, … 1000 cars per week).
        val_el = val._element
        for old in val_el.findall(qn('c:majorGridlines')):
            val_el.remove(old)
        v_gl = ET.Element(qn('c:majorGridlines'))
        sp = ET.SubElement(v_gl, qn('c:spPr'))
        ln = ET.SubElement(sp, qn('a:ln'))
        ln.set('w', '9525')
        ln.set('cap', 'flat'); ln.set('cmpd', 'sng'); ln.set('algn', 'ctr')
        fill = ET.SubElement(ln, qn('a:solidFill'))
        clr = ET.SubElement(fill, qn('a:srgbClr')); clr.set('val', 'C8CDD3')
        dash = ET.SubElement(ln, qn('a:prstDash')); dash.set('val', 'dash')
        v_axpos = val_el.find(qn('c:axPos'))
        v_axpos.addnext(v_gl)

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

        # Concept-explanation callout below the chart — cream-fill rounded
        # rect (the Convention-style format documented in the Teaching
        # CLAUDE.md).  Tightened on 2026-05-15: shorter height + smaller
        # internal padding so the cream chrome hugs the two text lines.
        banner_w = Inches(10.0)
        banner_h = Inches(0.78)
        banner_x = (SLIDE_W - banner_w) // 2
        banner_y = chart_top + chart_h + Inches(0.05)
        _add_convention_box(
            slide, banner_x, banner_y, banner_w, banner_h,
            runs=[
                ("Each curve flattens as L rises  (diminishing MPL)",
                 {'size': 17, 'bold': True, 'color': NAVY}),
                ("The vertical distance between curves narrows as K rises  (diminishing MPK)",
                 {'size': 17, 'bold': True, 'color': NAVY, 'newline': True}),
            ],
            size=17, align=PP_ALIGN.CENTER,
            pad_h=Inches(0.15), pad_v=Inches(0.04),
        )

    s = make_diagram_slide(
        prs, page_num=12,
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
    """Short Run: Marginal Product of Labor — concept intro.

    Mirrors the source deck's original slide 16: re-anchor the short-run
    framing (K fixed, L flexible), introduce the MPL concept name in
    accent blue, give the formal change/change definition with "change"
    flagged in accent red italic, and close with the canonical
    ΔQ/ΔL formula in big OMML below.
    """
    ACCENT_BLUE = RGBColor(0x00, 0x70, 0xC0)
    DARK_YELLOW = RGBColor(0xB8, 0x86, 0x0B)   # was red; per user request

    def _styled_run(p, text, *, size=24, bold=False, italic=False,
                    color=NAVY, font="Calibri"):
        r = p.add_run()
        r.text = text
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
        return r

    def _block(slide, left, top, width, height, runs, *,
               align=PP_ALIGN.LEFT):
        tb = slide.shapes.add_textbox(int(left), int(top),
                                        int(width), int(height))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
        tf.margin_top = Inches(0.0);  tf.margin_bottom = Inches(0.0)
        p = tf.paragraphs[0]
        p.alignment = align
        for text, opts in runs:
            _styled_run(p, text, **opts)
        return tb

    def draw(slide):
        indent = Inches(0.45)
        # 1) MAIN BULLET — "Typically in the short run:"  (28pt bold navy,
        #    prefixed with a navy ▪ glyph)
        _block(slide, MARGIN, Inches(1.85), RULE_W, Inches(0.45), [
            ("▪  ",                {'size': 28, 'bold': True}),
            ("Typically",          {'size': 28, 'bold': True}),
            (" in the short run:", {'size': 28, 'bold': True}),
        ])
        # 2) SUB-BULLETS — Capital is fixed (K̅) / Labor is flexible (L)
        #    Two paragraphs with inline OMML for the symbols.  Sub-text
        #    bumped 24 → 26 pt on 2026-05-15 (sub-bullets were too small).
        _add_mixed_textbox(slide,
                            MARGIN + indent, Inches(2.45),
                            RULE_W - indent, Inches(1.20),
                            [
                                ('text', "–  ", {'size': 26}),
                                ('text', "Capital is ", {'size': 26}),
                                ('text', "fixed", {'size': 26, 'bold': True}),
                                ('text', "  (", {'size': 26}),
                                ('omml', _omml_acc_overline('K'), {'size': 26}),
                                ('text', ")", {'size': 26}),
                                ('break', '', {}),
                                ('text', "–  ", {'size': 26}),
                                ('text', "Labor is flexible  (", {'size': 26}),
                                ('omml', _omml_run('L'), {'size': 26}),
                                ('text', ")", {'size': 26}),
                            ],
                            default_size=26, default_color=NAVY)

        # 3) MAIN BULLET — "Important Concept:  Marginal Product of Labor"
        _block(slide, MARGIN, Inches(4.00), RULE_W, Inches(0.45), [
            ("▪  ",                        {'size': 28, 'bold': True}),
            ("Important Concept:  ",       {'size': 28, 'bold': True}),
            ("Marginal Product of Labor",  {'size': 28, 'bold': True,
                                             'color': ACCENT_BLUE}),
        ])
        # 4) SUB-BULLET — formal definition with "change" emphasised,
        #    26pt (bumped from 24 — see same note as above), indented.
        def_tb = slide.shapes.add_textbox(
            int(MARGIN + indent), int(Inches(4.60)),
            int(RULE_W - indent), int(Inches(0.75)))
        def_tf = def_tb.text_frame
        def_tf.word_wrap = True
        def_tf.margin_left = Inches(0.05); def_tf.margin_right = Inches(0.05)
        def_tf.margin_top = Inches(0); def_tf.margin_bottom = Inches(0)
        p = def_tf.paragraphs[0]
        _styled_run(p, "–  ", size=26)
        _styled_run(p, "The ", size=26)
        _styled_run(p, "marginal product of labor ", size=26,
                    color=ACCENT_BLUE)
        _styled_run(p, "is the ", size=26)
        _styled_run(p, "change", size=26, italic=True, color=DARK_YELLOW)
        _styled_run(p, " in output due to a ", size=26)
        _styled_run(p, "change", size=26, italic=True, color=DARK_YELLOW)
        _styled_run(p, " in labor input:", size=26)

        # 5) Big OMML formula MPL = ΔQ / ΔL  (36pt, blue)
        mpl     = _omml_sub(_omml_run('MP'), _omml_run('L'))
        delta_q = _omml_text('Δ') + _omml_run('Q')
        delta_l = _omml_text('Δ') + _omml_run('L')
        frac    = _omml_frac(delta_q, delta_l)
        omml_full = mpl + _omml_text(' = ') + frac
        _add_math_equation(slide,
                            left=Inches(4.7), top=Inches(5.55),
                            width=Inches(4.0), height=Inches(1.25),
                            omml_content=omml_full,
                            size_pt=36, color=ACCENT_BLUE)

    s = make_diagram_slide(
        prs, page_num=13,
        section_tag=SECTION_TAG_P1,
        title="Short Run:  Marginal Product of Labor",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Re-anchor the short-run framing: capital is fixed (we write the "
        "bar over K to emphasise that K is held constant) while labor is "
        "flexible. Now introduce the marginal product of labor formally: "
        "MPL is the CHANGE in output due to a CHANGE in labor input. The "
        "formula MPL = ΔQ / ΔL is shorthand for \"how much extra output "
        "do I get from one more worker?\" – exactly the question every "
        "plant manager asks every shift.  Next slide: the actual numbers "
        "from the Rivian production function show MPL declining."
    ))


def slide_mpl_data(prs):
    """Numerical MPL example — matches original slide 17 of the source deck.

    Fixes K = 100 robots, walks L through {0, 500, 1k, 2k, 3k}, and
    builds the L | K | Q | ΔL | ΔQ | MPL table from the same Cobb-Douglas
    function as slides 10 and 11.  MPL falls strictly down the column,
    making the "diminishing marginal product" lesson visible numerically.
    A right-hand Convention box explains the step-by-step ΔL / ΔQ rule.
    """
    ACCENT_BLUE = RGBColor(0x00, 0x70, 0xC0)
    ACCENT_RED  = RGBColor(0xFF, 0x00, 0x00)
    MPL_FILL    = RGBColor(0xFF, 0xF5, 0xE0)   # cream highlight for MPL column
    CONV_FILL   = RGBColor(0xFD, 0xF6, 0xE6)   # softer cream for callout

    # Per-column body-cell colors (header row stays white-on-navy).
    # Column order on 2026-05-15: L | K | Q | ΔQ | ΔL | MPL  (ΔQ before
    # ΔL, swapped from the original ordering per user request).
    BLACK_NUM = RGBColor(0x00, 0x00, 0x00)
    RED_NUM   = RGBColor(0xC0, 0x00, 0x00)
    GREEN_NUM = RGBColor(0x1B, 0x5E, 0x20)        # darker / deeper green
    BLUE_NUM  = ACCENT_BLUE
    COL_COLORS = [BLACK_NUM, RED_NUM, BLACK_NUM, GREEN_NUM, GREEN_NUM, BLUE_NUM]

    K_FIX = 100
    L_GRID = [0, 500, 1000, 2000, 3000]

    def draw(slide):
        # Main bullet — replaces the old centred italic captions with a
        # proper bullet structure (per user request 2026-05-15).
        _add_mixed_textbox(slide,
                            MARGIN, Inches(1.85),
                            RULE_W, Inches(0.45),
                            [
                                ('text', "▪  ",
                                 {'size': 24, 'bold': True, 'color': NAVY}),
                                ('text', "Example:  MPL from Rivian Production function.",
                                 {'size': 24, 'bold': True, 'color': NAVY}),
                            ],
                            align=PP_ALIGN.LEFT,
                            default_size=24, default_color=NAVY)

        # Sub-bullet with inline OMML K̅
        _add_mixed_textbox(slide,
                            MARGIN + Inches(0.45), Inches(2.35),
                            RULE_W - Inches(0.45), Inches(0.40),
                            [
                                ('text', "–  ", {'size': 22, 'color': NAVY}),
                                ('text', "Fix capital at  ",
                                 {'size': 22, 'color': NAVY}),
                                ('omml', _omml_acc_overline('K'),
                                 {'size': 22}),
                                ('text', "  =  100",
                                 {'size': 22, 'bold': True, 'color': NAVY}),
                            ],
                            align=PP_ALIGN.LEFT,
                            default_size=22, default_color=NAVY)

        # ---- Table (6 columns, including the new K column) ----
        # The Δ-columns (ΔL, ΔQ, MPL) are rendered specially on
        # 2026-05-15: their cells are blank inside the table, and the
        # values are drawn as floating textboxes positioned at the
        # BOUNDARY between two adjacent rows.  This visually illustrates
        # the convention that each Δ is computed relative to the
        # previous (initial) point — values live "between" rows, not on
        # them.  MPL floats also get the cream MPL_FILL background;
        # ΔL/ΔQ floats are transparent (green numbers on white).
        Q = [_pf_value(K_FIX, L) for L in L_GRID]
        dL_values  = [None]
        dQ_values  = [None]
        mpl_values = [None]
        rows_data = [["L", "K", "Q", "ΔQ", "ΔL", "MPL"]]
        for i, L in enumerate(L_GRID):
            row = [f"{L:,}", f"{K_FIX}", f"{Q[i]:,}"]
            if i == 0:
                row += ["", "", ""]                       # all 3 Δ cells empty
            else:
                dL = L_GRID[i] - L_GRID[i-1]
                dQ = Q[i] - Q[i-1]
                mpl = dQ / dL
                row += ["", "", ""]                       # all 3 Δ cells empty
                dL_values.append(f"{dL:,}")
                dQ_values.append(f"{dQ}")
                mpl_values.append(f"{mpl:.3f}")
            rows_data.append(row)

        rows = len(rows_data); cols = len(rows_data[0])
        col_widths = [Inches(0.80), Inches(0.65),
                       Inches(0.80), Inches(0.85),
                       Inches(0.75), Inches(0.95)]
        tbl_w = sum(col_widths)
        tbl_h = Inches(2.55)
        tbl_top = Inches(2.85)
        # Table no longer centred – keep it on the LEFT so a Convention
        # callout fits to its right.
        tbl_left = Inches(0.80)
        _add_graphicframe_shadow(slide, tbl_left, tbl_top, tbl_w, tbl_h)
        tshape = slide.shapes.add_table(rows, cols, tbl_left, tbl_top,
                                          tbl_w, tbl_h)
        tbl = tshape.table
        for ci, w in enumerate(col_widths):
            tbl.columns[ci].width = w

        cell_pad_h = Inches(0.10)
        for r, row in enumerate(rows_data):
            for c, val in enumerate(row):
                cell = tbl.cell(r, c)
                cell.margin_left = cell_pad_h
                cell.margin_right = cell_pad_h
                cell.margin_top = Inches(0.03)
                cell.margin_bottom = Inches(0.03)
                cell.text = str(val)
                for p in cell.text_frame.paragraphs:
                    p.alignment = PP_ALIGN.CENTER
                    for run in p.runs:
                        run.font.name = "Calibri"
                        run.font.size = Pt(16)
                        if r == 0:
                            run.font.bold = True
                            run.font.color.rgb = WHITE
                        else:
                            # Per-column color scheme (2026-05-15):
                            # L/Q black, K red, ΔL/ΔQ green, MPL blue+bold.
                            run.font.color.rgb = COL_COLORS[c]
                            if c == cols - 1:
                                run.font.bold = True
                if r == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = NAVY
                else:
                    # All data cells: white background (the cream
                    # MPL_FILL now lives on the floating MPL textboxes
                    # below, not on the cells themselves).
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = WHITE

        # ---- Floating Δ values, drawn at row boundaries ----
        # ΔL, ΔQ and MPL all float between rows; MPL gets a cream-fill
        # rounded-rect background (was the column's fill before), while
        # ΔL and ΔQ are transparent green text on the underlying white
        # cell.  Assumes equal row heights (tbl_h / rows) — true here
        # since no row has multi-line content.
        row_h = tbl_h / rows
        col_left = [tbl_left + sum(col_widths[:c]) for c in range(cols + 1)]
        float_h = Inches(0.34)                            # ~cell height
        GREEN = COL_COLORS[3]                             # 0x008000

        def _float_value(text, c, i, *, color, bold=False, fill_rgb=None,
                          border=None, line_w=0.5):
            """Place ``text`` in column c at the boundary above row i+1."""
            r = i + 1
            boundary_y = tbl_top + r * row_h
            cell_x = col_left[c]
            cell_w = col_widths[c]
            top_y = int(boundary_y - float_h / 2)
            # Optional fill — draw a rounded rect behind the text.
            if fill_rgb is not None:
                # Inset the fill rect slightly inside the column so it
                # doesn't kiss the column-separator lines.
                pad = Inches(0.04)
                rect = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    int(cell_x + pad), top_y,
                    int(cell_w - 2 * pad), int(float_h),
                )
                try: rect.adjustments[0] = 0.18
                except Exception: pass
                rect.fill.solid()
                rect.fill.fore_color.rgb = fill_rgb
                if border is not None:
                    rect.line.color.rgb = border
                    rect.line.width = Pt(line_w)
                else:
                    rect.line.fill.background()
                rect.shadow.inherit = False
            tb = slide.shapes.add_textbox(
                int(cell_x), top_y, int(cell_w), int(float_h),
            )
            ttf = tb.text_frame
            ttf.word_wrap = True
            ttf.margin_left = Inches(0.02); ttf.margin_right = Inches(0.02)
            ttf.margin_top = Inches(0); ttf.margin_bottom = Inches(0)
            ttf.vertical_anchor = MSO_ANCHOR.MIDDLE
            pp = ttf.paragraphs[0]
            pp.alignment = PP_ALIGN.CENTER
            rr = pp.add_run()
            rr.text = text
            rr.font.name = "Calibri"
            rr.font.size = Pt(16)
            rr.font.bold = bold
            rr.font.color.rgb = color

        # Column order: c=3 → ΔQ, c=4 → ΔL, c=5 → MPL (swapped 2026-05-15).
        for i in range(1, len(L_GRID)):
            _float_value(dQ_values[i],  3, i, color=GREEN)
            _float_value(dL_values[i],  4, i, color=GREEN)
            _float_value(mpl_values[i], 5, i,
                          color=ACCENT_BLUE, bold=True,
                          fill_rgb=MPL_FILL)

        # ---- Green DOWN-arrows in the Q column, between adjacent rows ----
        # Mirrors the green connectors on the original slide 17 — visually
        # links Q[i] → Q[i+1] (the "we went from this Q to that Q" cue
        # that pairs with the ΔQ float between the same two rows).
        # Position fine-tuned on 2026-05-15: ~5 mm right of column-centre,
        # then shifted ~3 mm left so a small gap opens between the wavy
        # connector and the ΔQ digit it points at.
        q_arrow_x = (col_left[2] + int(col_widths[2] * 0.72)
                      + Inches(0.20) - Inches(0.12))
        arrow_h = Inches(0.36)                            # vertical span
        dq_col_center = col_left[3] + col_widths[3] // 2
        # Approximate width of one digit at 16 pt Calibri (used to find the
        # x-position of the FIRST digit inside a centred ΔQ value).
        char_w = Inches(0.105)
        for i in range(1, len(L_GRID)):
            # Boundary y between row i and row i+1
            r = i + 1
            boundary_y = tbl_top + r * row_h
            _add_arrow(slide,
                        (q_arrow_x, int(boundary_y - arrow_h / 2)),
                        (q_arrow_x, int(boundary_y + arrow_h / 2)),
                        color=GREEN, weight_pt=3.0, head=True)
            # Wavy green connector from the arrow midpoint across to the
            # ΔQ first-digit centre (shifted ~2 mm short so a small gap
            # remains between the line end and the digit).  Polyline
            # approximation of ~1.75 sine cycles for a gentle wave.
            n_chars = len(dQ_values[i])
            first_digit_x = dq_col_center - int((n_chars - 1) / 2 * char_w)
            line_end_x = first_digit_x - Inches(0.08)
            _add_wavy_line(slide,
                            q_arrow_x, line_end_x, boundary_y,
                            amplitude=Inches(0.02),
                            cycles=1.75, segments=36,
                            color=GREEN, weight_pt=1.5)

        # ---- Wide low-arc green line: first Q-arrow → Convention box ----
        # Cubic-Bezier inspired by the original slide 17.  Runs THROUGH
        # the empty horizontal band between the table header row and the
        # first row of floating ΔQ / ΔL / MPL numbers (i.e., the L = 0
        # row).  Apex sits inside that band — the curve is therefore a
        # very wide, very shallow inverted-U rather than a half-circle
        # arching over the whole table.  Stops ~0.05" before the
        # Convention box's left edge.
        arc_x_start = q_arrow_x
        arc_y_start = tbl_top + 2 * row_h                   # middle of first arrow
        arc_x_end = Inches(6.10)                            # 0.05" left of conv
        arc_y_end = Inches(3.50)                            # inside the empty band
        arc_apex_y = Inches(3.35)                           # apex inside L = 0 row
        bbox_left = int(min(arc_x_start, arc_x_end))
        bbox_top = int(arc_apex_y)
        bbox_w = int(abs(arc_x_end - arc_x_start))
        bbox_h = int(max(arc_y_start, arc_y_end) - arc_apex_y)
        # Normalized coords (0–100000 along each axis of the bounding box)
        start_lx = 0
        start_ly = int(round((arc_y_start - arc_apex_y) / (max(arc_y_start, arc_y_end) - arc_apex_y) * 100000))
        end_lx = 100000
        end_ly = int(round((arc_y_end - arc_apex_y) / (max(arc_y_start, arc_y_end) - arc_apex_y) * 100000))
        # Control points pulled to the TOP and very close to the side
        # edges → inverted-U shape (steeper sides, flatter top) rather
        # than a perfectly round arc.
        cp1 = (8000, 0)
        cp2 = (92000, 0)
        weight_emu = int(1.5 * 12700)
        green_hex = f'{GREEN[0]:02X}{GREEN[1]:02X}{GREEN[2]:02X}'
        P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        A_NS_LOCAL = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        arc_xml = (
            f'<p:sp xmlns:p="{P_NS}" xmlns:a="{A_NS_LOCAL}">'
            f'<p:nvSpPr><p:cNvPr id="0" name="HalfCircleArc"/>'
            f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr>'
            f'<a:xfrm>'
            f'<a:off x="{bbox_left}" y="{bbox_top}"/>'
            f'<a:ext cx="{bbox_w}" cy="{bbox_h}"/>'
            f'</a:xfrm>'
            f'<a:custGeom>'
            f'<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
            f'<a:rect l="0" t="0" r="0" b="0"/>'
            f'<a:pathLst>'
            f'<a:path w="100000" h="100000" fill="none">'
            f'<a:moveTo><a:pt x="{start_lx}" y="{start_ly}"/></a:moveTo>'
            f'<a:cubicBezTo>'
            f'<a:pt x="{cp1[0]}" y="{cp1[1]}"/>'
            f'<a:pt x="{cp2[0]}" y="{cp2[1]}"/>'
            f'<a:pt x="{end_lx}" y="{end_ly}"/>'
            f'</a:cubicBezTo>'
            f'</a:path>'
            f'</a:pathLst>'
            f'</a:custGeom>'
            f'<a:noFill/>'
            f'<a:ln w="{weight_emu}" cap="rnd">'
            f'<a:solidFill><a:srgbClr val="{green_hex}"/></a:solidFill>'
            f'</a:ln>'
            f'</p:spPr>'
            f'</p:sp>'
        )
        slide.shapes._spTree.append(ET.fromstring(arc_xml))

        # ---- Blue arrow: bottom of MPL column → Note below the table ----
        # Visually anchors the "Note: MPL is declining" callout to the
        # MPL data above it.  Same blue as the MPL numbers (ACCENT_BLUE).
        mpl_col_center = col_left[5] + col_widths[5] // 2
        tbl_bottom = tbl_top + tbl_h
        note_top_y = Inches(6.10)        # see note positioning below
        _add_arrow(slide,
                    (mpl_col_center, int(tbl_bottom + Inches(0.10))),
                    (mpl_col_center, int(note_top_y - Inches(0.05))),
                    color=ACCENT_BLUE, weight_pt=3.0, head=True)

        # ---- Convention callout to the right of the table ----
        # 2026-05-15: narrower (5.20" → 4.20") + larger font (17 → 19 pt);
        # ΔL and ΔQ now use the same green as the body-cell ΔQ/ΔL digits.
        # Later that day: added an "Interpretation:" second paragraph
        # spelling out the first MPL value — box height bumped to 1.60"
        # to accommodate the additional 2 lines of wrapped text, then
        # widened (4.20" → 5.80") so the Interpretation line breaks
        # cleanly into exactly two lines.
        conv_w = Inches(5.80)
        conv_h = Inches(1.60)
        conv_x = tbl_left + tbl_w + Inches(0.55)
        conv_y = tbl_top + (tbl_h - conv_h) // 2     # vertically centred
        conv_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            int(conv_x), int(conv_y), int(conv_w), int(conv_h),
        )
        conv_box.fill.solid()
        conv_box.fill.fore_color.rgb = CONV_FILL
        conv_box.line.color.rgb = NAVY
        conv_box.line.width = Pt(1.0)
        conv_box.shadow.inherit = False
        try: conv_box.adjustments[0] = 0.12
        except Exception: pass
        _add_mixed_textbox(slide,
                            conv_x + Inches(0.18),
                            conv_y + Inches(0.10),
                            conv_w - Inches(0.36),
                            conv_h - Inches(0.20),
                            [
                                ('text', "Convention:  ",
                                 {'size': 19, 'bold': True, 'color': NAVY}),
                                ('text', "Compute  ",
                                 {'size': 19, 'color': NAVY}),
                                ('omml',
                                 _omml_text('Δ', color=GREEN_NUM)
                                 + _omml_run('L', color=GREEN_NUM),
                                 {'size': 19}),
                                ('text', "  and  ",
                                 {'size': 19, 'color': NAVY}),
                                ('omml',
                                 _omml_text('Δ', color=GREEN_NUM)
                                 + _omml_run('Q', color=GREEN_NUM),
                                 {'size': 19}),
                                ('text', "  for each interval",
                                 {'size': 19, 'color': NAVY}),
                                ('break', '', {}),
                                ('text', "Interpretation:  ",
                                 {'size': 19, 'bold': True, 'color': NAVY}),
                                ('text',
                                 "Between 0 and 500 workers, MPL is "
                                 "approximately 0.224",
                                 {'size': 19, 'color': NAVY}),
                            ],
                            align=PP_ALIGN.LEFT,
                            default_size=19, default_color=NAVY)

        # ---- "MPL is declining as we add workers" — Convention-style box ----
        # 2026-05-15: the Note now lives inside the same cream Convention
        # callout chrome used elsewhere on the slide.  Centred horizontally
        # at y=6.10 (the blue MPL→Note arrow above still terminates just
        # before this box).
        note_w = Inches(8.20)
        note_h = Inches(0.75)
        note_x = (SLIDE_W - note_w) // 2
        note_y = Inches(6.10)
        _add_convention_box(
            slide, note_x, note_y, note_w, note_h,
            runs=[
                ("Note:  ",
                 {'size': 22, 'bold': True, 'color': NAVY}),
                ("MPL ",
                 {'size': 22, 'bold': True, 'italic': True,
                  'color': ACCENT_BLUE}),
                ("is ",
                 {'size': 22, 'bold': True, 'color': NAVY}),
                ("declining",
                 {'size': 22, 'bold': True, 'italic': True,
                  'color': ACCENT_BLUE}),
                (" as we add workers",
                 {'size': 22, 'bold': True, 'color': NAVY}),
            ],
            size=22, align=PP_ALIGN.CENTER,
        )

    s = make_diagram_slide(
        prs, page_num=14,
        section_tag=SECTION_TAG_P1,
        title="Marginal Product of Labor (MPL):  Calculation",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Same Cobb-Douglas function we used on slides 10 and 11, now with "
        "K = 100 robots fixed and L walked through {0, 500, 1000, 2000, "
        "3000} workers.  L grid is denser at the start so the diminishing-"
        "MPL effect shows up cleanly: MPL falls from 0.224 cars/worker "
        "(going from 0 to 500 workers) down to 0.050 cars/worker (going "
        "from 2,000 to 3,000 workers).  Convention: ΔL and ΔQ are step-"
        "by-step changes (current row minus previous row), NOT changes "
        "relative to L = 0.  This is the textbook MPL and it's what "
        "makes \"diminishing marginal product\" show up numerically."
    ))


def slide_13(prs):
    """Diminishing Marginal Product of Labor — mirrors original slide 18.

    Header + bullet wording restored to the original deck on 2026-05-15.
    Two dashed secant lines added to the LEFT chart to make MPL = slope
    of the production function visible; the "plot the slope" callout
    moved into the gap between the two charts so it reads as a bridge.
    """
    bullets = [
        ("We hold one input fixed (capital) and…", 0),
        ("Use more and more of a variable input (labor)", 0),
        ("Then total output will increase by less and less", 1),
        ("i.e., the Marginal Product of Labor (MPL) goes down", 1),
    ]

    K_FIX = 100
    BLUE = RGBColor(0x2E, 0x75, 0xB6)   # matches K=100 series on slide 11
    DASH_COLOR = GOLD                    # tangent lines (= MPL series color)

    def draw_pictures(slide):
        # Compute Q and MPL from the same Cobb-Douglas (Q = 0.5·√(K·L))
        # we use on slides 10 and 11.  Fix K = 100 robots.
        L_vals = PF_L_VALS                     # 0, 1000, 2000, …, 10000
        Q_vals = [_pf_value(K_FIX, L) for L in L_vals]
        MPL_L = L_vals[1:]                     # MPL evaluated for L ≥ 1000
        MPL_vals = [
            (Q_vals[i] - Q_vals[i-1]) / (L_vals[i] - L_vals[i-1])
            for i in range(1, len(L_vals))
        ]

        # ---- Chart frames ------------------------------------------------
        # 2026-05-15: chart-top moved down to 3.58" so the per-chart
        # caption can sit ABOVE the chart instead of below it.  Chart
        # heights unchanged → bottom = 6.43".
        cap_y = Inches(3.25)
        cap_h = Inches(0.30)
        left_chart_x  = Inches(0.40)
        left_chart_y  = Inches(3.58)
        left_chart_w  = Inches(5.90)
        left_chart_h  = Inches(2.85)
        right_chart_x = Inches(7.00)
        right_chart_y = Inches(3.58)
        right_chart_w = Inches(5.90)
        right_chart_h = Inches(2.85)

        # LEFT chart: Total output Q vs L
        left_chart_shape = _make_simple_line_chart(
            slide, left_chart_x, left_chart_y,
            left_chart_w, left_chart_h,
            categories=[f"{L:,}" for L in L_vals],
            values=Q_vals,
            line_color=BLUE,
            x_title="Workers (L)",
            y_title="Output (Q)",
            y_max=550, y_unit=50,
        )
        # RIGHT chart: MPL vs L — gold series so the curve visually
        # matches the gold tangent lines drawn on the left chart.
        right_chart_shape = _make_simple_line_chart(
            slide, right_chart_x, right_chart_y,
            right_chart_w, right_chart_h,
            categories=[f"{L:,}" for L in MPL_L],
            values=MPL_vals,
            line_color=DASH_COLOR,
            x_title="Workers (L)",
            y_title="MPL  (cars per worker)",
            y_max=0.18, y_unit=0.02,
        )

        # ---- Smooth both curves + pin the inner plot area to a known
        #      bounding box so the overlay tangent lines can be drawn at
        #      exact positions on the curve. -----------------------------
        INNER = ('0.15', '0.04', '0.80', '0.78')  # x, y, w, h fractions
        def _post_process_chart(chart_shape):
            chart_el = chart_shape.chart._element.find(qn('c:chart'))
            plot_el = chart_el.find(qn('c:plotArea')) if chart_el is not None else None
            if plot_el is not None:
                for old in plot_el.findall(qn('c:layout')):
                    plot_el.remove(old)
                layout = ET.Element(qn('c:layout'))
                ml = ET.SubElement(layout, qn('c:manualLayout'))
                lt = ET.SubElement(ml, qn('c:layoutTarget')); lt.set('val', 'inner')
                xm = ET.SubElement(ml, qn('c:xMode')); xm.set('val', 'edge')
                ym = ET.SubElement(ml, qn('c:yMode')); ym.set('val', 'edge')
                for tag, val in zip(('c:x', 'c:y', 'c:w', 'c:h'), INNER):
                    el = ET.SubElement(ml, qn(tag)); el.set('val', val)
                plot_el.insert(0, layout)
            # smooth=1 on every series so the curve is a smooth spline
            # through the data points instead of piecewise-linear segments.
            # NOTE: python-pptx also writes a CHART-LEVEL <c:smooth val="0">
            # directly under <c:lineChart>; PowerPoint honors that over
            # the series-level setting, so update / remove it too.
            for series in chart_shape.chart.series:
                ser_xml = series._element
                for sm in ser_xml.findall(qn('c:smooth')):
                    ser_xml.remove(sm)
                sm_el = ET.SubElement(ser_xml, qn('c:smooth'))
                sm_el.set('val', '1')
            # Chart-level smooth: live inside <c:plotArea>/<c:lineChart>
            line_chart = plot_el.find(qn('c:lineChart')) if plot_el is not None else None
            if line_chart is not None:
                for sm in line_chart.findall(qn('c:smooth')):
                    sm.set('val', '1')
        _post_process_chart(left_chart_shape)
        _post_process_chart(right_chart_shape)

        # ---- Dashed TANGENT lines on the LEFT chart ----------------------
        # Coordinates hand-tweaked in PowerPoint on 2026-05-15 against
        # the rendered smooth curve so each line visibly *kisses* the
        # production function at one point (the analytical Q = 5·√L
        # tangents drawn via _draw_tangent looked like secants because
        # PowerPoint's spline smoothing differs slightly from the true
        # √L curve).  Keep these exact endpoints — re-running the
        # analytical helper will re-introduce the visual mismatch.
        # Y-coords track the chart-top:
        #   chart_y = 3.40" (hand-edit baseline) → 3.58" (current, shift +0.18)
        # User hand-edit starts: 5.486, 4.512, 3.961, 3.570  (at chart_y=3.40)
        # → +0.18 shift gives: 5.666, 4.692, 4.141, 3.750
        # Steep / early tangent (≈ L = 1500 region):
        _add_arrow(slide,
                    start_xy=(Inches(1.471), Inches(5.666)),
                    end_xy=(Inches(2.536), Inches(4.692)),
                    color=DASH_COLOR, weight_pt=2.0,
                    head=False, dash='dash')
        # Flat / late tangent (≈ L = 8000 region):
        _add_arrow(slide,
                    start_xy=(Inches(4.464), Inches(4.141)),
                    end_xy=(Inches(6.103), Inches(3.750)),
                    color=DASH_COLOR, weight_pt=2.0,
                    head=False, dash='dash')

        # Captions ABOVE each chart (moved on 2026-05-15 from below the
        # chart to here; chart frames shifted down by 0.33" to make room).
        _add_text(slide, left_chart_x, cap_y, left_chart_w, cap_h,
                   "Total output  (rising, flattening)",
                   size=13, italic=True, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")
        _add_text(slide, right_chart_x, cap_y, right_chart_w, cap_h,
                   "Slope = MPL  (falling)",
                   size=13, italic=True, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")

        # ---- "plot the slope" callout BETWEEN the two charts -------------
        # Combined (box + block arrow) centred horizontally on the gap
        # midpoint (x = 6.65).  Sizes bumped ~30 % on 2026-05-15 and a
        # soft drop shadow added to both shapes for visual weight.
        cb_w = Inches(1.04)              # 0.80 × 1.30
        cb_h = Inches(0.72)              # 0.55 × ~1.30
        arr_w = Inches(0.72)             # 0.55 × ~1.30
        arr_h = Inches(0.39)             # 0.30 × 1.30
        gap_mid_x = (left_chart_x + left_chart_w + right_chart_x) // 2
        # 0.32" (≈ 0.8 cm = 0.5 cm + 3 mm) left of the gap midpoint
        # per user requests on 2026-05-15.
        cb_x = gap_mid_x - (cb_w + arr_w) // 2 - Inches(0.32)
        cb_y = left_chart_y + (left_chart_h - cb_h) // 2     # vert. centred
        cb_shape = _add_callout_box(slide,
                                      left=cb_x, top=cb_y,
                                      width=cb_w, height=cb_h,
                                      text="plot the slope",
                                      fill=GOLD, text_color=NAVY,
                                      size=13, bold=True)
        if cb_shape is not None:
            _add_drop_shadow(cb_shape)
        # Block right-arrow (MSO_SHAPE.RIGHT_ARROW) — much more visible
        # than a thin line connector.
        arr_shape = _add_arrow_shape(slide,
                                       left=cb_x + cb_w,
                                       top=cb_y + (cb_h - arr_h) // 2,
                                       width=arr_w, height=arr_h,
                                       direction="right", fill=GOLD)
        if arr_shape is not None:
            _add_drop_shadow(arr_shape)

        # Bottom takeaway bar — nudged from 6.40 → 6.55 since the chart
        # frames are now ~0.33" taller (captions moved above).  Bar
        # bottom = 7.10, footer rule at 7.135 → clear by 0.035".
        _add_takeaway_bar(slide,
                           "Note:  MPL is the slope  (dQ / dL)  of the output curve",
                           top=Inches(6.55), fill=NAVY,
                           width=Inches(9.5), size=18)

    s = make_content_bulleted(
        prs, page_num=15,
        section_tag=SECTION_TAG_P1,
        title="Diminishing Marginal Product of Labor",
        bullets=bullets,
        # Tightened spacing on 2026-05-15 so all 4 bullets sit ABOVE
        # the charts that start at y = 3.25".  bullets_top raised by
        # ~0.5 cm (1.85 → 1.65) per user request so they clear the
        # figures' top edges.
        size=22, sub_size=20,
        line_spacing_pts=2, sub_line_spacing_pts=0,
        bullets_top=Inches(1.53),
        extras=draw_pictures,
    )
    _set_notes(s, (
        "The headline of this section. Diminishing MPL is a near-universal "
        "feature of short-run production: each additional worker has to "
        "share the same fixed capital, so the marginal contribution shrinks. "
        "This isn't a quirk of Rivian – it's nearly always true."
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
        prs, page_num=16,
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
    """Rivian hiring scenario: ~$80k R1T, ~$40k materials, fixed K.

    Replaces the two Tesla photos (car carrier + Model 3 EMF) with a
    single Rivian R1T photo at the right, bullets at the left.  Layout
    matches the rest of the deck: hero image on one side, narrative on
    the other, gold takeaway bar at the bottom.
    """
    def draw(slide):
        # Bullets on the LEFT  (materials cost dropped per user decision —
        # we'll handle materials separately in the cost-side of the module).
        bullets = [
            ("Demand and output price are given", 0),
            ("Large number of R1T ordered at price of ~$80k", 1),
            ("(average transaction price, 2024–25)", 2),
            ("Short run:  capital (factory size, robots) is fixed", 0),
            ("The only way to expand production is to hire more workers", 0),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(8.0), height=Inches(4.4),
            items=bullets,
            size=22, sub_size=18, line_spacing_pts=8,
        )

        # Rivian R1T picture on the RIGHT (replaces the Tesla car carrier).
        rivian = OUT_DIR / "_rivian.jpg"
        if rivian.exists():
            pic = slide.shapes.add_picture(
                str(rivian),
                int(Inches(8.55)), int(Inches(1.95)),
                width=int(Inches(4.30)), height=int(Inches(3.0)),
            )
            _apply_picture_style(pic)
            # Small attribution
            _add_text(slide, Inches(8.55), Inches(5.05),
                       Inches(4.30), Inches(0.20),
                       "Rivian R1T  (CC BY-SA, Wikimedia)",
                       size=9, italic=True, color=GRAY,
                       align=PP_ALIGN.CENTER, font="Calibri")

        # Gold takeaway bar at the bottom – the key question
        _add_takeaway_bar(slide,
                           "How many workers should Rivian optimally hire?",
                           top=Inches(6.45), fill=GOLD, text_color=NAVY,
                           width=Inches(10.5))

    s = make_diagram_slide(
        prs, page_num=17,
        section_tag=SECTION_TAG_P1,
        title="Rivian Hiring Scenario:  ~$80k R1T, Fixed K",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "Setup for the next several slides.  We'll derive Rivian's optimal "
        "hiring level.  Given: Rivian R1T sells at ~$80K (average "
        "transaction price reported by Rivian for 2024–25), and capital "
        "(the plant and robot count) is fixed in the short run.  Materials "
        "and other variable costs are deliberately set aside here – we "
        "handle them on the cost side of the module.  Question: how many "
        "workers should Rivian hire?"
    ))


def slide_16(prs):
    """MRPL concept – merged from source slides 17 and 18.

    Establishes that this is a SHORT-RUN concept (K fixed) and gives the
    proper textbook definition MRPL = MR × MPL.  For a price-taker firm
    MR ≈ P, so MRPL ≈ P × MPL.  Materials are NOT netted out here – they
    belong on the cost side of the module (user decision: drop materials
    from the MRPL framing to avoid confusing MC = marginal cost).
    """
    def _add_styled_box(slide, left, top, width, height, *,
                          label, fill, text_color, size, corner_adj=0.06,
                          shadow_alpha=50000):
        """Filled, slightly-rounded box with a soft drop shadow.

        corner_adj small (0.05-0.08) gives a barely-rounded rectangle;
        shadow_alpha=50000 is 50% black at 4pt blur, 3pt offset.
        """
        left, top, width, height = int(left), int(top), int(width), int(height)
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                       left, top, width, height)
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
        shp.line.fill.background()
        shp.shadow.inherit = False
        try: shp.adjustments[0] = corner_adj
        except Exception: pass
        # Drop shadow via XML
        spPr = shp._element.spPr
        for old in spPr.findall(qn('a:effectLst')):
            spPr.remove(old)
        effectLst = ET.SubElement(spPr, qn('a:effectLst'))
        outerShdw = ET.SubElement(effectLst, qn('a:outerShdw'))
        outerShdw.set('blurRad', '50800')
        outerShdw.set('dist', '38100')
        outerShdw.set('dir', '2700000')
        outerShdw.set('algn', 'tl')
        outerShdw.set('rotWithShape', '0')
        rgb = ET.SubElement(outerShdw, qn('a:srgbClr'))
        rgb.set('val', '000000')
        alpha = ET.SubElement(rgb, qn('a:alpha'))
        alpha.set('val', str(int(shadow_alpha)))
        # Label centred
        tf = shp.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.15); tf.margin_right = Inches(0.15)
        tf.margin_top = Inches(0.05); tf.margin_bottom = Inches(0.05)
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = label
        r.font.name = "Calibri"
        r.font.size = Pt(size)
        r.font.bold = True
        r.font.color.rgb = text_color
        return shp

    def _add_shadow(shp, alpha=50000):
        spPr = shp._element.spPr
        for old in spPr.findall(qn('a:effectLst')):
            spPr.remove(old)
        effectLst = ET.SubElement(spPr, qn('a:effectLst'))
        outerShdw = ET.SubElement(effectLst, qn('a:outerShdw'))
        outerShdw.set('blurRad', '50800')
        outerShdw.set('dist', '38100')
        outerShdw.set('dir', '2700000')
        outerShdw.set('algn', 'tl')
        outerShdw.set('rotWithShape', '0')
        rgb = ET.SubElement(outerShdw, qn('a:srgbClr'))
        rgb.set('val', '000000')
        a = ET.SubElement(rgb, qn('a:alpha'))
        a.set('val', str(int(alpha)))

    def draw(slide):
        # Short-run framing (italic navy) — sets the scope
        _add_text(slide, MARGIN, Inches(1.85), RULE_W, Inches(0.35),
                   "In the short run  (capital K fixed):",
                   size=18, italic=True, bold=True, color=NAVY,
                   align=PP_ALIGN.CENTER, font="Calibri")

        # ---- HERO box (two-line): name on top, plain-English below ----
        hero_x = Inches(1.0); hero_y = Inches(2.25)
        hero_w = Inches(11.3); hero_h = Inches(1.30)
        hero = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                        int(hero_x), int(hero_y),
                                        int(hero_w), int(hero_h))
        hero.fill.solid(); hero.fill.fore_color.rgb = NAVY
        hero.line.fill.background()
        hero.shadow.inherit = False
        try: hero.adjustments[0] = 0.06
        except Exception: pass
        _add_shadow(hero)
        htf = hero.text_frame
        htf.vertical_anchor = MSO_ANCHOR.MIDDLE
        htf.margin_left = Inches(0.15); htf.margin_right = Inches(0.15)
        htf.margin_top = Inches(0.05);  htf.margin_bottom = Inches(0.05)
        htf.word_wrap = True
        # Concept-accent blue for the concept name (per course CLAUDE.md);
        # remaining text white bold.
        CONCEPT_ACCENT = RGBColor(0x9E, 0xC5, 0xF7)   # soft light-blue
        p1 = htf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        r0 = p1.add_run(); r0.text = "MRPL  =  "
        r0.font.name = "Calibri"; r0.font.size = Pt(28); r0.font.bold = True
        r0.font.color.rgb = WHITE
        r1 = p1.add_run(); r1.text = "Marginal Revenue Product of Labor"
        r1.font.name = "Calibri"; r1.font.size = Pt(28); r1.font.bold = True
        r1.font.color.rgb = CONCEPT_ACCENT
        p2 = htf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(4)
        r = p2.add_run()
        r.text = "the extra revenue from one more worker"
        r.font.name = "Calibri"; r.font.size = Pt(17); r.font.bold = False
        r.font.italic = True
        r.font.color.rgb = WHITE

        # ---- DECOMPOSITION box right below the HERO ----
        # Header → MR/MPL definitions → three bullets, all inside one box.
        dec_x = Inches(1.0); dec_y = Inches(3.70)
        dec_w = Inches(11.3); dec_h = Inches(2.85)
        dec = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                       int(dec_x), int(dec_y),
                                       int(dec_w), int(dec_h))
        dec.fill.solid(); dec.fill.fore_color.rgb = RGBColor(0xFD, 0xF6, 0xE6)
        dec.line.color.rgb = NAVY
        dec.line.width = Pt(0.75)
        dec.shadow.inherit = False
        try: dec.adjustments[0] = 0.06
        except Exception: pass
        _add_shadow(dec)
        dtf = dec.text_frame
        dtf.vertical_anchor = MSO_ANCHOR.TOP
        dtf.margin_left = Inches(0.30); dtf.margin_right = Inches(0.30)
        dtf.margin_top = Inches(0.12);  dtf.margin_bottom = Inches(0.08)
        dtf.word_wrap = True
        # Header
        ph = dtf.paragraphs[0]
        ph.alignment = PP_ALIGN.CENTER
        rh = ph.add_run()
        rh.text = "Decomposition:   MRPL  =  MR × MPL"
        rh.font.name = "Calibri"; rh.font.size = Pt(18); rh.font.bold = True
        rh.font.color.rgb = NAVY
        # MR definition (centred, italic, smaller)
        pmr = dtf.add_paragraph()
        pmr.alignment = PP_ALIGN.CENTER
        pmr.space_before = Pt(4)
        rmr = pmr.add_run()
        rmr.text = "MR:  marginal revenue from selling an extra item"
        rmr.font.name = "Calibri"; rmr.font.size = Pt(13)
        rmr.font.italic = True
        rmr.font.color.rgb = NAVY
        # MPL definition (centred)
        pmpl = dtf.add_paragraph()
        pmpl.alignment = PP_ALIGN.CENTER
        pmpl.space_before = Pt(2)
        rmpl = pmpl.add_run()
        rmpl.text = "MPL:  extra output (marginal product) from hiring one more worker"
        rmpl.font.name = "Calibri"; rmpl.font.size = Pt(13)
        rmpl.font.italic = True
        rmpl.font.color.rgb = NAVY
        # Bullet 1
        pb1 = dtf.add_paragraph()
        pb1.alignment = PP_ALIGN.LEFT
        pb1.space_before = Pt(10)
        rb = pb1.add_run()
        rb.text = "•  When MPL falls, MRPL falls"
        rb.font.name = "Calibri"; rb.font.size = Pt(15); rb.font.bold = True
        rb.font.color.rgb = NAVY
        # Sub of bullet 1
        pb1s = dtf.add_paragraph()
        pb1s.alignment = PP_ALIGN.LEFT
        pb1s.space_before = Pt(2)
        rb = pb1s.add_run()
        rb.text = "      Even at constant price, each additional worker is worth less"
        rb.font.name = "Calibri"; rb.font.size = Pt(13); rb.font.italic = True
        rb.font.color.rgb = NAVY
        # Bullet 2 (trimmed for 5–10 word target)
        pb2 = dtf.add_paragraph()
        pb2.alignment = PP_ALIGN.LEFT
        pb2.space_before = Pt(6)
        rb = pb2.add_run()
        rb.text = "•  Decreasing MPL  ⇒  each marginal hire is worth less"
        rb.font.name = "Calibri"; rb.font.size = Pt(15); rb.font.bold = True
        rb.font.color.rgb = NAVY
        # Bullet 3 — price-taker simplification (trimmed)
        pb3 = dtf.add_paragraph()
        pb3.alignment = PP_ALIGN.LEFT
        pb3.space_before = Pt(6)
        rb = pb3.add_run()
        rb.text = "•  Price-taker case:  MR = P,  so  MRPL = P × MPL"
        rb.font.name = "Calibri"; rb.font.size = Pt(15); rb.font.bold = True
        rb.font.color.rgb = NAVY

        # ---- DECISION RULE box at the bottom (GOLD) ----
        dr_x = Inches(1.0); dr_y = Inches(6.65)
        dr_w = Inches(11.3); dr_h = Inches(0.45)
        dr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      int(dr_x), int(dr_y),
                                      int(dr_w), int(dr_h))
        dr.fill.solid(); dr.fill.fore_color.rgb = GOLD
        dr.line.fill.background()
        dr.shadow.inherit = False
        try: dr.adjustments[0] = 0.06
        except Exception: pass
        _add_shadow(dr)
        drtf = dr.text_frame
        drtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        drtf.margin_left = Inches(0.15); drtf.margin_right = Inches(0.15)
        drtf.margin_top = Inches(0.05);  drtf.margin_bottom = Inches(0.05)
        drtf.word_wrap = True
        pdr = drtf.paragraphs[0]
        pdr.alignment = PP_ALIGN.CENTER
        rdr = pdr.add_run()
        rdr.text = "Decision rule:   If  MRPL > w (wage),   hire more workers"
        rdr.font.name = "Calibri"; rdr.font.size = Pt(22); rdr.font.bold = True
        rdr.font.color.rgb = NAVY

    s = make_diagram_slide(
        prs, page_num=18,
        section_tag=SECTION_TAG_P1,
        title="Hiring in the Short Run:  An Important Concept",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "MRPL — Marginal Revenue Product of Labor.  In plain terms: how "
        "much extra revenue does one more worker produce?  The textbook "
        "definition is MR × MPL — marginal revenue per unit times the "
        "marginal product of labor.  For a firm small enough that one "
        "more truck doesn't move the market price (price-taker), MR ≈ P, "
        "so MRPL ≈ P × MPL.  Since MPL falls as L grows (diminishing "
        "returns), MRPL also falls — which is why every firm has a finite "
        "optimal hiring level.  We're staying in the short run here: "
        "capital K is fixed, so the only lever is L.  Materials and other "
        "variable costs come back in the cost-side of the module."
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
        prs, page_num=19,
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
            "Rivian:  100 robots and 6,000 employees in its Normal, IL plant",
            "Price ~$80k per R1T,  approximately constant in quantity",
            ("What is MRPL?  (in $ per worker, per week)", 0),
            ("Hint:", 0),
            ("MRPL  =  MR × MPL  ≈  P × MPL", 1),
        ]
        normalized = [(b, 0) if isinstance(b, str) else b for b in bullets]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(7.8), height=Inches(4.4),
            items=normalized,
            size=22, sub_size=18, line_spacing_pts=10,
        )

        # ---- Compact production-function table (same data as slide 10) ----
        _add_compact_pf_table(slide,
                               tbl_left=Inches(9.55), tbl_top=Inches(2.30))

        # Discussion-break badge (rounded-parallelogram, bottom-right)
        _add_discussion_break(slide, top=Inches(6.55), width=Inches(4.8))

    s = make_diagram_slide(
        prs, page_num=19,
        section_tag=SECTION_TAG_P1,
        title="Example:  Calculate MRPL at 6,000 Employees and 100 Robots",
        draw_diagram=draw,
    )
    _set_notes(s, (
        "A specific number to anchor the concept. Walk through it: at 6,000 "
        "workers and 100 robots, find MPL from the production-function "
        "table, then multiply by the sale price (~$80K per R1T)."
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
        prs, page_num=20,
        section_tag=SECTION_TAG_P1,
        title="What Is Rivian's MRPL at 6,000 Employees?",
        draw_diagram=draw,
    )
    _draw_poll_pill(s)
    _set_notes(s, (
        "Quick PollEv – compute the MRPL at 6,000 employees and submit. "
        "Give them 30 seconds. The point isn't to get the number perfectly, "
        "it's to make sure everyone is doing the calculation in their head."
    ))


def slide_20(prs):
    """Solution: MRPL of Rivian — uses actual values from slide 10's table.

    Applies slide 13's "initial-point" convention: ΔQ and ΔL are computed
    relative to the previous row of the L grid.  Currently at L = 6,000
    workers; the next 1,000-worker step takes us to L = 7,000.
    """
    bullets = [
        ("Currently at L = 6,000 workers (K = 100 robots)", 0),
        ("From the production-function table on slide 10:", 0),
        ("Q (6,000)  =  387 R1T per week", 1),
        ("Q (7,000)  =  418 R1T per week", 1),
        ("Per the slide-13 convention:  MPL  =  ΔQ / ΔL", 0),
        ("MPL  =  (418 − 387) / 1,000  =  0.031 cars per worker per week", 1),
        ("MRPL  =  MPL × P  ≈  0.031 × $80,000  ≈  $2,480 per worker per week", 0),
        ("Compare to the weekly wage of ONE additional worker — not the whole workforce", 1),
    ]
    s = make_content_bulleted(
        prs, page_num=21,
        section_tag=SECTION_TAG_P1,
        title="Solution:  MRPL of Rivian",
        bullets=bullets,
        size=20, sub_size=18, line_spacing_pts=8,
    )
    _set_notes(s, (
        "Reveal the answer.  Pull Q(6,000) = 387 and Q(7,000) = 418 from "
        "slide 10's production-function table (K = 100 robots).  Apply "
        "the slide-13 convention: ΔQ = 31, ΔL = 1,000, so MPL ≈ 0.031 "
        "cars per worker per week.  At a sale price of ~$80K per R1T, "
        "that gives MRPL ≈ $2,480 per worker per week.  The most common "
        "slip in this calculation is comparing MRPL to the TOTAL wage "
        "bill instead of the weekly wage of ONE more worker."
    ))


def slide_21(prs):
    """Hire when MRPL > wage; stop when MRPL = wage."""
    bullets = [
        ("Should Rivian hire more workers?", 0),
        ("Suppose the weekly gross wage is $1,400 per worker", 1),
        ("Yes — MRPL > wage", 1),
        ("Hiring one more worker:", 0),
        ("Revenue rises by MRPL", 1),
        ("Wage bill rises by w", 1),
        ("Profit rises whenever MRPL > w", 1),
    ]
    s = make_content_bulleted(
        prs, page_num=22,
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
    """The optimal hiring rule: MRPL = w   (merged with old slide 21).

    Bullets on the left walk through the marginal-analysis logic;
    the chart on the right shows MRPL declining with L, the wage as
    a horizontal line, and the intersection at L*.
    """
    import math
    def draw(slide):
        # ---- Bullets (merged from slides 21 + 22) ----
        bullets = [
            ("Should Rivian hire more workers?", 0),
            ("Suppose the weekly wage is $1,400 per worker", 1),
            ("Hiring one more worker:", 0),
            ("Revenue rises by MRPL", 1),
            ("Wage bill rises by w", 1),
            ("Profit rises whenever MRPL > w", 1),
            ("Optimum:  hire L*  where  MRPL = w", 0),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(6.0), height=Inches(4.30),
            items=bullets,
            size=20, sub_size=16, line_spacing_pts=8,
        )

        # ---- Native MRPL / wage chart on the right ----
        # Continuous MPL formula for K = 100:  MPL = 0.25·√K / √L  =  2.5/√L
        # MRPL = $80,000 × MPL = $200,000 / √L  (in dollars / worker / week).
        # L grid: 1k, 3k, 5k, …, 25k → 13 points.  Wage $1,400 constant.
        L_vals = list(range(1000, 25001, 2000))
        K_FIX = 100
        mrpl_vals = [int(round(80000 * 0.25 * math.sqrt(K_FIX) / math.sqrt(L)))
                      for L in L_vals]
        wage_vals = [1400] * len(L_vals)
        cats = [f"{L//1000}k" for L in L_vals]

        _make_multi_line_chart(
            slide,
            Inches(6.55), Inches(1.85),
            Inches(6.50), Inches(4.30),
            categories=cats,
            series=[
                ("MRPL", mrpl_vals, NAVY, 'circle'),
                ("Wage (w)", wage_vals, GOLD, 'square'),
            ],
            x_title="L   (workers)",
            y_title="$ per worker per week",
            y_min=0, y_max=7000, y_unit=1000,
            legend_pos=('0.74', '0.06', '0.22', '0.20'),
        )

        # ---- Bottom: MB = MC anchor + rule statement ----
        star_w = Inches(1.6)
        star_h = Inches(1.05)
        star_x = MARGIN
        star_y = Inches(6.30)
        _add_anchor_burst(
            slide, star_x, star_y, star_w, star_h,
            top_text="MB = MC",
            bottom_text="(of labor)",
            top_size=14, bottom_size=11,
        )

        bar_x = star_x + star_w + Inches(0.25)
        bar_w = Inches(10.6)
        _add_filled_box(
            slide, bar_x, Inches(6.50), bar_w, Inches(0.55),
            "Optimum:  L*  where  MRPL  =  w",
            fill=NAVY, text_color=WHITE, size=20, bold=True,
        )
        _add_arrow(slide,
                    (star_x + star_w, star_y + star_h // 2),
                    (bar_x, Inches(6.50) + Inches(0.275)),
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
        title="Long Run:  Rivian Builds a New Georgia Plant",
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
            pic = slide.shapes.add_picture(
                str(rivian),
                int(Inches(8.7)), int(Inches(2.4)),
                width=int(Inches(4.3)),
            )
            _add_drop_shadow(pic)
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
        "Rivian announced a second US assembly plant in Social Circle, "
        "Georgia, in late 2022 – a ~$5B project meant to add ~400,000 "
        "vehicles per year of capacity once it ramps.  We'll apply the "
        "bang-for-the-buck rule to a stylised version of those plans:  "
        "robots vs. workers,  given each input's price and marginal "
        "product.  The numbers I'll use are illustrative, calibrated to "
        "give a clean teaching example;  the strategic point is what to "
        "do when the two MP/price ratios are not yet equal."
    ))


def slide_36(prs):
    """Is Rivian's current plan optimal? (production function).

    Uses the SAME compact production-function table as slide 18, so all
    downstream calculations (MPL, MPK, the bang-for-the-buck ratio on
    slide 39) read off one consistent data source.  At K = 200, L = 5,000
    the table gives Q = 500 vehicles/week – matching the slide narrative.
    """
    def draw(slide):
        bullets = [
            ("The production function at Rivian's new plant:", 0),
            ("Current mix to produce 500 vehicles/week:  200 robots, 5,000 workers", 0),
            ("Read Q at K = 200, L = 5,000 in the table  →  Q = 500", 1),
            ("Other input mixes are possible — which is best?", 0),
        ]
        _add_hierarchical_bullets(
            slide,
            left=MARGIN, top=Inches(1.85),
            width=Inches(7.0), height=Inches(3.5),
            items=bullets,
            size=20, sub_size=18, line_spacing_pts=10,
        )

        # Same compact production-function table as slide 18 — locks the
        # downstream calculations to the slide-10 data.
        _add_compact_pf_table(slide,
                               tbl_left=Inches(9.55), tbl_top=Inches(2.30))

        _add_takeaway_bar(slide,
                           "Compare MP per dollar across inputs at the current mix",
                           top=Inches(6.55), fill=NAVY, text_color=WHITE,
                           width=Inches(11.0))

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
        "Quick PollEv.  Looking at the numbers on the previous slide – "
        "200 robots and 5,000 workers producing 500 R1Ts/week, with the "
        "MP values given – is the current mix optimal?  Give them 30 "
        "seconds to think through the bang-for-the-buck ratios.  Some "
        "will say yes, some no;  reveal in the next slide.  The point "
        "isn't the vote count, it's the active calculation."
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
        title="Three Cost Types,  Three Different Decision Rules",
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

        # Cost breakdown — 4 cost boxes positioned UNDER the "Your own car"
        # option (per-mile costs only apply to that scenario, not to the
        # company-car alternative).
        _add_text(slide, start_x, Inches(4.05), opt_w, Inches(0.35),
                  "Costs associated with your car  (per mile driven):",
                  size=14, italic=True, color=GRAY, font="Calibri",
                  align=PP_ALIGN.CENTER)
        costs = [
            ("20¢", "insurance"),
            ("20¢", "maintenance"),
            ("15¢", "electricity"),
            ("45¢", "lease on the vehicle"),
        ]
        # 2 × 2 grid under the LEFT option box (Your own car)
        cost_w = (opt_w - Inches(0.10)) // 2
        cost_h = Inches(0.75)
        for i, (amt, lbl) in enumerate(costs):
            row, col = divmod(i, 2)
            cx = start_x + col * (cost_w + Inches(0.10))
            cy = Inches(4.45) + row * (cost_h + Inches(0.10))
            _add_filled_box(
                slide, cx, cy, cost_w, cost_h,
                f"{amt}   {lbl}",
                fill=NAVY, text_color=WHITE,
                size=15, bold=True,
            )

        # Question + Discussion-break badge in the corner
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
        # Three decision points during Waterworld's production – using the
        # original-deck data (sunk vs. expected-additional cost at each
        # point, with $150M expected revenue throughout).
        # Header (date) | Sunk | Additional | Overall | Revenue | Profit | Decision
        scenarios = [
            ("June 1994",       16,  84, 100, 150, "+50", "Make!"),
            ("September 1994", 100,  40, 140, 150, "+10", "Make!"),
            ("December 1994",  140,  35, 175, 150, "−25", "Make!"),
        ]
        col_w = Inches(3.9)
        col_gap = Inches(0.2)
        col_x0 = (SLIDE_W - col_w * 3 - col_gap * 2) // 2

        # Small row-label column on the LEFT (outside the 3-column grid)
        row_labels = [
            "",
            "Sunk cost  ($M)",
            "Expected additional cost  ($M)",
            "Overall cost incl. sunk  ($M)",
            "Expected revenue  ($M)",
            "Expected profit  ($M)",
            "Decision",
        ]
        # The 3-column scenario block fits; row labels appear as a faint
        # caption strip running down between the divider and the leftmost
        # column.  Display them as italic gray text positioned to the left.
        label_w = Inches(0.05)   # not used – we render labels at the leftmost
        row_y = [Inches(1.85), Inches(2.40), Inches(3.10),
                  Inches(3.80), Inches(4.50), Inches(5.20),
                  Inches(5.90)]

        # Render the 3-column data
        for j, sc in enumerate(scenarios):
            x = col_x0 + (col_w + col_gap) * j
            # Column header (date band)
            _add_filled_box(
                slide, x, row_y[0], col_w, Inches(0.50),
                sc[0], fill=NAVY, text_color=WHITE,
                size=20, bold=True,
            )
            # 5 number rows
            cells = [str(sc[1]), str(sc[2]), str(sc[3]), str(sc[4]), sc[5]]
            for i, v in enumerate(cells):
                _add_outlined_box(
                    slide, x, row_y[i + 1], col_w, Inches(0.60),
                    v, fill=WHITE, line=NAVY, text_color=NAVY,
                    size=18, bold=False, line_w=1.0,
                )
            # Decision band (gold)
            _add_filled_box(
                slide, x, row_y[6], col_w, Inches(0.55),
                sc[6], fill=GOLD, text_color=NAVY,
                size=20, bold=True,
            )

        # Row-label captions on the left (width is already in EMU)
        label_w = col_x0 - MARGIN - Inches(0.10)
        for i, lbl in enumerate(row_labels):
            if not lbl: continue
            _add_text(slide, MARGIN, row_y[i], label_w,
                       Inches(0.60),
                       lbl, size=12, italic=True, color=GRAY,
                       align=PP_ALIGN.RIGHT,
                       anchor=MSO_ANCHOR.MIDDLE, font="Calibri")

        # Bottom takeaway
        _add_takeaway_bar(
            slide,
            "Sunk costs are sunk  —  continue whenever  revenue  >  additional cost",
            top=Inches(6.60), fill=NAVY, text_color=WHITE,
            width=Inches(11.5), size=18,
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
    """Dictionary of costs – native three-card cheat sheet.

    Same formulas and labels as the source image, restyled to the deck's
    NAVY/GOLD palette with OMML rendering for each headline formula.
    """
    def draw(slide):
        # Three-card row: Total / Average / Marginal cost.
        col_w = Inches(4.05)
        col_gap = Inches(0.15)
        col_x0 = (SLIDE_W - col_w * 3 - col_gap * 2) // 2
        hdr_y = Inches(1.85)
        hdr_h = Inches(0.55)
        formula_y = hdr_y + hdr_h + Inches(0.10)
        formula_h = Inches(0.95)
        subs_y = formula_y + formula_h + Inches(0.15)

        # Card 1 — Total Cost
        x = col_x0
        _add_filled_box(slide, x, hdr_y, col_w, hdr_h,
                        "Total Cost", fill=NAVY, text_color=WHITE,
                        size=22, bold=True)
        _add_math_equation(
            slide, x, formula_y, col_w, formula_h,
            _omml_text('TC = TFC + TVC'),
            size_pt=22, color=NAVY, fill=RGBColor(0xFD, 0xF6, 0xE6),
            line=NAVY,
        )
        _add_hierarchical_bullets(
            slide,
            left=x + Inches(0.15), top=subs_y,
            width=col_w - Inches(0.30), height=Inches(2.80),
            items=[
                ("TFC = Total Fixed Cost", 0),
                ("(ignore sunk costs)", 1),
                ("TVC = Total Variable Cost", 0),
            ],
            size=15, sub_size=12, line_spacing_pts=8,
        )

        # Card 2 — Average Cost
        x = col_x0 + col_w + col_gap
        _add_filled_box(slide, x, hdr_y, col_w, hdr_h,
                        "Average Cost", fill=NAVY, text_color=WHITE,
                        size=22, bold=True)
        _add_math_equation(
            slide, x, formula_y, col_w, formula_h,
            _omml_text('ATC = ') + _omml_frac(_omml_text('TC'), _omml_text('Q')),
            size_pt=22, color=NAVY, fill=RGBColor(0xFD, 0xF6, 0xE6),
            line=NAVY,
        )
        _add_hierarchical_bullets(
            slide,
            left=x + Inches(0.15), top=subs_y,
            width=col_w - Inches(0.30), height=Inches(2.80),
            items=[
                ("AFC = TFC / Q", 0),
                ("AVC = TVC / Q", 0),
                ("ATC = AFC + AVC", 0),
            ],
            size=15, sub_size=12, line_spacing_pts=8,
        )

        # Card 3 — Marginal Cost
        x = col_x0 + 2 * (col_w + col_gap)
        _add_filled_box(slide, x, hdr_y, col_w, hdr_h,
                        "Marginal Cost", fill=NAVY, text_color=WHITE,
                        size=22, bold=True)
        _add_math_equation(
            slide, x, formula_y, col_w, formula_h,
            _omml_text('MC = ') + _omml_frac(_omml_text('Δ') + _omml_text('TC'),
                                                _omml_text('Δ') + _omml_text('Q')),
            size_pt=22, color=NAVY, fill=RGBColor(0xFD, 0xF6, 0xE6),
            line=NAVY,
        )
        _add_hierarchical_bullets(
            slide,
            left=x + Inches(0.15), top=subs_y,
            width=col_w - Inches(0.30), height=Inches(2.80),
            items=[
                ("= ΔTVC / ΔQ  (TFC is constant)", 0),
                ("Derivative form:", 0),
                ("MC = dTC / dQ", 1),
                ("    = dTVC / dQ", 1),
            ],
            size=15, sub_size=12, line_spacing_pts=8,
        )

        _add_text(slide, MARGIN, Inches(6.60), RULE_W, Inches(0.4),
                  "Cheat sheet to refer back to for the rest of the module",
                  size=14, italic=True, color=GRAY, font="Calibri",
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
        # Reconstruct the original 10-K snippet layout: a row-label column
        # (rId3), a 2022 numbers column (rId4), a 2021 numbers column (rId5),
        # and two red "≈ TVC" / "≈ TC" overlay annotations (rId6 / rId7).
        labels_x = Inches(0.6)
        labels_y = Inches(1.90)
        labels_w = Inches(3.40)             # row labels
        labels_h = Inches(2.20)
        # Match piece-to-piece native aspect: image43 235×152 (~1.55:1)
        _add_source_image(slide, 50, "rId3",
                           left=labels_x, top=labels_y,
                           width=labels_w, height=labels_h,
                           shadow=False)

        # 2022 numbers column (image44, 173×489 ≈ 0.354:1)
        col_w = Inches(1.75)
        col_h = Inches(4.40)
        col22_x = labels_x + labels_w + Inches(0.10)
        _add_source_image(slide, 50, "rId4",
                           left=col22_x, top=labels_y,
                           width=col_w, height=col_h,
                           shadow=False)

        # 2021 numbers column (image45, 141×201 ≈ 0.70:1)
        col21_x = col22_x + col_w + Inches(0.05)
        col21_w = Inches(1.70)
        col21_h = Inches(1.85)
        _add_source_image(slide, 50, "rId5",
                           left=col21_x, top=labels_y,
                           width=col21_w, height=col21_h,
                           shadow=False)

        # "≈ TVC" red overlay near the COGS line (image230, rId6)
        _add_source_image(slide, 50, "rId6",
                           left=labels_x + Inches(0.30),
                           top=labels_y + Inches(0.95),
                           width=Inches(1.50), height=Inches(0.45),
                           shadow=False)
        # "≈ TC" red overlay near the Total-costs line (image240, rId7)
        _add_source_image(slide, 50, "rId7",
                           left=labels_x + Inches(0.30),
                           top=labels_y + Inches(1.80),
                           width=Inches(1.40), height=Inches(0.40),
                           shadow=False)

        # Right-side commentary — fixed vs variable mapping
        _add_text(slide, Inches(9.20), Inches(1.95), Inches(3.90), Inches(0.40),
                  "Mapping to textbook concepts",
                  size=18, italic=True, bold=True, color=NAVY,
                  font="Calibri")
        _add_filled_box(
            slide, Inches(9.20), Inches(2.45), Inches(3.90), Inches(0.65),
            "Cost of goods sold  ≈  TVC",
            fill=NAVY, text_color=WHITE, size=15, bold=True,
        )
        _add_filled_box(
            slide, Inches(9.20), Inches(3.20), Inches(3.90), Inches(0.65),
            "Total costs and expenses  ≈  TC",
            fill=NAVY, text_color=WHITE, size=15, bold=True,
        )
        _add_outlined_box(
            slide, Inches(9.20), Inches(3.95), Inches(3.90), Inches(0.70),
            "SG&A  ≈  Mix of fixed (rent, depreciation) and variable (wages)",
            fill=WHITE, line=NAVY, text_color=NAVY, size=12, bold=False,
            line_w=1.0,
        )
        _add_outlined_box(
            slide, Inches(9.20), Inches(4.75), Inches(3.90), Inches(0.70),
            "Interest expense  ≈  Fixed (long-term debt service)",
            fill=WHITE, line=NAVY, text_color=NAVY, size=12, bold=False,
            line_w=1.0,
        )

        _add_takeaway_bar(
            slide,
            "Every 10-K can be read as a fixed-vs-variable split",
            top=Inches(6.55), fill=GOLD, text_color=NAVY,
            width=Inches(10.5),
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

        # Clean ChatGPT logo on the right (replaces the busy phone-and-laptop
        # photo).  Logo from Wikimedia Commons (PD-ineligible-trademark).
        chatgpt = OUT_DIR / "_chatgpt_logo.png"
        if chatgpt.exists():
            pic = slide.shapes.add_picture(
                str(chatgpt),
                int(Inches(9.85)), int(Inches(2.40)),
                width=int(Inches(3.00)), height=int(Inches(1.70)),
            )
            # Logo: flat, no shadow / rounding (per CLAUDE.md exceptions)
        else:
            _add_source_image(slide, 51, "rId5",
                               left=Inches(9.6), top=Inches(2.0),
                               height=Inches(3.5))
        _add_text(slide, Inches(9.6), Inches(4.25), Inches(3.5), Inches(0.30),
                   "ChatGPT logo  (OpenAI;  PD on Wikimedia)",
                   size=10, italic=True, color=GRAY,
                   align=PP_ALIGN.CENTER, font="Calibri")

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
        "Quick PollEv.  Compute the marginal cost of adding the second "
        "user to a ChatGPT Team plan.  The common trap: students see "
        "the $25 Team rate and answer $25.  But the FIRST user gets "
        "re-priced from $20 to $25 when they switch plans — so the true "
        "MC of the 2nd user is $5 (the re-pricing) + $25 (the new fee) = "
        "$30.  This is the canonical example for MC ≠ AC.  Give them 30 "
        "seconds, then reveal the solution on the next slide."
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

        # Cost function (OMML).  Quadratic form fit to the Excel data:
        # TC = 800,000 + 200 · Q²  (Q in vehicles/week, TC in $).
        eq_xml = (
            _omml_text('TC') +
            _omml_text(' = ') +
            _omml_text('800,000') +
            _omml_text(' + ') +
            _omml_text('200') +
            _omml_text(' · ') +
            _omml_sup(_omml_run('Q'), _omml_text('2'))
        )
        _add_math_equation(
            slide, (SLIDE_W - Inches(9.0)) // 2, Inches(2.85),
            Inches(9.0), Inches(0.70),
            eq_xml, size_pt=24, color=WHITE, fill=NAVY,
        )

        # Native TC curve.  Y-axis in $K so labels stay clean (10–3,220).
        cats = [str(q) for q in COST_Q_VALS]
        tc_vals_K = [_cost_tc(q) / 1000 for q in COST_Q_VALS]
        _make_simple_line_chart(
            slide, Inches(2.50), Inches(3.75),
            Inches(8.30), Inches(2.65),
            categories=cats, values=tc_vals_K,
            line_color=NAVY,
            x_title="Q   (vehicles per week)",
            y_title="TC   ($K)",
            y_min=0, y_max=3500, y_unit=500,
        )

        _add_takeaway_bar(
            slide,
            "Fixed plus a convex quadratic term — cost rises faster than output",
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
    """Rivian's Georgia plant – Cost Components (TC = TFC + TVC).

    Native python-pptx chart with three series, driven by the same data
    as slides 55 and 57 (TC = 800k + 200·Q²).
    """
    def draw(slide):
        cats = [str(q) for q in COST_Q_VALS]
        tc_vals  = [_cost_tc(q)  / 1000 for q in COST_Q_VALS]
        tfc_vals = [COST_TFC     / 1000 for _ in COST_Q_VALS]
        tvc_vals = [_cost_tvc(q) / 1000 for q in COST_Q_VALS]

        _make_multi_line_chart(
            slide, Inches(0.50), Inches(1.85),
            Inches(12.30), Inches(4.55),
            categories=cats,
            series=[
                ("TC",  tc_vals,  NAVY,                        'circle'),
                ("TFC", tfc_vals, GOLD,                        'square'),
                ("TVC", tvc_vals, RGBColor(0xC0, 0x50, 0x4D),  'triangle'),  # warm red — distinct from TC
            ],
            x_title="Q   (vehicles per week)",
            y_title="Cost   ($K)",
            y_min=0, y_max=3500, y_unit=500,
            legend_pos=('0.10', '0.08', '0.18', '0.22'),
        )

        _add_takeaway_bar(
            slide,
            "Fixed costs dominate at low Q;  the quadratic TVC overtakes at scale",
            top=Inches(6.55), fill=NAVY, text_color=WHITE,
            width=Inches(11.5), size=18,
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
    """Rivian's Georgia plant – Per-Unit Costs (ATC, AVC, MC).

    Native chart, three series, same Excel data (TC = 800k + 200·Q²).
    Demonstrates: MC = 400Q − 200·dQ; AVC = 200·Q (linear rising);
    ATC = TFC/Q + 200·Q (U-shape).  MC crosses ATC at the ATC minimum.
    """
    def draw(slide):
        cats = [str(q) for q in COST_Q_VALS]
        # Per-unit values in $K so the y axis stays readable.
        atc_vals = [_cost_atc(q) / 1000 for q in COST_Q_VALS]
        avc_vals = [_cost_avc(q) / 1000 for q in COST_Q_VALS]
        mc_vals  = [_cost_mc(q)  / 1000 for q in COST_Q_VALS]

        _make_multi_line_chart(
            slide, Inches(0.50), Inches(1.85),
            Inches(12.30), Inches(4.55),
            categories=cats,
            series=[
                ("ATC", atc_vals, NAVY,                        'circle'),
                ("AVC", avc_vals, GOLD,                        'square'),
                ("MC",  mc_vals,  RGBColor(0xC0, 0x50, 0x4D),  'triangle'),  # warm red — distinct from ATC/AVC
            ],
            x_title="Q   (vehicles per week)",
            y_title="Per-unit cost   ($K)",
            y_min=0, y_max=90, y_unit=10,
            legend_pos=('0.78', '0.08', '0.18', '0.22'),
        )

        _add_takeaway_bar(
            slide,
            "MC crosses ATC at the ATC minimum  —  the textbook U-shape",
            top=Inches(6.55), fill=GOLD, text_color=NAVY,
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
        "Quick PollEv.  Given the iPhone 17 teardown numbers (display, "
        "logic board, chips, battery, casing, labour & assembly), what "
        "is the AVC of one unit?  Most students overestimate because "
        "they confuse retail price with marginal cost.  The teardown "
        "estimates put it around $580 — roughly half the retail price "
        "of ~$1,200.  That gap (price minus MC) is the contribution "
        "margin Apple keeps on each handset.  Reveal the answer next."
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
        title="Naïve Linear Cost Function:  Total Cost View",
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
        title="Naïve Linear Cost Function:  Per-Unit View",
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

        # Three plant-size U-shape SAC bands.  Each is drawn as a smooth
        # custGeom curve (two cubic Béziers meeting at the minimum) so the
        # parabolic feel of the original deck comes through, rather than
        # piecewise straight segments.
        def draw_sac(label, x_min, x_max, y_min, y_left, y_right,
                      color=GRAY):
            x_mid = (x_min + x_max) // 2
            bb_left = x_min
            bb_top = min(y_left, y_right)
            bb_right = x_max
            bb_bottom = y_min
            bb_w = bb_right - bb_left
            bb_h = bb_bottom - bb_top
            if bb_w <= 0 or bb_h <= 0:
                return
            # Convert to path-coord space (100000 × 100000)
            def px(real_x):
                return int(100000 * (real_x - bb_left) / bb_w)
            def py(real_y):
                return int(100000 * (real_y - bb_top) / bb_h)

            # Cubic-Bezier control points for a smooth U:
            #   - tangent at the endpoints is steeply VERTICAL (down on the
            #     left side, up on the right side) — produces the
            #     parabolic descent into / out of the trough;
            #   - tangent at the minimum is HORIZONTAL — flat bottom of U.
            # Previous placement (CP1 horizontal from P0) collapsed the
            # curve to an L-shape; this placement gives a proper U.
            P0L = (px(x_min),  py(y_left))
            PMD = (px(x_mid),  py(y_min))
            P3R = (px(x_max),  py(y_right))
            seg1_cp1 = (P0L[0] + (PMD[0] - P0L[0]) // 10,
                         P0L[1] + (PMD[1] - P0L[1]) * 7 // 10)
            seg1_cp2 = (PMD[0] - (PMD[0] - P0L[0]) * 3 // 10, PMD[1])
            seg2_cp1 = (PMD[0] + (P3R[0] - PMD[0]) * 3 // 10, PMD[1])
            seg2_cp2 = (P3R[0] - (P3R[0] - PMD[0]) // 10,
                         P3R[1] + (PMD[1] - P3R[1]) * 7 // 10)

            r_hex = f'{color[0]:02X}{color[1]:02X}{color[2]:02X}'
            custgeom_xml = (
                f'<a:custGeom xmlns:a="{A_NS}">'
                f'<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
                f'<a:rect l="0" t="0" r="100000" b="100000"/>'
                f'<a:pathLst>'
                f'<a:path w="100000" h="100000" fill="none" stroke="1">'
                f'<a:moveTo><a:pt x="{px(x_min)}" y="{py(y_left)}"/></a:moveTo>'
                f'<a:cubicBezTo>'
                f'<a:pt x="{seg1_cp1[0]}" y="{seg1_cp1[1]}"/>'
                f'<a:pt x="{seg1_cp2[0]}" y="{seg1_cp2[1]}"/>'
                f'<a:pt x="{px(x_mid)}" y="{py(y_min)}"/>'
                f'</a:cubicBezTo>'
                f'<a:cubicBezTo>'
                f'<a:pt x="{seg2_cp1[0]}" y="{seg2_cp1[1]}"/>'
                f'<a:pt x="{seg2_cp2[0]}" y="{seg2_cp2[1]}"/>'
                f'<a:pt x="{px(x_max)}" y="{py(y_right)}"/>'
                f'</a:cubicBezTo>'
                f'</a:path></a:pathLst>'
                f'</a:custGeom>'
            )

            shp = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                int(bb_left), int(bb_top), int(bb_w), int(bb_h),
            )
            shp.fill.background()
            shp.line.color.rgb = color
            shp.line.width = Pt(2.25)
            shp.shadow.inherit = False
            spPr = shp._element.spPr
            for old in spPr.findall(qn('a:prstGeom')):
                spPr.remove(old)
            custgeom = ET.fromstring(custgeom_xml)
            xfrm = spPr.find(qn('a:xfrm'))
            if xfrm is not None:
                xfrm.addnext(custgeom)
            else:
                spPr.insert(0, custgeom)

            # Label above the minimum
            _add_text(slide,
                       x_mid - Inches(0.9),
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

        # Lower-envelope LAC — passes EXACTLY through each SAC minimum
        # so the three U-curves visibly touch the envelope at their lowest
        # point.  Two straight segments connect the three (mid_x, y_min)
        # points;  pedagogically this is the textbook envelope rendering.
        b1_x_mid = (b1_x_min + b1_x_max) // 2
        b2_x_mid = (b2_x_min + b2_x_max) // 2
        b3_x_mid = (b3_x_min + b3_x_max) // 2
        _add_arrow(slide,
                    (b1_x_mid, b1_y_min),
                    (b2_x_mid, b2_y_min),
                    color=GOLD, weight_pt=3.0, head=False)
        _add_arrow(slide,
                    (b2_x_mid, b2_y_min),
                    (b3_x_mid, b3_y_min),
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


def build_deck(output_name="Module 3_clean.pptx"):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # Front matter
    slide_1(prs)
    slide_2(prs)
    slide_announcements(prs)     # page 3 — midterm logistics (reintroduced 2026-05-15)
    slide_3(prs)                 # page 4 onwards
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
    slide_mpl_data(prs)          # page 13 — MPL data (matches original slide 17)
    slide_13(prs)                # page 14 onwards
    slide_14(prs)
    slide_15(prs)
    slide_16(prs)
    # slide_17(prs)  — MERGED into slide_16; function kept for reference only
    slide_18(prs)
    slide_19(prs)
    slide_20(prs)
    # slide_21(prs)  — MERGED into slide_22; function kept for reference
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

    out = OUT_DIR / output_name
    prs.save(out)
    strip_unused_layouts(out)
    return out


if __name__ == "__main__":
    import sys
    name = sys.argv[1] if len(sys.argv) > 1 else "Module 3_clean.pptx"
    out = build_deck(name)
    print(f"Wrote {out}")
