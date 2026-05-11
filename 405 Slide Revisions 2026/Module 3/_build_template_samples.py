"""
Build the 405 Slides Template.

Outputs "405 Slides Template.pptx" — a clean, MBA-style slide template
with a navy top bar, a section tag in white, an action-title format,
and two gold left-edge segments (under the title and above the footer)
forming a parallel rhythm. All elements live within 0.7 cm side margins.

Built from scratch with python-pptx using only fonts that ship with
Windows / Office. No external assets, no attribution required, no
watermarks.
"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

OUT_DIR = Path(__file__).parent

# Standard 16:9 slide
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

TITLE_TEXT = "Sunk costs should never drive decisions"
SECTION_TAG = "Module 3 · Part 2 · Costs"
BULLETS = [
    ("Fixed costs", "don't depend on quantity produced (Q)"),
    ("Sunk costs", "a fixed cost that cannot be recovered"),
    ("→ Decision rule", "ignore sunk costs when choosing what to do next"),
    ("Variable costs", "depend on volume produced (Q)"),
]
FOOTER = "Management 405  ·  Module 3  ·  Production and Costs"


def _blank_slide(prs):
    blank_layout = prs.slide_layouts[6]
    return prs.slides.add_slide(blank_layout)


def _add_text(slide, left, top, width, height, text, *,
              size=18, bold=False, color=RGBColor(0, 0, 0),
              font="Calibri", align=PP_ALIGN.LEFT, italic=False):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box


def _add_rect(slide, left, top, width, height, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
    shp.shadow.inherit = False
    return shp


def build_template():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    navy = RGBColor(0x0B, 0x2B, 0x4E)
    gray = RGBColor(0x55, 0x5B, 0x66)
    rule = RGBColor(0xC8, 0xCD, 0xD3)
    gold = RGBColor(0xE0, 0x9F, 0x3E)
    white = RGBColor(0xFF, 0xFF, 0xFF)

    # 0.7 cm margin for all main horizontal elements
    margin_cm = Inches(0.7 / 2.54)  # ≈ 0.276"
    rule_w = SLIDE_W - margin_cm * 2
    gold_w = Inches(2.2)

    slide = _blank_slide(prs)

    # Top navy bar
    bar_h = Inches(0.34)
    _add_rect(slide, 0, 0, SLIDE_W, bar_h, navy)

    # Section tag — white, vertically centered, aligned to the 0.7cm margin
    tag_box = _add_text(slide, margin_cm, 0, Inches(12), bar_h,
                        SECTION_TAG.upper(), size=11, bold=True,
                        color=white, font="Calibri")
    tag_box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # Action title — spans the 0.7cm margins, sits close to the top bar
    _add_text(slide, margin_cm, Inches(0.55), rule_w, Inches(0.7),
              TITLE_TEXT, size=32, bold=True, color=navy, font="Calibri")

    # Divider under title — gray rule with gold left-segment on top
    _add_rect(slide, margin_cm, Inches(1.25), rule_w,
              Inches(0.02), rule)
    _add_rect(slide, margin_cm, Inches(1.235), gold_w,
              Inches(0.05), gold)

    # Body — two-column "concept | definition" layout, spans 0.7cm margins
    top = Inches(1.75)
    row_h = Inches(0.95)
    concept_w = Inches(3.2)
    col_gap = Inches(0.2)
    def_x = margin_cm + concept_w + col_gap
    def_w = rule_w - concept_w - col_gap
    for i, (concept, defn) in enumerate(BULLETS):
        y = top + row_h * i
        _add_text(slide, margin_cm, y, concept_w, row_h,
                  concept, size=18, bold=True, color=navy)
        _add_text(slide, def_x, y, def_w, row_h,
                  defn, size=18, color=gray)

    # Footer rule + matching gold left-segment (mirrors the top)
    _add_rect(slide, 0, Inches(7.15), SLIDE_W, Inches(0.02), rule)
    _add_rect(slide, margin_cm, Inches(7.135), gold_w,
              Inches(0.05), gold)

    # Footer text
    _add_text(slide, margin_cm, Inches(7.22), Inches(11), Inches(0.3),
              FOOTER, size=10, color=gray)
    _add_text(slide, Inches(12.6), Inches(7.22), Inches(0.5), Inches(0.3),
              "1", size=10, color=gray, align=PP_ALIGN.RIGHT)

    prs.save(OUT_DIR / "405 Slides Template.pptx")


if __name__ == "__main__":
    build_template()
    print(f"Wrote {OUT_DIR / '405 Slides Template.pptx'}")
