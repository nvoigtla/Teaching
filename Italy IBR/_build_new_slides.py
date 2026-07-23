"""
Build the 3 new economic-history slides into a temp deck (_newslides.pptx),
reusing the canonical deck's chrome/bullet helpers so they match exactly.
_add_slides.py then splices these into the canonical deck.

Order in temp deck: 1=Geography, 2=First Integrated Market, 3=The Fall.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import _build_Italy_Class1 as B  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR  # noqa: E402

HERE = Path(__file__).parent


def takeaway(slide, text):
    y, h = int(Inches(6.45)), int(Inches(0.52))
    B._add_rect(slide, int(B.MARGIN), y, int(B.RULE_W), h, B.GOLD)
    B._add_text(slide, int(B.MARGIN), y, int(B.RULE_W), h, text, size=18,
                bold=True, italic=True, color=B.NAVY, font="Calibri",
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def geo(prs):
    s = B._blank_slide(prs)
    B._top_bar(s, "Early History")
    B._action_title(s, "Mountains Push Italy Toward the Sea")
    B._hbullets(s, [
        (0, "A long peninsula — the Alps to the north, the Apennine spine down the middle"),
        (0, "Rugged terrain made overland transport slow and costly"),
        (0, "Far cheaper to move goods by sea → the economy turned outward to the Mediterranean"),
        (0, "The same mountains split the interior into many separate regions and cities"),
    ], left=int(B.MARGIN), top=int(Inches(1.85)), width=int(Inches(7.3)),
        height=int(Inches(4.1)))
    B._place_image(s, HERE / "Images" / "italy_relief.jpg",
                   cx=int(Inches(10.6)), cy=int(Inches(3.55)),
                   max_w=int(Inches(4.0)), max_h=int(Inches(3.9)),
                   shadow=True, rounded=False)
    B._add_text(s, int(Inches(8.7)), int(Inches(5.62)), int(Inches(3.9)),
                int(Inches(0.3)),
                "Relief map: E. Gaba & NordNordWest, CC BY-SA 3.0 (Wikimedia)",
                size=11, italic=True, color=B.GRAY, font="Calibri",
                align=PP_ALIGN.CENTER)
    takeaway(s, "Geography set two constants: openness to sea trade, and internal fragmentation")
    B._footer(s, 15)


def market(prs):
    s = B._blank_slide(prs)
    B._top_bar(s, "The Roman Empire")
    B._action_title(s, "Rome: The First Integrated Market")
    B._hbullets(s, [
        (0, "One currency (the denarius) + Roman roads + safe sea lanes"),
        (1, "→ a single Mediterranean market (“Mare Nostrum”)"),
        (0, "Roman law: property and enforceable contracts"),
        (1, "the foundation later European business would run on"),
        (0, "Publicani — tax-farming firms with tradeable shares: proto-corporations"),
        (0, "Grain, wine, oil, and luxuries move freely across the empire"),
    ], top=int(Inches(1.95)), height=int(Inches(4.0)))
    takeaway(s, "Prosperity came from integration — one market, common rules, infrastructure, security")
    B._footer(s, 24)


def fall(prs):
    s = B._blank_slide(prs)
    B._top_bar(s, "The Roman Empire")
    B._action_title(s, "The Fall = The Market Fragments")
    B._hbullets(s, [
        (0, "Overextension, currency debasement, and insecurity broke the integrated market"),
        (0, "Money, trade, and cities localized; Italy splintered into many small economies"),
        (0, "For ~1,000 years, prosperity would be rebuilt city by city"),
        (1, "→ next: the medieval communes"),
    ], top=int(Inches(2.1)), height=int(Inches(3.6)))
    takeaway(s, "What integration builds, dis-integration destroys")
    B._footer(s, 26)


def main():
    prs = Presentation()
    prs.slide_width = B.SLIDE_W
    prs.slide_height = B.SLIDE_H
    geo(prs)
    market(prs)
    fall(prs)
    prs.save(HERE / "_newslides.pptx")
    print("wrote _newslides.pptx (3 slides)")


if __name__ == "__main__":
    main()
