"""Build the 'Roman Empire at its Height' slide into a temp deck for splicing."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import _build_Italy_Class1 as B  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402
from pptx.enum.text import PP_ALIGN  # noqa: E402

HERE = Path(__file__).parent


def main():
    prs = Presentation()
    prs.slide_width = B.SLIDE_W
    prs.slide_height = B.SLIDE_H
    s = B._blank_slide(prs)
    B._top_bar(s, "The Roman Empire")
    B._action_title(s, "The Roman Empire at Its Height (~117 AD)")
    B._place_image(s, HERE / "Images" / "roman_empire_117.png",
                   cx=int(B.SLIDE_W // 2), cy=int(Inches(4.15)),
                   max_w=int(Inches(11.6)), max_h=int(Inches(5.2)),
                   shadow=True, rounded=True)
    B._add_text(s, int(B.MARGIN), int(Inches(6.82)), int(B.RULE_W),
                int(Inches(0.3)),
                "Map: Tataryn, CC BY-SA 3.0 (Wikimedia) — empire (red) and client states (pink) at Trajan’s death",
                size=11, italic=True, color=B.GRAY, font="Calibri",
                align=PP_ALIGN.CENTER)
    B._footer(s, 26)
    prs.save(HERE / "_empireslide.pptx")
    print("wrote _empireslide.pptx")


if __name__ == "__main__":
    main()
