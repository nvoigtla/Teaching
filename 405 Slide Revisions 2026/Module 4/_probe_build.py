"""Scratch harness: render whatever slide functions are named on the
command line into _probe.pptx, then export them to PNG via PowerPoint COM.

Never writes the canonical deck.  Usage:
    python _probe_build.py slide_06_market_structures slide_07_market_power
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

import _build_Module4 as B

OUT = Path(__file__).parent / "_probe.pptx"


def main():
    names = sys.argv[1:]
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    for i, name in enumerate(names, 1):
        fn = getattr(B, name)
        try:
            fn(prs, i)
        except TypeError:
            fn(prs)
    prs.save(str(OUT))
    print(f"{len(prs.slides.__iter__.__self__._sldIdLst)} slides -> {OUT}")


if __name__ == "__main__":
    main()
