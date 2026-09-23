# -*- coding: utf-8 -*-
"""Refuse to overwrite "Module 1 - Revised.pptx" from the retired pipeline.

**Module 1 - Revised.pptx IS THE SOURCE OF TRUTH** (2026-09-23, Nico).

On 2026-09-23 the in-class deck was merged into it: 22 slides were copied
across wholesale - including slide 66's 15-shape opportunity-cost table,
slide 83's 23-click exercise chart and slide 92's rebuilt backup slide -
and three poll slides were deleted, taking the deck from 95 to 92 slides.
None of that exists in `_build_Module1.py`, which still scaffolds the old
95-slide content. **Running the pipeline would silently discard the whole
merge**, and the rolling `_t-1` / `_t-2` backups only go back two saves.

This is the same failure that destroyed a morning's work in Module 3 on
2026-08-27, which is why Teaching CLAUDE.md now asks for a guard on every
pass that rewrites a deck. Here the guard is on the TARGET, not the script:
a pass may still write to a side path, so experimenting is unaffected - it
refuses only when it is about to clobber the canonical deck.

Override with `--force` on the command line, and only when you mean it.
"""
import sys
from pathlib import Path

CANONICAL = "module 1 - revised.pptx"
FORCE = "--force"

_WHY = """
REFUSING TO WRITE "%s".

  That deck is the SOURCE OF TRUTH for Module 1 (Nico, 2026-09-23).
  It carries the in-class merge - 22 slides copied across by hand,
  three deleted, 95 -> 92 slides - and NONE of it is in the build
  script. Running %s would discard the merge, and the rolling
  _t-1 / _t-2 backups only go back two saves.

  If you need to change Module 1, edit the .pptx (the way
  _slide33_edit.py / _slide50_shade.py do: targeted zip + lxml
  surgery), or build to a SIDE PATH and diff before promoting.

  Re-run with %s if you really mean to overwrite it.
"""


def refuse_if_canonical(target, argv=None):
    """Exit unless `target` is something other than the canonical deck.

    Returns True when the caller may proceed.
    """
    argv = sys.argv if argv is None else argv
    name = Path(target).name.lower()
    if name != CANONICAL:
        return True
    script = Path(argv[0]).name if argv and argv[0] else "this pass"
    if FORCE in argv:
        sys.stderr.write("WARNING: overwriting the canonical Module 1 deck "
                         "because %s was passed.\n" % FORCE)
        return True
    sys.exit(_WHY % (Path(target).name, script, FORCE))


def refuse_already_applied(what, argv=None):
    """For the one-off merge scripts: they have run, and re-running them
    would act on stale slide numbers."""
    argv = sys.argv if argv is None else argv
    if FORCE in argv:
        sys.stderr.write("WARNING: re-running %s because %s was passed.\n"
                         % (what, FORCE))
        return True
    sys.exit(
        '\n%s HAS ALREADY BEEN APPLIED (2026-09-23) and is kept as the\n'
        '  record of what was done. Its slide numbers refer to the deck as\n'
        '  it was BEFORE the merge, so re-running it would act on the wrong\n'
        '  slides. Pass %s only if you have restored a pre-merge deck.\n'
        % (what, FORCE))
