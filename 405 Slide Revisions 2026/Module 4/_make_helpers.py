"""One-off: carve the reusable helper layer out of Module 1's build script
into `_m4_helpers.py`.

Module 1's `_build_Module1.py` carries the proven chrome / box / bullet /
OMML / chart / table / figure primitives (themselves inherited from
Module 7 and Module 3) in its first ~2900 lines, mixed in with Module-1
slide functions further down.  This script copies the generic part
verbatim and appends the named helpers that live further down, so Module 4
builds on exactly the same primitives without a cross-folder import.

Run once, by path.  Re-running overwrites `_m4_helpers.py`.
"""
import ast
import re
from pathlib import Path

SRC = Path(__file__).parent.parent / "Module 1" / "_build_Module1.py"
DST = Path(__file__).parent / "_m4_helpers.py"

# everything up to and including `_add_math_equation` is generic
CONTIGUOUS_END_FUNC = "_add_math_equation"

# named top-level nodes to lift out of the Module-1 section below it
WANTED = [
    "CREAM", "DIM", "RED", "RED_FF", "GREEN_DK", "BLUE_PED", "STEEL",
    "DARKRED", "NB_BLUE",
    "_add_slidenum_field", "_draw_footer", "_add_media_image",
    "_set_cell_borders", "_add_styled_table",
    "POLLBREAK_XY", "POLLBREAK_WH", "POLLBREAK_SLANT",
    "_PB_S", "_PB_R", "_PB_D", "_PB_PATH", "_PB_SHADOW", "_PB_XML",
    "_add_pollbreak_badge", "make_stub",
    "SimpleFig", "_fig_axes", "_fig_line", "_fig_guide", "_fig_ylab",
    "_fig_xlab", "_fig_curve_label",
    "content_slide", "_highlight_texts", "_link_runs",
    "JUMP_BTN_FILL", "JUMP_BTN_GLYPH", "EXT_BTN_FILL", "EXT_BTN_GLYPH",
    "EXT_LINK_SHAPES",
    "_add_jump_button", "_add_ext_link_button", "_add_jump_pill",
    "PS_GLYPH", "PS_BOX_XY", "_add_ps_pointer",
    "RED_MW", "TITLE_LOWER", "_tc_word", "_title_case",
    "DIM_DROP",
    "SYMBOL_RE", "SUBSCRIPT_BASELINE", "_split_symbol_runs",
    "apply_symbol_subscripts", "_iter_text_frames",
]

BANNER = '''# ==========================================================================
#  _m4_helpers.py — shared primitive layer for the Module 4 rebuild.
#
#  Carved out of "Module 1/_build_Module1.py" by _make_helpers.py; that
#  file in turn carries the Module 7 / Module 3 helper layer verbatim.
#  Nothing here is Module-4 specific — palette, chrome, boxes, bullets,
#  OMML math, charts, tables, figures, badges, pointers, title case, and
#  the deck-wide symbol-subscript pass.
#
#  Module-4 content lives in _build_Module4.py.
# ==========================================================================

import uuid          # used by the live slide-number field footer

'''


def node_src(lines, node):
    start = min([node.lineno] + [d.lineno for d in
                                 getattr(node, "decorator_list", [])]) - 1
    # include a preceding comment block
    while start > 0 and lines[start - 1].lstrip().startswith("#"):
        start -= 1
    return "".join(lines[start:node.end_lineno])


def main():
    src = SRC.read_text(encoding="utf-8")
    lines = src.splitlines(keepends=True)
    tree = ast.parse(src)

    end_node = next(n for n in tree.body
                    if isinstance(n, ast.FunctionDef)
                    and n.name == CONTIGUOUS_END_FUNC)
    head = "".join(lines[: end_node.end_lineno])

    # drop the Module-1 slide-order import and its comment
    head = re.sub(r"# Canonical slide order.*?import _m1_order as _ORDER\n",
                  "", head, flags=re.S)
    # the docstring banner is Module-1's; replace it
    head = head[head.index("import copy"):]

    # LAST definition wins: Module 1 redefines _draw_footer further down to
    # emit a live <a:fld type="slidenum"> page number instead of static
    # text, and that override is the one Module 4 needs.
    by_name = {}
    for n in tree.body:
        if isinstance(n, (ast.FunctionDef, ast.ClassDef)):
            by_name[n.name] = n
        elif isinstance(n, ast.Assign):
            t = n.targets[0]
            if isinstance(t, ast.Name):
                by_name[t.id] = n

    tail_parts, missing = [], []
    for name in WANTED:
        node = by_name.get(name)
        if node is None or node.lineno <= end_node.end_lineno:
            if node is None:
                missing.append(name)
            continue
        tail_parts.append(node_src(lines, node))

    out = BANNER + head + "\n\n" + "\n\n".join(tail_parts) + "\n"
    DST.write_text(out, encoding="utf-8")
    print(f"wrote {DST.name}: {out.count(chr(10))} lines, "
          f"{len(tail_parts)} appended nodes")
    if missing:
        print("NOT FOUND:", ", ".join(missing))


if __name__ == "__main__":
    main()
