# -*- coding: utf-8 -*-
"""Map each display slide of the rebuilt deck to the SOURCE slide it came
from, by reading `main()` in _build_Module6.py and each slide function's
docstring tag ("[NV 45] · [V5 s4]", "[NVapp 48]", "[NEW – PG1 43]").

Read-only: it parses the build script with `ast` and never imports it, per
the Teaching CLAUDE.md rule about verifying passes statically.

    python _map_sources.py             # the whole deck
    python _map_sources.py 44 45 46    # only those display slides
"""
import ast
import io
import re
import sys
from pathlib import Path

SKIP = {"_raise_corner_marks", "_apply_written_notes",
        "_strip_unused_layouts", "apply_symbol_subscripts"}

HERE = Path(__file__).parent
SRC = HERE / "_build_Module6.py"

# "[NV 45]", "[NVapp 48]", "[V5 s4]", "[PG1 43]", "[PG2 16]"
TAG = re.compile(r"\[(?:NEW\s*[–—-]\s*)?"
                 r"(NVapp|NV|V\d|PG\d)\s*s?(\d+)\]")


def slide_functions(tree):
    """{function name: [(deck, slide), ...]} from each docstring."""
    out = {}
    for node in ast.walk(tree):
        if not isinstance(node, ast.FunctionDef):
            continue
        doc = ast.get_docstring(node) or ""
        out[node.name] = TAG.findall(doc)
    return out


def call_order(tree):
    """[(display_number, function name)] in the order main() calls them."""
    main = next((n for n in ast.walk(tree)
                 if isinstance(n, ast.FunctionDef) and n.name == "main"), None)
    if main is None:
        raise SystemExit("no main() in _build_Module6.py")
    order, n = [], 0
    for node in ast.walk(main):
        if not isinstance(node, ast.Call):
            continue
        f = node.func
        name = getattr(f, "id", None) or getattr(f, "attr", None)
        if not name or not name.startswith(("s_", "video_card", "make_m6_",
                                            "slide_")):
            continue
        order.append(name)
    return order


def main():
    tree = ast.parse(io.open(SRC, encoding="utf-8").read())
    docs = slide_functions(tree)
    # main() is written as one statement per slide, in display order, so the
    # statement list is the ordering -- ast.walk would not preserve it
    m = next(n for n in ast.walk(tree)
             if isinstance(n, ast.FunctionDef) and n.name == "main")
    seq = []
    for stmt in ast.walk(m):
        pass
    disp = 0
    for stmt in m.body:
        for node in ast.walk(stmt):
            if isinstance(node, ast.Call):
                name = getattr(node.func, "id", None)
                # the deck-wide post-processing passes at the end of
                # main() are calls too, and they make no slide
                if name in SKIP:
                    break
                if name in docs:
                    disp += 1
                    seq.append((disp, name, docs[name]))
                    break
    want = {int(a) for a in sys.argv[1:] if a.isdigit()}
    for d, name, tags in seq:
        if want and d not in want:
            continue
        srcs = ", ".join("%s %s" % (k, v) for k, v in tags) or "-"
        print("%3d  %-28s %s" % (d, name, srcs))
    print("# %d slides mapped" % len(seq))


if __name__ == "__main__":
    main()
