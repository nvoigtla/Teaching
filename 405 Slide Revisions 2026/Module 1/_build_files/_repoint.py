# -*- coding: utf-8 -*-
"""One-off (2026-09-24): repoint every script in `_build_files/` after the
step-1b cleanup moved it down a level.

The convention, from the project CLAUDE.md:
  HERE   = the script's own folder (`_build_files/`) - `_source_images/`, the
           poll sidecars and the JSON plans moved with the scripts, so those
           references do not change.
  MODULE = HERE.parent - every reference to a DECK or to `Recorded Video
           Slides/` goes through it. In PowerShell,
           `$folder = Split-Path $PSScriptRoot -Parent`.

Writer passes keep pointing at the deleted `Module 1 - Revised.pptx`
deliberately, so an accidental run fails loudly instead of rewriting the
deliverable; they only need their OUT_DIR moved up.
"""
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent

# (file, old, new) - every substitution must hit exactly once
SUBS = [
    # --- the step-1 / step-2 pipeline: MODULE moves up ------------------
    ("_full_inventory.py",
     'MODULE = HERE                     # set to HERE.parent under _build_files/',
     'MODULE = HERE.parent              # _build_files/ since 2026-09-24'),
    ("_full_verify.py",
     'MODULE = HERE                     # set to HERE.parent under _build_files/',
     'MODULE = HERE.parent              # _build_files/ since 2026-09-24'),
    ("_final_inventory.py",
     'MODULE = HERE                     # set to HERE.parent under _build_files/',
     'MODULE = HERE.parent              # _build_files/ since 2026-09-24'),
    ("_fix_candidates_chrome.py",
     'MODULE = HERE                     # set to HERE.parent under _build_files/',
     'MODULE = HERE.parent              # _build_files/ since 2026-09-24'),
    # --- writer passes and builders: OUT_DIR / HERE move up -------------
    ("_animate.py", "HERE = Path(__file__).parent",
     "HERE = Path(__file__).parent.parent   # the module folder; _build_files/ since 2026-09-24"),
    ("_group_pass.py", "HERE = Path(__file__).parent",
     "HERE = Path(__file__).parent.parent   # the module folder; _build_files/ since 2026-09-24"),
    ("_splice_media.py", "HERE = Path(__file__).parent",
     "HERE = Path(__file__).parent.parent   # the module folder; _build_files/ since 2026-09-24"),
    ("_build_Module1.py", "OUT_DIR = Path(__file__).parent",
     "OUT_DIR = Path(__file__).parent.parent   # the module folder; _build_files/ since 2026-09-24"),
    ("_build_M1_candidates.py", "OUT_DIR = Path(__file__).parent",
     "OUT_DIR = Path(__file__).parent.parent   # the module folder; _build_files/ since 2026-09-24"),
    ("_build_template_samples.py", "OUT_DIR = Path(__file__).parent",
     "OUT_DIR = Path(__file__).parent.parent   # the module folder; _build_files/ since 2026-09-24"),
    ("_diff2.py", "HERE = os.path.dirname(os.path.abspath(__file__))",
     "HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))"),
    ("_fetch_web_images.py", "HERE = os.path.dirname(os.path.abspath(__file__))",
     "HERE = os.path.dirname(os.path.abspath(__file__))"),   # images moved too
    # --- PowerShell: $folder becomes the module folder ------------------
    ("_assemble_full.ps1", "$folder = $PSScriptRoot",
     "$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-24"),
    ("_assemble_final.ps1", "$folder = $PSScriptRoot",
     "$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-24"),
    ("_full_verify_anim.ps1", "$folder = $PSScriptRoot",
     "$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-24"),
    ("_final_verify_anim.ps1", "$folder = $PSScriptRoot",
     "$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-24"),
    ("_slideshow_probe.ps1", "$folder = $PSScriptRoot",
     "$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-24"),
    ("_slideshow_probe.ps1", '$outDir = Join-Path $folder "_probe"',
     '$outDir = Join-Path $PSScriptRoot "_probe"   # captures stay beside the script'),
    ("_export_probe.ps1", "$folder = $PSScriptRoot",
     "$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-24"),
]

# Header note required of any tool that reads a deck the cleanup deleted.
NOTE = ('# NOTE (2026-09-24): this reads "Module 1 - Revised.pptx", which the\n'
        '# step-1b cleanup deleted. To use it again, restore that deck from git\n'
        '# into the MODULE folder first:\n'
        '#   git checkout <commit> -- "405 Slide Revisions 2026/Module 1/Module 1 - Revised.pptx"\n')
NEEDS_NOTE = ["_ic_vs_rev.py", "_style_diff.py", "_xml_diff.py"]


def main():
    for name, old, new in SUBS:
        p = HERE / name
        s = p.read_text(encoding="utf8")
        if old == new:
            print("  unchanged  %s" % name)
            continue
        if new in s and old not in s:
            print("  already    %s" % name)
            continue
        n = s.count(old)
        if n != 1:
            raise SystemExit("%s: expected 1 occurrence of %r, found %d"
                             % (name, old[:48], n))
        p.write_text(s.replace(old, new, 1), encoding="utf8")
        print("  repointed  %-26s %s" % (name, new.split("#")[0].strip()[:46]))

    for name in NEEDS_NOTE:
        p = HERE / name
        s = p.read_text(encoding="utf8")
        if "step-1b cleanup deleted" in s:
            continue
        # insert after the module docstring
        m = re.search(r'^("""(?:.|\n)*?""")\n', s)
        if not m:
            raise SystemExit("%s: no module docstring to insert after" % name)
        s = s[:m.end()] + NOTE + s[m.end():]
        p.write_text(s, encoding="utf8")
        print("  noted      %-26s reads the deleted Revised deck" % name)


if __name__ == "__main__":
    main()
