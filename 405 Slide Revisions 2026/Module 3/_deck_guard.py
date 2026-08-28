# -*- coding: utf-8 -*-
"""Refuse to rewrite a deck that git has no copy of.

Every pass in this folder edits `Module 3 - Revised.pptx` IN PLACE. The
rolling `_t-1` / `_t-2` backups are in .gitignore and were deleted when
the module was finished, so **git is the only way back**. If the deck
carries uncommitted changes - a hand edit made in PowerPoint, say - and a
pass runs over it, that work is gone for good.

So each pass calls `require_committed(DECK)` before it writes. It costs
one `git status` and turns "your hand edits are gone" into "commit first,
then re-run".

Why this exists (2026-08-27): `_retrofit_agenda.py` was executed
accidentally and rewrote all seven agenda slides and every top-bar tag.
The deck survived only because it happened to have been committed twenty
minutes earlier. That was luck, not a safety net.

Pass `--force` on the command line to override, for the case where you
deliberately want to re-run a pass over an uncommitted deck.
"""
import subprocess
import sys
from pathlib import Path

FORCE_FLAG = "--force"


def _git(args, cwd):
    try:
        p = subprocess.run(["git"] + args, cwd=str(cwd),
                           capture_output=True, text=True)
    except (OSError, ValueError) as e:      # git missing / bad args
        return None, str(e)
    if p.returncode != 0:
        return None, (p.stderr or p.stdout).strip()
    return p.stdout, None


def require_committed(deck, argv=None):
    """Exit unless `deck` is clean in git (or --force was passed).

    Returns True when the pass may proceed.
    """
    deck = Path(deck)
    argv = sys.argv if argv is None else argv
    forced = FORCE_FLAG in argv

    out, err = _git(["status", "--porcelain", "--", deck.name], deck.parent)

    if out is None:
        msg = "cannot ask git about %s (%s)" % (deck.name, err)
    elif out.strip():
        state = out.strip().splitlines()[0][:2].strip()
        if state == "??":
            msg = ("%s is NOT TRACKED by git - there is no copy to fall "
                   "back on if this pass gets it wrong" % deck.name)
        else:
            msg = ("%s has UNCOMMITTED CHANGES - a hand edit would be "
                   "destroyed by this pass with no way back" % deck.name)
    else:
        return True

    if forced:
        sys.stderr.write("WARNING: %s\n         proceeding anyway (%s)\n"
                         % (msg, FORCE_FLAG))
        return True

    sys.exit(
        "REFUSING TO RUN: %s.\n"
        "  This pass rewrites the deck in place and the _t-1 / _t-2 backups\n"
        "  are gitignored, so git is the only way back.\n"
        "  Commit the deck first:\n"
        "      git add \"%s\" && git commit -m \"...\"\n"
        "  or re-run with %s if you mean to overwrite it anyway."
        % (msg, deck.name, FORCE_FLAG))
