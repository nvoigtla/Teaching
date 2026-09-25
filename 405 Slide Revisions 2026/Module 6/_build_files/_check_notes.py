# -*- coding: utf-8 -*-
"""Audit the speaker notes against the slides they sit on.

Written 2026-09-15 after the zoo rescale left slides 70 and 72-74 with
notes quoting the OLD demand curve, and after the Module 6 wrap-up
podcast turned out to have sourced a claim from a note rather than from a
slide.  Both are the same failure: the notes are edited in a different
file from the slides, so nothing forces them to agree.

Three checks, none of which a render can catch:

  NONOTE   a content slide with no speaker note at all.
  NUMBER   a figure that appears in the note but on no shape of the
           slide.  Spelled-out numbers are folded to digits first, which
           is what catches "four dollars" against a slide reading $8.
  EMPTY    a note that is only whitespace or a stub.

NUMBER is a TRIAGE list, not a defect list: a note may legitimately name
a number the slide does not show (a date, "three conditions", a figure
carried over from the previous slide).  Read each hit rather than fixing
it blind.

    python _check_notes.py [deck.pptx]
"""
import re
import sys
from pathlib import Path

from pptx import Presentation

DECK = (sys.argv[1] if len(sys.argv) > 1 else
        str(Path(__file__).resolve().parents[1] / "Module 6 - Full.pptx"))

# slides that carry no note by design
SKIP_TITLES = ("",)

_UNITS = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14,
    "fifteen": 15, "sixteen": 16, "seventeen": 17, "eighteen": 18,
    "nineteen": 19,
}
_TENS = {"twenty": 20, "thirty": 30, "forty": 40, "fifty": 50,
         "sixty": 60, "seventy": 70, "eighty": 80, "ninety": 90}


def _words_to_numbers(text):
    """Every number the prose spells out, as a set of floats.

    Handles "four", "thirty-two", "twenty four", and the scale words that
    follow them ("fifty dollars", "sixteen times four over two").  Scale
    words alone ("a hundred") are not counted -- they carry no digit the
    slide could disagree with.
    """
    out = set()
    toks = re.findall(r"[a-z]+", text.lower())
    i = 0
    while i < len(toks):
        t = toks[i]
        val = None
        if t in _TENS:
            val = _TENS[t]
            if i + 1 < len(toks) and toks[i + 1] in _UNITS \
                    and 0 < _UNITS[toks[i + 1]] < 10:
                val += _UNITS[toks[i + 1]]
                i += 1
        elif t in _UNITS:
            val = _UNITS[t]
        if val is not None:
            # a following scale word multiplies it
            if i + 1 < len(toks):
                nxt = toks[i + 1]
                if nxt == "hundred":
                    val *= 100
                    i += 1
                elif nxt in ("thousand",):
                    val *= 1000
                    i += 1
            out.add(float(val))
        i += 1
    # hyphenated forms ("thirty-two") are already split by the \w+ scan,
    # but the pair has to be re-joined when the tens word came first
    for a, b in re.findall(r"([a-z]+)-([a-z]+)", text.lower()):
        if a in _TENS and b in _UNITS:
            out.add(float(_TENS[a] + _UNITS[b]))
    return out


_NUM = re.compile(r"(?<![\w.])(\d[\d,]*(?:\.\d+)?)")


def _digits(text):
    out = set()
    for m in _NUM.finditer(text):
        try:
            out.add(float(m.group(1).replace(",", "")))
        except ValueError:
            pass
    return out


def _walk(shapes, out):
    for sh in shapes:
        if sh.shape_type == 6:                      # GROUP
            _walk(sh.shapes, out)
            continue
        # a native TABLE is a graphicFrame with no text_frame of its own,
        # so its cells are invisible to a plain shape walk -- which made
        # every figure in slide 61's Netflix plan table read as drift
        if getattr(sh, "has_table", False):
            try:
                for row in sh.table.rows:
                    out.append(" ".join(c.text for c in row.cells))
            except Exception:
                pass
            continue
        if sh.has_text_frame and sh.text_frame.text.strip():
            out.append(sh.text_frame.text)


def main():
    prs = Presentation(DECK)
    slides = list(prs.slides)
    nonote, number, empty = [], [], []
    for i, s in enumerate(slides, 1):
        texts = []
        _walk(s.shapes, texts)
        body = "\n".join(texts)
        title = ""
        for t in texts:
            t = t.strip()
            if t and not t.startswith("Module 6 ·") \
                    and "Management 405" not in t and t != str(i):
                title = t.split("\n")[0]
                break

        note = ""
        if s.has_notes_slide:
            tf = s.notes_slide.notes_text_frame
            note = (tf.text or "").strip() if tf is not None else ""

        if not note:
            nonote.append((i, title))
            continue
        if len(note) < 40:
            empty.append((i, title, note))
            continue
        if "Poll Title:" in note:                   # PollEv machinery
            continue

        on_slide = _digits(body)
        in_note = _digits(note) | _words_to_numbers(note)
        # a note may restate a slide figure scaled by a unit word; allow
        # the common 1,000 / million steps rather than crying wolf
        missing = set()
        for v in in_note:
            if v in on_slide:
                continue
            if any(abs(v * f - o) < 1e-6 or abs(v - o * f) < 1e-6
                   for o in on_slide for f in (1000.0, 1e6, 100.0)):
                continue
            missing.add(v)
        # numbers under 4 are almost always "the three conditions"
        missing = {v for v in missing if v >= 4}
        if missing:
            number.append((i, title, sorted(missing), sorted(on_slide)))

    def head(tag, n):
        print("\n== %s : %d ==" % (tag, n))

    head("NONOTE (no speaker note)", len(nonote))
    for i, t in nonote:
        print("   s%-4d %s" % (i, t[:66]))
    head("EMPTY (stub note)", len(empty))
    for i, t, n in empty:
        print("   s%-4d %-40s %r" % (i, t[:40], n[:40]))
    head("NUMBER (figure in the note, not on the slide)", len(number))
    for i, t, miss, on in number:
        print("   s%-4d %s" % (i, t[:60]))
        print("         note has: %s" % ", ".join("%g" % v for v in miss))
        print("         slide has: %s"
              % (", ".join("%g" % v for v in on[:14]) or "(none)"))
    print("\ntotal: %d no-note, %d stub, %d number-drift"
          % (len(nonote), len(empty), len(number)))


if __name__ == "__main__":
    main()
