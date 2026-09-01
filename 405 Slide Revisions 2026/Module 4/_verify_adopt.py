"""Verify the result of `_adopt_videos.py`.

Checks, on the produced deck:
  1. slide count and the two dividers' position and wording
  2. every duplicated slide is byte-identical to its original apart from
     the cached page-number field
  3. every cached page number equals the slide's display position
     (for the slides that carry one)
  4. PollEverywhere integrity: a slide carrying a `tags` part also carries
     a notes part (a tag without notes crashes the slideshow deck-wide),
     and no embed id appears on two slides
  5. every relationship target resolves to a part that exists
"""
import re
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

import _video_map as VM

HERE = Path(__file__).resolve().parent
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R_ID = "{%s}id" % NS_R

SLIDENUM_RE = re.compile(
    r'<a:fld[^>]*type="slidenum"[^>]*>.*?<a:t>(\d*)</a:t>', re.S)


def rels_of(z, part):
    head, tail = part.rsplit("/", 1)
    rn = "%s/_rels/%s.rels" % (head, tail)
    if rn not in z.namelist():
        return []
    return [(c.get("Id"), c.get("Type").rsplit("/", 1)[1], c.get("Target"),
             c.get("TargetMode")) for c in ET.fromstring(z.read(rn))]


def resolve(target, from_part):
    base = from_part.rsplit("/", 1)[0]
    t = target
    while t.startswith("../"):
        t = t[3:]
        base = base.rsplit("/", 1)[0]
    return (base + "/" + t) if base else t


def texts(z, part):
    return " ".join(t.text or "" for t in
                    ET.fromstring(z.read(part)).iter(A_NS + "t"))


def main():
    path = HERE / (sys.argv[1] if len(sys.argv) > 1
                   else "Module 4 - Revised.pptx")
    z = zipfile.ZipFile(str(path))
    names = set(z.namelist())
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rmap = {r[0]: r[2] for r in rels_of(z, "ppt/presentation.xml")}
    parts = ["ppt/" + rmap[s.get(R_ID)].replace("../", "").lstrip("/")
             for s in pres.find("{%s}sldIdLst" % NS_P)]
    n = len(parts)
    fails = []

    def ok(cond, msg):
        print(("  PASS  " if cond else "  FAIL  ") + msg)
        if not cond:
            fails.append(msg)

    print("deck: %s   %d slides" % (path.name, n))

    # ---- 1. count and dividers
    print("\n1. slide count and dividers")
    expected = 90 + 1 + len(VM.DUPLICATE_AT_END) + 1
    ok(n == expected, "slide count is %d (expected %d)" % (n, expected))
    div_a, div_b = 91, n
    ta = " ".join(texts(z, parts[div_a - 1]).split())
    tb = " ".join(texts(z, parts[div_b - 1]).split())
    ok("Slides Not Included in the Videos" in ta,
       "slide %d is the 'not included' divider: %r" % (div_a, ta[:60]))
    ok("Some Applications of the Material" in tb,
       "slide %d is the 'applications' divider: %r" % (div_b, tb[:70]))

    # ---- 2. duplicates match originals
    print("\n2. duplicated slides vs their originals")
    bad = []
    for i, r in enumerate(VM.DUPLICATE_AT_END):
        dup_disp = 92 + i
        a = z.read(parts[r - 1]).decode("utf-8")
        b = z.read(parts[dup_disp - 1]).decode("utf-8")
        a2 = SLIDENUM_RE.sub("FLD", a)
        b2 = SLIDENUM_RE.sub("FLD", b)
        if a2 != b2:
            bad.append((r, dup_disp))
    ok(not bad, "all %d duplicates identical to source (ignoring page "
                "number); mismatches: %s" % (len(VM.DUPLICATE_AT_END), bad))

    # ---- 3. cached page numbers
    print("\n3. cached page numbers")
    wrong = []
    carry = 0
    for i, part in enumerate(parts, 1):
        m = SLIDENUM_RE.search(z.read(part).decode("utf-8"))
        if not m:
            continue
        carry += 1
        if m.group(1) != str(i):
            wrong.append((i, m.group(1)))
    ok(not wrong, "%d slides carry a page number, all cached correctly; "
                  "wrong: %s" % (carry, wrong[:10]))

    # ---- 4. poll integrity
    print("\n4. PollEverywhere integrity")
    polls = []
    for i, part in enumerate(parts, 1):
        rr = rels_of(z, part)
        tag = [x for x in rr if x[1] == "tags"]
        note = [x for x in rr if x[1] == "notesSlide"]
        if not tag:
            continue
        tx = z.read(resolve(tag[0][2], part)).decode("utf-8")
        eid = re.search(r'__PE_POLL_EMBED_ID" val="([^"]+)"', tx)
        note_txt = texts(z, resolve(note[0][2], part)) if note else ""
        polls.append((i, eid.group(1) if eid else None, bool(note),
                      "Poll Title" in note_txt))
    for i, eid, has_note, has_boiler in polls:
        ok(has_note and has_boiler,
           "slide %d: embed %s has its notes part (required)"
           % (i, (eid or "?")[:8]))
    ids = [p[1] for p in polls]
    ok(len(ids) == len(set(ids)),
       "no embed id used twice (%d live poll(s): %s)"
       % (len(polls), [p[0] for p in polls]))
    # a slide with PollEv boilerplate notes but no tag is harmless but odd
    orphan_notes = []
    for i, part in enumerate(parts, 1):
        rr = rels_of(z, part)
        if any(x[1] == "tags" for x in rr):
            continue
        note = [x for x in rr if x[1] == "notesSlide"]
        if note and "Poll Title: Do not modify" in \
                texts(z, resolve(note[0][2], part)):
            orphan_notes.append(i)
    ok(not orphan_notes,
       "no slide keeps PollEv boilerplate notes without a tag: %s"
       % orphan_notes)

    # ---- 5. every rel target exists
    print("\n5. relationship targets resolve")
    missing = []
    for part in [p for p in names if p.endswith(".xml")
                 and "/_rels/" not in p and "/" in p]:
        for rid, typ, tgt, mode in rels_of(z, part):
            if mode == "External":
                continue
            t = resolve(tgt, part)
            if t not in names:
                missing.append((part, tgt))
    ok(not missing, "all internal rel targets exist; missing: %s"
       % missing[:6])

    z.close()
    print("\n%s  (%d failure(s))"
          % ("ALL CHECKS PASS" if not fails else "FAILURES PRESENT",
             len(fails)))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
