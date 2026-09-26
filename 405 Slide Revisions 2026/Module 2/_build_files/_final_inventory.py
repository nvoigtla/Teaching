# -*- coding: utf-8 -*-
"""Create the final slide version, step 2 (read-only): pair every slide of
`Module X - In Class.pptx` with a slide of `Module X - Full.pptx` by content
fingerprint (never by number), and classify what step 2 has to do with it.

The step-1 sibling is `_full_inventory.py`; this is the same idea one layer
up. Writes `_final_pairing.json` and prints a table. No deck is written.

Rule 4 of step 2: the video blocks of Full are FROZEN. A match that lands
inside one is reported and never acted on, so an in-class edit of a taped
slide can only ever touch that slide's in-class copy.

Module-agnostic: the module folder is the script's own folder (move it to
`_build_files/` and set MODULE = HERE.parent).
"""
# NOTE (2026-09-25): "Module 2 - Full.pptx" and "Module 2 - Revised.pptx"
# were deleted by the step-1b cleanup. Full was a byte copy of Revised, so
# to use this again restore Revised from git and copy it:
#   git checkout b5484ae5 -- "405 Slide Revisions 2026/Module 2/Module 2 - Revised.pptx"
#   copy "Module 2 - Revised.pptx" "Module 2 - Full.pptx"
import hashlib
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).resolve().parent
MODULE = HERE.parent              # _build_files/ since 2026-09-25
FULL = next(iter(sorted(MODULE.glob("Module * - Full.pptx"))))
INCLASS = next(iter(sorted(MODULE.glob("Module * - In Class.pptx"))))

# Slides of Full that came from a taped video deck (step 1's own map).
# Everything else came from Revised and is in scope for step 2.
VIDEO_BLOCK = set(range(72, 117))   # Module 2: Videos 1-3 at Full 72-116

NP = "http://schemas.openxmlformats.org/presentationml/2006/main"
NA = "http://schemas.openxmlformats.org/drawingml/2006/main"
NR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _rels(z, part):
    head, tail = part.rsplit("/", 1)
    rp = head + "/_rels/" + tail + ".rels"
    if rp not in z.namelist():
        return {}
    return {r.get("Id"): (r.get("Type").rsplit("/", 1)[-1], r.get("Target"))
            for r in ET.fromstring(z.read(rp))}


def _texts(xml):
    out = []
    for p in ET.fromstring(xml).iter("{%s}p" % NA):
        s = "".join(t.text or "" for t in p.iter("{%s}t" % NA)).strip()
        if s:
            out.append(s)
    return out


def deck(path):
    z = zipfile.ZipFile(path)
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    prels = _rels(z, "ppt/presentation.xml")
    out = []
    for i, s in enumerate(pres.find("{%s}sldIdLst" % NP), 1):
        part = "ppt/" + prels[s.get("{%s}id" % NR)][1].lstrip("/").replace("../", "")
        xml = z.read(part)
        rels = _rels(z, part)
        notes = ""
        for typ, tgt in rels.values():
            if typ == "notesSlide":
                notes = "\n".join(
                    t for t in _texts(z.read("ppt/" + tgt.replace("../", "")))
                    if not re.fullmatch(r"\d+", t))
        x = xml.decode("utf-8")
        out.append(dict(
            n=i, texts=_texts(xml), notes=notes,
            hidden=ET.fromstring(xml).get("show") == "0",
            clicks=len(re.findall(r'nodeType="clickEffect"', x)),
            jumps=len(re.findall(r'action="ppaction://hlinksldjump"', x)),
            pics=len(re.findall(r"<p:pic>", x)),
            tags=any(t == "tags" for t, _ in rels.values()),
            xmlhash=hashlib.md5(xml).hexdigest()))
    z.close()
    return out


def norm(s):
    s = s.lower().replace("’", "'")
    return [w for w in re.sub(r"[^a-z0-9$%.= ]+", " ", s).split() if w]


def fp(sl):
    """fingerprint, with the live page-number run dropped"""
    return " ".join(norm(" ".join(t for t in sl["texts"]
                                  if not re.fullmatch(r"\d+", t))))


def main():
    full, ic = deck(FULL), deck(INCLASS)
    ffp = {s["n"]: fp(s) for s in full}
    ftok = {n: set(f.split()) for n, f in ffp.items()}

    pairing, used = [], {}
    for s in ic:
        f = fp(s)
        exact = [n for n, g in ffp.items() if g == f and f]
        if exact:
            # rule 4: a copy outside the video blocks is always the target
            outside = [n for n in exact if n not in VIDEO_BLOCK]
            fresh = [n for n in (outside or exact) if n not in used]
            best, score = (fresh or outside or exact)[0], 1.0
        else:
            t = set(f.split())
            best, score = None, 0.0
            for n, ft in ftok.items():
                if not ft or not t:
                    continue
                j = len(t & ft) / max(len(t | ft), 1)
                if j > score:
                    best, score = n, j
            if best in VIDEO_BLOCK:
                # prefer a near-match outside the frozen blocks
                alt, alts = None, 0.0
                for n, ft in ftok.items():
                    if n in VIDEO_BLOCK or not ft or not t:
                        continue
                    j = len(t & ft) / max(len(t | ft), 1)
                    if j > alts:
                        alt, alts = n, j
                if alts >= 0.5:
                    best, score = alt, alts
        if score < 0.5:
            best, score = None, score
        if best:
            used.setdefault(best, []).append(s["n"])
        r = full[best - 1] if best else None
        diff = []
        if r:
            if r["texts"] != s["texts"]:
                diff.append("text")
            if r["notes"] != s["notes"]:
                diff.append("notes")
            if r["clicks"] != s["clicks"]:
                diff.append("clicks %d->%d" % (r["clicks"], s["clicks"]))
            if r["pics"] != s["pics"]:
                diff.append("pics %d->%d" % (r["pics"], s["pics"]))
            if r["hidden"] != s["hidden"]:
                diff.append("hidden")
            if r["xmlhash"] != s["xmlhash"] and not diff:
                diff.append("xml")
        pairing.append(dict(ic=s["n"], full=best, score=round(score, 3),
                            diff=diff, tags=s["tags"], jumps=s["jumps"],
                            pics=s["pics"], empty=not s["texts"],
                            frozen=best in VIDEO_BLOCK if best else False,
                            title=(s["texts"] or ["<no text>"])[0][:58]))

    dropped = [s["n"] for s in full
               if s["n"] not in used and s["n"] not in VIDEO_BLOCK]
    (HERE / "_final_pairing.json").write_text(json.dumps(
        dict(pairing=pairing, dropped=dropped,
             full=[dict(n=s["n"], title=" | ".join(s["texts"][:2])[:90],
                        tags=s["tags"], jumps=s["jumps"], hidden=s["hidden"])
                   for s in full]),
        indent=1, ensure_ascii=False), encoding="utf-8")

    print("In Class %d slides  ->  Full %d slides\n" % (len(ic), len(full)))
    for p in pairing:
        if p["full"] is None:
            flag = "  <-- NEW (rule 5)"
        elif p["frozen"]:
            flag = "  <-- IN A VIDEO BLOCK (rule 4)"
        elif p["score"] < 0.999:
            flag = "  <-- CHECK"
        else:
            flag = ""
        print("IC%-3d -> F%-5s %.2f %-22s %-58s%s"
              % (p["ic"], p["full"], p["score"], ",".join(p["diff"])[:22],
                 p["title"], flag))

    print("\nFull slides in scope that In Class DROPPED (rule 6 - kept):")
    for n in dropped:
        print("  F%-3d %s" % (n, " | ".join(full[n - 1]["texts"][:2])[:84]))
    dup = {n: v for n, v in used.items() if len(v) > 1}
    print("\nFull slides matched more than once:", dup or "none")
    empties = [p["ic"] for p in pairing if p["empty"]]
    print("empty / text-less In-Class slides (rule 10):", empties or "none")


if __name__ == "__main__":
    main()
