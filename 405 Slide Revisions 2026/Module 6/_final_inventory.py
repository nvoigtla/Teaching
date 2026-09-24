# -*- coding: utf-8 -*-
"""Finalize routine, step 1 (read-only): inventory the taped video decks
in `Recorded Video Slides/` and pair every taped slide with a slide of
`Module 6 - Revised.pptx` by content fingerprint (never by number).

Writes `_final_pairing.json` and prints a table.  Reads only -- no deck is
written.  Build input for `_finalize.py`; keep.
"""
import hashlib, json, re, sys, zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).resolve().parent
REVISED = HERE / "Module 6 - Revised.pptx"
VDIR = HERE / "Recorded Video Slides"
NP = "http://schemas.openxmlformats.org/presentationml/2006/main"
NA = "http://schemas.openxmlformats.org/drawingml/2006/main"
NR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NREL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _rels(z, part):
    rp = part.rsplit("/", 1)
    rp = rp[0] + "/_rels/" + rp[1] + ".rels"
    if rp not in z.namelist():
        return {}
    root = ET.fromstring(z.read(rp))
    return {r.get("Id"): (r.get("Type").rsplit("/", 1)[-1], r.get("Target"),
                          r.get("TargetMode")) for r in root}


def _texts(xml):
    root = ET.fromstring(xml)
    out = []
    for p in root.iter("{%s}p" % NA):
        s = "".join(t.text or "" for t in p.iter("{%s}t" % NA))
        if s.strip():
            out.append(s.strip())
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
        texts = _texts(xml)
        notes = ""
        for typ, tgt, _ in rels.values():
            if typ == "notesSlide":
                npart = "ppt/" + tgt.replace("../", "")
                nt = _texts(z.read(npart))
                notes = "\n".join(t for t in nt if not re.fullmatch(r"\d+", t))
        root = ET.fromstring(xml)
        hidden = root.get("show") == "0"
        x = xml.decode("utf-8")
        clicks = len(re.findall(r'nodeType="clickEffect"', x))
        jumps = re.findall(r'action="ppaction://hlinksldjump"', x)
        pics = len(re.findall(r"<p:pic>", x))
        out.append(dict(n=i, id=s.get("id"), part=part, texts=texts,
                        notes=notes, hidden=hidden, clicks=clicks,
                        jumps=len(jumps), pics=pics,
                        tags=any(t == "tags" for t, _, _ in rels.values()),
                        media=sorted(t for t, _, _ in rels.values()
                                     if t in ("video", "media", "audio")),
                        xmlhash=hashlib.md5(xml).hexdigest()))
    return out


def norm(s):
    s = s.lower().replace("’", "'")
    s = re.sub(r"[^a-z0-9$%.= ]+", " ", s)
    return [w for w in s.split() if w]


def fp(sl):
    return " ".join(norm(" ".join(sl["texts"])))


def main():
    rev = deck(REVISED)
    vdecks = sorted(p for p in VDIR.glob("Module 6 - Video *.pptx"))
    vdecks.sort(key=lambda p: int(re.search(r"Video (\d+)", p.name).group(1)))
    rfp = {s["n"]: fp(s) for s in rev}
    rtok = {n: set(f.split()) for n, f in rfp.items()}
    pairing, used = [], {}
    for vp in vdecks:
        k = int(re.search(r"Video (\d+)", vp.name).group(1))
        for s in deck(vp):
            f = fp(s)
            exact = [n for n, g in rfp.items() if g == f and f]
            if exact:
                best, score = exact[0], 1.0
                if len(exact) > 1:           # same text twice: prefer unused,
                    fresh = [n for n in exact if n not in used]   # then notes
                    same_notes = [n for n in (fresh or exact)
                                  if rev[n - 1]["notes"] == s["notes"]]
                    best = (same_notes or fresh or exact)[0]
            else:
                t = set(f.split())
                best, score = None, 0.0
                for n, rt in rtok.items():
                    if not rt or not t:
                        continue
                    j = len(t & rt) / max(len(t | rt), 1)
                    if j > score:
                        best, score = n, j
            used.setdefault(best, []).append((k, s["n"]))
            r = rev[best - 1] if best else None
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
            pairing.append(dict(video=k, deck=vp.name, vn=s["n"],
                                title=(s["texts"] or [""])[0][:60],
                                rev=best, score=round(score, 3), diff=diff,
                                tags=s["tags"], jumps=s["jumps"],
                                media=s["media"], hidden=s["hidden"]))
    unpaired = [s["n"] for s in rev if s["n"] not in used]
    (HERE / "_final_pairing.json").write_text(json.dumps(
        dict(pairing=pairing, unpaired=unpaired,
             revised=[dict(n=s["n"], title=" | ".join(s["texts"][:2])[:90],
                           tags=s["tags"], jumps=s["jumps"],
                           hidden=s["hidden"]) for s in rev]),
        indent=1, ensure_ascii=False), encoding="utf-8")
    for p in pairing:
        flag = "" if p["score"] >= 0.999 else ("  <-- CHECK" if p["score"] >= 0.5 else "  <-- NEW?")
        print("V%d.%-3d -> R%-4s %.2f %-28s %s%s" % (
            p["video"], p["vn"], p["rev"], p["score"], ",".join(p["diff"]),
            p["title"], flag))
    print("\nrevised slides in no video deck:", unpaired)
    for n in unpaired:
        s = rev[n - 1]
        print("  R%d: %s" % (n, " | ".join(s["texts"][:2])[:100]))
    dup = {n: v for n, v in used.items() if len(v) > 1}
    print("revised slides used more than once:", dup)


if __name__ == "__main__":
    main()
