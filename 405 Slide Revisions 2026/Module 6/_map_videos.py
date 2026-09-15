# -*- coding: utf-8 -*-
"""Match every slide of the 8 original video decks to a slide of the main
deck, so the video block boundaries in the rebuild rest on evidence.

Matching is on the slide's own text, normalised: the title alone is not
enough (several titles repeat), so a slide is fingerprinted by its whole
text content and scored against every main-deck slide by token overlap.
Build input -- keep.
"""
import re, sys
from pathlib import Path

def slides(path):
    txt = Path(path).read_text(encoding="utf-8")
    out = {}
    for blk in re.split(r"^## Slide ", txt, flags=re.M)[1:]:
        n = int(re.match(r"\d+", blk).group())
        body = " ".join(re.findall(r"^- TXT \[[^\]]*\]: (.*)$", blk, flags=re.M))
        out[n] = body
    return out

def norm(s):
    s = s.lower().replace("\u23ce", " ")
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return [w for w in s.split() if len(w) > 2]

main = slides("_source_inventory_NV.md")
mtok = {n: set(norm(t)) for n, t in main.items()}

print("| Video | Slide | Video-deck title | Best main-deck match | Overlap |")
print("|---|---|---|---|---|")
for v in range(1, 9):
    vs = slides(f"_source_inventory_V{v}.md")
    for n, body in sorted(vs.items()):
        vt = set(norm(body))
        if not vt:
            continue
        best, score = None, 0.0
        for m, mt in mtok.items():
            if not mt:
                continue
            j = len(vt & mt) / max(1, len(vt))
            if j > score:
                best, score = m, j
        title = re.split(r"\u23ce", body)[0][:52]
        mtitle = re.split(r"\u23ce", main.get(best, ""))[0][:46] if best else "-"
        flag = "" if score >= 0.60 else ("  <-- CHECK" if score >= 0.30 else "  <-- NO MATCH")
        print(f"| V{v} | {n} | {title} | NV {best}: {mtitle} | {score:.2f}{flag} |")
