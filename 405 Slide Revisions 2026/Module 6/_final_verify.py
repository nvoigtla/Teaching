# -*- coding: utf-8 -*-
"""Finalize routine, verification (read-only).

Checks "Module 6 - Final.pptx" against its sources:
  * every taped slide against its video-deck slide -- the slide XML with
    ids / rIds / cached page numbers normalised away, plus notes, click
    count, pictures, poll tags and hidden flag;
  * every in-class copy against its main-deck original (tag aside);
  * every slide-jump link resolved to the slide it lands on.
"""
import json, re, sys, zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import _final_inventory as FI

HERE = Path(__file__).resolve().parent
FINAL = HERE / "Module 6 - Final.pptx"
plan = json.loads((HERE / "_final_plan.json").read_text(encoding="utf-8"))


def raw(path):
    z = zipfile.ZipFile(path)
    return z, FI.deck(path)


def normxml(z, part):
    x = z.read(part).decode("utf-8")
    x = re.sub(r'<\?xml[^>]*\?>', '', x)
    x = re.sub(r'\s(r:embed|r:id|r:link|r:pict|id|spid|creationId|val)="[^"]*"', '', x)
    x = re.sub(r'<a:fld[^>]*type="slidenum"[^>]*>.*?</a:fld>', '<fld/>', x, flags=re.S)
    x = re.sub(r'<p14:creationId[^>]*/>', '', x)
    x = re.sub(r'<p:extLst>.*?</p:extLst>', '', x, flags=re.S)
    x = re.sub(r'\s(smtClean|dirty|err)="[^"]*"', '', x)
    x = re.sub(r'<a:endParaRPr[^>]*/>', '', x)
    return x


def clean(t):
    return [s for s in t if not re.fullmatch(r"\d+", s)]


fz, fin = raw(FINAL)
rz, rev = raw(FI.REVISED)
print("Final slides:", len(fin))

# expected layout of Final, by origin
exp = []                                    # (kind, key)
pairs = {x["r"]: x for x in plan["pairs"]}
for n in range(2, 95):
    exp.append(("taped", pairs[n]) if n in pairs else ("rev", n))
exp.append(("divider", None))
for n in plan["copies"]:
    exp.append(("copy", n))
for n in plan["move"]:
    exp.append(("rev", n))
assert len(exp) == len(fin), (len(exp), len(fin))

vcache = {}
bad = 0
for i, (kind, key) in enumerate(exp):
    f = fin[i]
    if kind == "taped":
        vp = FI.VDIR / key["deck"]
        if vp not in vcache:
            vcache[vp] = raw(vp)
        vz, vd = vcache[vp]
        s, sz_ = vd[key["vn"] - 1], vz
    elif kind in ("rev", "copy"):
        s, sz_ = rev[key - 1], rz
        # a copy of a TAPED slide is compared with the taped version
        if key in pairs:
            vp = FI.VDIR / pairs[key]["deck"]
            if vp not in vcache:
                vcache[vp] = raw(vp)
            vz, vd = vcache[vp]
            s, sz_ = vd[pairs[key]["vn"] - 1], vz
    else:
        print("%3d  divider: %s" % (i + 1, " / ".join(f["texts"])))
        continue
    probs = []
    ft, st = clean(f["texts"]), clean(s["texts"])
    if kind == "copy":
        ft = [re.sub(r"In Class · Examples · ", "", t) for t in ft]
        st = [re.sub(r"(Video \d+|In Class · Examples) · ", "", t) for t in st]
    if ft != st:
        probs.append("text")
    for k in ("notes", "clicks", "pics", "tags", "hidden", "media"):
        if f[k] != s[k]:
            probs.append(k)
    if kind != "copy":
        a, b = normxml(fz, f["part"]), normxml(sz_, s["part"])
        if a != b:
            probs.append("xml")
    else:
        a = normxml(fz, f["part"]); b = normxml(sz_, s["part"])
        a = re.sub(r"<a:t>Module 6 · [^<]*</a:t>", "", a)
        b = re.sub(r"<a:t>Module 6 · [^<]*</a:t>", "", b)
        if a != b:
            probs.append("xml")
    tag = next((t for t in f["texts"] if t.startswith("Module 6 ·")), "")
    if probs:
        bad += 1
    if probs or kind == "copy":
        print("%3d  %-5s %-12s %-18s %s" % (i + 1, kind,
              ("V%s.%d" % (re.search(r"Video (\d+)", key["deck"]).group(1), key["vn"])) if kind == "taped" else "R%d" % key,
              ",".join(probs) or "ok", tag[:70]))
print("slides with differences:", bad)

# links
NR = FI.NR
pres = ET.fromstring(fz.read("ppt/presentation.xml"))
prels = FI._rels(fz, "ppt/presentation.xml")
order = ["ppt/" + prels[s.get("{%s}id" % NR)][1].replace("../", "").lstrip("/")
         for s in pres.find("{%s}sldIdLst" % FI.NP)]
for i, part in enumerate(order, 1):
    x = fz.read(part).decode("utf-8")
    rels = FI._rels(fz, part)
    for rid in re.findall(r'<a:hlinkClick r:id="(rId\d+)"[^>]*action="ppaction://hlinksldjump"', x):
        tgt = "ppt/slides/" + rels[rid][1].split("/")[-1]
        print("link: slide %d -> slide %d (%s)" % (i, order.index(tgt) + 1,
              " | ".join(fin[order.index(tgt)]["texts"][:2])[:60]))
