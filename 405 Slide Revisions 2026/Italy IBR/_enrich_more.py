"""
Enrichments (2026-07-20, remaining sections) — economic-history bullets from
the podcast research. In-place OOXML; by current display->part. Run
_resize_bullets.py apply afterward.
  31 Dark Ages, 33 Communes, 46 Decline, 48 Napoleon, 59 & 60 Unification,
  97 Strong Brands.
"""
import os
import shutil
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
LVL_MARL = {0: 342900, 1: 731520, 2: 1097280}
LVL_CHAR = {0: "▪", 1: "–", 2: "·"}


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def para_xml(level, text):
    lvl = f' lvl="{level}"' if level else ""
    return (f'<a:p xmlns:a="{A}"><a:pPr marL="{LVL_MARL[level]}" indent="-274320"{lvl}>'
            f'<a:buClr><a:srgbClr val="0B2B4E"/></a:buClr>'
            f'<a:buFont typeface="Calibri"/><a:buChar char="{LVL_CHAR[level]}"/></a:pPr>'
            f'<a:r><a:rPr lang="en-US" sz="2400"><a:solidFill><a:srgbClr val="0B2B4E"/></a:solidFill>'
            f'<a:latin typeface="Calibri"/></a:rPr><a:t>{escape(text)}</a:t></a:r></a:p>')


def body_box(root):
    best, best_n = None, 0
    for sp in root.iter(q(P, "sp")):
        tb = sp.find(q(P, "txBody"))
        if tb is None:
            continue
        n = sum(1 for pp in tb.findall(q(A, "p"))
                if pp.find(q(A, "pPr") + "/" + q(A, "buChar")) is not None)
        if n > best_n:
            best, best_n = tb, n
    return best


def enrich(data, part, *, edits=None, inserts=None, appends=None):
    root = ET.fromstring(data[f"ppt/slides/{part}"])
    tb = body_box(root)
    for old, new in (edits or []):
        for t in tb.iter(q(A, "t")):
            if t.text and old in t.text:
                t.text = t.text.replace(old, new)
                break
    for after_needle, text, level in (inserts or []):
        target = None
        for pp in tb.findall(q(A, "p")):
            if after_needle in "".join(t.text or "" for t in pp.iter(q(A, "t"))):
                target = pp
                break
        newp = ET.fromstring(para_xml(level, text))
        (target.addnext(newp) if target is not None else tb.append(newp))
    for text, level in (appends or []):
        tb.append(ET.fromstring(para_xml(level, text)))
    data[f"ppt/slides/{part}"] = ser(root)


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target") for r in ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    order = [os.path.basename(rid2t[s.get(q(R, "id"))]) for s in pres.find(q(P, "sldIdLst"))]

    enrich(data, order[30],  # 31 Dark Ages
           inserts=[("Muslim rule in the South",
                     "Arab Sicily: a farming revolution — citrus, sugar, irrigation", 1)],
           appends=[("North self-governs; the South is ruled top-down — the deep root of the North–South divide", 0)])
    enrich(data, order[32],  # 33 Communes
           edits=[("Peace Treaty:", "Peace of Constance (1183):"),
                  ("Treaty:", "Peace of Constance (1183):")],
           inserts=[("Commercial Revolution",
                     "Crusades (from 1095) reopened Eastern trade routes", 1)],
           appends=[("Maritime republics — Venice, Genoa, Pisa, Amalfi — dominated Mediterranean trade", 0)])
    enrich(data, order[45],  # 46 Decline
           inserts=[("Discovery of America",
                     "Columbus (a Genoese) sailed for Spain — trade shifted to the Atlantic", 1),
                    ("occupied by Spain",
                     "Refeudalization: barons and latifundia entrenched in the South", 1)],
           appends=[("The 1630 plague devastated the northern cities", 0)])
    enrich(data, order[47],  # 48 Napoleon
           inserts=[("Code Civil",
                     "Suppressed guilds, metric system, meritocratic administration", 1)])
    enrich(data, order[58],  # 59 Italy After Unification (1)
           appends=[("Unification built a national market — but on northern terms (law, currency, tariffs)", 0)])
    enrich(data, order[59],  # 60 Italy After Unification (2)
           appends=[("The 1887 tariff helped northern industry but hurt southern farming", 0),
                    ("Millions emigrated (1876–1914), mostly from the South — a remittance economy", 0)])
    enrich(data, order[96],  # 97 Strong Brands
           appends=[("Family firms & mid-sized “pocket multinationals” — the “fourth capitalism”", 0)])

    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print("enriched 31,33,46,48,59,60,97")


if __name__ == "__main__":
    main()
