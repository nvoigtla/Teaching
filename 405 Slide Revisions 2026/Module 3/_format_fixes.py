# -*- coding: utf-8 -*-
"""Module 3 formatting fixes approved by Nico on 2026-08-26, after the
deck-wide audit in _audit_format.py:

  1. BACK BUTTONS (displays 80-81) — the original deck's gray left-arrow
     "back" shapes become the deck-standard navy rounded pill with white
     bold "← Back" at the fixed lower-right position. The jump action
     lives on the shape (cNvPr/hlinkClick), so it survives the restyle.
  2. OFF-PALETTE COLORS — every stray mapped onto the palette. The warm
     off-whites all collapse onto cream; the deliberate red/green cost
     coding on display 44 keeps its coding but moves to the course red
     and green. One stray is kept on purpose - see COLOR_MAP.
  3. TRAILING PERIODS on displays 10, 16 and 25.
  4. FLAT FILLED BOXES — the concept map's boxes (display 7), the note
     bar on display 17 and the three-card comparison on display 44 get
     slight rounding and the soft shade; display 17's cream callout gets
     the shade it was missing.
  5. SLIDE TITLES — the five with a stray lower-case word are raised to
     title case (capitalisation only, no rewording). Display 61 is left
     alone: it is written as a full sentence, not a heading.

Two audit findings are deliberately NOT changed, and are reported
instead: display 25's navy "Wage" cell stays flat (it is the header of a
table column, and its neighbour, the main table's own header row, is
flat too — rounding one of the pair would break it), and the box+text
pairs on display 7 stay ungrouped (grouping invalidates a slide's
animations, which is a separate pass).

Rerunnable: every step is idempotent.
"""
import copy
import os
import re
import sys
import zipfile
import shutil
from pathlib import Path

from lxml import etree as ET

sys.stdout.reconfigure(encoding="utf-8")
HERE = Path(__file__).parent
DECK = HERE / "Module 3 - Revised.pptx"

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0

NAVY, WHITE, CREAM = "0B2B4E", "FFFFFF", "FDF6E6"

# --- 2. palette mapping ----------------------------------------------------
COLOR_MAP = {
    "FFF5E0": CREAM,      # table-cell number highlights (16, 25)
    "F4F1EA": CREAM,      # card / panel fills (18, 40, 41)
    "FDECDB": CREAM,      # question band (55)
    "EEECE1": CREAM,      # backup-slide title banners (80, 81)
    "FFF2CC": "F6E8C9",   # highlighted answer cells (47) -> pale gold
    "8B1A1A": "C00000",   # sunk-cost red (44) -> course red
    "2E7D32": "1B5E20",   # variable-cost green (44) -> course green
    "4F6128": "1B5E20",   # olive label + leader (38) -> course green
    # NOT remapped: #9EC5F7 on display 20 is a light tint of concept
    # blue sitting on the NAVY hero box. Concept blue itself (#0070C0)
    # is too dark on that ground - checked against a render - so the
    # tint stays as a deliberate deck accent.
    "404040": "555B66",   # secondary text (18) -> palette gray
    "C8CED6": "C8CDD3",   # hairline rules (55) -> rule gray
    "C0C0C0": "000000",   # a silver drop shadow (12) -> standard black
}

# --- 5. title case: stray lower-case words, capitalisation only ------------
TITLE_FIXES = {
    28: ["likes", "being", "treated", "unequally"],
    47: ["over"],
    54: ["additional"],          # "per" stays lower case (preposition)
    66: ["is"],
    71: ["than"],
}

# --- 3. trailing periods ---------------------------------------------------
# Named line by line: two of the three sit INSIDE a box+text group, and a
# blanket "line ends in a period" rule would also strip captions and
# running prose, which keep their punctuation.
PERIOD_LINES = {
    10: ["A production function transforms inputs into outputs."],
    16: ["▪  Example:  MPL from Rivian Production function.",
         "Example:  MPL from Rivian Production function."],
    25: ["Optimal hiring falls into the interval between 3,000 and "
         "3,500 workers."],
}

# --- 4. flat boxes ---------------------------------------------------------
ROUND_ADJ = "12000"          # ~12 % corner, the deck's callout rounding
SHADOW = ('<a:outerShdw xmlns:a="%s" blurRad="50800" dist="38100" '
          'dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000">'
          '<a:alpha val="45000"/></a:srgbClr></a:outerShdw>' % A)

# --- 1. back buttons -------------------------------------------------------
BACK_XY = (11.72, 6.60, 1.55, 0.46)      # the deck-wide fixed position
BACK_LABEL = "←  Back"


def q(ns, t):
    return "{%s}%s" % (ns, t)


def order_of(data):
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target") for r in
             ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    return [os.path.basename(rid2t[s.get(q(R, "id"))])
            for s in pres.find(q(P, "sldIdLst"))]


def geom(sp):
    off = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "off"))
    ext = sp.find(q(P, "spPr") + "/" + q(A, "xfrm") + "/" + q(A, "ext"))
    if off is None or ext is None:
        return None
    return (int(off.get("x")) / EMU, int(off.get("y")) / EMU,
            int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU)


def add_shadow(sp):
    """Soft drop shadow, inserted in schema order (…ln -> effectLst)."""
    spPr = sp.find(q(P, "spPr"))
    if spPr is None or spPr.find(q(A, "effectLst") + "/"
                                 + q(A, "outerShdw")) is not None:
        return False
    for old in spPr.findall(q(A, "effectLst")):
        spPr.remove(old)
    eff = ET.SubElement(spPr, q(A, "effectLst"))
    eff.append(ET.fromstring(SHADOW))
    ln = spPr.find(q(A, "ln"))
    anchor = ln if ln is not None else spPr.find(q(A, "solidFill"))
    if anchor is not None:
        spPr.remove(eff)
        anchor.addnext(eff)
    return True


def round_corners(sp):
    pg = sp.find(q(P, "spPr") + "/" + q(A, "prstGeom"))
    if pg is None or pg.get("prst") != "rect":
        return False
    pg.set("prst", "roundRect")
    for old in pg.findall(q(A, "avLst")):
        pg.remove(old)
    av = ET.SubElement(pg, q(A, "avLst"))
    gd = ET.SubElement(av, q(A, "gd"))
    gd.set("name", "adj")
    gd.set("fmla", "val " + ROUND_ADJ)
    return True


def runs_of(sp):
    return list(sp.iter(q(A, "t")))


z = zipfile.ZipFile(DECK)
data = {n: z.read(n) for n in z.namelist()}
z.close()
order = order_of(data)
log = []

for i, base in enumerate(order, 1):
    part = "ppt/slides/" + base
    tree = ET.fromstring(data[part])
    spTree = tree.find(q(P, "cSld") + "/" + q(P, "spTree"))
    touched = False

    # 2. palette --------------------------------------------------------
    for clr in tree.iter(q(A, "srgbClr")):
        new = COLOR_MAP.get(clr.get("val").upper())
        if new:
            clr.set("val", new)
            touched = True
            log.append((i, "color", "#%s -> #%s"
                        % (clr.get("val"), new)))

    for sp in spTree.findall(q(P, "sp")):
        g = geom(sp)
        text = "".join(t.text or "" for t in runs_of(sp))

        # 1. back buttons ----------------------------------------------
        if text.strip().lower() == "back" and g and g[1] > 6.0:
            pg = sp.find(q(P, "spPr") + "/" + q(A, "prstGeom"))
            xfrm = sp.find(q(P, "spPr") + "/" + q(A, "xfrm"))
            off, ext = xfrm.find(q(A, "off")), xfrm.find(q(A, "ext"))
            off.set("x", str(int(BACK_XY[0] * EMU)))
            off.set("y", str(int(BACK_XY[1] * EMU)))
            ext.set("cx", str(int(BACK_XY[2] * EMU)))
            ext.set("cy", str(int(BACK_XY[3] * EMU)))
            pg.set("prst", "rect")          # normalise, then round below
            for old in pg.findall(q(A, "avLst")):
                pg.remove(old)
            ET.SubElement(pg, q(A, "avLst"))
            round_corners(sp)
            pg.find(q(A, "avLst") + "/" + q(A, "gd")).set(
                "fmla", "val 28000")        # pill, not a soft-cornered box
            fill = sp.find(q(P, "spPr") + "/" + q(A, "solidFill") + "/"
                           + q(A, "srgbClr"))
            fill.set("val", NAVY)
            add_shadow(sp)
            for rpr in sp.iter(q(A, "rPr")):
                rpr.set("sz", "1500")
                c = rpr.find(q(A, "solidFill") + "/" + q(A, "srgbClr"))
                if c is not None:
                    c.set("val", WHITE)
            ts = runs_of(sp)
            ts[0].text = BACK_LABEL
            for extra in ts[1:]:
                extra.text = ""
            touched = True
            log.append((i, "back button", "navy '← Back' pill at %.2f,%.2f"
                        % BACK_XY[:2]))
            continue

        # 4. flat filled boxes -----------------------------------------
        fill = sp.find(q(P, "spPr") + "/" + q(A, "solidFill") + "/"
                       + q(A, "srgbClr"))
        shadowed = sp.find(q(P, "spPr") + "/" + q(A, "effectLst") + "/"
                           + q(A, "outerShdw")) is not None
        if i == 7 and fill is not None and g and not shadowed \
                and 1.5 < g[1] < 7.0 and g[2] < 12.0:
            round_corners(sp)
            add_shadow(sp)
            touched = True
            log.append((i, "box", "rounded + shade  [%.2f,%.2f]" % g[:2]))
        # display 44's three-card comparison: the header bars and the
        # body cards were square while the "Examples" cards under them
        # were rounded — square corners on a lifted card is the same
        # defect as a missing shade (CLAUDE.md: header cells of a
        # multi-column comparison are rounded). The white BACKING cards
        # behind charts and tables elsewhere stay square by design.
        if i == 44 and fill is not None and shadowed and round_corners(sp):
            touched = True
            log.append((i, "box", "rounded  [%.2f,%.2f]" % g[:2]))
        if i == 17 and fill is not None and g and not shadowed \
                and 6.4 < g[1] < 7.0:
            round_corners(sp)
            add_shadow(sp)
            touched = True
            log.append((i, "box", "rounded + shade  [%.2f,%.2f]" % g[:2]))
        if i == 17 and fill is not None and g and not shadowed \
                and fill.get("val") == CREAM:
            add_shadow(sp)
            touched = True
            log.append((i, "box", "shade on cream callout  [%.2f,%.2f]"
                        % g[:2]))

        # 5. titles ------------------------------------------------------
        if i in TITLE_FIXES and g and 0.40 < g[1] < 0.95 and g[2] > 8:
            for t in runs_of(sp):
                s = t.text or ""
                for w in TITLE_FIXES[i]:
                    s = re.sub(r"\b%s\b" % w, w.capitalize(), s)
                if s != (t.text or ""):
                    t.text = s
                    touched = True
            log.append((i, "title", "".join(t.text or ""
                                            for t in runs_of(sp))[:60]))

    # 3. trailing periods (also inside box+text groups) -----------------
    for sp in spTree.iter(q(P, "sp")):
        for para in sp.iter(q(A, "p")):
            ts = [t for t in para.iter(q(A, "t"))]
            if not ts:
                continue
            line = "".join(t.text or "" for t in ts).strip()
            if line not in PERIOD_LINES.get(i, ()):
                continue
            for t in reversed(ts):
                if (t.text or "").rstrip().endswith("."):
                    t.text = (t.text or "").rstrip()[:-1]
                    touched = True
                    log.append((i, "period", line[:60]))
                    break

    # 4b. cream callouts inside groups still need their shade -----------
    if i == 17:
        for sp in spTree.iter(q(P, "sp")):
            fill = sp.find(q(P, "spPr") + "/" + q(A, "solidFill") + "/"
                           + q(A, "srgbClr"))
            if fill is not None and fill.get("val") == CREAM \
                    and add_shadow(sp):
                touched = True
                log.append((i, "box", "shade on the cream callout"))

    if touched:
        data[part] = ET.tostring(tree, xml_declaration=True,
                                 encoding="UTF-8", standalone=True)

by_kind = {}
for i, kind, detail in log:
    by_kind.setdefault(kind, []).append((i, detail))
for kind in ("back button", "title", "period", "box", "color"):
    hits = by_kind.get(kind, [])
    if kind == "color":
        print("\ncolor  (%d runs/fills remapped on slides %s)"
              % (len(hits), sorted({h[0] for h in hits})))
        continue
    print("\n%s  (%d)" % (kind, len(hits)))
    for i, detail in hits:
        print("  %2d  %s" % (i, detail))

out = DECK.with_suffix(".fix_tmp.pptx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
    for name, blob in data.items():
        zout.writestr(name, blob)
shutil.move(str(out), str(DECK))
print("\nsaved %s" % DECK.name)
