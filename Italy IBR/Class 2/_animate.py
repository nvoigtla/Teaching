# -*- coding: utf-8 -*-
"""
Inject Fade / on-click build animations into "Class 2 - Revised.pptx".

The timing primitives (`effect_par`, `click_group`, `timing_xml`) are the ones
proven on the Class 1 deck -- byte-pattern identical to what PowerPoint itself
writes.  What differs here is the plan: this deck is small and every slide has
named shapes, so the beats are declared explicitly per slide instead of being
inferred by heuristics.

Rules applied (Teaching CLAUDE.md):
  - Fade entrance, 0.5 s, on click; chrome never animates.
  - Bullet slides: the FIRST top-level bullet (with its sub-bullets) shows with
    the slide; the build starts from the second.
  - A box and the text inside it are one shape here, so each callout is one beat.
  - Static (no animation): everything except slide 3. The title, roadmap/agenda,
    poll, dividers, Thank You and backup slides never animated; the debate and
    company slides (7-13) were built step by step until Nico asked for them to
    land complete, which is how the presentation half of the deck now behaves.

Run:  python _animate.py all apply      (or a slide list, e.g. "3")
Verify click STRUCTURE (not just effect counts) via PowerPoint COM:
MainSequence.Item(i).Timing.TriggerType -- 1 = on click, 2 = with previous.
"""
import os
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 2 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# 1 title · 2 roadmap · 4 poll · 5 divider · 6 line-up · 14 thank-you
# 15 backup divider · 16 trip schedule
SKIP = {1, 2, 4, 5, 6, 14, 15, 16}

# ---------------------------------------------------------------- per slide
# "bullets": build a bullet box by top-level paragraph groups (first group stays
#            visible with the slide).
# "beats":   explicit list of beats; each beat is a list of shape names revealed
#            together on one click.
# Slides 7 onward carry NO animations (Nico, 2026-08-22): once the class hands
# over to the students, everything on a presentation slide should be on screen
# from the moment it appears.  They are deliberately left OUT of SKIP so that a
# rerun still strips any timing a previous pass may have written.
PLAN = {
    3:  dict(bullets="TextBox 6"),
}


def q(ns, t):
    return "{%s}%s" % (ns, t)


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


class Counter:
    def __init__(self, start=3):
        self.n = start

    def __call__(self):
        v = self.n
        self.n += 1
        return v


def sptgt(spid, prg):
    if prg is None:
        return '<p:spTgt spid="%s"/>' % spid
    return ('<p:spTgt spid="%s"><p:txEl><p:pRg st="%d" end="%d"/></p:txEl>'
            "</p:spTgt>" % (spid, prg[0], prg[1]))


def effect_par(spid, prg, node_type, ids):
    a, b, c = ids(), ids(), ids()
    tgt = sptgt(spid, prg)
    return (
        '<p:par><p:cTn id="%d" presetID="10" presetClass="entr" presetSubtype="0" '
        'fill="hold" grpId="0" nodeType="%s"><p:stCondLst><p:cond delay="0"/>'
        "</p:stCondLst><p:childTnLst>"
        '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold"><p:stCondLst>'
        '<p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl>%s</p:tgtEl>'
        "<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>"
        "</p:cBhvr><p:to><p:strVal val=\"visible\"/></p:to></p:set>"
        '<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="%d" '
        'dur="500"/><p:tgtEl>%s</p:tgtEl></p:cBhvr></p:animEffect>'
        "</p:childTnLst></p:cTn></p:par>"
        % (a, node_type, b, tgt, c, tgt))


def click_group(beat, ids):
    outer, inner = ids(), ids()
    effs = "".join(
        effect_par(spid, prg, "clickEffect" if i == 0 else "withEffect", ids)
        for i, (spid, prg) in enumerate(beat))
    return (
        '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst>'
        '<p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
        '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst><p:cond delay="0"/>'
        "</p:stCondLst><p:childTnLst>%s</p:childTnLst></p:cTn></p:par>"
        "</p:childTnLst></p:cTn></p:par>" % (outer, inner, effs))


def timing_xml(beats, para_boxes):
    ids = Counter(3)
    groups = "".join(click_group(b, ids) for b in beats)
    bld = "".join('<p:bldP spid="%s" grpId="0" build="p"/>' % s for s in para_boxes)
    bldlst = "<p:bldLst>%s</p:bldLst>" % bld if bld else ""
    return (
        '<p:timing xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:tnLst><p:par>'
        '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        '<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" '
        'dur="indefinite" nodeType="mainSeq"><p:childTnLst>%s</p:childTnLst>'
        "</p:cTn>"
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/>'
        "</p:tgtEl></p:cond></p:prevCondLst>"
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/>'
        "</p:tgtEl></p:cond></p:nextCondLst>"
        "</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>%s</p:timing>"
        % (A, R, P, groups, bldlst))


def shape_info(el):
    cnv = el.find(".//" + q(P, "cNvPr"))
    paras = []
    for i, pp in enumerate(el.iter(q(A, "p"))):
        pPr = pp.find(q(A, "pPr"))
        lvl = int(pPr.get("lvl")) if (pPr is not None and pPr.get("lvl")) else 0
        txt = "".join(t.text or "" for t in pp.iter(q(A, "t"))).strip()
        paras.append((i, lvl, bool(txt)))
    return dict(id=int(cnv.get("id")), name=cnv.get("name") or "", paras=paras)


def bullet_groups(box):
    """Top-level paragraph groups: [[idx, ...], ...] in document order."""
    groups, cur = [], None
    for (idx, lvl, ne) in box["paras"]:
        if lvl == 0:
            cur = []
            groups.append(cur)
        elif cur is None:
            cur = []
            groups.append(cur)
        if ne:
            cur.append(idx)
    return [g for g in groups if g]


def plan_slide(disp, shapes):
    spec = PLAN.get(disp)
    if not spec:
        return [], []
    by_name = {}
    for s in shapes:
        by_name.setdefault(s["name"], []).append(s)

    if "bullets" in spec:
        box = by_name[spec["bullets"]][0]
        groups = bullet_groups(box)[1:]          # first group shows with the slide
        beats = [[(box["id"], (i, i)) for i in g] for g in groups]
        return beats, [box["id"]]

    beats = []
    for beat in spec["beats"]:
        members = []
        for name in beat:
            for s in by_name.get(name, []):
                members.append((s["id"], None))
        if members:
            beats.append(members)
    return beats, []


def main():
    args = sys.argv[1:]
    apply = "apply" in args
    sel = [a for a in args if a != "apply"]
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()

    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target")
             for r in ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    order = [os.path.basename(rid2t[s.get(q(R, "id"))])
             for s in pres.find(q(P, "sldIdLst"))]

    if sel in ([], ["all"]):
        todo = [d for d in range(1, len(order) + 1) if d not in SKIP]
    else:
        todo = [int(x) for x in sel]

    report = []
    for disp in todo:
        part = order[disp - 1]
        root = ET.fromstring(data["ppt/slides/" + part])
        tree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
        shapes = [shape_info(el) for el in tree
                  if ET.QName(el).localname in ("sp", "pic", "graphicFrame", "grpSp")]
        beats, para_ids = plan_slide(disp, shapes)
        for old in root.findall(q(P, "timing")):
            root.remove(old)
        if beats:
            root.append(ET.fromstring(timing_xml(beats, para_ids)))
        data["ppt/slides/" + part] = ser(root)
        report.append((disp, len(beats)))

    for disp, n in report:
        print("slide %3d: %d click-beats" % (disp, n))
    print("== %d slides, %d click-beats ==" % (len(report),
                                               sum(n for _, n in report)))

    if apply:
        tmp = DECK.with_suffix(".pptx.tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for name, blob in data.items():
                out.writestr(name, blob)
        with zipfile.ZipFile(tmp) as chk:
            assert chk.testzip() is None
        shutil.move(str(tmp), str(DECK))
        print("APPLIED")
    else:
        print("DRY RUN (pass 'apply' to write)")


if __name__ == "__main__":
    main()
