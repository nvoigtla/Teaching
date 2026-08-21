# -*- coding: utf-8 -*-
"""Inject Fade / on-click build animations via OOXML <p:timing> —
Module 2 VIDEO PART (2026-08-16). Engine from Module 7 / Italy IBR.

Teaching CLAUDE.md rules implemented:
  - Fade entrance, 0.5 s, on click; one style deck-wide. Chrome (top
    bar, tag, title, rules, footer, page number, chart axes/backings,
    video-link boxes, back buttons) never animates.
  - Text slides: first top-level bullet (with subs) shows with the
    slide; build from the second, one top-level bullet per click.
    Figures + captions ride on the click of the bullet they support.
  - Diagram/chart slides: custom story plans (PLANS below) — curve +
    its label on one click, guides before regions, callouts in story
    order, the gold equilibrium/takeaway element last. Unlisted shapes
    stay visible from the start (axes, ticks that anchor the setup).
  - Skips: title, roadmap, outline/agenda slides, section cards, poll
    and video slides, the practice-video index.

Pipeline:  _build_Module2InClass.py -> _splice_media.py -> _animate.py all apply
Verify via PowerPoint COM (MainSequence counts) + eyeball.

Selector language for PLANS beats:
  "t:PREFIX"    first UNUSED text shape whose text starts with PREFIX
  "t:PREFIX#n"  the n-th (1-based, doc order) text match, used or not
  "pic:N" / "cxn:N" / "grp:N" / "osp:N"   n-th picture / connector /
                group / textless shape in document order
  "pr:PREFIX:st:end"  paragraph range st..end of the text box whose
                text starts with PREFIX (adds a bldP for the box)
"""
import os
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
import sys as _sys
_deckarg = [a for a in _sys.argv[1:]
            if a.endswith(".pptx")]
DECK = Path(_deckarg[0]) if _deckarg else     HERE / "Module 2 - Video Part Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
EMU = 914400.0

SKIP_TITLE = {1, 19, 32}
SKIP_AGENDA = {2, 20, 33, 57}
SKIP_MEDIA = {9, 28}
SKIP_STATIC = {40}                 # data table stays static
SKIP = SKIP_TITLE | SKIP_AGENDA | SKIP_MEDIA | SKIP_STATIC

FIG_GROUP_DEFAULT = 0
FIG_GROUP = {}
PLANS_PRE = {}
STATIC = {}

PLANS = {
    3: [
        ["cxn:2"],
        ["grp:0"],
        ["t:This is the demand"],
    ],
    4: [
        ["cxn:4", "t:D"],
        ["cxn:5", "cxn:6"],
        ["osp:3", "t:Total revenue"],
        ["cxn:7", "cxn:8"],
    ],
    5: [
        ["cxn:3", "t:P₀", "cxn:5", "t:Q₀", "t:TR₀"],
        ["cxn:4", "t:P₁", "cxn:6", "t:Q₁", "t:TR₁"],
        ["osp:3", "grp:0", "cxn:7"],
        ["osp:4", "grp:1", "cxn:8"],
        ["t:Which area is bigger?", "t:Total revenue"],
    ],
    6: [
        ["t:Eᴅ < −1"],
        ["t:−1 < Eᴅ"],
        ["osp:3", "t:Eᴅ = −1"],
        ["osp:4", "t:TR#2", "cxn:5"],
        ["grp:1"],
        ["grp:2"],
    ],
    7: [
        ["pr:Total Revenues:1:1"],
        ["t:To assess"],
        ["grp:0"],
        ["grp:1"],
    ],
    8: [
        ["pr:Information:1:1"],
        ["pr:Information:2:2"],
        ["pr:Information:3:3"],
    ],
    10: [
        ["pr:Demand is:1:2"],
    ],
    11: [
        ["pr:Novo Nordisk:1:1"],
        ["pr:Novo Nordisk:2:2"],
    ],
    12: [
        ["t:Demand is"],
        ["pr:A price cut:0:0"],
        ["pr:A price cut:1:1"],
        ["pr:A price cut:2:2"],
    ],
    13: [
        ["pr:The price decrease:0:0"],
        ["pr:The price decrease:1:1"],
        ["pr:The price decrease:2:2"],
        ["pr:The price decrease:3:3"],
    ],
    14: [
        ["pr:Is the percent:1:1"],
    ],
    15: [
        ["pr:In 2014:1:1"],
        ["pr:In 2014:2:5"],
    ],
    16: [
        ["t:E D"],
        ["t:where"],
        ["pr:→:0:0"],
        ["pr:→:1:1"],
        ["pr:→:2:2"],
    ],
    17: [
        ["pr:What does this suggest:1:1"],
        ["pr:What does this suggest:2:2"],
        ["pr:What does this suggest:3:3"],
    ],
    18: [
        ["pr:April 2025:1:1"],
        ["pr:April 2025:2:2"],
        ["pr:April 2025:3:3"],
        ["grp:0"],
    ],
    22: [
        ["t:MR"],
        ["t:Derivative", "cxn:0"],
        ["t:How to compute"],
    ],
    23: [
        ["pr:::1:1"],
        ["pr:::2:3"],
        ["pr:::4:4"],
        ["pr:::5:5"],
        ["t:→"],
    ],
    24: [
        ["t:Step 1", "t:P"],
        ["t:Step 2", "t:TR"],
        ["t:Step 3", "t:MR"],
    ],
    25: [
        ["pr:1.  Calculate inverse:2:3"],
        ["pr:1.  Calculate inverse:4:5"],
        ["t:→"],
    ],
    26: [
        ["cxn:5", "t:MR"],
        ["osp:3", "t:Eᴅ = −1"],
        ["t:Elastic portion", "t:Inelastic portion"],
        ["osp:4", "t:TR#2", "cxn:6"],
        ["grp:1", "cxn:7"],
    ],
    27: [
        ["pr:MR:1:1"],
        ["pr:MR:2:2"],
        ["pr:MR:3:3"],
        ["pr:MR:4:4"],
    ],
    29: [
        ["t:1.", "t:P"],
        ["t:2.", "t:TR"],
        ["t:3.", "t:MR"],
    ],
    30: [
        ["pr:What is the revenue:0:0"],
        ["pr:What is the revenue:1:1"],
    ],
    31: [
        ["pr:1.  Inverse demand:0:2"],
        ["t:MR"],
        ["pr:Why revenue maximization:0:0"],
    ],
    38: [
        ["pr:Description::1:1"],
        ["pr:Description::2:2"],
    ],
    39: [
        ["t:Slope", "cxn:0"],
        ["t:where:"],
        ["t:E"],
    ],
    41: [
        ["osp:3", "osp:4", "osp:5", "osp:6", "osp:7", "osp:8", "osp:9",
         "osp:10", "osp:11", "osp:12", "osp:13", "osp:14", "osp:15",
         "osp:16", "osp:17", "osp:18", "t:Scatterplot"],
    ],
    42: [
        ["cxn:3", "cxn:4"],
        ["grp:0"],
        ["cxn:5", "cxn:6", "cxn:7"],
        ["cxn:2"],
    ],
    43: [
        ["grp:0"],
        ["t:Demand Equation:"],
        ["t:Note: Now"],
        ["t:→"],
    ],
    44: [
        ["pr:R Square:1:1"],
        ["pr:R Square:2:3"],
        ["pr:R Square:4:4", "t:e.g.,"],
    ],
    45: [
        ["pr:We have estimated:1:1"],
        ["pr:We have estimated:2:2"],
        ["t:Q = 151"],
        ["t:E"],
    ],
    46: [
        ["t:where:"],
    ],
    47: [
        ["grp:0"],
        ["t:Demand Equation:"],
    ],
    48: [
        ["pr:Currently:1:1"],
        ["pr:Currently:2:2"],
        ["pr:Currently:3:3"],
    ],
    49: [
        ["pr:Q = 479:1:2"],
        ["t:E"],
        ["t:→"],
    ],
    50: [
        ["t:1.", "t:P"],
        ["t:2.", "t:TR"],
        ["t:3.", "t:MR#1"],
        ["t:MR#2"],
    ],
    51: [
        ["pr:MR:1:1"],
        ["cxn:3", "t:MR#2", "osp:3", "t:Eᴅ = −1", "t:Elastic portion",
         "t:Inelastic portion"],
    ],
    53: [
        ["pic:0"],
        ["t:Actual finding"],
    ],
    54: [
        ["pic:0", "pic:1", "pic:2"],
    ],
    56: [
        ["pr:Randomizing:1:1"],
    ],
}


def q(ns, t):
    return "{%s}%s" % (ns, t)


class Counter:
    def __init__(self, start=3):
        self.n = start

    def __call__(self):
        v = self.n
        self.n += 1
        return v


def sptgt(spid, prg):
    if prg is None:
        return '<p:spTgt spid="%d"/>' % spid
    return ('<p:spTgt spid="%d"><p:txEl><p:pRg st="%d" end="%d"/>'
            '</p:txEl></p:spTgt>' % (spid, prg[0], prg[1]))


def effect_par(spid, prg, node_type, ids):
    a, b, c = ids(), ids(), ids()
    tgt = sptgt(spid, prg)
    return (
        '<p:par><p:cTn id="%d" presetID="10" presetClass="entr" '
        'presetSubtype="0" fill="hold" grpId="0" nodeType="%s">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '<p:tgtEl>%s</p:tgtEl><p:attrNameLst><p:attrName>'
        'style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
        '<p:to><p:strVal val="visible"/></p:to></p:set>'
        '<p:animEffect transition="in" filter="fade"><p:cBhvr>'
        '<p:cTn id="%d" dur="500"/><p:tgtEl>%s</p:tgtEl></p:cBhvr>'
        '</p:animEffect></p:childTnLst></p:cTn></p:par>'
        % (a, node_type, b, tgt, c, tgt))


def click_group(beat, ids):
    outer, inner = ids(), ids()
    effs = "".join(
        effect_par(spid, prg, "clickEffect" if i == 0 else "withEffect",
                   ids)
        for i, (spid, prg) in enumerate(beat))
    return (
        '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst>'
        '<p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
        '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst>'
        '<p:cond delay="0"/></p:stCondLst><p:childTnLst>%s'
        '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'
        % (outer, inner, effs))


def timing_xml(beats, para_boxes):
    ids = Counter(3)
    groups = "".join(click_group(b, ids) for b in beats)
    bld = "".join('<p:bldP spid="%d" grpId="0" build="p"/>' % spid
                  for spid in sorted(set(para_boxes)))
    bldlst = '<p:bldLst>%s</p:bldLst>' % bld if bld else ""
    return (
        '<p:timing xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:tnLst>'
        '<p:par><p:cTn id="1" dur="indefinite" restart="never" '
        'nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" '
        'nextAc="seek"><p:cTn id="2" dur="indefinite" '
        'nodeType="mainSeq"><p:childTnLst>%s</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl>'
        '<p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst>'
        '<p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl>'
        '</p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn>'
        '</p:par></p:tnLst>%s</p:timing>' % (A, R, P, groups, bldlst))


def shape_info(el, kinds):
    cnv = el.find(".//" + q(P, "cNvPr"))
    xf = el.find(".//" + q(A, "xfrm"))
    x = y = w = h = 0.0
    if xf is not None:
        o = xf.find(q(A, "off"))
        e = xf.find(q(A, "ext"))
        if o is not None and e is not None:
            x, y = int(o.get("x")) / EMU, int(o.get("y")) / EMU
            w, h = int(e.get("cx")) / EMU, int(e.get("cy")) / EMU
    paras = []
    for idx, pp in enumerate(el.iter(q(A, "p"))):
        lvl = 0
        pPr = pp.find(q(A, "pPr"))
        if pPr is not None and pPr.get("lvl"):
            lvl = int(pPr.get("lvl"))
        txt = "".join(t.text or "" for t in pp.iter(q(A, "t"))).strip()
        paras.append((idx, lvl, bool(txt)))
    tag = ET.QName(el).localname
    txt = " ".join(t.text or "" for t in el.iter(q(A, "t")))
    txt = (txt + " " + " ".join(t.text or ""
                                for t in el.iter(q(M, "t")))).strip()
    info = {
        "id": int(cnv.get("id")), "name": cnv.get("name") or "",
        "tag": tag, "x": x, "y": y, "w": w, "h": h, "text": txt,
        "paras": paras,
        "bul0": sum(1 for (_, l, ne) in paras if l == 0 and ne),
    }
    key = {"pic": "pic", "cxnSp": "cxn", "grpSp": "grp"}.get(tag)
    if key is None:
        key = "osp" if not txt else None
    if key:
        info["idx"] = "%s:%d" % (key, kinds[key])
        kinds[key] += 1
    else:
        info["idx"] = None
    return info


def collect_shapes(spTree):
    kinds = {"pic": 0, "cxn": 0, "grp": 0, "osp": 0}
    shapes = []
    for el in spTree:
        tag = ET.QName(el).localname
        if tag == "AlternateContent":
            ch = el.find(q(MC, "Choice"))
            if ch is not None and len(ch):
                el = ch[0]
                tag = ET.QName(el).localname
        if tag == "graphicFrame":
            cnv = el.find(".//" + q(P, "cNvPr"))
            xf = el.find(q(P, "xfrm"))
            x = y = 0.0
            if xf is not None and xf.find(q(A, "off")) is not None:
                o = xf.find(q(A, "off"))
                x, y = int(o.get("x")) / EMU, int(o.get("y")) / EMU
            shapes.append({"id": int(cnv.get("id")), "name": "",
                           "tag": "graphicFrame", "x": x, "y": y,
                           "w": 0, "h": 0, "text": "", "paras": [],
                           "bul0": 0,
                           "idx": "gf:%d" % kinds.setdefault("gf", 0)})
            kinds["gf"] += 1
            continue
        if tag not in ("sp", "cxnSp", "pic", "grpSp"):
            continue
        shapes.append(shape_info(el, kinds))
    return shapes


def is_chrome(s):
    t = s["text"].strip()
    if s["y"] + s["h"] <= 1.32 and not (1.3 < s["y"]):
        return True                      # top bar / tag / title / rule
    if s["y"] >= 7.05:
        return True                      # footer rules / text / page no
    if t.startswith("Management 405"):
        return True
    if s["x"] > 12.4 and s["tag"] == "sp" and len(t) <= 3:
        return True                      # page number
    if t.startswith("\u25b6"):
        return True                      # practice-video link boxes
    if t.startswith("\u2190 Back") or t.startswith("Tickets"):
        return True                      # navigation pills
    if (s["tag"] == "sp" and not t and s["w"] > 4 and s["h"] > 2.5):
        return True                      # big white chart backings
    return False


def resolve(sel, shapes, used):
    """Resolve one selector to (spid, prg) or None."""
    if sel.startswith("pr:"):
        rest = sel.split(":", 1)[1]
        prefix, st, end = rest.rsplit(":", 2)
        for s in shapes:
            if s["text"].startswith(prefix):
                # one effect per paragraph, so PowerPoint keeps them in
                # this click group (a single multi-para range effect
                # gets re-expanded into separate clicks on open)
                targets = [(s["id"], (i, i))
                           for i in range(int(st), int(end) + 1)]
                return targets, s["id"]
        raise KeyError(sel)
    if sel.startswith("t:"):
        body = sel[2:]
        nth = None
        if "#" in body:
            body, n = body.rsplit("#", 1)
            nth = int(n)
        matches = [s for s in shapes
                   if s["text"].strip().startswith(body)]
        if nth is not None:
            if len(matches) < nth:
                raise KeyError(sel)
            return (matches[nth - 1]["id"], None), None
        for s in matches:
            if s["id"] not in used:
                used.add(s["id"])
                return (s["id"], None), None
        raise KeyError(sel)
    # indexed selectors
    for s in shapes:
        if s["idx"] == sel:
            return (s["id"], None), None
    raise KeyError(sel)


def para_beats(box):
    groups = []
    cur = None
    for (idx, lvl, ne) in box["paras"]:
        if lvl == 0:
            cur = [(idx, ne)]
            groups.append(cur)
        else:
            if cur is None:
                cur = []
                groups.append(cur)
            cur.append((idx, ne))
    clean = [[i for (i, ne) in g if ne] for g in groups]
    clean = [g for g in clean if g]
    if not clean:
        return [], []
    return clean[0], clean[1:]


def default_plan(shapes, disp):
    """Bullet-driven default: first top-level bullet static, one click
    per further top-level bullet; figures + captions ride on the
    FIG_GROUP bullet; a gold takeaway bar (if any) last."""
    content = [s for s in shapes if not is_chrome(s)]
    for sel in STATIC.get(disp, ()):
        body = sel[2:]
        content = [s for s in content
                   if not s["text"].strip().startswith(body)]
    if not content:
        return [], []
    boxes = [s for s in content if s["tag"] == "sp" and s["bul0"] >= 2]
    bullets = max(boxes, key=lambda s: s["bul0"]) if boxes else None
    beats, para_ids = [], []
    takeaways = [s for s in content
                 if s["name"].startswith("Rounded Rectangle")
                 and s["y"] > 5.8 and s["text"].strip().startswith("\u2192")]
    t_ids = {s["id"] for s in takeaways}
    if bullets is not None:
        _static, groups = para_beats(bullets)
        bullet_beats = [[(bullets["id"], (i, i)) for i in g]
                        for g in groups]
        para_ids.append(bullets["id"])
        figs = [s for s in content
                if s["id"] != bullets["id"] and s["id"] not in t_ids]
        figs.sort(key=lambda s: (0 if s["tag"] in ("pic", "grpSp")
                                 else 1, s["y"], s["x"]))
        fig_targets = [(s["id"], None) for s in figs]
        if fig_targets:
            if not bullet_beats:
                bullet_beats.append(fig_targets)
            else:
                gi = FIG_GROUP.get(disp, FIG_GROUP_DEFAULT)
                gi = max(0, min(gi if gi >= 0
                                else len(bullet_beats) + gi,
                                len(bullet_beats) - 1))
                bullet_beats[gi] += fig_targets
        beats.extend(bullet_beats)
    else:
        # no bullet box: reveal pictures (with nearest caption), then
        # remaining text boxes top-to-bottom
        pics = [s for s in content if s["tag"] in ("pic", "grpSp")]
        caps = [s for s in content if s["tag"] == "sp"
                and s["text"].strip() and s["id"] not in t_ids]
        cap_of = {}
        for c in caps:
            if pics:
                best = min(pics, key=lambda p: (p["x"] + p["w"] / 2
                                                - c["x"] - c["w"] / 2) ** 2
                           + (p["y"] + p["h"] / 2 - c["y"] - c["h"] / 2) ** 2)
                if abs(best["y"] + best["h"] - c["y"]) < 0.4 or \
                   abs(c["y"] + c["h"] - best["y"]) < 0.4:
                    cap_of.setdefault(best["id"], []).append(c["id"])
        capped = {cid for v in cap_of.values() for cid in v}
        for p_ in sorted(pics, key=lambda s: (round(s["y"] / 1.4), s["x"])):
            beats.append([(p_["id"], None)]
                         + [(cid, None) for cid in cap_of.get(p_["id"], [])])
        for s in sorted([c for c in caps if c["id"] not in capped],
                        key=lambda s: s["y"]):
            beats.append([(s["id"], None)])
    for t in takeaways:
        beats.append([(t["id"], None)])
    return beats, para_ids


def custom_plan(shapes, disp):
    used = set()
    beats, para_ids = [], []
    for beat_sel in PLANS[disp]:
        beat = []
        for sel in beat_sel:
            (target, pbox) = resolve(sel, shapes, used)
            if isinstance(target, list):
                beat.extend(target)
            else:
                beat.append(target)
            if pbox is not None:
                para_ids.append(pbox)
        beats.append(beat)
    return beats, para_ids


def main():
    args = sys.argv[1:]
    apply = "apply" in args
    sel = [a for a in args
           if a != "apply" and not a.endswith(".pptx")]
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target") for r in
             ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    order = [os.path.basename(rid2t[s.get(q(R, "id"))])
             for s in pres.find(q(P, "sldIdLst"))]

    if sel == ["all"] or not sel:
        todo = [d for d in range(1, 58) if d not in SKIP]
    else:
        todo = [int(x) for x in sel]

    PLANS.update(PLANS_PRE)
    for disp in todo:
        part = "ppt/slides/" + order[disp - 1]
        tree = ET.fromstring(data[part])
        spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
        shapes = collect_shapes(spTree)
        if disp in PLANS:
            beats, para_ids = custom_plan(shapes, disp)
        else:
            beats, para_ids = default_plan(shapes, disp)
        if not beats:
            print("s%02d: no beats — skipped" % disp)
            continue
        # strip existing timing, then append the new block
        for t_el in tree.findall(q(P, "timing")):
            tree.remove(t_el)
        timing = ET.fromstring(timing_xml(beats, para_ids))
        tree.append(timing)
        data[part] = ET.tostring(tree, xml_declaration=True,
                                 encoding="UTF-8", standalone=True)
        print("s%02d: %d clicks (%d effects)"
              % (disp, len(beats), sum(len(b) for b in beats)))

    if apply:
        tmp = DECK.with_suffix(".anim_tmp.pptx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, blob in data.items():
                zout.writestr(name, blob)
        os.replace(tmp, DECK)
        print("written:", DECK)
    else:
        print("(dry run — pass 'apply' to write)")


if __name__ == "__main__":
    main()
