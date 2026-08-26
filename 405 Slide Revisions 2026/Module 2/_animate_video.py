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

# 2026-08-24: Wrigley example (old 8-10) deleted, so every display
# from 11 on moved down by three.
# 2026-08-24: CT adoption of 30-52 (CT 34-60) — our extra
# elasticity slide dropped, CT's two regression slides merged,
# and six poll placeholders inserted.
SKIP_TITLE = {1, 13, 23}
SKIP_AGENDA = {14, 24}      # slide 2 is no longer static: in
                                # his final Video 1 deck the
                                # teaching-note link fades in on
                                # its own click (2026-08-25)
SKIP_MEDIA = {4, 5, 6, 19, 21}
                                # 4/5/6/22 arrive from the final
                                # decks with their own <p:timing>
SKIP_STATIC = {11, 31, 39, 40, 43}
                                # 36: the data table; 11/40/41/54
                                # were made fully static in his
                                # final decks (2026-08-25)
SKIP = SKIP_TITLE | SKIP_AGENDA | SKIP_MEDIA | SKIP_STATIC

FIG_GROUP_DEFAULT = 0
FIG_GROUP = {}
PLANS_PRE = {}
STATIC = {}

PLANS = {
    2: [   # the outline itself is static; only the teaching-note
           # link is revealed, on one click
        ["t:▶  Teaching Note"],
    ],
    3: [   # 2026-08-25: the two functions first, then the curve with
           # its labels as ONE object, then the note underneath
        ["grp:1"],
        ["grp:0"],
        ["t:Note: this is"],
    ],
    7: [   # 2026-08-25: the question and the FIRST card land together
        ["t:Total Revenues"],
        ["t:To assess", "grp:0"],
        ["grp:1"],
    ],
    8: [  # the revenue forecast, then the question with all four of
          # its answer options on one click (2026-08-25)
        ["pr:Novo Nordisk:1:1"],
        ["pr:Novo Nordisk:2:6"],
    ],
    9: [
        ["t:Demand is"],
        ["pr:A price cut:0:0"],
        ["pr:A price cut:1:1"],
        ["pr:A price cut:2:2"],
    ],
    10: [  # 2026-08-25: Profit = Revenue - Costs is revealed first
        ["t:Profit"],
        ["pr:We know:0:0"],
        ["pr:We know:1:1"],
        ["pr:We know:2:2"],
        ["pr:We know:3:3"],
    ],
    12: [  # 2026-08-25: first question + its options are up front;
           # one click brings the second question and its options
        ["pr:What does this suggest:4:8"],
    ],
    15: [  # CT 21: objective, its rule bar, the firm's objective, its
        # rule bar, then today's question
        ["t:General objective"],
        ["t:Net benefits are maximized"],
        ["t:Firms’ objective"],
        ["t:Produce where Marginal Revenue"],
        ["t:Today: how to compute"],
    ],
    16: [
        ["t:MR"],
        ["t:Derivative", "cxn:0"],
        ["t:How to compute"],
    ],
    17: [  # calculus refresher, relaid out: definition card, the
        # general rule, then the same rule on numbers
        ["grp:0"],
        ["t:The general rule", "t:If"],
        ["t:A worked example", "t:With a = 1"],
        ["t:→"],
    ],
    18: [  # three steps, each with its sub-line and its formula box
        ["t:Step 1", "t:P"],
        ["t:Step 2", "t:TR"],
        ["t:Step 3", "t:MR"],
        ["t:✎  Problem Set 2", "t:▤  Teaching Note"],
    ],
    22: [
        ["pr:MR:1:1"],
        ["pr:MR:2:2"],
        ["pr:MR:3:3"],
        ["pr:MR:4:4"],
    ],
    20: [
        ["t:1.", "t:P"],
        ["t:2.", "t:TR"],
        ["t:3.", "t:MR"],
    ],
    28: [  # the product, then the price history with ALL of its native
           # axis labels on the same beat (a figure and its labels are
           # one reveal)
        ["grp:0"],
        ["grp:1", "t:$90", "t:$70", "t:$50", "t:$30",
         "t:May", "t:Jun", "t:Jul", "t:Aug", "t:Sep",
         "t:$83.99", "t:$62.98", "t:$35.00"],
    ],
    # ---- Video 3 -----------------------------------------------------
    30: [  # only the elasticity formula is revealed
        ["t:E D  ="],
    ],
    32: [  # the gold line, then the blue one, then the question
        ["cxn:3"],
        ["cxn:2"],
        ["grp:0"],
    ],
    33: [  # just the two vertical residuals
        ["cxn:3"],
        ["cxn:4"],
    ],
    34: [  # the table is up front; equation, then the hand-off line
        ["grp:1"],
        ["t:Now that we have"],
    ],
    35: [  # the three questions the polls ask, one at a time
        ["pr:Use the estimated:1:1"],
        ["pr:Use the estimated:2:2"],
        ["pr:Use the estimated:3:3"],
    ],
    37: [  # each step of the 3-step method, then the answer
        ["pr:3-step method:1:1"],
        ["pr:3-step method:2:2"],
        ["pr:3-step method:3:3"],
        ["pr:3-step method:4:5"],
    ],
    41: [  # the airline example, then each of the two questions
        ["pr:Transaction data:2:4"],
        ["pr:Transaction data:6:6"],
        ["pr:Transaction data:8:8"],
    ],
    42: [  # the causal reading lands with both photos and the arrow
        ["t:Causal interpretation", "pic:0", "pic:1", "cxn:0"],
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
        "custgeom": el.find(".//" + q(A, "custGeom")) is not None,
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
    # a freeform CURVE is not a backing: a wide, tall Bezier bbox was
    # being swallowed here (video slide 6's TR hill), which silently
    # dropped it from every plan (2026-08-24)
    if (s["tag"] == "sp" and not t and s["w"] > 4 and s["h"] > 2.5
            and not s.get("custgeom")):
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
        todo = [d for d in range(1, len(order) + 1) if d not in SKIP]
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
