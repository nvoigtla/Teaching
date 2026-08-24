# -*- coding: utf-8 -*-
"""Inject Fade / on-click build animations via OOXML <p:timing> —
Module 2 In-Class (2026-08-15). Engine from Module 7 / Italy IBR.

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
DECK = Path(_deckarg[0]) if _deckarg else     HERE / "Module 2 - In Class Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
EMU = 914400.0

SKIP_TITLE = {1}
SKIP_AGENDA = {6, 7, 8, 25}          # roadmap + outline/section agendas
SKIP_MEDIA = {4, 5, 11, 12, 13, 32, 33, 37, 38, 42, 43, 49, 50,
              61, 62, 69, 70}        # PollEv pairs + live Excel slide
SKIP_STATIC = {48,                   # single-thought poll setup (Q=10−P)
               16, 17}                # Gates quote + Inglehart panel —
                                      # Nico removed the builds by hand
                                      # (2026-08-23)
SKIP = SKIP_TITLE | SKIP_AGENDA | SKIP_MEDIA | SKIP_STATIC

# figures ride on this animated-bullet index by default (0 = first
# built bullet); per-slide overrides:
FIG_GROUP_DEFAULT = 0
FIG_GROUP = {
    23: -1,   # network-effects screenshot -> last bullet
}

PLANS_PRE = {}

STATIC = {
    10: ["t:Poll Break"],
    36: ["t:Poll Break"],
}

# ---------------------------------------------------------------------------
# Custom story plans (selector language in the header docstring)
# ---------------------------------------------------------------------------
PLANS = {
    9: [   # law of demand (Nico's boxed sections, 2026-08-15; re-cut by
        # hand 2026-08-23): assumption box static; law box; then the whole
        # demand-curve figure as one grouped object; reasons in one beat
        ["grp:1"],
        ["grp:2"],
        ["pr:Reasons:0:2"],
    ],
    15: [  # staircase: bar + price + label per click; guide + takeaway last
        ["osp:3", "t:$12", "t:MPV of 1st slice"],
        ["osp:4", "t:$9", "t:MPV of 2nd slice"],
        ["osp:5", "t:$6", "t:MPV of 3rd slice"],
        ["osp:6", "t:$3", "t:MPV of 4th slice"],
        ["t:MPV of 5th slice", "t:$0"],
        ["cxn:2", "t:Diminishing Marginal Personal Value"],
    ],
    18: [  # consumer optimization (2026-08-16 boxed redesign; re-cut by
        # hand 2026-08-23): recall box static, then the MB = MC star as
        # the general rule, then the MPV version of it as the payoff
        ["grp:2"],
        ["grp:1"],
    ],
    19: [  # movies, re-cut by hand 2026-08-23: MPV curve + its labels,
        # then the "MPV is the demand curve" callout (now grouped with its
        # arrow), then the upward-sloping MC, then Q*
        ["osp:3", "t:MPV#1", "t:MPV diminishes"],
        ["grp:0"],
        ["cxn:2", "t:MC (incl."],
        ["cxn:3", "t:Q*"],
    ],
    20: [  # aggregation, re-cut by hand 2026-08-23: consumer 1, consumer
        # 2, the plus signs, then one horizontal sum per row (each "="
        # grouped with the aggregate dot it produces), aggregate curve last
        ["cxn:2", "osp:3", "osp:5", "osp:7", "osp:9", "osp:11",
         "t:Consumer 1 demand"],
        ["cxn:3", "osp:4", "osp:6", "osp:8", "osp:10", "osp:12",
         "t:Consumer 2 demand"],
        ["grp:3"],
        ["grp:0", "grp:5"],
        ["grp:1"],
        ["grp:2"],
        ["grp:4"],
        ["cxn:4"],
    ],
    21: [  # factors, re-cut by hand 2026-08-23: the demand-shift figure
        # first (falling, then rising — each curve with its own arrow and
        # label as one group), then the bullets one at a time, with the
        # cartoon riding on the Ryanair sub-bullet
        ["grp:1"],
        ["grp:0"],
        ["pr:Income:0:0"],
        ["pr:Income:1:3"],
        ["pr:Income:4:4", "pic:0"],
        ["pr:Income:5:5"],
        ["pr:Income:6:6"],
        ["pr:Income:7:7"],
        ["pr:Income:8:8"],
        ["t:Anything else?"],
    ],
    22: [  # snob news: the concept box, then Ferrari panel (+source),
        # then Birkin panel
        ["grp:0"],
        ["pic:0", "pic:1", "t:Source: The Wall Street"],
        ["pic:2", "pic:3"],
    ],
    27: [  # Netflix: chart static; implications, key concept, video pointer
        ["pr:Netflix frequently raises:1:3"],
        ["pr:Netflix frequently raises:4:4"],
        ["pr:Netflix frequently raises:5:5"],
    ],
    28: [  # what is elasticity: formula static; unit-free; %-label last
        ["pr:Definition::2:2"],
        ["pr:Definition::3:3", "t:percentage change", "cxn:0"],
    ],
    29: [  # three types: one card per click, recipe line last
        ["t:Own-price elasticity", "t:E D"],
        ["t:Income elasticity", "t:E I"],
        ["t:Cross-price elasticity", "t:E X"],
        ["t:All three follow"],
    ],
    30: [  # own-price definition: formula static; bullets build
        ["pr:% change in quantity:1:1"],
        ["pr:% change in quantity:2:2"],
    ],
    31: [  # water: photo static; box -> arrow+box -> arrow+question box
        ["t:Consulting company"],
        ["cxn:0", "t:LADWP wants"],
        ["cxn:1", "t:By how much"],
    ],
    34: [  # water solution: equation lines one per click; takeaway last
        ["t:E D"],
        ["t:−0.4"],
        ["t:%Δ P"],
        ["t:Raise the price"],
    ],
    35: [  # categories: formula static; one category row per click
        ["t:If", "t::"],
        ["t:If", "t::"],
        ["t:If", "t::"],
    ],
    39: [  # yoga solution: equation, question, answer + interpretation
        ["t:E d"],
        ["pr:Is demand for yoga:0:0"],
        ["pr:Is demand for yoga:1:2"],
    ],
    40: [  # method 1: formula + convention; caveats; PS pointer
        ["t:E D", "grp:0"],
        ["pr:This method approximates:0:1"],
        ["t:→"],
    ],
    41: [  # e-books: quote + logo static; question builds
        ["pr:Amazon said:2:2"],
    ],
    44: [  # e-books solution
        ["pr:Problem states:1:1"],
        ["t:E D"],
        ["t:→"],
    ],
    45: [  # Mega Millions: facts, equation, caution + source
        ["pr:April 2025:1:1"],
        ["t:E D"],
        ["grp:0", "t:Source: Hansen"],
    ],
    46: [  # method 2: formula + slope callout; interpretation; linear note
        ["t:E D", "grp:0"],
        ["pr:ΔQ/ΔP is the slope:0:2"],
        ["pr:ΔQ/ΔP is the slope:3:3"],
    ],
    47: [  # point elasticity: step + its result box per click; note last
        ["t:Step 1", "t:Q"],
        ["t:Step 2", "t:dQ dP"],
        ["t:Step 3", "t:Q"],
        ["t:Step 4", "t:E D"],
        ["t:#Note#1"],
    ],
    51: [  # Q=10−P solution: steps, formula with step 4, verdict
        ["pr:Answer is:1:1"],
        ["pr:Answer is:2:2"],
        ["pr:Answer is:3:3"],
        ["pr:Answer is:4:4", "t:E D"],
        ["t:→"],
    ],
    52: [  # linear case: D, explanation, regions, unit point, intercepts,
        # "don't operate there" payoff last
        ["cxn:2", "t:D"],
        ["t:Linear demand curve"],
        ["t:Eᴅ < −1"],
        ["t:−1 < Eᴅ < 0"],
        ["osp:4", "t:Eᴅ = −1"],
        ["t:Eᴅ = −∞", "t:Eᴅ = 0"],
        ["grp:0", "cxn:3"],
    ],
    54: [  # Uber: chart + logo static; elasticity callout builds
        ["grp:0", "cxn:0"],
    ],
    55: [  # special cases: intro static; left panel, then right panel
        ["cxn:2", "t:D#1", "t:Perfectly Elastic", "t:E D"],
        ["cxn:5", "t:D#2", "t:Perfectly Inelastic", "t:E D"],
    ],
    56: [  # determinants: substitutes block static; then remaining factors
        ["pr:Availability of substitutes:3:3"],
        ["pr:Availability of substitutes:4:4"],
    ],
    57: [  # market vs firm: first bullet static; conditions + photo together
        ["pr:The elasticity a company:1:4", "grp:0"],
    ],
    58: [  # re-anchor: own-price (faded) static; other two cards; next line
        ["t:Income elasticity", "t:E I"],
        ["t:Cross-price elasticity", "t:E X"],
        ["t:Own-price elasticity ✓"],
    ],
    59: [  # income elasticity definition
        ["pr:% change in quantity:1:1"],
    ],
    60: [  # Rivian example: photo + badge static
        ["pr:Average income:1:1"],
        ["pr:Average income:2:2"],
    ],
    63: [  # R3 solution
        ["pr:Answer is:1:1"],
        ["pr:Answer is:2:4"],
        ["t:E I"],
        ["t:Does this seem"],
    ],
    64: [  # income categories: row + its picture + caption per click
        ["pr:,:0:0", "grp:0"],
        ["pr:,:1:1", "grp:1"],
        ["pr:,:2:2", "grp:2"],
    ],
    65: [  # recession retailers: lines static; recession marker; then the
        # two stories (bullet + its curve label together)
        ["cxn:2", "t:U.S. economy"],
        ["pr:In 2007:1:1", "t:Walmart’s stock rose"],
        ["pr:In 2007:2:2", "t:Target’s stock fell"],
    ],
    66: [  # inferior-goods news: whole panel one beat
        ["pic:0", "pic:1", "pic:2", "t:Source: The Wall Street"],
    ],
    67: [  # cross-price definition: formula static; cases with pictures
        ["pr:% change in the quantity:1:1"],
        ["pr:% change in the quantity:2:3", "grp:0"],
        ["pr:% change in the quantity:4:4", "grp:1"],
    ],
    68: [  # popcorn example: photo + badge static; question builds
        ["pr:When a movie theater:1:1"],
    ],
    71: [  # popcorn solution
        ["pr:Unit free:1:2"],
        ["t:E X"],
        ["pr:Popcorn and movie tickets:0:0"],
    ],
    72: [  # cross-price news: whole panel one beat
        ["pic:0", "pic:1", "pic:2", "t:Source: The Wall Street"],
    ],
    73: [  # cereal table: table + boxes static; takeaway line builds
        ["t:Larger cross-price"],
    ],
    74: [  # cheat sheet: definition, categories, computing — 3 beats
        ["t:Definition", "t:E D", "t:Intuitively"],
        ["t:Categories", "t:Inelastic"],
        ["t:Computing elasticity", "t:Two observed", "t:E d",
         "t:Whole demand", "t:E d", "t:ΔQ/ΔP = slope"],
    ],
    75: [  # post-work videos box + caption + arrow
        ["t:▶"],
    ],
    76: [  # post-work PS2 box, then the estimation note
        ["t:▶"],
        ["grp:0"],
    ],
}


# ---------------------------------------------------------------------------
# 2026-08-15 a bookend law-of-demand recap was inserted at display 14 and
# the config above (PRE-bookend numbering) was shifted +1 from 14 on.
# 2026-08-23 Nico deleted that slide again, so the config numbering is
# live as written and no shift is applied.
# ---------------------------------------------------------------------------


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
