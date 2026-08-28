# -*- coding: utf-8 -*-
"""Inject Fade / on-click build animations via OOXML <p:timing> —
Module 4 (2026-08-28). Engine from Module 1 / Module 7 / Italy IBR.

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

Pipeline:  _build_Module4.py -> _splice_media.py -> _group_pass.py
           -> _animate.py all apply
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
DECK = Path(_deckarg[0]) if _deckarg else HERE / "Module 4 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
EMU = 914400.0

# --------------------------------------------------------------------------
#  Module 4 configuration
# --------------------------------------------------------------------------
SKIP_TITLE = {1}                             # the deck title slide
SKIP_AGENDA = {3, 4, 5, 8, 12, 42, 46,       # roadmap + the 8 outline slides
               55, 76}
SKIP_MEDIA = {21, 32, 39, 52, 66, 70}        # the spliced PollEverywhere slides
# single-thought slides: one table, one map, or deliberately blank
SKIP_STATIC = {9, 53, 54, 62, 72}
SKIP = SKIP_TITLE | SKIP_AGENDA | SKIP_MEDIA | SKIP_STATIC

# figures ride on this animated-bullet index by default (0 = first
# built bullet); per-slide overrides:
FIG_GROUP_DEFAULT = 0
FIG_GROUP = {}

PLANS_PRE = {}

STATIC = {}

# --------------------------------------------------------------------------
#  Custom story plans (selector language in the header docstring).
#  Authored per slide from the inventory printed by:
#      python _animate.py inventory <display>...
#  Anything not selected stays visible from the start, which is how chart
#  axes and tick labels stay static.
# --------------------------------------------------------------------------
# NOTE: the numeric t: selectors below (60,000 / 400 / 358 / 210 / 187.5 /
# 662.5) are chart labels printed from the cost constants in
# _build_Module4.py.  Change those constants and these have to change with
# them - the engine raises KeyError on a selector it cannot resolve, so a
# stale one fails the run rather than animating the wrong shape.
PLANS = {
    # --- 1 - introduction ------------------------------------------------
    7: [["t:(Price takers)"],
        ["t:(Price searchers)"],
        ["grp:0"],
        ["grp:1"],
        ["pic:0", "grp:2"],
        ["grp:3"],
        ["osp:6", "t:TODAY"],
        ["t:Today: firms"]],

    # --- 2 - perfect competition -----------------------------------------
    10: [["cxn:4", "t:S"],
         ["cxn:5", "t:D"],
         ["osp:3", "t:Market equilibrium"],
         ["cxn:7", "t:P*"],
         ["cxn:6"],
         ["cxn:8", "t:P*", "t:Demand curve"]],

    # --- 2a - short-run profit maximization -------------------------------
    14: [["grp:0"], ["grp:1"], ["grp:2"], ["grp:3"]],

    15: [["osp:3", "t:TC#2"],
         ["t:TC#1"],
         ["cxn:2", "t:60,000", "t:TFC"]],

    18: [["grp:0"],
         ["cxn:2", "t:TR ="],
         ["cxn:3", "cxn:4", "t:10", "t:30"],
         ["cxn:7", "t:3", "t:d ="],
         ["grp:1"]],

    19: [["cxn:2", "t:TR#1"],
         ["osp:3", "t:TC#1"],
         ["cxn:3", "t:TFC"],
         ["t:TR slope"],
         ["t:TC slope"],
         ["cxn:4", "cxn:5", "osp:4", "osp:5", "t:Q*"],
         ["grp:0"]],

    25: [["cxn:2", "t:MC"],
         ["osp:3", "t:ATC"],
         ["cxn:3", "t:P = MR", "t:400"],
         ["cxn:4", "t:Q*", "osp:5"],
         ["osp:4", "t:Profit", "t:358"]],

    27: [["t:Positive Profit", "cxn:2", "cxn:3", "osp:3", "osp:4", "osp:5",
          "t:MC#1", "t:ATC#1", "t:P#2", "t:Q*#1", "t:Profit,"],
         ["t:Zero Profit", "cxn:6", "cxn:7", "osp:6", "osp:7",
          "t:MC#2", "t:ATC#2", "t:P#4", "t:Q*#2", "t:P = ATC"],
         ["t:Negative Profit", "cxn:10", "cxn:11", "osp:8", "osp:9",
          "osp:10", "t:MC#3", "t:ATC#3", "t:P#6", "t:Q*#3", "t:Loss,"],
         ["t:Profit = ("]],

    30: [["t:Option 1", "t:Profit#1", "t:The fixed"],
         ["t:Option 2", "t:Profit#2", "t:Operating adds"],
         ["t:Difference"],
         ["t:Operate if"]],

    35: [["osp:4", "t:MC"],
         ["osp:5", "t:ATC"],
         ["osp:6", "t:AVC"],
         ["cxn:2", "t:P high"],
         ["cxn:3", "t:Q*"],
         ["osp:3", "t:Profit"],
         ["grp:0"]],

    36: [["osp:4", "t:MC"],
         ["osp:5", "t:ATC"],
         ["osp:6", "t:AVC"],
         ["cxn:2", "t:P low"],
         ["cxn:3", "t:Q*"],
         ["osp:3", "t:Loss if"],
         ["grp:0"]],

    37: [["osp:4", "t:MC"],
         ["osp:5", "t:ATC"],
         ["osp:6", "t:AVC"],
         ["cxn:2", "t:P very"],
         ["cxn:3", "t:Q*"],
         ["osp:3", "t:Loss if"],
         ["grp:0"]],

    # --- 2b - firm and market supply --------------------------------------
    43: [["cxn:2", "t:MC", "cxn:14", "t:AVC"],
         ["cxn:3", "cxn:4", "t:210#1", "t:187.5#1", "osp:3"],
         ["cxn:5", "cxn:6", "t:400#1", "t:662.5#1", "osp:4"],
         ["cxn:9", "t:S"],
         ["cxn:10", "cxn:11", "t:210#2", "t:187.5#2", "osp:5"],
         ["cxn:12", "cxn:13", "t:400#2", "t:662.5#2", "osp:6"],
         ["t:Read the quantity"]],

    44: [["cxn:2", "t:MC0"],
         ["cxn:4", "t:P = MR"],
         ["cxn:5", "t:q0", "osp:3"],
         ["t:High-yield", "cxn:3", "t:MC1"],
         ["cxn:6", "t:q1", "osp:4", "cxn:7"]],

    45: [["cxn:4", "t:S0", "cxn:6", "t:D"],
         ["cxn:7", "cxn:8", "t:Q0", "t:P0", "osp:3"],
         ["cxn:11", "t:MC0", "cxn:13", "t:MR0"],
         ["cxn:15", "t:q0", "osp:5"],
         ["cxn:12", "t:MC1"],
         ["cxn:17", "t:q1", "osp:7"],
         ["cxn:5", "t:S1"],
         ["cxn:9", "cxn:10", "t:Q1", "t:P1", "osp:4"],
         ["cxn:14", "t:MR1"],
         ["cxn:16", "t:q2", "osp:6"],
         ["t:Everyone's costs"]],

    # --- 2c - long run -----------------------------------------------------
    48: [["osp:3", "t:LAC"],
         ["osp:4", "t:LMC"],
         ["cxn:4", "t:MR = P#2", "t:PLR", "cxn:5", "t:QLR", "osp:6"],
         ["cxn:2", "t:MR = P#1", "t:P1", "cxn:3", "t:Q1", "osp:5"],
         ["cxn:8", "t:Entry"],
         ["cxn:6", "t:MR = P#3", "t:P2", "cxn:7", "t:Q2", "osp:7"],
         ["cxn:9", "t:Exit"],
         ["t:At QLR"]],

    # --- 3 - market distortions -------------------------------------------
    57: [["cxn:2", "t:D"],
         ["cxn:3", "t:S"],
         ["cxn:4", "cxn:5", "t:P*", "t:Q*", "osp:4"],
         ["osp:3", "t:CS#1"],
         ["grp:0"]],

    58: [["cxn:2", "t:D"],
         ["cxn:3", "t:S"],
         ["cxn:4", "cxn:5", "t:P*", "t:Q*", "osp:4"],
         ["osp:3", "t:PS#1"],
         ["grp:0"]],

    60: [["cxn:2", "t:S#1"],
         ["cxn:4", "t:D#1"],
         ["cxn:6", "t:P0", "cxn:9", "t:Q0"],
         ["cxn:3", "t:S\u2019"],
         ["cxn:5", "t:PB", "cxn:7", "t:PS", "cxn:8", "t:Q1"],
         ["osp:3", "t:A", "osp:4", "t:C"],
         ["osp:5", "t:B", "osp:6", "t:D#2"],
         ["grp:0"],
         ["t:Buyers lose"],
         ["t:A tax on gasoline"]],

    63: [["cxn:3", "t:SLabor"],
         ["cxn:2", "t:DLabor"],
         ["cxn:4", "cxn:5", "t:w*", "t:L*"],
         ["cxn:6", "t:wmin"],
         ["cxn:7", "t:Ld", "cxn:8", "t:Ls"],
         ["cxn:9", "t:Ls \u2212"],
         ["osp:3", "t:A"],
         ["osp:4", "t:B", "osp:5", "t:C"],
         ["t:Welfare effects"],
         ["grp:0"]],

    68: [["cxn:3", "t:S"],
         ["cxn:2", "t:D"],
         ["cxn:4", "cxn:5", "t:P*", "t:Q*"],
         ["cxn:6", "t:Pmax"],
         ["cxn:7", "t:Qs", "cxn:8", "t:Qd"],
         ["cxn:9", "t:Qd \u2212"],
         ["osp:3", "t:A"],
         ["osp:4", "t:B", "osp:5", "t:C"],
         ["t:Welfare effects"]],

    74: [["grp:0"],
         ["grp:1"],
         ["grp:2", "t:Source:"],
         ["t:Rents rose"]],

    # --- 4 - externalities -------------------------------------------------
    79: [["cxn:2", "t:Demand"],
         ["cxn:3", "t:Supply = internal"],
         ["cxn:5", "cxn:6", "t:QMarket", "osp:3"],
         ["cxn:4", "t:MCI + EMC"],
         ["cxn:9", "t:EMC ="],
         ["cxn:7", "cxn:8", "t:QExt", "osp:4"],
         ["grp:0"]],

    80: [["cxn:2", "t:Demand"],
         ["cxn:3", "t:Supply = internal"],
         ["cxn:5", "cxn:6", "t:QMarket", "osp:3"],
         ["cxn:4", "t:MCI + EMC"],
         ["cxn:9", "t:Tax ="],
         ["cxn:7", "cxn:8", "t:QExt", "osp:4"],
         ["grp:0"],
         ["grp:1"]],

    81: [["grp:0"],
         ["pic:0"]],

    82: [["cxn:2", "t:D (MB"],
         ["cxn:3", "t:Pilots"],
         ["cxn:4", "t:MCI + tax"],
         ["cxn:5", "t:Noise tax"],
         ["osp:3", "t:Q* = 0"],
         ["t:SMC lies"]],
}

PLANS.update({
    11: [["pr:Germany cut:1:1"],
         ["grp:0"],
         ["grp:1"]],

    16: [["t:Key questions"],
         ["t:1", "t:Should the Yi"],
         ["t:2", "t:How much should"],
         ["pr:We start with:0:0"],
         ["pr:We start with:1:1"]],

    26: [["grp:0"],
         ["t:Cost of goods"],
         ["t:The accounting"]],

    38: [["pr:Your firm:1:1"],
         ["grp:0"],
         ["t:Should you continue"]],

    49: [["pr:In the 1920s:1:1"],
         ["grp:0"],
         ["pr:In the 1920s:2:2"],
         ["t:Profit in a market"]],

    59: [["pr:The inefficiency:1:1"],
         ["pr:The inefficiency:2:2"],
         ["t:Deadweight loss", "t:with welfare"]],

    61: [["grp:0"],
         ["pr:Who hands:1:1"],
         ["pr:Who hands:2:2"],
         ["pr:Who hands:3:3"],
         ["t:The less price"]],

    64: [["pr:Theory predicts:1:6"],
         ["pic:0"],
         ["pic:1"],
         ["pic:2", "t:One way firms"],
         ["pr:Theory predicts:7:7"],
         ["grp:0"]],

    65: [["pic:0"],
         ["pr:1 bed:0:4"]],

    67: [["pic:0"],
         ["grp:0"],
         ["pic:1"]],

    71: [["pr:Non-price:1:1"],
         ["pr:Non-price:2:2"],
         ["grp:0"],
         ["grp:1"]],

    75: [["pr:Argentina repealed:1:1"],
         ["pr:Argentina repealed:2:2"]],

    77: [["pr:Negative externality:1:2"],
         ["pr:Negative externality:3:4"]],

    78: [["t:Price mechanisms", "t:Tax the activity"],
         ["t:Quantity mechanisms", "pr:Negative externality:0:0"],
         ["pr:Negative externality:1:1"]],
})

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
    txt = " ".join(txt.split())      # normalize run-boundary whitespace
    # runs are joined with a space above, which splits symbols the
    # subscript pass broke into "P" + "0". Keep a concatenated variant so
    # t:/pr: selectors can still name them (2026-08-23).
    tight = "".join(t.text or "" for t in el.iter(q(A, "t")))
    tight = (tight + "".join(t.text or "" for t in el.iter(q(M, "t"))))
    tight = " ".join(tight.split())
    # a curve is a custGeom outline with no fill; a chart backing card is
    # a preset rect with a solid fill.  is_chrome needs to tell them apart.
    spPr = el.find(q(P, "spPr"))
    custgeom = spPr is not None and spPr.find(q(A, "custGeom")) is not None
    solidfill = spPr is not None and spPr.find(q(A, "solidFill")) is not None
    info = {
        "custgeom": custgeom, "solidfill": solidfill,
        "id": int(cnv.get("id")), "name": cnv.get("name") or "",
        "tag": tag, "x": x, "y": y, "w": w, "h": h, "text": txt,
        "tight": tight,
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
                           "w": 0, "h": 0, "text": "", "tight": "",
                           "paras": [], "bul0": 0,
                           "custgeom": False, "solidfill": False,
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
    if (s["tag"] == "sp" and not t and s["w"] > 4 and s["h"] > 2.5
            and s.get("solidfill") and not s.get("custgeom")):
        return True                      # big white chart backings
    return False


def resolve(sel, shapes, used):
    """Resolve one selector to (spid, prg) or None.

    Text selectors search the CONTENT shapes only: a slide title or a
    top-bar tag must never be animated, and it would otherwise win a
    prefix match ahead of the label the plan meant (Sales Tax / "S").
    """
    if sel.startswith(("t:", "pr:")):
        shapes = [s for s in shapes if not is_chrome(s)]
    if sel.startswith("pr:"):
        rest = sel.split(":", 1)[1]
        prefix, st, end = rest.rsplit(":", 2)
        for s in shapes:
            if (s["text"].startswith(prefix)
                    or s.get("tight", "").startswith(prefix)):
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
                   if (s["text"].strip().startswith(body)
                       or s.get("tight", "").strip().startswith(body))]
        if nth is not None:
            if len(matches) < nth:
                raise KeyError(sel)
            return (matches[nth - 1]["id"], None), None
        for s in matches:
            if s["id"] not in used:
                used.add(s["id"])
                return (s["id"], None), None
        raise KeyError(sel)
    if sel.startswith("n:"):
        # exact shape-name selector (names emitted by _sd_chart and the
        # grouping pass) — immune to index shifts from grouping
        want = sel[2:]
        for s in shapes:
            if s["name"] == want and s["id"] not in used:
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


def inventory(order, data, disps):
    """Print the animatable shapes of each slide, in document order, with
    the selector that reaches them.  This is what the PLANS below are
    authored from; chrome is marked so it can be ignored."""
    for disp in disps:
        part = "ppt/slides/" + order[disp - 1]
        tree = ET.fromstring(data[part])
        spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
        print("===== slide %d =====" % disp)
        for s_ in collect_shapes(spTree):
            txt = (s_["text"] or "").strip().replace("\n", " / ")[:52]
            print("  %-9s %-16s [%5.2f,%5.2f %4.2fx%4.2f]%s %s"
                  % (s_["idx"], s_["tag"] + ":" + s_["name"][:14],
                     s_["x"], s_["y"], s_["w"], s_["h"],
                     "  CHROME" if is_chrome(s_) else "        ", txt))


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

    if sel and sel[0] == "inventory":
        rest = [int(x) for x in sel[1:]]
        inventory(order, data, rest or range(1, len(order) + 1))
        return

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
