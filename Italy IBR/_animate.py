# -*- coding: utf-8 -*-
"""
Inject Fade / on-click build animations into the deck via OOXML <p:timing>.
Rules (Teaching CLAUDE.md):
  - Fade entrance, 0.5 s, on click. Chrome never animates.
  - Text slides: first top-level bullet (with its subs) shows with the slide;
    build from the second, one top-level bullet (+ its subs) per click.
  - Side-figure content slides: the figure + caption stay visible from the start
    (context), and the bullets build.
  - Gallery slides (pictures are the content): open empty, reveal each picture
    (with its nearest caption) one per click, in reading order.
  - Two-column cards (102): reveal each column as one beat, left then right.
  - Featured-research slides: reveal the paper title, then the authors.
  - Takeaway bar: its own final click.
Skips: title (1,2), agenda/dividers, poll slides, backups, blanks.

Verify via PowerPoint COM (effect counts/targets) + eyeball. Run with a slide
list (e.g. "16 20 44") or "all", plus "apply" to write.
"""
import os
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0

SKIP_LOGISTICS = set()   # logistics slides 3-9 now animate (build after each main bullet)
SKIP_AGENDA = {10, 12, 18, 30, 32, 41, 46, 56, 65, 68, 92}
SKIP_POLL = {11, 13, 21, 35, 48, 69, 137}
SKIP_TITLE = {1, 2}
SKIP_BACKUP = set(range(109, 140))
SKIP_FEATURED = {36, 51, 78, 88, 96}  # featured-research dividers: no animations (2026-07-23)
SKIP_CUSTOM = {54, 59, 61, 72, 106}  # 105 = GBU ten-years-on (question static, changes on clicks)  # hand-built timing: fig+label+legend co-reveal (2026-07-24); rebuild wipes it
SKIP = SKIP_LOGISTICS | SKIP_AGENDA | SKIP_POLL | SKIP_TITLE | SKIP_BACKUP | SKIP_FEATURED | SKIP_CUSTOM

# For bullet slides with a side figure: which ANIMATED top-level bullet the
# figure fades in with (0-based index among the animated bullet groups; -1 = the
# last one). Default is 0 (fades in with the first built bullet, which introduces
# the topic); listed here are only the overrides where the figure matches a later
# bullet.
FIG_GROUP_DEFAULT = 0
FIG_GROUP = {
    16: 3,    # Sarcophagus -> "Advanced art and written language"
    90: 1,    # Fiat 500 -> "Icons: the Vespa, the Fiat 500 ..."
    47: -1,   # Naples-plague painting -> "The 1630 plague devastated the northern cities"
}
# When each picture pairs with its own bullet: map pictures (sorted top-to-bottom)
# to the ANIMATED bullet index they fade in with. Overrides FIG_GROUP for the slide.
PIC_BULLET = {
    6: [0, 1],   # top photo -> "Hotel in Milan"; second photo -> "Hotel in Turin"
}


def q(ns, t):
    return f"{{{ns}}}{t}"


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
        return f'<p:spTgt spid="{spid}"/>'
    return (f'<p:spTgt spid="{spid}"><p:txEl><p:pRg st="{prg[0]}" end="{prg[1]}"/>'
            f'</p:txEl></p:spTgt>')


def effect_par(spid, prg, node_type, ids):
    a, b, c = ids(), ids(), ids()
    tgt = sptgt(spid, prg)
    return (
        f'<p:par><p:cTn id="{a}" presetID="10" presetClass="entr" presetSubtype="0" '
        f'fill="hold" grpId="0" nodeType="{node_type}"><p:stCondLst><p:cond delay="0"/>'
        f'</p:stCondLst><p:childTnLst>'
        f'<p:set><p:cBhvr><p:cTn id="{b}" dur="1" fill="hold"><p:stCondLst>'
        f'<p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl>{tgt}</p:tgtEl>'
        f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
        f'<p:to><p:strVal val="visible"/></p:to></p:set>'
        f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{c}" dur="500"/>'
        f'<p:tgtEl>{tgt}</p:tgtEl></p:cBhvr></p:animEffect>'
        f'</p:childTnLst></p:cTn></p:par>')


def click_group(beat, ids):
    outer, inner = ids(), ids()
    effs = "".join(effect_par(spid, prg, "clickEffect" if i == 0 else "withEffect", ids)
                   for i, (spid, prg) in enumerate(beat))
    return (
        f'<p:par><p:cTn id="{outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
        f'</p:stCondLst><p:childTnLst>'
        f'<p:par><p:cTn id="{inner}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f'<p:childTnLst>{effs}</p:childTnLst></p:cTn></p:par>'
        f'</p:childTnLst></p:cTn></p:par>')


def timing_xml(beats, para_boxes):
    ids = Counter(3)
    groups = "".join(click_group(b, ids) for b in beats)
    bld = "".join(f'<p:bldP spid="{spid}" grpId="0" build="p"/>' for spid in para_boxes)
    bldlst = f'<p:bldLst>{bld}</p:bldLst>' if bld else ""
    return (
        f'<p:timing xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"><p:tnLst><p:par>'
        f'<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        f'<p:childTnLst>{groups}</p:childTnLst></p:cTn>'
        f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        f'</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>{bldlst}</p:timing>')


def is_chrome(sh):
    n = sh["name"]
    if n in ("TextBox 2", "TextBox 3"):
        return True
    if n.startswith("Rectangle "):          # plain bars / rules / gold strips
        return True
    if sh["x"] > 11.5 and sh["tag"] == "sp":  # page-number box
        return True
    t = sh["text"].strip()
    if t.startswith("International Business Residential") or t.startswith("Management 405"):
        return True
    return False


def shape_info(el):
    cnv = el.find(".//" + q(P, "cNvPr"))
    xf = el.find(".//" + q(A, "xfrm"))
    x = y = w = h = 0.0
    if xf is not None:
        o = xf.find(q(A, "off")); e = xf.find(q(A, "ext"))
        if o is not None and e is not None:
            x, y = int(o.get("x")) / EMU, int(o.get("y")) / EMU
            w, h = int(e.get("cx")) / EMU, int(e.get("cy")) / EMU
    paras = []
    idx = 0
    for pp in el.iter(q(A, "p")):
        lvl = 0
        pPr = pp.find(q(A, "pPr"))
        if pPr is not None and pPr.get("lvl"):
            lvl = int(pPr.get("lvl"))
        txt = "".join(t.text or "" for t in pp.iter(q(A, "t"))).strip()
        paras.append((idx, lvl, bool(txt)))
        idx += 1
    bul0 = sum(1 for (_, l, ne) in paras if l == 0 and ne)
    return {
        "id": int(cnv.get("id")), "name": cnv.get("name") or "", "tag": ET.QName(el).localname,
        "x": x, "y": y, "w": w, "h": h,
        "text": " ".join(t.text or "" for t in el.iter(q(A, "t"))),
        "paras": paras, "bul0": bul0,
    }


def para_beats(box):
    """Group paragraphs into top-level beats; return (static_indices, [beat_index_lists])."""
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
    # keep only non-empty indices in each group
    clean = [[i for (i, ne) in g if ne] for g in groups]
    clean = [g for g in clean if g]
    if not clean:
        return [], []
    return clean[0], clean[1:]


def nearest_pic(cap, pics):
    cx, cy = cap["x"] + cap["w"] / 2, cap["y"] + cap["h"] / 2
    best, bd = None, 1e9
    for p in pics:
        px, py = p["x"] + p["w"] / 2, p["y"] + p["h"] / 2
        d = (px - cx) ** 2 + (py - cy) ** 2
        if d < bd:
            best, bd = p["id"], d
    return best


def plan(shapes, disp=None):
    """Return (beats, para_box_ids). beats = list of [(spid, prg_or_None), ...]."""
    content = [s for s in shapes if not is_chrome(s)]
    if not content:
        return [], []
    title = next((s for s in shapes if s["name"] == "TextBox 3"), None)
    title_txt = (title["text"].strip() if title else "")

    boxes = [s for s in content if s["tag"] == "sp" and s["bul0"] >= 2]
    bullets = max(boxes, key=lambda s: s["bul0"]) if boxes else None
    # takeaway = the grouped bottom bar (fallback: an ungrouped rounded rect at bottom)
    takeaway_ids = [s["id"] for s in content if s["name"] == "TakeawayGroup"]
    if not takeaway_ids:
        takeaway_ids = [s["id"] for s in content if s["tag"] == "sp"
                        and s["name"].startswith("Rounded Rectangle") and s["y"] > 5.8]
    # cards = grouped body columns (fallback: ungrouped rounded rects)
    cards = [s for s in content if s["name"] == "CardGroup"]
    if not cards:
        cards = [s for s in content if s["tag"] == "sp"
                 and s["name"].startswith("Rounded Rectangle") and s["y"] < 5.8 and 4.5 < s["w"] < 7.5]
    # figures / galleries exclude the takeaway & card groups (handled explicitly)
    _handled = set(takeaway_ids) | {c["id"] for c in cards}
    pics = [s for s in content if s["tag"] in ("pic", "grpSp") and s["id"] not in _handled]

    beats = []
    para_ids = []

    # ---- two-column cards (e.g. slide 102): each card is now one group ----
    if len(cards) >= 2:
        for card in sorted(cards, key=lambda s: s["x"]):
            beats.append([(card["id"], None)])
        if takeaway_ids:
            beats.append([(t, None) for t in takeaway_ids])
        return beats, para_ids

    # ---- featured-research slides ----
    if title_txt == "Featured research":
        tbs = [s for s in content if s["tag"] == "sp" and s["text"].strip()]
        for s in sorted(tbs, key=lambda s: s["y"]):
            beats.append([(s["id"], None)])
        return beats, para_ids

    # ---- bullet-driven content slides ----
    if bullets is not None:
        _static, groups = para_beats(bullets)
        bullet_beats = [[(bullets["id"], (i, i)) for i in g] for g in groups]
        para_ids.append(bullets["id"])
        # figure(s) + caption/link fade in ON THE CLICK of the bullet they illustrate
        fig = [s for s in content
               if s["id"] != bullets["id"] and s["id"] not in takeaway_ids]
        fig.sort(key=lambda s: (0 if s["tag"] in ("pic", "grpSp") else 1, s["y"], s["x"]))
        fig_targets = [(s["id"], None) for s in fig]
        if fig_targets:
            if not bullet_beats:
                bullet_beats.append(fig_targets)
            elif disp in PIC_BULLET:
                # each picture (top-to-bottom) pairs with its own bullet
                def clamp(i):
                    return max(0, min(i, len(bullet_beats) - 1))
                idxs = PIC_BULLET[disp]
                pics_f = [s for s in fig if s["tag"] in ("pic", "grpSp")]
                for i, pic in enumerate(pics_f):
                    bullet_beats[clamp(idxs[i] if i < len(idxs) else idxs[-1])].append((pic["id"], None))
                # any captions / other bits ride with the last mapped bullet
                for s in fig:
                    if s["tag"] not in ("pic", "grpSp"):
                        bullet_beats[clamp(idxs[-1])].append((s["id"], None))
            else:
                gi = FIG_GROUP.get(disp, FIG_GROUP_DEFAULT)
                if gi < 0:
                    gi = len(bullet_beats) + gi
                gi = max(0, min(gi, len(bullet_beats) - 1))
                bullet_beats[gi] += fig_targets
        beats.extend(bullet_beats)
        if takeaway_ids:
            beats.append([(t, None) for t in takeaway_ids])
        return beats, para_ids

    # ---- gallery / picture-driven slides (open empty, reveal each) ----
    if pics:
        caps = [s for s in content if s["tag"] == "sp" and s["text"].strip()
                and s["id"] not in takeaway_ids]
        pic_caps = {p["id"]: [] for p in pics}
        for c in caps:
            pid = nearest_pic(c, pics)
            if pid is not None:
                pic_caps[pid].append(c["id"])
        for p in sorted(pics, key=lambda s: (round(s["y"] / 1.4), s["x"])):
            members = [(p["id"], None)] + [(cid, None) for cid in pic_caps[p["id"]]]
            beats.append(members)
        if takeaway_ids:
            beats.append([(t, None) for t in takeaway_ids])
        return beats, para_ids

    # ---- fallback: reveal remaining content text boxes top-to-bottom ----
    for s in sorted([s for s in content if s["text"].strip()], key=lambda s: s["y"]):
        beats.append([(s["id"], None)])
    return beats, para_ids


def main():
    args = [a for a in sys.argv[1:]]
    apply = "apply" in args
    sel = [a for a in args if a != "apply"]
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target")
             for r in ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    order = [os.path.basename(rid2t[s.get(q(R, "id"))]) for s in pres.find(q(P, "sldIdLst"))]

    if sel == ["all"] or not sel:
        todo = [d for d in range(1, len(order) + 1) if d not in SKIP]
    else:
        todo = [int(x) for x in sel]

    report = []
    for disp in todo:
        part = order[disp - 1]
        root = ET.fromstring(data[f"ppt/slides/{part}"])
        tree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
        shapes = [shape_info(el) for el in tree
                  if ET.QName(el).localname in ("sp", "pic", "graphicFrame", "grpSp")]
        beats, para_ids = plan(shapes, disp)
        if not beats:
            report.append((disp, 0))
            continue
        # remove any existing timing, then append fresh
        for old in root.findall(q(P, "timing")):
            root.remove(old)
        root.append(ET.fromstring(timing_xml(beats, para_ids)))
        data[f"ppt/slides/{part}"] = ser(root)
        report.append((disp, len(beats)))

    for disp, n in report:
        print(f"slide {disp:3}: {n} click-beats")
    total = sum(n for _, n in report)
    print(f"== {len(report)} slides, {total} click-beats ==")

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
