# -*- coding: utf-8 -*-
"""Audit a deck's BUILD ORDER against the Teaching CLAUDE.md animation
rules.  The click counts _animate.py prints say how many beats a slide has;
they say nothing about whether the beats are the right ones, and the
generic auto-rollout reveals one body shape per click with no idea which
shapes belong to one thought.  This finds the cases the rules name:

  PANEL    a picture and the caption / header that belongs to it land on
           DIFFERENT clicks ("a whole labelled panel reveals as one beat")
  BAR      the gold takeaway bar is not the LAST click ("the takeaway bar
           gets its own final click so the punchline lands last")
  FIRST    the first top-level bullet of a bullet slide is animated ("the
           FIRST top-level bullet is visible the moment the slide appears")
  CHROME   a corner mark, poll badge or footer element is animated
  AXIS     an axis line or axis title is animated (it is chrome)
  NEWS     a newspaper clipping on an "In the News" slide is animated
  CURVE    a curve and the label naming it land on different clicks, or a
           DERIVED curve (MR) is revealed before the primary curve it is
           derived from (D) or before the other primary line (MC, S)

Read-only.

    python _check_anim.py                      # the canonical deck
    python _check_anim.py "<deck>.pptx"
"""
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0

HERE = Path(__file__).parent
MODULE = HERE.parent  # _build_files/ since 2026-09-24
DECK = Path(sys.argv[1]) if len(sys.argv) > 1 else MODULE / "Module 6 - Full.pptx"

GOLD = "E09F3E"


def q(ns, t):
    return "{%s}%s" % (ns, t)


def shapes_of(tree, out, parent=None):
    """Flatten the spTree to (id, kind, x, y, w, h, text), decoding groups.

    Each shape carries the id of the group that CONTAINS it: a grouped
    curve is animated through its group, so a member-level check has to be
    able to find the beat its group was given.
    """
    for el in tree:
        tag = el.tag.split('}')[-1]
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        cid = None
        for path in ("nvSpPr", "nvPicPr", "nvGrpSpPr", "nvCxnSpPr",
                     "nvGraphicFramePr"):
            c = el.find("./%s/%s" % (q(P, path), q(P, "cNvPr")))
            if c is not None:
                cid = c
                break
        xf = el.find(".//%s" % q(A, "xfrm"))
        geo = None
        if xf is not None:
            o, e = xf.find(q(A, "off")), xf.find(q(A, "ext"))
            if o is not None and e is not None:
                geo = (int(o.get("x")) / EMU, int(o.get("y")) / EMU,
                       int(e.get("cx")) / EMU, int(e.get("cy")) / EMU)
        txt = "".join(t.text or "" for t in el.iter(q(A, "t"))).strip()
        fill = el.find(".//%s/%s" % (q(P, "spPr"), q(A, "solidFill")))
        hexv = ""
        if fill is not None:
            c2 = fill.find(q(A, "srgbClr"))
            if c2 is not None:
                hexv = c2.get("val", "")
        if cid is not None and geo is not None:
            out.append({"id": cid.get("id"), "name": cid.get("name") or "",
                        "kind": tag, "x": geo[0], "y": geo[1],
                        "w": geo[2], "h": geo[3], "text": txt, "fill": hexv,
                        "parent": parent})
        if tag == "grpSp":
            # index members too, pointing each at the group that animates it
            shapes_of(el, out, cid.get("id") if cid is not None else parent)
    return out


def clicks_of(root):
    """[[spid, ...], ...] -- one list per on-click beat, in order."""
    tm = root.find(q(P, "timing"))
    if tm is None:
        return []
    seq = next(iter(tm.iter(q(P, "seq"))), None)
    if seq is None:
        return []
    beats, cur = [], None
    for par in seq.iter(q(P, "par")):
        ct = par.find("./%s" % q(P, "cTn"))
        if ct is None:
            continue
        kind = ct.get("nodeType")
        if kind not in ("clickEffect", "withEffect", "afterEffect"):
            continue
        tgt = par.find(".//%s" % q(P, "spTgt"))
        if tgt is None:
            continue
        prg = tgt.find(".//%s/%s" % (q(P, "txEl"), q(P, "pRg")))
        st = int(prg.get("st")) if prg is not None else None
        if kind == "clickEffect":
            cur = [(tgt.get("spid"), st)]
            beats.append(cur)
        elif cur is not None:
            cur.append((tgt.get("spid"), st))
    return beats


def main():
    z = zipfile.ZipFile(str(DECK))
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = {r.get("Id"): r.get("Target") for r in
            ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    order = ["ppt/" + rels[s.get(q(R, "id"))].replace("../", "")
             for s in pres.find(q(P, "sldIdLst"))]

    findings = {"PANEL": [], "BAR": [], "FIRST": [], "CHROME": [],
                "CURVE": [], "NEWS": [], "AXIS": []}
    for n, part in enumerate(order, 1):
        root = ET.fromstring(z.read(part))
        beats = clicks_of(root)
        if not beats:
            continue
        tree = root.find(q(P, "cSld")).find(q(P, "spTree"))
        shp = shapes_of(tree, [])
        by = {s["id"]: s for s in shp}
        # which beat is each shape on?  0-based; None = static
        beat_of = {}
        for i, b in enumerate(beats):
            for sp, _st in b:
                beat_of[sp] = i
        # a grouped shape is revealed on its GROUP's click, so give every
        # member the beat of the nearest ancestor that has one.  Without
        # this every curve-level check silently passes: the grouping pass
        # pairs each curve with its label, so the only id the timing ever
        # names is the pair group's.
        for sh in shp:
            if sh["id"] in beat_of:
                continue
            anc = sh.get("parent")
            while anc is not None and anc not in beat_of:
                anc = (by.get(anc) or {}).get("parent")
            if anc is not None:
                beat_of[sh["id"]] = beat_of[anc]

        # --- CHROME -------------------------------------------------
        for sp in beat_of:
            s = by.get(sp)
            if s is None:
                continue
            if (s["h"] >= 0.25 and s["y"] < 7.15 < s["y"] + s["h"]
                    and s["kind"] != "pic" and s["w"] <= 6.5):
                findings["CHROME"].append((n, s["name"], s["text"][:30]))

        # --- PANEL: a picture and its caption on different clicks ----
        pics = [s for s in shp if s["kind"] == "pic" and s["id"] in beat_of]
        texts = [s for s in shp if s["kind"] == "sp" and s["text"]
                 and s["id"] in beat_of]
        for p in pics:
            for t in texts:
                # a caption sits directly under (or a header directly over)
                # the picture and is horizontally centred on it
                cx_p = p["x"] + p["w"] / 2.0
                cx_t = t["x"] + t["w"] / 2.0
                # A caption is about as wide as the picture it belongs to.
                # A slide-wide takeaway bar under the bottom picture, or a
                # header that spans the whole content area above the first
                # of several pictures, is neither -- and both were being
                # reported as broken panels (s38, s62).
                if t["fill"].upper() == GOLD:
                    continue
                if t["w"] > p["w"] * 1.6:
                    continue
                # a caption or header is a LINE or two, not a block: slide
                # 80's purchase-options list sits 0.33" above the ice-cream
                # photo and is the slide's data, not the photo's label
                if t["h"] > 1.0:
                    continue
                if abs(cx_p - cx_t) > max(p["w"], t["w"]) * 0.40:
                    continue
                gap_below = t["y"] - (p["y"] + p["h"])
                gap_above = p["y"] - (t["y"] + t["h"])
                if not (-0.05 <= gap_below <= 0.45 or
                        -0.05 <= gap_above <= 0.45):
                    continue
                if beat_of[p["id"]] != beat_of[t["id"]]:
                    findings["PANEL"].append(
                        (n, "pic@%.1f,%.1f" % (p["x"], p["y"]),
                         t["text"][:34], beat_of[p["id"]] + 1,
                         beat_of[t["id"]] + 1))

        # --- AXIS: an axis and its title are chrome ------------------
        for sp in beat_of:
            sh = by.get(sp)
            if sh and sh["name"].startswith(("sdaxis:", "sdaxistitle:")):
                findings["AXIS"].append(
                    (n, "%s on click %d" % (sh["name"], beat_of[sp] + 1)))

        # --- NEWS: a clipping at the top of the slide is chrome ------
        # Scoped by TITLE, deliberately: "a wide picture near the top" also
        # describes a product screenshot the bullets go on to explain, and
        # those do animate.  The slide that must not is the one whose
        # clipping the class reads while the question is put to them.
        title = ""
        for sh in shp:
            if 0.45 < sh["y"] < 0.95 and sh["text"]:
                title = sh["text"]
        if title.strip().lower().startswith("in the news"):
            for sh in shp:
                if (sh["kind"] == "pic" and sh["id"] in beat_of
                        and sh["y"] < 1.95 and sh["w"] > 7.5):
                    findings["NEWS"].append(
                        (n, "clipping animated on click %d"
                         % (beat_of[sh["id"]] + 1)))

        # --- CURVE: order, and curve-with-its-own-label --------------
        # The build helpers stamp a line "sdcurve:<key>" and its label
        # "sdcurvelab:<key>", so the pair is known by NAME and does not
        # have to be guessed from proximity.  Two rules (2026-09-15, Nico):
        # a curve and its label are one beat, and the DERIVED curve comes
        # after the primary ones -- MR is read off D, and a dashed guide at
        # MR = 0 cannot precede the MR it marks.
        curve_beat = {}
        for sp, b in beat_of.items():
            nm = (by.get(sp) or {}).get("name", "")
            for pre in ("sdcurve:", "sdcurvelab:"):
                if nm.startswith(pre):
                    key = nm[len(pre):].split(":")[0]
                    curve_beat.setdefault(key, {})[pre] = b
        for key, seen in curve_beat.items():
            a = seen.get("sdcurve:")
            t = seen.get("sdcurvelab:")
            if a is not None and t is not None and a != t:
                findings["CURVE"].append(
                    (n, "%s: line on click %d, label on click %d"
                     % (key, a + 1, t + 1)))
        def _first(key):
            v = curve_beat.get(key, {})
            got = [x for x in (v.get("sdcurve:"), v.get("sdcurvelab:"))
                   if x is not None]
            return min(got) if got else None
        d_b, mr_b, mc_b = _first("D"), _first("MR"), _first("MC")
        if mr_b is not None and d_b is not None and mr_b < d_b:
            findings["CURVE"].append(
                (n, "MR (derived) on click %d, before D on click %d"
                 % (mr_b + 1, d_b + 1)))
        if mr_b is not None and mc_b is not None and mr_b < mc_b:
            findings["CURVE"].append(
                (n, "MR (derived) on click %d, before MC on click %d"
                 % (mr_b + 1, mc_b + 1)))

        # --- BAR: the gold takeaway bar must be the last beat --------
        for s in shp:
            if (s["fill"].upper() == GOLD and s["text"] and s["w"] > 4.0
                    and s["h"] < 1.0 and s["y"] > 5.0
                    and not (s["y"] < 7.15 < s["y"] + s["h"])):
                b = beat_of.get(s["id"])
                if b is None:
                    findings["BAR"].append((n, s["text"][:40], "static"))
                elif b != len(beats) - 1:
                    findings["BAR"].append(
                        (n, s["text"][:40], "click %d of %d" % (b + 1,
                                                                len(beats))))

        # --- FIRST: the opening bullet of a bullet slide -------------
        # a bullet slide = one big text box with >= 3 paragraphs and no
        # picture / chart of its own
        if not pics:
            bigs = [s for s in shp if s["kind"] == "sp" and s["w"] > 7.0
                    and s["h"] > 2.0 and s["text"]]
            # The rule is about the first PARAGRAPH, not the first shape:
            # _animate.py animates a bullet box by paragraph range, and a
            # plan that starts at pRg 1 already leaves the opening bullet
            # on screen.  Testing the shape flagged four compliant slides
            # (2, 41, 59, 84), which is exactly the crying-wolf an audit
            # must not do.
            if len(bigs) == 1 and beats:
                starts = [st for sp, st in beats[0]
                          if sp == bigs[0]["id"] and st is not None]
                if starts and min(starts) == 0:
                    findings["FIRST"].append((n, bigs[0]["text"][:46]))

    for k in ("CHROME", "AXIS", "NEWS", "CURVE", "PANEL", "BAR", "FIRST"):
        rows = findings[k]
        print("\n== %s : %d ==" % (k, len(rows)))
        for r in rows:
            print("   s%-4d %s" % (r[0], "  ".join(str(x) for x in r[1:])))
    print("\ntotal findings: %d" % sum(len(v) for v in findings.values()))


if __name__ == "__main__":
    main()
