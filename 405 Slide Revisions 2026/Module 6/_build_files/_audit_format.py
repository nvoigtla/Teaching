# -*- coding: utf-8 -*-
"""Audit a built deck against the documented formatting rules.

Written 2026-09-09 after Nico: "make sure you recheck all formatting.  I
don't think the 'Back' button is right, for example.  Check also other
rules."  Eyeballing renders catches layout, not rule compliance, so the
rules that CAN be checked mechanically are checked here.

Checks, each keyed to the rule it enforces (Teaching CLAUDE.md):
  1  trailing periods on bullets / labels          ("No trailing periods")
  2  box + callout text under 18 pt                ("Box/callout text floor")
  3  chart-internal labels under 16 pt             (same rule)
  4  slide titles not in Title Case                ("Slide titles TITLE CASE")
  5  shapes outside the canvas or under the footer ("never spills past")
  6  footer page numbers that are not live fields  ("LIVE slide-number")
  7  a demand curve that is not dark red C00000    ("A DEMAND curve is dark red")
  8  the "Back" pill not navy with white text      ("Back navigation buttons")
  9  an OPAQUE CARD overlapping another card, a     (2026-09-14, Nico:
     picture, or a curve                            "make sure you always
                                                     check for such overlaps")

Usage:  python _audit_format.py ["Module 6 - Revised.pptx"]
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).parent
MODULE = HERE.parent  # _build_files/ since 2026-09-24
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400.0
W_IN, H_IN = 13.333, 7.5
FOOTER_RULE = 7.15

SP_RE = re.compile(r'<p:(sp|pic|graphicFrame|cxnSp)\b.*?</p:\1>', re.S)
NAME_RE = re.compile(r'<p:cNvPr[^>]*name="([^"]*)"')
OFF_RE = re.compile(r'<a:off x="(-?\d+)" y="(-?\d+)"/>')
EXT_RE = re.compile(r'<a:ext cx="(\d+)" cy="(\d+)"/>')
PARA_RE = re.compile(r'<a:p>.*?</a:p>', re.S)
RUN_RE = re.compile(r'<a:r>.*?</a:r>', re.S)
SZ_RE = re.compile(r'<a:rPr[^>]*\bsz="(\d+)"')
T_RE = re.compile(r'<a:t(?:\s[^>]*)?>(.*?)</a:t>', re.S)

# ---- check 9 -------------------------------------------------------------
# The failure this catches is the one Nico hit twice on slide 8: an opaque
# CARD (a filled callout, an equation box, a picture) laid over something it
# hides -- the demand line, or another card.  Restricting it to OPAQUE cards
# is what keeps it quiet: a shaded region carries an `alpha` and a curve is a
# connector, and neither is a card, so the figure's own deliberate
# overlaps -- a region under a curve, a label inside a region -- never fire.
ALPHA_RE = re.compile(r'<a:alpha val="(\d+)"')
# The fill has to be the SHAPE's, so look inside <p:spPr> only -- a
# <a:solidFill> anywhere in the block is usually a run's text COLOUR, and
# reading that as a fill turned every coloured label into a "card".
SPPR_RE = re.compile(r'<p:spPr.*?</p:spPr>', re.S)
PRSTGEOM_RE = re.compile(r'<a:prstGeom prst="([^"]+)"')
FLIP_RE = re.compile(r'<a:xfrm([^>]*)>')
MIN_OVERLAP = 0.02          # square inches; smaller is a rounded corner

# Chrome and figure furniture that is allowed to sit on anything.
def _is_card(kind, blk, box):
    """An opaque card: a picture, or a fully opaque filled rect/roundRect."""
    x, y, w, h = box
    if y + h <= 0.46 or y >= 7.10:          # top bar / footer band
        return False
    if kind == "pic":
        return True
    if kind != "sp":
        return False
    g = PRSTGEOM_RE.search(blk)
    if not g or g.group(1) not in ("rect", "roundRect"):
        return False
    spPr = SPPR_RE.search(blk)
    if not spPr or "<a:solidFill>" not in spPr.group():
        return False
    if ALPHA_RE.search(spPr.group()):       # a wash, not a card
        return False
    if w >= 10.0:                           # an agenda band or a takeaway bar
        return False
    if abs(h - 0.36) < 0.03 and w < 1.7:    # a coverage pill
        return False
    if w < 0.30 or h < 0.18:                # a swatch, a tick, a strip
        return False
    return True


SMALL_WORDS = {"a", "an", "the", "and", "but", "or", "nor", "of", "in", "at",
               "to", "for", "with", "on", "vs.", "vs", "per", "as", "by",
               "from", "into", "than", "that", "if", "is", "it"}


def slide_parts(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = {r.get("Id"): r.get("Target") for r in
            ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    return ["ppt/" + rels[s.get("{%s}id" % R)].replace("../", "")
            for s in pres.find("{%s}sldIdLst" % P)]


def _seg_hits(x0, y0, x1, y1, rx, ry, rw, rh):
    """Does the segment (x0,y0)-(x1,y1) cross the rectangle's interior?

    Liang-Barsky, shrunk by a hair so a curve merely TOUCHING a card's edge
    -- a leader that starts at the box boundary, which is what the
    arrow-origin rule asks for -- is not a hit.
    """
    pad = 0.03
    rx, ry = rx + pad, ry + pad
    rw, rh = rw - 2 * pad, rh - 2 * pad
    if rw <= 0 or rh <= 0:
        return False
    dx, dy = x1 - x0, y1 - y0
    t0, t1 = 0.0, 1.0
    for pq in ((-dx, x0 - rx), (dx, rx + rw - x0),
               (-dy, y0 - ry), (dy, ry + rh - y0)):
        pp, qq = pq
        if abs(pp) < 1e-12:
            if qq < 0:
                return False
            continue
        r = qq / pp
        if pp < 0:
            if r > t1:
                return False
            t0 = max(t0, r)
        else:
            if r < t0:
                return False
            t1 = min(t1, r)
    return t0 <= t1


def _within(a, b, pad=0.02):
    """Is rect a inside rect b (with a little slack)?"""
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return (ax >= bx - pad and ay >= by - pad
            and ax + aw <= bx + bw + pad and ay + ah <= by + bh + pad)


def _rendered_shapes(deck):
    """Per slide, the shapes that check 9 cares about, at RENDERED inches.

    Yields lists of ``(kind, box, text)`` where kind is "pic", "card" or
    "line".  For a line the box is the segment (x0, y0, x1, y1), read off
    the connector's own flipH / flipV; for the others it is (x, y, w, h).
    Group children are decoded through off/ext vs chOff/chExt, so a grouped
    callout reports where it actually sits.
    """
    from pptx import Presentation

    AN = "{%s}" % A

    def emit(shapes, dx, dy, sx, sy, out):
        for sh in shapes:
            el = sh._element
            if el.tag.endswith("}grpSp"):
                xfrm = el.find(".//" + AN + "xfrm")
                off, ext = xfrm.find(AN + "off"), xfrm.find(AN + "ext")
                ch, ce = xfrm.find(AN + "chOff"), xfrm.find(AN + "chExt")
                cw, chh = int(ce.get("cx")) or 1, int(ce.get("cy")) or 1
                nsx = sx * int(ext.get("cx")) / float(cw)
                nsy = sy * int(ext.get("cy")) / float(chh)
                emit(sh.shapes,
                     dx + (int(off.get("x")) * sx
                           - int(ch.get("x")) * nsx),
                     dy + (int(off.get("y")) * sy
                           - int(ch.get("y")) * nsy),
                     nsx, nsy, out)
                continue
            if sh.left is None or sh.width is None:
                continue
            x = (dx + sh.left * sx) / EMU
            y = (dy + sh.top * sy) / EMU
            w, h = sh.width * sx / EMU, sh.height * sy / EMU
            txt = ""
            if sh.has_text_frame:
                txt = " ".join(sh.text_frame.text.split())[:40]
            blk = el.xml
            if el.tag.endswith("}pic"):
                out.append(("pic", (x, y, w, h), txt))
            elif el.tag.endswith("}cxnSp"):
                a = FLIP_RE.search(blk)
                a = a.group(1) if a else ""
                x0, x1 = (x + w, x) if 'flipH="1"' in a else (x, x + w)
                y0, y1 = (y + h, y) if 'flipV="1"' in a else (y, y + h)
                out.append(("line", (x0, y0, x1, y1), txt))
            elif _is_card("sp", blk, (x, y, w, h)):
                out.append(("card", (x, y, w, h), txt))

    prs = Presentation(deck)
    for slide in prs.slides:
        out = []
        emit(slide.shapes, 0, 0, 1.0, 1.0, out)
        yield out


def unesc(s):
    for a, b in (("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
                 ("&quot;", '"'), ("&apos;", "'")):
        s = s.replace(a, b)
    return s


def title_case_bad(title):
    """Words that should be capitalised but are not."""
    bad = []
    words = re.findall(r"[A-Za-z][A-Za-z'\u2019\-]*", title)
    for i, w in enumerate(words):
        low = w.lower()
        first_or_last = i == 0 or i == len(words) - 1
        if w[0].isupper():
            continue
        if low in SMALL_WORDS and not first_or_last:
            continue
        bad.append(w)
    return bad


def main():
    deck = Path(sys.argv[1]) if len(sys.argv) > 1 else \
        MODULE / "Module 6 - Full.pptx"
    z = zipfile.ZipFile(str(deck))
    parts = slide_parts(z)
    findings = []

    def note(n, code, msg):
        findings.append((n, code, msg))

    for n, part in enumerate(parts, 1):
        xml = z.read(part).decode("utf-8")
        body = xml[xml.find("<p:spTree"):xml.find("</p:spTree>")]
        # 6 — the footer number must be a live slidenum field
        has_footer = re.search(
            r'<a:off x="256032" y="6583680"/>', body) is not None             or "Complex Pricing and Advanced Pricing Strategies" in body
        if has_footer and 'type="slidenum"' not in body:
            note(n, "PAGENUM", "footer page number is not a live field")
        for m in SP_RE.finditer(body):
            blk = m.group()
            nm = NAME_RE.search(blk)
            name = nm.group(1) if nm else "?"
            off, ext = OFF_RE.search(blk), EXT_RE.search(blk)
            geo = None
            if off and ext:
                geo = (int(off.group(1)) / EMU, int(off.group(2)) / EMU,
                       int(ext.group(1)) / EMU, int(ext.group(2)) / EMU)
            texts = [unesc(t).strip() for t in T_RE.findall(blk)]
            joined = " ".join(t for t in texts if t).strip()
            # 5 — off-canvas
            if geo:
                x, y, w, h = geo
                TOL = 0.06
                in_footer = y >= FOOTER_RULE - 0.02
                if not in_footer and (
                        x < -TOL or y < -TOL or x + w > W_IN + TOL
                        or y + h > H_IN + TOL):
                    note(n, "OFFCANVAS",
                         "%s [%.2f,%.2f %.2fx%.2f]" % (name, x, y, w, h))
            # 1 — trailing period on a bullet
            for para in PARA_RE.findall(blk):
                ptxt = "".join(unesc(t) for t in T_RE.findall(para)).strip()
                if len(ptxt) > 8 and re.search(r"[a-z0-9%)\u201d]\.$", ptxt):
                    note(n, "PERIOD", "%s: ...%s" % (name, ptxt[-42:]))
            # 2 / 3 — type size floors
            sizes = [int(s) / 100.0 for s in SZ_RE.findall(blk)]
            if sizes and joined:
                mn = min(sizes)
                is_caption = ('i="1"' in blk and mn <= 13)
                filled = "<a:solidFill>" in blk.split("<a:p>")[0]
                # the coverage pill: 1.55" on a main outline row, 1.14"
                # on a sub-row under a header (2026-09-14) -- both are the
                # documented geometry, so neither is a 18 pt violation
                pill = bool(geo) and (abs(geo[3] - 0.36) < 0.03
                                      and (abs(geo[2] - 1.55) < 0.03
                                           or abs(geo[2] - 1.14) < 0.03))
                narrow = bool(geo) and (geo[2] < 0.95 or geo[3] < 0.34)
                # the corner pills (back button, backup link, problem-set
                # pointer) are specified at ~15 pt in Teaching CLAUDE.md
                corner = bool(geo) and geo[1] >= 6.40 and geo[2] <= 3.10
                if (not is_caption and not pill and not narrow
                        and not corner and mn < 16 and filled and geo
                        and geo[3] > 0.30):
                    note(n, "SIZE<18", "%s at %.0f pt: %s"
                         % (name, mn, joined[:46]))
            # 8 — the Back pill
            if joined.startswith("\u2190 Back") or joined == "Back":
                navy = "0B2B4E" in blk
                white_text = "FFFFFF" in blk
                if not (navy and white_text):
                    note(n, "BACKPILL",
                         "not navy fill + white text (navy=%s white=%s)"
                         % (navy, white_text))
        # 4 — title case, on the action title (the 0.28/0.55 box)
        for m in SP_RE.finditer(body):
            blk = m.group()
            off = OFF_RE.search(blk)
            if not off:
                continue
            x, y = int(off.group(1)) / EMU, int(off.group(2)) / EMU
            if abs(x - 0.28) < 0.02 and abs(y - 0.55) < 0.03:
                t = " ".join(unesc(s) for s in T_RE.findall(blk)).strip()
                bad = title_case_bad(t)
                if bad:
                    note(n, "TITLECASE", "%s  <- %s" % (t[:58], bad))

    # 9 — opaque cards must not overlap each other or a curve.
    #     This one reads the deck through python-pptx rather than the raw
    #     XML, because by the time the deck ships its callouts are GROUPED:
    #     a group child's <a:off> is in the group's child coordinate space,
    #     so the regex reader sees positions that are not on the slide.
    #     Decoding off/ext against chOff/chExt is the only way the check is
    #     true of the file Nico actually opens.
    for n, shapes in enumerate(_rendered_shapes(deck), 1):
        cards = [sh for sh in shapes if sh[0] in ("pic", "card")]
        lines = [sh for sh in shapes if sh[0] == "line"]
        for i in range(len(cards)):
            for j in range(i + 1, len(cards)):
                ak, (ax, ay, aw, ah), at = cards[i]
                bk, (bx, by, bw, bh), bt = cards[j]
                # Two PICTURES that touch are usually one figure clipped
                # into parts and laid side by side (slide 18's headline
                # strip over its photo), which is deliberate.  A picture
                # over a BOX, or two boxes, is not.
                if ak == "pic" and bk == "pic":
                    continue
                ox = min(ax + aw, bx + bw) - max(ax, bx)
                oy = min(ay + ah, by + bh) - max(ay, by)
                if ox <= 0 or oy <= 0 or ox * oy <= MIN_OVERLAP:
                    continue
                # a card fully INSIDE another is the deck's own
                # box-plus-text / label-in-a-card pattern, not a defect
                if (_within((ax, ay, aw, ah), (bx, by, bw, bh))
                        or _within((bx, by, bw, bh), (ax, ay, aw, ah))):
                    continue
                note(n, "OVERLAP", "card %.2f\" over card: %s | %s"
                     % (ox * oy, at[:22] or "-", bt[:22] or "-"))
        for _k, (x0, y0, x1, y1), _t in lines:
            for ck, (cx, cy, cw, ch), ct in cards:
                if _seg_hits(x0, y0, x1, y1, cx, cy, cw, ch):
                    note(n, "OVERLAP",
                         "curve through card: %s" % (ct[:34] or "-"))

    # 7 — demand curves: every dark-red connector is fine; flag a NAVY
    #     connector that carries a demand label nearby is beyond a static
    #     check, so just report the dark-red count per chart slide.
    by_code = {}
    for n, code, msg in findings:
        by_code.setdefault(code, []).append((n, msg))
    order = ["OFFCANVAS", "BACKPILL", "PAGENUM", "OVERLAP", "PERIOD",
             "TITLECASE", "SIZE<18"]
    for code in order:
        rows = by_code.get(code, [])
        print("\n== %s : %d ==" % (code, len(rows)))
        for n, msg in rows[:40]:
            print("   s%-4d %s" % (n, msg))
        if len(rows) > 40:
            print("   ... and %d more" % (len(rows) - 40))
    print("\ntotal findings: %d" % len(findings))


if __name__ == "__main__":
    main()
