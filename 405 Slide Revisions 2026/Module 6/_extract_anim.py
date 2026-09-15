# -*- coding: utf-8 -*-
"""Extract Nico's ORIGINAL animation choreography from the Module 6 source
decks, click by click, so the rebuild can reproduce it beat for beat.

Teaching CLAUDE.md: "Never regenerate animations from the generic
guidelines when the source deck carries my hand-tuned choreography."  48 of
the 76 main-deck slides carry a timing tree, so most of Module 6 does.

For each slide this prints the main sequence as CLICK groups.  Every effect
is resolved from its ``spid`` to a shape SIGNATURE — kind, rendered
position/size in inches, and the first of its text — because shape ids are
not stable across a rebuild and only the signature can be matched against
the new deck.

Usage:
    python _extract_anim.py NV               # the main deck
    python _extract_anim.py NV 36 42         # only those slides
    python _extract_anim.py APP | V3         # the other source decks
"""
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'p': P_NS, 'a': A_NS, 'r': R_NS}
EMU = 914400.0

SRC = Path(__file__).parent / "Module 6 - NV Slides"
DECKS = {
    "NV": "Module 6.pptx",
    "APP": "Module 6 -- Slides On-Campus Applications with Solutions.pptx",
}
for _k in range(1, 9):
    DECKS["V%d" % _k] = "Module 6 - Video %d.pptx" % _k

# The entrance/exit presets PowerPoint writes, mapped to something readable.
FILTER_NAME = {
    "fade": "FADE", "wipe(up)": "WIPE-UP", "wipe(down)": "WIPE-DOWN",
    "wipe(left)": "WIPE-LEFT", "wipe(right)": "WIPE-RIGHT",
    "dissolve": "DISSOLVE", "barn(inVertical)": "BARN",
    "checkerboard(across)": "CHECKER", "blinds(horizontal)": "BLINDS",
    "circle": "CIRCLE", "plus": "PLUS", "wedge": "WEDGE",
    "randombar(horizontal)": "RANDOMBAR", "strips(downLeft)": "STRIPS",
}


def slide_parts(z):
    """Display order -> part name, via presentation.xml (NEVER assume
    slideN.xml is display slide N -- PowerPoint renumbers parts on save)."""
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = {r.get('Id'): r.get('Target')
            for r in ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    return ["ppt/" + rels[s.get('{%s}id' % R_NS)].replace("../", "")
            for s in pres.find('p:sldIdLst', NS)]


def _xfrm(el):
    x = el.find('.//' + '{%s}xfrm' % A_NS)
    if x is None:
        return None
    off, ext = x.find('{%s}off' % A_NS), x.find('{%s}ext' % A_NS)
    if off is None or ext is None:
        return None
    return (int(off.get('x')) / EMU, int(off.get('y')) / EMU,
            int(ext.get('cx')) / EMU, int(ext.get('cy')) / EMU)


def shape_index(spTree, parent_xf=None, out=None):
    """id -> signature, walking into groups and decoding group transforms
    so a child's RENDERED position is what gets reported."""
    if out is None:
        out = {}
    for el in spTree:
        tag = el.tag.split('}')[1]
        if tag == 'grpSp':
            gx = el.find('{%s}grpSpPr/{%s}xfrm' % (P_NS, A_NS))
            child_xf = None
            if gx is not None:
                off = gx.find('{%s}off' % A_NS)
                ext = gx.find('{%s}ext' % A_NS)
                cho = gx.find('{%s}chOff' % A_NS)
                che = gx.find('{%s}chExt' % A_NS)
                if None not in (off, ext, cho, che):
                    child_xf = (int(off.get('x')), int(off.get('y')),
                                int(ext.get('cx')), int(ext.get('cy')),
                                int(cho.get('x')), int(cho.get('y')),
                                int(che.get('cx')), int(che.get('cy')))
            shape_index(el, child_xf, out)
            continue
        nv = el.find('.//{%s}cNvPr' % P_NS)
        if nv is None:
            continue
        sid = nv.get('id')
        geo = _xfrm(el)
        if geo and parent_xf:
            ox, oy, cx, cy, chx, chy, chcx, chcy = parent_xf
            sx = cx / float(chcx or 1)
            sy = cy / float(chcy or 1)
            geo = ((ox + (geo[0] * EMU - chx) * sx) / EMU,
                   (oy + (geo[1] * EMU - chy) * sy) / EMU,
                   geo[2] * sx, geo[3] * sy)
        txt = " ".join(t.text or "" for t in el.iter('{%s}t' % A_NS)).strip()
        out[sid] = {
            'kind': {'sp': 'SHAPE', 'pic': 'PIC', 'graphicFrame': 'FRAME',
                     'cxnSp': 'CXN'}.get(tag, tag.upper()),
            'name': nv.get('name', ''),
            'geo': geo,
            'text': (txt[:60] + '…') if len(txt) > 60 else txt,
        }
    return out


def effect_of(node):
    """(direction, preset) for one effect node's children."""
    for ch in node.iter():
        t = ch.tag.split('}')[1]
        if t == 'animEffect':
            filt = ch.get('filter', '')
            return (ch.get('transition', 'in'),
                    FILTER_NAME.get(filt, filt or '?'))
        if t == 'anim' and ch.get('calcmode'):
            return ('in', 'MOTION')
        if t == 'animMotion':
            return ('in', 'MOTION')
    for ch in node.iter('{%s}set' % P_NS):
        for attr in ch.iter('{%s}attrName' % P_NS):
            if attr.text == 'style.visibility':
                for v in ch.iter('{%s}strVal' % P_NS):
                    return (('in' if v.get('val') == 'visible' else 'out'),
                            'APPEAR')
    return ('in', '?')


def spid_of(node):
    for tgt in node.iter('{%s}spTgt' % P_NS):
        return tgt.get('spid')
    return None


def prange(node):
    """A text range target, when the effect animates one paragraph."""
    for tr in node.iter('{%s}txEl' % P_NS):
        pr = tr.find('{%s}pRg' % P_NS)
        if pr is not None:
            return (int(pr.get('st')), int(pr.get('end')))
    return None


def main_sequence(root):
    """The main sequence.  NOTE: nodeType="mainSeq" sits on the <p:seq>'s
    CHILD <p:cTn>, not on the <p:seq> itself."""
    for seq in root.iter('{%s}seq' % P_NS):
        ctn = seq.find('{%s}cTn' % P_NS)
        if ctn is not None and ctn.get('nodeType') == 'mainSeq':
            return seq
    return None


def clicks_of(seq):
    """The main sequence as a list of clicks, each a list of effects."""
    ctn = seq.find('{%s}cTn' % P_NS)
    lst = ctn.find('{%s}childTnLst' % P_NS) if ctn is not None else None
    if lst is None:
        return []
    out = []
    for click_par in lst:
        effects = []
        for node in click_par.iter('{%s}cTn' % P_NS):
            nt = node.get('nodeType')
            if nt not in ('clickEffect', 'withEffect', 'afterEffect'):
                continue
            sid = spid_of(node)
            if sid is None:
                continue
            direction, preset = effect_of(node)
            effects.append({'trigger': nt, 'spid': sid, 'dir': direction,
                            'preset': preset, 'prg': prange(node),
                            'dur': node.get('dur')})
        if effects:
            out.append(effects)
    return out


def report(tag, wanted=None):
    z = zipfile.ZipFile(str(SRC / DECKS[tag]))
    parts = slide_parts(z)
    print("# Original animation choreography: %s" % DECKS[tag])
    print("#   %d slides; effects resolved from spid to a shape signature."
          % len(parts))
    print()
    for n, part in enumerate(parts, 1):
        if wanted and n not in wanted:
            continue
        root = ET.fromstring(z.read(part))
        seq = main_sequence(root)
        if seq is None:
            continue
        clicks = clicks_of(seq)
        if not clicks:
            continue
        idx = shape_index(root.find('.//{%s}spTree' % P_NS))
        print("## %s slide %d  —  %d clicks" % (tag, n, len(clicks)))
        for ci, effects in enumerate(clicks, 1):
            for ei, e in enumerate(effects):
                s = idx.get(e['spid'], {})
                geo = s.get('geo')
                pos = ("[%5.2f,%5.2f %5.2f x%5.2f]" % geo) if geo else "[?]"
                lead = ("  click %2d:" % ci) if ei == 0 else "          +"
                rng = (" pRg %d-%d" % e['prg']) if e['prg'] else ""
                print("%s %-4s %-7s %-6s %s %-6s %s%s"
                      % (lead, e['dir'].upper(), e['preset'], s.get('kind',
                                                                    '?'),
                         pos, s.get('name', '')[:16], s.get('text', ''), rng))
        print()


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        raise SystemExit(__doc__)
    tag = args[0]
    wanted = {int(a) for a in args[1:]} or None
    report(tag, wanted)
