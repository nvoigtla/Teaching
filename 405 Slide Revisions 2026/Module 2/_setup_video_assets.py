# -*- coding: utf-8 -*-
"""Setup for the Video Part rebuild: extract images (Nico deck ->
_source_images_video/, CT adopted slides with ct_ prefix), write
_source_inventory_video.md (shapes + runs + OMML + tables), and dump
verbatim notes to _video_notes.py keyed by OLD slide number."""
import os
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
HERE = Path(__file__).parent
NICO = HERE / "Module 2 - Video Part.pptx"
CT = HERE / "CT Module 2 - Video Part.pptx"
IMG = HERE / "_source_images_video"
IMG.mkdir(exist_ok=True)

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}
EMU = 914400.0
SPLICED = {9, 20}


def slide_order(z):
    pres = ET.fromstring(z.read('ppt/presentation.xml'))
    rels = ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    relmap = {r.get('Id'): r.get('Target') for r in rels}
    return ['ppt/' + relmap[s.get('{%s}id' % NS['r'])].lstrip('/')
            for s in pres.find('p:sldIdLst', NS)]


def rels_of(z, part):
    base = part.split('/')[-1]
    try:
        return ET.fromstring(z.read('ppt/slides/_rels/%s.rels' % base))
    except KeyError:
        return None


zn = zipfile.ZipFile(NICO)
order = slide_order(zn)
media_names = {}
extracted = set()
for disp, part in enumerate(order, 1):
    rels = rels_of(zn, part)
    if rels is None:
        continue
    for r in rels:
        if r.get('Type').endswith('/image'):
            tgt = r.get('Target').replace('../', '')
            if 'NULL' in tgt or r.get('TargetMode') == 'External':
                continue
            fname = tgt.split('/')[-1]
            media_names[(disp, r.get('Id'))] = fname
            if disp not in SPLICED and fname not in extracted:
                (IMG / fname).write_bytes(zn.read('ppt/' + tgt))
                extracted.add(fname)
print("nico images:", len(extracted))

zc = zipfile.ZipFile(CT)
corder = slide_order(zc)
CT_WANT = {10: 'ozempic', 16: 'netflix', 18: 'mcdonalds', 31: 'insideout',
           35: 'abtest', 36: 'amazonexp', 37: 'amazonrecent'}
ct_media_names = {}
for disp, part in enumerate(corder, 1):
    rels = rels_of(zc, part)
    if rels is None:
        continue
    for r in rels:
        if r.get('Type').endswith('/image'):
            tgt = r.get('Target').replace('../', '')
            ct_media_names[(disp, r.get('Id'))] = tgt.split('/')[-1]
            if disp in CT_WANT:
                fname = 'ct_%s_%s' % (CT_WANT[disp], tgt.split('/')[-1])
                if not (IMG / fname).exists():
                    (IMG / fname).write_bytes(zc.read('ppt/' + tgt))
                    print("ct image:", fname)


def fmt_in(v):
    return "%.2f" % (int(v) / EMU) if v is not None else "?"


def shape_geom(sp, tag):
    if tag == 'graphicFrame':
        xf = sp.find('p:xfrm', NS)
    else:
        xf = sp.find('.//a:xfrm', NS)
    if xf is None:
        return "(no xfrm)"
    off = xf.find('a:off', NS)
    ext = xf.find('a:ext', NS)
    if off is None or ext is None:
        return "(partial)"
    return "x=%s y=%s w=%s h=%s" % (
        fmt_in(off.get('x')), fmt_in(off.get('y')),
        fmt_in(ext.get('cx')), fmt_in(ext.get('cy')))


def walk_para(p_el):
    out = []
    for el in p_el.iter():
        if el.tag == '{%s}t' % NS['a'] and el.text:
            out.append(el.text)
        elif el.tag == '{%s}t' % NS['m'] and el.text:
            out.append('⟪%s⟫' % el.text)
    return ''.join(out)


def run_fmt(r):
    rPr = r.find('a:rPr', NS)
    fl = []
    if rPr is not None:
        if rPr.get('b') == '1':
            fl.append('B')
        if rPr.get('i') == '1':
            fl.append('I')
        if rPr.get('u') and rPr.get('u') != 'none':
            fl.append('U')
        if rPr.get('sz'):
            fl.append(str(int(rPr.get('sz')) // 100))
        c = rPr.find('.//a:srgbClr', NS)
        if c is not None:
            fl.append('#' + c.get('val'))
    t = r.find('a:t', NS)
    return (t.text or '') if t is not None else '', ','.join(fl)


def dump_deck(z, order_, names, out, skip=(), which=None):
    for disp, part in enumerate(order_, 1):
        if which and disp not in which:
            continue
        out.append('\n## Slide %d  [%s]%s' % (
            disp, part.split('/')[-1],
            '  [SPLICED]' if disp in skip else ''))
        if disp in skip:
            continue
        root = ET.fromstring(z.read(part))
        tree = root.find('p:cSld/p:spTree', NS)
        for ch in tree:
            tag = ch.tag.split('}')[1]
            if tag not in ('sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'):
                continue
            nv = ch.find('.//p:cNvPr', NS)
            nm = nv.get('name') if nv is not None else '?'
            if tag == 'pic':
                blip = ch.find('.//a:blip', NS)
                rid = (blip.get('{%s}embed' % NS['r'])
                       if blip is not None else '?')
                out.append('- pic "%s" %s media=%s' % (
                    nm, shape_geom(ch, tag), names.get((disp, rid), '?')))
                continue
            if tag == 'graphicFrame':
                out.append('- graphicFrame "%s" %s' % (
                    nm, shape_geom(ch, tag)))
                for tr in ch.iter('{%s}tr' % NS['a']):
                    cells = []
                    for tc in tr.findall('{%s}tc' % NS['a']):
                        cells.append(''.join(
                            t.text or '' for t in
                            tc.iter('{%s}t' % NS['a'])))
                    out.append('    ROW| ' + ' | '.join(cells))
                continue
            out.append('- %s "%s" %s' % (tag, shape_geom(ch, tag), nm)
                       if False else
                       '- %s "%s" %s' % (tag, nm, shape_geom(ch, tag)))
            for p_el in ch.iter('{%s}p' % NS['a']):
                pPr = p_el.find('a:pPr', NS)
                lvl = int(pPr.get('lvl', '0')) if pPr is not None else 0
                bu = (pPr is not None
                      and pPr.find('a:buNone', NS) is not None)
                runs = []
                has_math = any(e.tag.endswith('}oMath')
                               for e in p_el.iter())
                for r in p_el.findall('a:r', NS):
                    txt, f = run_fmt(r)
                    runs.append(('⟨%s⟩%s' % (f, txt)) if f else txt)
                line = ''.join(runs)
                math = walk_para(p_el) if has_math else ''
                if line.strip() or math.strip():
                    out.append('    L%d%s| %s%s' % (
                        lvl, ' buNone' if bu else '', line[:180],
                        ('  MATH: ' + math[:180]) if has_math else ''))


out = ["# Module 2 VIDEO PART source inventory",
       "# Nico deck on 10x7.5 canvas; CT on 13.33x7.5.",
       "\n" + "=" * 70 + "\n#  NICO DECK\n" + "=" * 70]
dump_deck(zn, order, media_names, out, skip=SPLICED)
out.append("\n" + "=" * 70 + "\n#  CT DECK (adopted/reference slides)\n"
           + "=" * 70)
dump_deck(zc, corder, ct_media_names, out,
          which={9, 10, 13, 14, 16, 17, 18, 19, 22, 25, 31, 32, 35, 36,
                 37, 41, 42, 49, 55, 60, 61})
(HERE / "_source_inventory_video.md").write_text('\n'.join(out),
                                                 encoding='utf-8')
print("inventory:", len(out), "lines")

# ---- verbatim notes sidecar (old slide number -> notes text) ----
notes = {}
for disp, part in enumerate(order, 1):
    if disp in SPLICED:
        continue
    rels = rels_of(zn, part)
    if rels is None:
        continue
    for r in rels:
        if r.get('Type').endswith('/notesSlide'):
            npart = 'ppt/' + r.get('Target').replace('../', '')
            nroot = ET.fromstring(zn.read(npart))
            # body placeholder paragraphs only
            body = None
            for sp in nroot.iter('{%s}sp' % NS['p']):
                ph = sp.find('.//p:ph', NS)
                if ph is not None and ph.get('type') == 'body':
                    body = sp
                    break
            if body is None:
                continue
            paras = []
            for p_el in body.iter('{%s}p' % NS['a']):
                paras.append(walk_para(p_el))
            txt = '\n'.join(paras).strip()
            if len(txt) > 20:
                notes[disp] = txt
lines = ["# -*- coding: utf-8 -*-",
         '"""Verbatim speaker notes from "Module 2 - Video Part.pptx",',
         'keyed by OLD slide number (auto-generated — do not edit)."""',
         "", "NOTES = {"]
for k in sorted(notes):
    lines.append("    %d: %r," % (k, notes[k]))
lines.append("}")
(HERE / "_video_notes.py").write_text('\n'.join(lines), encoding='utf-8')
print("notes sidecar:", sorted(notes))
