# -*- coding: utf-8 -*-
"""One-time setup for the Module 2 In-Class rebuild:
1. copy the proven pipeline modules from Module 7,
2. extract source images from Nico's deck into _source_images/,
3. extract CT-adopted images (news clippings, gas station) with ct_ prefix,
4. write _source_inventory.md: per-slide shape geometry + run formatting
   for the slides the build script must reproduce faithfully.
"""
import os
import shutil
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

HERE = Path(__file__).parent
M7 = HERE.parent / "Module 7"
NICO = HERE / "Module 2 - In Class with Solutions.pptx"
CT = HERE / "CT Module 2 - In Class.pptx"
IMG = HERE / "_source_images"
IMG.mkdir(exist_ok=True)

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
EMU = 914400.0

# ---- 1. pipeline modules -------------------------------------------------
for f in ("_build_template_samples.py", "_animate.py"):
    if not (HERE / f).exists():
        shutil.copy(M7 / f, HERE / f)
        print("copied", f)

# ---- helpers ---------------------------------------------------------------

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


# ---- 2. Nico's images ------------------------------------------------------
POLL_SLIDES = {4, 5, 11, 12, 30, 31, 35, 36, 40, 41, 46, 47, 57, 58, 64, 65}
SPLICED = POLL_SLIDES | {13}

zn = zipfile.ZipFile(NICO)
order = slide_order(zn)
slide_media = {}          # disp -> [(rid, medianame)]
for disp, part in enumerate(order, 1):
    rels = rels_of(zn, part)
    if rels is None:
        continue
    entries = []
    for r in rels:
        if r.get('Type').endswith('/image'):
            tgt = r.get('Target').replace('../', '')
            entries.append((r.get('Id'), tgt.split('/')[-1]))
    if entries:
        slide_media[disp] = entries

extracted = set()
for disp, entries in sorted(slide_media.items()):
    if disp in SPLICED:
        continue
    for rid, fname in entries:
        if fname in extracted:
            continue
        data = zn.read('ppt/media/' + fname)
        (IMG / fname).write_bytes(data)
        extracted.add(fname)
print(f"extracted {len(extracted)} images from Nico's deck")

# ---- 3. CT adopted images --------------------------------------------------
CT_WANT = {19: 'snob', 45: 'gas', 48: 'recession_ref', 49: 'inferior',
           55: 'crossprice'}
zc = zipfile.ZipFile(CT)
corder = slide_order(zc)
for disp, tag in CT_WANT.items():
    rels = rels_of(zc, corder[disp - 1])
    for r in rels:
        if r.get('Type').endswith('/image'):
            tgt = r.get('Target').replace('../', '')
            fname = 'ct_%s_%s' % (tag, tgt.split('/')[-1])
            if not (IMG / fname).exists():
                (IMG / fname).write_bytes(zc.read('ppt/' + tgt))
                print("ct image:", fname)

# ---- 4. inventory ----------------------------------------------------------

def fmt_in(v):
    return "%.2f" % (int(v) / EMU) if v is not None else "?"


def shape_geom(sp):
    xfrm = sp.find('.//a:xfrm', NS)
    if xfrm is None:
        return "(no xfrm)"
    off = xfrm.find('a:off', NS)
    ext = xfrm.find('a:ext', NS)
    if off is None or ext is None:
        return "(partial xfrm)"
    return "x=%s y=%s w=%s h=%s" % (
        fmt_in(off.get('x')), fmt_in(off.get('y')),
        fmt_in(ext.get('cx')), fmt_in(ext.get('cy')))


def run_fmt(r):
    rPr = r.find('a:rPr', NS)
    flags = []
    if rPr is not None:
        if rPr.get('b') == '1':
            flags.append('B')
        if rPr.get('i') == '1':
            flags.append('I')
        if rPr.get('u') and rPr.get('u') != 'none':
            flags.append('U')
        if rPr.get('sz'):
            flags.append(str(int(rPr.get('sz')) // 100) + 'pt')
        clr = rPr.find('.//a:srgbClr', NS)
        if clr is not None:
            flags.append('#' + clr.get('val'))
    t = r.find('a:t', NS)
    txt = t.text if t is not None and t.text else ''
    return txt, ','.join(flags)


def dump_txbody(sp, out, indent='    '):
    tx = sp.find('.//p:txBody', NS)
    if tx is None:
        return
    for p in tx.findall('a:p', NS):
        pPr = p.find('a:pPr', NS)
        lvl = int(pPr.get('lvl', '0')) if pPr is not None else 0
        parts = []
        for r in p.findall('a:r', NS):
            txt, fl = run_fmt(r)
            parts.append('⟨%s⟩%s' % (fl, txt) if fl else txt)
        if p.find('a14:m/..', {'a14': 'http://schemas.microsoft.com/office/drawing/2010/main'}) is not None:
            parts.append('[OMML-MATH]')
        # detect math anywhere in paragraph
        if any(el.tag.endswith('}oMath') for el in p.iter()):
            parts.append('[OMML]')
        line = ''.join(parts)
        if line.strip():
            out.append('%sL%d| %s' % (indent, lvl, line))


def dump_slide(z, part, disp, out, media_names):
    root = ET.fromstring(z.read(part))
    tree = root.find('p:cSld/p:spTree', NS)
    out.append('\n## Slide %d  [%s]' % (disp, part.split('/')[-1]))
    for child in tree:
        tag = child.tag.split('}')[1]
        if tag in ('nvGrpSpPr', 'grpSpPr'):
            continue
        name_el = child.find('.//p:cNvPr', NS)
        nm = name_el.get('name') if name_el is not None else '?'
        if tag == 'sp':
            geom = shape_geom(child)
            prst = child.find('.//a:prstGeom', NS)
            kind = prst.get('prst') if prst is not None else 'custom/ph'
            out.append('- sp "%s" %s [%s]' % (nm, geom, kind))
            dump_txbody(child, out)
        elif tag == 'pic':
            blip = child.find('.//a:blip', NS)
            rid = blip.get('{%s}embed' % NS['r']) if blip is not None else '?'
            fname = media_names.get((disp, rid), '?')
            out.append('- pic "%s" %s media=%s' % (nm, shape_geom(child), fname))
        elif tag == 'graphicFrame':
            out.append('- graphicFrame "%s" %s' % (nm, shape_geom(child)))
        elif tag == 'grpSp':
            out.append('- grpSp "%s" %s {' % (nm, shape_geom(child)))
            for sub in child:
                stag = sub.tag.split('}')[1]
                if stag in ('sp', 'cxnSp', 'pic', 'grpSp'):
                    snm_el = sub.find('.//p:cNvPr', NS)
                    snm = snm_el.get('name') if snm_el is not None else '?'
                    out.append('  - %s "%s" %s' % (stag, snm, shape_geom(sub)))
                    if stag == 'sp':
                        dump_txbody(sub, out, indent='      ')
            out.append('  }')
        elif tag == 'cxnSp':
            out.append('- cxnSp "%s" %s' % (nm, shape_geom(child)))
    # notes
    base = part.split('/')[-1]
    rels = rels_of(z, part)
    if rels is not None:
        for r in rels:
            if r.get('Type').endswith('/notesSlide'):
                npart = 'ppt/' + r.get('Target').replace('../', '')
                nroot = ET.fromstring(z.read(npart))
                txts = [t.text for t in nroot.iter('{%s}t' % NS['a'])
                        if t.text]
                note = ' '.join(txts).strip()
                if note and len(note) > 3:
                    out.append('NOTES: %s' % note)


out = ["# Module 2 source inventory (generated by _setup_assets.py)",
       "# Geometry in inches on the ORIGINAL 10x7.5 canvas (Nico) /",
       "# 13.33x7.5 (CT). Runs: ⟨B,I,sz,#color⟩text. L0/L1 = outline level."]

media_names = {}
for disp, entries in slide_media.items():
    for rid, fname in entries:
        media_names[(disp, rid)] = fname

out.append("\n" + "=" * 70 + "\n#  NICO DECK — Module 2 - In Class with Solutions.pptx\n" + "=" * 70)
for disp, part in enumerate(order, 1):
    if disp in SPLICED:
        out.append('\n## Slide %d  [SPLICED VERBATIM — not scripted]' % disp)
        continue
    dump_slide(zn, part, disp, out, media_names)

ct_media_names = {}
for disp, part in enumerate(corder, 1):
    rels = rels_of(zc, part)
    if rels is None:
        continue
    for r in rels:
        if r.get('Type').endswith('/image'):
            tgt = r.get('Target').replace('../', '')
            ct_media_names[(disp, r.get('Id'))] = tgt.split('/')[-1]

CT_REF = [2, 3, 5, 12, 18, 19, 22, 24, 27, 32, 33, 37, 38, 45, 48, 49, 55, 56]
out.append("\n" + "=" * 70 + "\n#  CT DECK (reference for adopted slides) — CT Module 2 - In Class.pptx\n" + "=" * 70)
for disp in CT_REF:
    dump_slide(zc, corder[disp - 1], disp, out, ct_media_names)

(HERE / "_source_inventory.md").write_text('\n'.join(out), encoding='utf-8')
print("wrote _source_inventory.md,", len(out), "lines")
