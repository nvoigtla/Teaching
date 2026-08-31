# -*- coding: utf-8 -*-
"""Splice the 6 live PollEverywhere slides into the built Module 4 deck -
the phase-3 media step, RERUNNABLE so _build_Module4.py never freezes.

The source is a SIDECAR, `_handoff_polls_M4.pptx`, carved out of
"Module 4.pptx" by _make_poll_sidecar.ps1 with PowerPoint itself, so every
poll slide kept its tags part, its notes part and its screenshot.  The
sidecar is a BUILD INPUT - never delete it, and never round-trip it through
python-pptx.

The poll NOTES travel with the slide: the PollEverywhere add-in scans the
deck at slideshow start, finds its __PE_POLL_EMBED_ID tag and reads the
poll data out of that slide's notes.  A poll slide whose tag is present but
whose notes part is missing crashes the slideshow renderer DECK-WIDE.

Pure zip + ElementTree surgery: python-pptx would strip the poll `tags`
relationships.  The original 4:3 content is shifted +1.667" right to centre
it on the 16:9 canvas.

Pipeline:  python _build_Module4.py -> python _splice_media.py [deck]

The engine below is copied verbatim from Module 1/_splice_media.py
(2026-08-27); only the source and the splice map are Module 4's.
"""
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).parent
ORIGINAL = HERE / "_handoff_polls_M4.pptx"

NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
R_ID = '{%s}id' % NS_R
X_SHIFT_EMU = 1524000          # +1.667" to centre 10" content on 13.33"

CT_BY_EXT = {
    'png': 'image/png', 'jpeg': 'image/jpeg', 'jpg': 'image/jpg',
    'emf': 'image/x-emf', 'gif': 'image/gif',
    'xlsx': ('application/vnd.openxmlformats-officedocument.'
             'spreadsheetml.sheet'),
    'vml': 'application/vnd.openxmlformats-officedocument.vmlDrawing',
}
CT_TAGS = ('application/vnd.openxmlformats-officedocument.'
           'presentationml.tags+xml')
CT_NOTES = ('application/vnd.openxmlformats-officedocument.'
            'presentationml.notesSlide+xml')

SOURCES = {"M4": ORIGINAL}

# target display in "Module 4 - Revised.pptx" -> (sidecar, slide number IN
# THE SIDECAR).  The sidecar keeps the poll slides in their original
# relative order, so its 1-6 are the original deck's 20, 30, 37, 50, 62, 66.
SPLICE_MAP = {
    24: ("M4", 1),    # optimal quantity at P = 230 (question view)
}

def display_to_part(z):
    pres = ET.fromstring(z.read('ppt/presentation.xml'))
    rels = ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    relmap = {r.get('Id'): r.get('Target') for r in rels}
    sldlst = pres.find('{%s}sldIdLst' % NS_P)
    return ['ppt/' + relmap[s.get(R_ID)].lstrip('/') for s in sldlst]


def _copy_part_tree(src, src_name, disp, new_parts, tgt_names,
                    ct_overrides, ct_defaults):
    """Copy src part under a pe{disp}_ name; if it has a .rels file, copy
    its internal targets too (recursively) and rewrite the rels. Returns
    the new part name. Idempotent per (disp, src_name)."""
    folder = src_name.rsplit('/', 2)[-2]
    new_name = 'ppt/%s/pe%02d_%s' % (folder, disp, src_name.split('/')[-1])
    if new_name in new_parts or new_name in tgt_names:
        return new_name
    ext = src_name.split('.')[-1].lower()
    data = src.read(src_name)
    new_parts[new_name] = data
    if folder == 'tags':
        ct_overrides.append(('/' + new_name, CT_TAGS))
    elif ext in CT_BY_EXT:
        ct_defaults[ext] = CT_BY_EXT[ext]
    # dependent rels?
    base_dir = src_name.rsplit('/', 1)[0]
    rels_name = '%s/_rels/%s.rels' % (base_dir, src_name.split('/')[-1])
    if rels_name in set(src.namelist()):
        rels = ET.fromstring(src.read(rels_name))
        for r in rels:
            if r.get('TargetMode') == 'External':
                continue
            child_src = 'ppt/' + r.get('Target').replace('../', '')
            child_new = _copy_part_tree(src, child_src, disp, new_parts,
                                        tgt_names, ct_overrides,
                                        ct_defaults)
            child_folder = child_new.rsplit('/', 2)[-2]
            r.set('Target', '../%s/%s' % (child_folder,
                                          child_new.split('/')[-1]))
        ET.register_namespace('', NS_REL)
        new_parts['ppt/%s/_rels/%s.rels'
                  % (folder, new_name.split('/')[-1])] = ET.tostring(
            rels, xml_declaration=True, encoding='UTF-8')
    return new_name


def splice(deck_path):
    sources = {}
    for key, path in SOURCES.items():
        z = zipfile.ZipFile(path)
        sources[key] = (z, display_to_part(z))

    tmp = deck_path.with_suffix('.splice_tmp.pptx')
    with zipfile.ZipFile(deck_path) as tgt:
        tgt_parts = display_to_part(tgt)
        items = {n: tgt.read(n) for n in tgt.namelist()}
    tgt_names = set(items)

    new_parts = {}
    ct_overrides = []
    ct_defaults = {}

    # our deck's notesMaster (needed for copied notes parts)
    notes_master = next((n for n in tgt_names
                         if re.match(r'ppt/notesMasters/notesMaster\d+'
                                     r'\.xml$', n)), None)
    assert notes_master, "built deck has no notesMaster part"

    for disp, (src_key, s_disp) in sorted(SPLICE_MAP.items()):
        src, src_parts = sources[src_key]
        t_part = tgt_parts[disp - 1]
        t_base = t_part.split('/')[-1]
        t_rels = ET.fromstring(items['ppt/slides/_rels/%s.rels' % t_base])

        s_part = src_parts[s_disp - 1]
        s_base = s_part.split('/')[-1]
        slide_xml = src.read(s_part).decode('utf-8')
        s_rels = ET.fromstring(
            src.read('ppt/slides/_rels/%s.rels' % s_base))

        # reuse OUR stub's layout rel
        layout_tgt = None
        for r in t_rels:
            if r.get('Type').endswith('slideLayout'):
                layout_tgt = r.get('Target')
        assert layout_tgt, 'stub %d has no layout rel' % disp

        out = ET.Element('{%s}Relationships' % NS_REL)
        rid_map = {}
        n = 1

        def add_rel(rtype, target, mode=None):
            nonlocal n
            rid = 'rId%d' % n
            n += 1
            e = ET.SubElement(out, '{%s}Relationship' % NS_REL)
            e.set('Id', rid)
            e.set('Type', rtype)
            e.set('Target', target)
            if mode:
                e.set('TargetMode', mode)
            return rid

        add_rel('http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/slideLayout', layout_tgt)

        for r in s_rels:
            typ = r.get('Type')
            suffix = typ.split('/')[-1]
            tgt_rel = r.get('Target')
            old_id = r.get('Id')
            if suffix == 'slideLayout':
                continue
            if suffix == 'notesSlide':
                # copy the SOURCE notes part (poll mechanism lives there)
                src_notes = 'ppt/' + tgt_rel.replace('../', '')
                new_notes = 'ppt/notesSlides/pe%02d_notes.xml' % disp
                new_parts[new_notes] = src.read(src_notes)
                nrels = ET.Element('{%s}Relationships' % NS_REL)
                e = ET.SubElement(nrels, '{%s}Relationship' % NS_REL)
                e.set('Id', 'rId1')
                e.set('Type', 'http://schemas.openxmlformats.org/'
                              'officeDocument/2006/relationships/'
                              'notesMaster')
                e.set('Target', '../notesMasters/'
                                + notes_master.split('/')[-1])
                e = ET.SubElement(nrels, '{%s}Relationship' % NS_REL)
                e.set('Id', 'rId2')
                e.set('Type', 'http://schemas.openxmlformats.org/'
                              'officeDocument/2006/relationships/slide')
                e.set('Target', '../slides/' + t_base)
                ET.register_namespace('', NS_REL)
                new_parts['ppt/notesSlides/_rels/pe%02d_notes.xml.rels'
                          % disp] = ET.tostring(
                    nrels, xml_declaration=True, encoding='UTF-8')
                ct_overrides.append(('/' + new_notes, CT_NOTES))
                rid_map[old_id] = add_rel(typ, '../notesSlides/'
                                          + new_notes.split('/')[-1])
                continue
            if r.get('TargetMode') == 'External':
                rid_map[old_id] = add_rel(typ, tgt_rel, 'External')
                continue
            # internal part: copy under a collision-safe name, RECURSIVELY
            # carrying each copied part's own .rels (the vmlDrawing part of
            # the slide-13 OLE embed references the EMF preview through its
            # own rels file — dropping it breaks the preview image).
            src_name = 'ppt/' + tgt_rel.replace('../', '')
            new_name = _copy_part_tree(src, src_name, disp, new_parts,
                                       tgt_names, ct_overrides, ct_defaults)
            folder = new_name.rsplit('/', 2)[-2]
            rel_target = '../%s/%s' % (folder, new_name.split('/')[-1])
            rid_map[old_id] = add_rel(typ, rel_target)

        for old, new in rid_map.items():
            slide_xml = slide_xml.replace('"%s"' % old, '"%s"' % new)

        # center 4:3 content on the 16:9 canvas (string-level edit; an
        # ElementTree round-trip would break mc:AlternateContent)
        slide_xml = re.sub(
            r'<a:off x="(-?\d+)"',
            lambda m: '<a:off x="%d"' % (int(m.group(1)) + X_SHIFT_EMU),
            slide_xml)

        items[t_part] = slide_xml.encode('utf-8')
        ET.register_namespace('', NS_REL)
        items['ppt/slides/_rels/%s.rels' % t_base] = ET.tostring(
            out, xml_declaration=True, encoding='UTF-8')
        print('spliced display %d <- %s source %d (%s)'
              % (disp, src_key, s_disp, s_base))

    # content types
    ET.register_namespace('', NS_CT)
    ct = ET.fromstring(items['[Content_Types].xml'])
    have_defaults = {d.get('Extension').lower()
                     for d in ct.findall('{%s}Default' % NS_CT)}
    have_overrides = {o.get('PartName')
                      for o in ct.findall('{%s}Override' % NS_CT)}
    for ext, typ in ct_defaults.items():
        if ext not in have_defaults:
            e = ET.SubElement(ct, '{%s}Default' % NS_CT)
            e.set('Extension', ext)
            e.set('ContentType', typ)
    for part, typ in ct_overrides:
        if part not in have_overrides:
            e = ET.SubElement(ct, '{%s}Override' % NS_CT)
            e.set('PartName', part)
            e.set('ContentType', typ)
    items['[Content_Types].xml'] = ET.tostring(
        ct, xml_declaration=True, encoding='UTF-8')

    # 2026-08-26: a spliced slide arrives with the SOURCE deck's cached
    # footer number, which undoes the build's renumber pass for exactly
    # these slides. The number is a live `slidenum` field, so PowerPoint
    # recomputes it — but the cached value is what shows until it does, so
    # rewrite it here from the slide's real position. String-level, because
    # an ElementTree round-trip would break mc:AlternateContent.
    n_pg = 0
    for disp, t_part in enumerate(tgt_parts, start=1):
        xml = items[t_part].decode('utf-8')
        new_xml, n = re.subn(
            r'(<a:fld[^>]*type="slidenum"[^>]*>.*?<a:t>)\d+(</a:t>)',
            lambda m: m.group(1) + str(disp) + m.group(2), xml, flags=re.S)
        if n and new_xml != xml:
            items[t_part] = new_xml.encode('utf-8')
            n_pg += n
    if n_pg:
        print('cached page number(s) corrected: %d' % n_pg)

    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)
        for name, data in new_parts.items():
            zout.writestr(name, data)
    shutil.move(str(tmp), str(deck_path))
    print('saved', deck_path)


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    deck = Path(args[0]) if args else HERE / 'Module 4 - Revised.pptx'
    splice(deck)
