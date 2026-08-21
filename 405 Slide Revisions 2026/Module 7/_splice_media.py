# -*- coding: utf-8 -*-
"""Splice the live PollEverywhere / video slides from the original
"Module 7.pptx" into the built deck — phase-3 media step, kept as a
RERUNNABLE script so _build_Module7.py never freezes.

Pipeline:  python _build_Module7.py  ->  python _splice_media.py [deck]

Default: splices the lightweight poll slides only (slide 33 after the
2026-08-10 concept-map insert). --with-video additionally splices slide 15, whose video is an EMBEDDED
11.5 MB mp4 — run that only for the final/teaching copy so the working
deck (and each git commit of it) stays small.

Pure zip + lxml/ElementTree surgery: python-pptx would strip the poll
`tags` relationships and NULL-target video rels that make these slides
live. Original content is shifted +1.667" right to center the 4:3
layout on our 16:9 canvas. Rerunning after every rebuild is the
intended workflow; the script always starts from the fresh build
output, so nothing accumulates.
"""
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).parent
ORIGINAL = HERE / "Module 7.pptx"

NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
R_ID = '{%s}id' % NS_R
X_SHIFT_EMU = 1524000          # +1.667" to center 10" content on 13.33"

CT_BY_EXT = {
    'png': 'image/png', 'jpeg': 'image/jpeg', 'jpg': 'image/jpeg',
    'mp4': 'video/mp4', 'wmv': 'video/x-ms-wmv',
}
CT_TAGS = ('application/vnd.openxmlformats-officedocument.'
           'presentationml.tags+xml')

# rel Types (suffix) -> handling: 'copy' parts, 'keep' external, 'ours'
REL_SLIDELAYOUT = 'slideLayout'
REL_NOTES = 'notesSlide'


def display_to_part(z):
    pres = ET.fromstring(z.read('ppt/presentation.xml'))
    rels = ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    relmap = {r.get('Id'): r.get('Target') for r in rels}
    sldlst = pres.find('{%s}sldIdLst' % NS_P)
    return ['ppt/' + relmap[s.get(R_ID)] for s in sldlst]


def splice(deck_path, displays):
    src = zipfile.ZipFile(ORIGINAL)
    src_parts = display_to_part(src)

    tmp = deck_path.with_suffix('.splice_tmp.pptx')
    with zipfile.ZipFile(deck_path) as tgt:
        tgt_parts = display_to_part(tgt)
        tgt_names = set(tgt.namelist())
        items = {n: tgt.read(n) for n in tgt.namelist()}

    new_parts = {}          # name -> bytes to add
    ct_overrides = []       # (partname, content_type)
    ct_defaults = {}        # ext -> content type

    for disp in displays:
        t_part = tgt_parts[disp - 1]
        t_base = t_part.split('/')[-1]
        t_rels = ET.fromstring(items['ppt/slides/_rels/%s.rels' % t_base])

        if disp == 33:
            # Nico's UPDATED PollEv slide (inserted by hand 2026-08-10,
            # $120-anchor bands, embed ID a7ca8c79-…), preserved as
            # sidecar files — spliced verbatim, already 16:9-sized.
            # Sidecar XML references rId1 = tags, rId4 = image; the
            # rels file below must keep those exact ids.
            slide_xml = (HERE / '_handoff_s32_poll.xml').read_text(
                encoding='utf-8')
            layout_tgt = None
            for r in t_rels:
                if r.get('Type').endswith(REL_SLIDELAYOUT):
                    layout_tgt = r.get('Target')
            out = ET.Element('{%s}Relationships' % NS_REL)
            # CRITICAL: rId3 notes part with the poll title + URL. The
            # PollEverywhere add-in scans the deck at SLIDESHOW start,
            # finds its __PE_POLL_EMBED_ID tag, and reads the poll data
            # from the slide NOTES — a poll slide WITHOUT this notes
            # part crashes the slideshow renderer deck-wide ("The slide
            # failed to open properly" on slide 1; root-caused
            # 2026-08-11 via slideshow-window bisection).
            spec = [('rId1', 'tags', '../tags/pe32_tag.xml'),
                    ('rId2', REL_SLIDELAYOUT, layout_tgt),
                    ('rId3', REL_NOTES, '../notesSlides/pe33_notes.xml'),
                    ('rId4', 'image', '../media/pe32_poll.png')]
            for rid, suffix, tgt_rel in spec:
                e = ET.SubElement(out, '{%s}Relationship' % NS_REL)
                e.set('Id', rid)
                e.set('Type', 'http://schemas.openxmlformats.org/'
                              'officeDocument/2006/relationships/'
                              + suffix)
                e.set('Target', tgt_rel)
            new_parts['ppt/tags/pe32_tag.xml'] = (
                HERE / '_handoff_s32_tag.xml').read_bytes()
            new_parts['ppt/media/pe32_poll.png'] = (
                HERE / '_source_images/_s32_pollev.png').read_bytes()
            # build the notes part from an existing notesSlide template
            import re as _re
            tmpl_name = next(n for n in sorted(items)
                             if _re.match(
                                 r'ppt/notesSlides/notesSlide\d+\.xml$',
                                 n))
            tmpl = items[tmpl_name].decode('utf-8')
            tmpl_rels = items['ppt/notesSlides/_rels/%s.rels'
                              % tmpl_name.split('/')[-1]].decode('utf-8')
            note_lines = (HERE / '_handoff_s32_notes.txt').read_text(
                encoding='utf-8').splitlines()

            def _esc(t):
                return t.replace('&', '&amp;').replace('<', '&lt;')

            paras = ''.join(
                ('<a:p><a:r><a:rPr lang="en-US" dirty="0"/>'
                 '<a:t>%s</a:t></a:r></a:p>' % _esc(l)) if l
                else '<a:p/>' for l in note_lines)
            mbody = _re.search(
                r'(<p:sp>(?:(?!</p:sp>).)*?type="body"'
                r'(?:(?!</p:sp>).)*?</p:sp>)', tmpl, _re.S)
            body = mbody.group(1)
            newbody = _re.sub(
                r'(<p:txBody>.*?<a:lstStyle/>).*?(</p:txBody>)',
                lambda m: m.group(1) + paras + m.group(2),
                body, flags=_re.S)
            assert newbody != body, 'notes template body not rewritten'
            new_parts['ppt/notesSlides/pe33_notes.xml'] = tmpl.replace(
                body, newbody).encode('utf-8')
            new_parts['ppt/notesSlides/_rels/pe33_notes.xml.rels'] = \
                _re.sub(r'(Type="[^"]*/slide"[^>]*Target=")[^"]*(")',
                        lambda m: (m.group(1) + '../slides/' + t_base
                                   + m.group(2)),
                        tmpl_rels).encode('utf-8')
            ct_overrides.append(
                ('/ppt/notesSlides/pe33_notes.xml',
                 'application/vnd.openxmlformats-officedocument.'
                 'presentationml.notesSlide+xml'))
            ct_overrides.append(('/ppt/tags/pe32_tag.xml', CT_TAGS))
            ct_defaults['png'] = CT_BY_EXT['png']
            items[t_part] = slide_xml.encode('utf-8')
            ET.register_namespace('', NS_REL)
            items['ppt/slides/_rels/%s.rels' % t_base] = ET.tostring(
                out, xml_declaration=True, encoding='UTF-8')
            print('spliced display 33 from sidecar (Nico poll '
                  'a7ca8c79)')
            continue

        s_part = src_parts[disp - 1]
        s_base = s_part.split('/')[-1]
        slide_xml = src.read(s_part).decode('utf-8')
        s_rels = ET.fromstring(
            src.read('ppt/slides/_rels/%s.rels' % s_base))

        # our stub's layout rel is reused; everything else is rebuilt
        layout_tgt = None
        for r in t_rels:
            if r.get('Type').endswith(REL_SLIDELAYOUT):
                layout_tgt = r.get('Target')
        assert layout_tgt, 'stub has no layout rel'

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
        # keep OUR notes part (stub notes) if the stub has one
        for r in t_rels:
            if r.get('Type').endswith(REL_NOTES):
                add_rel(r.get('Type'), r.get('Target'))

        for r in s_rels:
            typ = r.get('Type')
            suffix = typ.split('/')[-1]
            tgt_rel = r.get('Target')
            old_id = r.get('Id')
            if suffix in (REL_SLIDELAYOUT, REL_NOTES):
                continue
            if r.get('TargetMode') == 'External':
                rid_map[old_id] = add_rel(typ, tgt_rel, 'External')
                continue
            # internal part: copy under a collision-safe name
            src_name = 'ppt/' + tgt_rel.replace('../', '')
            ext = src_name.split('.')[-1].lower()
            folder = src_name.rsplit('/', 2)[-2]
            new_name = 'ppt/%s/pe%02d_%s' % (folder, disp,
                                             src_name.split('/')[-1])
            if new_name not in new_parts and new_name not in tgt_names:
                new_parts[new_name] = src.read(src_name)
                if folder == 'tags':
                    ct_overrides.append(('/' + new_name, CT_TAGS))
                elif ext in CT_BY_EXT:
                    ct_defaults[ext] = CT_BY_EXT[ext]
            rel_target = '../%s/%s' % (folder, new_name.split('/')[-1])
            rid_map[old_id] = add_rel(typ, rel_target)

        for old, new in rid_map.items():
            slide_xml = slide_xml.replace('"%s"' % old, '"%s"' % new)

        # center the 4:3 content on the 16:9 canvas: shift every
        # <a:off> x by +1.667". String-level edit — an ElementTree
        # round-trip would re-prefix namespaces and break the
        # mc:AlternateContent Requires="p14" binding. NOTE: this also
        # shifts offsets inside groups (chOff-relative); fine for the
        # group-less poll slides — recheck when enabling slide 14.
        import re
        slide_xml = re.sub(
            r'<a:off x="(-?\d+)"',
            lambda m: '<a:off x="%d"' % (int(m.group(1)) + X_SHIFT_EMU),
            slide_xml)

        items[t_part] = slide_xml.encode('utf-8')
        ET.register_namespace('', NS_REL)
        items['ppt/slides/_rels/%s.rels' % t_base] = ET.tostring(
            out, xml_declaration=True, encoding='UTF-8')
        print('spliced display %d: %s -> %s' % (disp, s_part, t_part))

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

    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)
        for name, data in new_parts.items():
            zout.writestr(name, data)
    shutil.move(str(tmp), str(deck_path))
    print('saved', deck_path)


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    deck = Path(args[0]) if args else HERE / 'Module 7 - Revised.pptx'
    displays = [33]                      # polls (concrete PollEv)
    if '--with-video' in sys.argv:
        displays.append(15)              # embedded 11.5 MB mp4 — final
    splice(deck, displays)
