# -*- coding: utf-8 -*-
"""Splice the live PollEverywhere slides + the pizza Excel-embed slide from
"Module 2 - In Class with Solutions.pptx" into the built deck — phase-3
media step, RERUNNABLE so _build_Module2InClass.py never freezes.

Pipeline:  python _build_Module2InClass.py -> python _splice_media.py [deck]

Adapted from Module 7/_splice_media.py (2026-08-15). Two M2 changes:
 1. POLL NOTES TRAVEL WITH THE SLIDE. The PollEv add-in reads the poll
    title + URL from the slide NOTES at slideshow start; a poll slide
    without its notes part crashes the slideshow renderer deck-wide.
    The M7 generic branch kept the stub's notes — here we copy the
    SOURCE notes part for every spliced slide.
 2. OLE-EMBED CONTENT TYPES. Slide 13 carries an embedded Excel
    worksheet (xlsx + vmlDrawing + emf preview); the copy loop registers
    the extra content-type defaults from the source deck.

Pure zip + ElementTree surgery: python-pptx would strip the poll `tags`
relationships and the OLE/VML parts. Original 4:3 content is shifted
+1.667" right to center on the 16:9 canvas. Rerunning after every
rebuild is the intended workflow.
"""
import re
import shutil
import sys
import uuid
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).parent
ORIGINAL = HERE / "Module 2 - In Class with Solutions.pptx"

from _notes_m2 import SPLICED_NOTES

NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
R_ID = '{%s}id' % NS_R
X_SHIFT_EMU = 1524000          # +1.667" to center 10" content on 13.33"

# ---------------------------------------------------------------------------
# Deck chrome for the ONE spliced slide that needs it (2026-08-28, Nico:
# slide 13's header "looks strange").  A spliced slide is copied verbatim
# from the original deck, so it never went through _draw_top_bar_tc /
# _draw_action_title / _draw_footer.  Rather than hand-tune a second
# style here, these EMU values are copied from slide 10 - slide 13's
# sibling, built by those helpers - so the two slides are identical:
#   navy bar   (0, 0)              13.333 x 0.42
#   tag        (0.276, 0)          12.00  x 0.42   Calibri 16 pt bold white
#   title      (0.276, 0.55)       12.782 x 0.70   Calibri 32 pt bold navy
#   rule       (0.276, 1.25)       12.782 x 0.02   C8CDD3
#   gold strip (0.276, 1.235)       2.20  x 0.05   E09F3E
#   footer rule/strip/text at y 7.15 / 7.135 / 7.20
# The page number is NOT emitted here: every spliced slide already gets a
# live slidenum field from the generic block further down.
# The tag comes from the build script's own constant (it is import-safe:
# all deck building sits behind its __main__ guard), so a tag rename
# still has ONE home.
from _build_Module2InClass import (                       # noqa: E402
    TAG_LAW as TAG_LAW_13,
    FOOTER_TEXT as FOOTER_TEXT_13,
)
TITLE_13 = '1. MORE Customers Buy the Product'

_STYLE = ('<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>'
          '<a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>'
          '<a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>'
          '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>'
          '</p:style>')


def _chrome_rect(sid, name, x, y, cx, cy, fill):
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/>'
            '<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm>'
            '<a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:ln><a:noFill/></a:ln><a:effectLst/></p:spPr>%s'
            '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            '<a:p><a:pPr algn="ctr"/><a:endParaRPr/></a:p></p:txBody>'
            '</p:sp>' % (sid, name, x, y, cx, cy, fill, _STYLE))


def _chrome_text(sid, name, x, y, cx, cy, text, sz, color,
                 bold=False, middle=False):
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/>'
            '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm>'
            '<a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:noFill/></p:spPr><p:txBody>'
            '<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"%s>'
            '<a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="l"/>'
            '<a:r><a:rPr sz="%d" b="%d" i="0" u="none">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="Calibri"/></a:rPr><a:t>%s</a:t></a:r>'
            '</a:p></p:txBody></p:sp>'
            % (sid, name, x, y, cx, cy,
               ' anchor="ctr"' if middle else '',
               sz, 1 if bold else 0, color, text))


def _chrome_xml(tag_text, title_text):
    """The deck's standard bar / tag / title / rules / footer, as XML."""
    return ''.join([
        _chrome_rect(9810, 'TopBar', 0, 0, 12191695, 384048, '0B2B4E'),
        _chrome_text(9811, 'SectionTag', 251999, 0, 10972800, 384048,
                     tag_text, 1600, 'FFFFFF', bold=True, middle=True),
        _chrome_text(9812, 'ActionTitle', 251999, 502920, 11687697,
                     640080, title_text, 3200, '0B2B4E', bold=True),
        _chrome_rect(9813, 'TitleRule', 251999, 1143000, 11687697,
                     18288, 'C8CDD3'),
        _chrome_rect(9814, 'TitleGold', 251999, 1129284, 2011680,
                     45720, 'E09F3E'),
        _chrome_rect(9815, 'FooterRule', 0, 6537960, 12191695, 18288,
                     'C8CDD3'),
        _chrome_rect(9816, 'FooterGold', 251999, 6524244, 2011680,
                     45720, 'E09F3E'),
        _chrome_text(9817, 'FooterText', 251999, 6583680, 10058400,
                     292608, FOOTER_TEXT_13, 1200, '555B66'),
    ])

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

# new-deck display -> source-deck display
# 2026-08-15: displays shifted +1 from 14 on (bookend insert);
# 2026-08-23: bookend deleted by Nico, so the shift is back to zero
# 2026-08-25: Nico kept only the CHART slide of the water and yoga
# poll pairs (sources 31 and 36) and dropped the e-book pair (40/41)
# with the rest of that example.
SPLICE_MAP = {
    4: 4, 5: 5, 11: 11, 12: 12, 13: 13,
    32: 31, 36: 36,
    45: 46, 46: 47,
    57: 57, 58: 58, 64: 64, 65: 65,
}

# Per-slide source override: (deck, source display, apply the 4:3 -> 16:9
# x-shift).  Slide 13 (the live Excel embed) comes from its own 16:9
# SIDECAR rather than the original 4:3 deck (2026-08-28).  Why: Nico
# resized the OLE frame in PowerPoint from the source deck's squashed
# 7034213 x 4767263 to 12290425 x 5330825 EMU, which is the embed's
# native aspect (imgW/imgH = 2.305) and is what makes the table and the
# chart read correctly.  PowerPoint REGENERATED the cached EMF preview
# at that size when he saved; stretching the original deck's preview to
# the same frame instead renders a blown-up crop of the chart.  So the
# slide is carried whole - his frame, his preview, his workbook - from a
# one-slide sidecar carved out of his deck with PowerPoint via COM.
# The sidecar is a BUILD INPUT: never delete it, never round-trip it
# through python-pptx.
EXCEL_SIDECAR = HERE / "_handoff_excel_s13.pptx"
SRC_FOR = {
    13: (EXCEL_SIDECAR, 1, False),
}


def display_to_part(z):
    pres = ET.fromstring(z.read('ppt/presentation.xml'))
    rels = ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    relmap = {r.get('Id'): r.get('Target') for r in rels}
    sldlst = pres.find('{%s}sldIdLst' % NS_P)
    return ['ppt/' + relmap[s.get(R_ID)].lstrip('/') for s in sldlst]


X_SHIFT_PT = X_SHIFT_EMU / 12700.0     # 120 pt — same recentering shift


def _with_notes_text(data, text):
    """Write ``text`` into the body placeholder of a copied notes part.

    Used only for slide 13 (the live Excel embed), which is spliced in
    from the original deck and therefore cannot get its notes from the
    build script.  POLL slides must never come through here: their notes
    carry the PollEverywhere payload and are copied verbatim.
    """
    if not text:
        return data
    tree = ET.fromstring(data)
    for sp in tree.iter('{%s}sp' % NS_P):
        ph = sp.find('.//{%s}ph' % NS_P)
        if ph is None or ph.get('type') != 'body':
            continue
        tx = sp.find('{%s}txBody' % NS_P)
        if tx is None:
            continue
        for p in tx.findall('{%s}p' % NS_A):
            tx.remove(p)
        for para in text.split('\n\n'):
            p = ET.SubElement(tx, '{%s}p' % NS_A)
            r = ET.SubElement(p, '{%s}r' % NS_A)
            rPr = ET.SubElement(r, '{%s}rPr' % NS_A)
            rPr.set('lang', 'en-US')
            rPr.set('dirty', '0')
            t = ET.SubElement(r, '{%s}t' % NS_A)
            t.text = para
        break
    return ET.tostring(tree, xml_declaration=True, encoding='UTF-8')


def _shift_vml(data):
    """Recenter the legacy VML shapes of an OLE embed by the same amount
    as the slide's <a:off x>.

    PowerPoint pairs a <p:oleObj spid="_x0000_sNNNN"> with the VML shape
    of that id.  When only the graphicFrame moves, the pair no longer
    coincides and PowerPoint renders the VML shape as a SEPARATE picture
    lying on top of the OLE frame — which swallows the double-click, so
    the embedded workbook can no longer be opened (Nico, slide 13,
    2026-08-23).  Shifting the VML `left:` in lockstep keeps them one
    object.
    """
    txt = data.decode('utf-8')

    def bump(m):
        return 'left:%gpt' % (float(m.group(1)) + X_SHIFT_PT)

    return re.sub(r'left:(-?[\d.]+)pt', bump, txt).encode('utf-8')


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
    if ext == 'vml':
        data = _shift_vml(data)
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
    # one open handle per distinct source deck (slide 13 comes from its
    # own sidecar, everything else from the original In-Class deck)
    _zips = {}

    def _source(path):
        key = str(path)
        if key not in _zips:
            z = zipfile.ZipFile(path)
            _zips[key] = (z, display_to_part(z), set(z.namelist()))
        return _zips[key]

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

    for disp, s_disp in sorted(SPLICE_MAP.items()):
        s_deck, s_disp, do_shift = SRC_FOR.get(disp,
                                               (ORIGINAL, s_disp, True))
        src, src_parts, src_names = _source(s_deck)

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
                new_parts[new_notes] = _with_notes_text(
                    src.read(src_notes), SPLICED_NOTES.get(disp))
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

        # SINGLE-PASS remap.  A sequential replace() per entry clobbers
        # itself whenever the id space overlaps: on slide 13 the old
        # rId6 (EMF preview) became rId4, and the later rId4 -> rId6 pass
        # rewrote that same string, leaving the OLE fallback <a:blip>
        # pointing at the notes part.  PowerPoint then could not draw the
        # embed's preview and fell back to the loose VML picture, which
        # sat on top of the object and swallowed the double-click
        # (Nico: "I cannot open the underlying excel file", 2026-08-23).
        slide_xml = re.sub(
            r'"(rId\d+)"',
            lambda m: '"%s"' % rid_map.get(m.group(1), m.group(1)),
            slide_xml)

        # center 4:3 content on the 16:9 canvas (string-level edit; an
        # ElementTree round-trip would break mc:AlternateContent).  A
        # slide taken from a 16:9 sidecar is already positioned and must
        # NOT be shifted again.
        if do_shift:
            slide_xml = re.sub(
                r'<a:off x="(-?\d+)"',
                lambda m: '<a:off x="%d"' % (int(m.group(1))
                                             + X_SHIFT_EMU),
                slide_xml)
        # ...but NOT the spTree's own <p:grpSpPr> transform: the blanket
        # regex hits it too, and while the modern renderer ignores it
        # (ext = 0), the legacy VML path on the slide-13 OLE embed honours
        # it and shifts the whole slide a second time (2026-08-23).
        slide_xml = re.sub(
            r'(<p:grpSpPr><a:xfrm><a:off x=")(-?\d+)(")',
            lambda m: '%s0%s' % (m.group(1), m.group(3)), slide_xml)

        # slide 13 (live Excel): give the spliced slide the deck's own
        # chrome (2026-08-28, Nico: "the header looks strange").  It
        # arrived from the original deck with three defects:
        #   * a bare title in a layout PLACEHOLDER, no navy top bar, no
        #     section tag, no rule / gold strip and no footer, so it did
        #     not look like any other content slide;
        #   * an EMPTY body placeholder sitting at 0,0 in the top-left
        #     corner (invisible, but it is a real shape);
        #   * a number placeholder inheriting the stub layout's spot.
        # All three placeholders are dropped and the chrome is APPENDED
        # after the OLE frame, so the thin rule draws over the frame's
        # top edge (the frame starts at y 1.26, the rule sits at 1.25)
        # rather than under it.  Geometry is copied from slide 10, this
        # slide's sibling, so the two are pixel-identical.  The OLE
        # object itself is untouched — the embed keeps working.
        if disp == 13:
            for _ph in ('<p:ph type="title"', '<p:ph type="body"',
                        '<p:ph type="sldNum"'):
                _i = slide_xml.find(_ph)
                if _i == -1:
                    continue
                _j0 = slide_xml.rindex('<p:sp>', 0, _i)
                _j1 = slide_xml.index('</p:sp>', _i) + len('</p:sp>')
                slide_xml = slide_xml[:_j0] + slide_xml[_j1:]
            # the sidecar already carries a PageNum textbox (this pass
            # put it there on an earlier run); drop it so the generic
            # block below does not add a second one
            _i = slide_xml.find('name="PageNum"')
            if _i != -1:
                _j0 = slide_xml.rindex('<p:sp>', 0, _i)
                _j1 = slide_xml.index('</p:sp>', _i) + len('</p:sp>')
                slide_xml = slide_xml[:_j0] + slide_xml[_j1:]
            slide_xml = slide_xml.replace(
                '</p:spTree>', _chrome_xml(TAG_LAW_13, TITLE_13)
                + '</p:spTree>')

        # live page-number field on every spliced slide (2026-08-16:
        # "slide numbers throughout") — same look/position as
        # _add_slidenum_field in the build script
        guid = str(uuid.uuid5(uuid.NAMESPACE_DNS,
                              'm2ic-splice-num-%d' % disp)).upper()
        num_sp = (
            '<p:sp><p:nvSpPr><p:cNvPr id="9802" name="PageNum"/>'
            '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="11475720" y="6583680"/>'
            '<a:ext cx="502920" cy="292608"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:noFill/></p:spPr><p:txBody>'
            '<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0"/>'
            '<a:lstStyle/><a:p><a:pPr algn="r"/>'
            '<a:fld id="{%s}" type="slidenum">'
            '<a:rPr lang="en-US" sz="1200" dirty="0">'
            '<a:solidFill><a:srgbClr val="555B66"/></a:solidFill>'
            '<a:latin typeface="Calibri"/></a:rPr>'
            '<a:t>%d</a:t></a:fld></a:p></p:txBody></p:sp>'
            % (guid, disp))
        slide_xml = slide_xml.replace('</p:spTree>',
                                      num_sp + '</p:spTree>')

        items[t_part] = slide_xml.encode('utf-8')
        ET.register_namespace('', NS_REL)
        items['ppt/slides/_rels/%s.rels' % t_base] = ET.tostring(
            out, xml_declaration=True, encoding='UTF-8')
        print('spliced display %d <- source %d (%s)'
              % (disp, s_disp, s_base))

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
    deck = Path(args[0]) if args else HERE / 'Module 2 - In Class Revised.pptx'
    splice(deck)
