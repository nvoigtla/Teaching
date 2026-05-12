"""
Replace slide 1 cleanly:

  1. Build a clean sidecar deck via python-pptx containing only the
     template's title slide.
  2. Delete the current slide 1 in Module 3_NEW_draft.pptx
     (drop XML, rels, notesSlide, Content_Types Override, presentation
     Relationship, sldIdLst entry).
  3. Inject the sidecar's title slide into the deck as a NEW slide
     (gets appended to the end), then move it to position 1.
  4. Re-add the speaker note for slide 1.

This avoids any inheritance from the destination master (no
showMasterSp tricks, no <p:bg> overrides) because the slide is a
fresh part with a fresh rels file.
"""

import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from _build_template_samples import (  # noqa: E402
    SLIDE_W, SLIDE_H, slide_title,
)
from _speaker_notes import NOTES  # noqa: E402
from _polish_v2 import (  # noqa: E402
    NS, Q, parse, serialize, _resolve_relative,
    REL_SLIDE, REL_LAYOUT, REL_IMAGE, REL_NOTES, REL_NOTES_MASTER,
    CT_NOTES,
    get_slide_path_at_position, next_rId,
    build_notes_slide_xml, _xml_escape,
)

TARGET = HERE / "Module 3_NEW_draft.pptx"
SIDECAR = HERE / "_sidecar_title_v3.pptx"


# ============================================================
# Cut: drop a slide (and its notesSlide) from the package
# ============================================================

def cut_slide_by_position(members, pos_1_based):
    print(f"  cutting slide at position {pos_1_based}")
    slide_path = get_slide_path_at_position(members, pos_1_based)
    slide_dir, slide_file = slide_path.rsplit('/', 1)
    slide_rels_path = f'{slide_dir}/_rels/{slide_file}.rels'

    # Collect referenced parts via slide's rels (notesSlide, images, etc.)
    parts_to_drop = {slide_path, slide_rels_path}
    if slide_rels_path in members:
        srels = parse(members[slide_rels_path])
        for rel in srels.findall(Q('rel', 'Relationship')):
            rtype = rel.get('Type', '')
            target = rel.get('Target', '')
            if rtype == REL_NOTES:
                ns_path = _resolve_relative(slide_dir, target)
                parts_to_drop.add(ns_path)
                ns_dir, ns_file = ns_path.rsplit('/', 1)
                parts_to_drop.add(f'{ns_dir}/_rels/{ns_file}.rels')
            # Don't drop images: they may be shared with other slides.
            # (Layout, master, etc. definitely shared — never drop.)

    # Find the presentation-level rId for the slide
    pres_rels = parse(members['ppt/_rels/presentation.xml.rels'])
    rId_to_drop = None
    for rel in list(pres_rels.findall(Q('rel', 'Relationship'))):
        if rel.get('Type') == REL_SLIDE:
            t = ('ppt/' + rel.get('Target')).replace('\\', '/')
            if t == slide_path:
                rId_to_drop = rel.get('Id')
                pres_rels.remove(rel)
                break
    members['ppt/_rels/presentation.xml.rels'] = serialize(pres_rels)
    print(f"    dropped pres rel {rId_to_drop}")

    # Drop sldId from sldIdLst
    pres = parse(members['ppt/presentation.xml'])
    sldIdLst = pres.find(Q('p', 'sldIdLst'))
    for sld in list(sldIdLst.findall(Q('p', 'sldId'))):
        if sld.get(Q('r', 'id')) == rId_to_drop:
            sldIdLst.remove(sld)
    members['ppt/presentation.xml'] = serialize(pres)

    # Drop Content_Types Overrides for the dropped parts
    ct = parse(members['[Content_Types].xml'])
    for ov in list(ct.findall(Q('ct', 'Override'))):
        pn = ov.get('PartName', '').lstrip('/')
        if pn in parts_to_drop:
            ct.remove(ov)
    members['[Content_Types].xml'] = serialize(ct)

    # Drop the actual files
    for p in parts_to_drop:
        members.pop(p, None)
    print(f"    dropped parts: {sorted(parts_to_drop)}")


# ============================================================
# Build sidecar + inject + move-to-position-1
# ============================================================

def build_sidecar():
    prs = Presentation()
    prs.slide_width = Emu(int(SLIDE_W))
    prs.slide_height = Emu(int(SLIDE_H))
    slide_title(prs)
    prs.save(SIDECAR)


def next_slide_filename(members):
    used = set()
    for name in members:
        if name.startswith('ppt/slides/slide') and name.endswith('.xml'):
            tail = name[len('ppt/slides/slide'):-len('.xml')]
            try:
                used.add(int(tail))
            except ValueError:
                pass
    n = 1
    while n in used:
        n += 1
    return f'ppt/slides/slide{n}.xml'


def inject_title_slide_at_position_1(members):
    # 1) Build sidecar with template's title slide
    build_sidecar()
    with zipfile.ZipFile(SIDECAR, 'r') as zf:
        title_xml = zf.read('ppt/slides/slide1.xml')

    # 2) Pick filename + add to package
    new_slide_path = next_slide_filename(members)
    slide_dir, slide_file = new_slide_path.rsplit('/', 1)
    new_rels_path = f'{slide_dir}/_rels/{slide_file}.rels'
    print(f"  injecting new title slide as {new_slide_path}")
    members[new_slide_path] = title_xml

    # 3) Build minimal rels (just layout)
    rels_root = etree.Element(Q('rel', 'Relationships'),
                              nsmap={None: NS['rel']})
    layout_rel = etree.SubElement(rels_root, Q('rel', 'Relationship'))
    layout_rel.set('Id', 'rId1')
    layout_rel.set('Type', REL_LAYOUT)
    layout_rel.set('Target', '../slideLayouts/slideLayout2.xml')
    members[new_rels_path] = serialize(rels_root)

    # 4) Add Content_Types Override
    ct = parse(members['[Content_Types].xml'])
    ov = etree.SubElement(ct, Q('ct', 'Override'))
    ov.set('PartName', '/' + new_slide_path)
    ov.set('ContentType',
           'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
    members['[Content_Types].xml'] = serialize(ct)

    # 5) Add Relationship + sldIdLst entry at position 0 (front)
    pres = parse(members['ppt/presentation.xml'])
    pres_rels = parse(members['ppt/_rels/presentation.xml.rels'])
    new_rId = next_rId(pres_rels)
    new_rel = etree.SubElement(pres_rels, Q('rel', 'Relationship'))
    new_rel.set('Id', new_rId)
    new_rel.set('Type', REL_SLIDE)
    new_rel.set('Target', new_slide_path[len('ppt/'):])
    members['ppt/_rels/presentation.xml.rels'] = serialize(pres_rels)

    sldIdLst = pres.find(Q('p', 'sldIdLst'))
    # Compute next unique sldId numeric id
    existing_ids = [int(s.get('id'))
                    for s in sldIdLst.findall(Q('p', 'sldId'))]
    new_id = max(existing_ids) + 1 if existing_ids else 256
    new_sld = etree.Element(Q('p', 'sldId'))
    new_sld.set('id', str(new_id))
    new_sld.set(Q('r', 'id'), new_rId)
    # Insert at front so it becomes slide 1
    sldIdLst.insert(0, new_sld)
    members['ppt/presentation.xml'] = serialize(pres)

    print(f"    new slide rId = {new_rId}, sldId id = {new_id}, "
          f"inserted at sldIdLst position 0")


# ============================================================
# Re-add the speaker note for the new slide 1
# ============================================================

def next_notesSlide_filename(members):
    used = set()
    for name in members:
        if name.startswith('ppt/notesSlides/notesSlide') and name.endswith('.xml'):
            tail = name[len('ppt/notesSlides/notesSlide'):-len('.xml')]
            try:
                used.add(int(tail))
            except ValueError:
                pass
    n = 1
    while n in used:
        n += 1
    return f'ppt/notesSlides/notesSlide{n}.xml'


def add_notes_for_slide_1(members):
    note_text = NOTES[1]
    slide_path = get_slide_path_at_position(members, 1)
    slide_dir, slide_file = slide_path.rsplit('/', 1)
    slide_rels_path = f'{slide_dir}/_rels/{slide_file}.rels'

    new_notes_path = next_notesSlide_filename(members)
    new_notes_filename = new_notes_path.rsplit('/', 1)[1]
    new_notes_rels_path = f'ppt/notesSlides/_rels/{new_notes_filename}.rels'

    members[new_notes_path] = build_notes_slide_xml(note_text).encode('utf-8')

    # notesSlide rels: notesMaster + back to slide
    rels_root = etree.Element(Q('rel', 'Relationships'),
                              nsmap={None: NS['rel']})
    master_rel = etree.SubElement(rels_root, Q('rel', 'Relationship'))
    master_rel.set('Id', 'rId1')
    master_rel.set('Type', REL_NOTES_MASTER)
    master_rel.set('Target', '../notesMasters/notesMaster1.xml')
    back_rel = etree.SubElement(rels_root, Q('rel', 'Relationship'))
    back_rel.set('Id', 'rId2')
    back_rel.set('Type', REL_SLIDE)
    back_rel.set('Target', f'../slides/{slide_file}')
    members[new_notes_rels_path] = serialize(rels_root)

    # Content_Types Override
    ct = parse(members['[Content_Types].xml'])
    ov = etree.SubElement(ct, Q('ct', 'Override'))
    ov.set('PartName', '/' + new_notes_path)
    ov.set('ContentType', CT_NOTES)
    members['[Content_Types].xml'] = serialize(ct)

    # Slide -> notesSlide rel
    slide_rels = parse(members[slide_rels_path])
    new_rel_id = next_rId(slide_rels)
    new_rel = etree.SubElement(slide_rels, Q('rel', 'Relationship'))
    new_rel.set('Id', new_rel_id)
    new_rel.set('Type', REL_NOTES)
    new_rel.set('Target', f'../notesSlides/{new_notes_filename}')
    members[slide_rels_path] = serialize(slide_rels)

    print(f"  added speaker note as {new_notes_path}")


# ============================================================
# Main
# ============================================================

def main():
    print(f"Reading {TARGET.name}")
    with zipfile.ZipFile(TARGET, 'r') as zf:
        members = {n: zf.read(n) for n in zf.namelist()}

    # 1) Cut current slide 1
    cut_slide_by_position(members, 1)

    # 2) Inject fresh template title slide at position 1
    inject_title_slide_at_position_1(members)

    # 3) Re-add the speaker note for the new slide 1
    add_notes_for_slide_1(members)

    tmp = TARGET.with_suffix('.pptx.tmp')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zf:
        for n, d in members.items():
            zf.writestr(n, d)
    tmp.replace(TARGET)
    print(f"\nWrote {TARGET.name}")


if __name__ == "__main__":
    main()
