"""
Replace the three logo overlays with real photographs:

  Slide 26 – Anthropic logo  → Demis Hassabis Nobel Prize photo
              (the DeepMind chief whose researcher Anthropic is poaching)
  Slide 51 – ChatGPT logo    → Smartphone running the ChatGPT app
  Slide 72 – Alphabet logo   → Googleplex (Google HQ) campus photo

Each existing logo shape (named "ExampleImage" by the previous overlay
step) is removed, and a new <p:pic> is added with a real photograph and
an attribution caption.
"""

import sys
import zipfile
from pathlib import Path

from lxml import etree

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from _polish_v2 import (  # noqa: E402
    NS, Q, parse, serialize,
    REL_IMAGE,
    get_slide_path_at_position, next_rId,
)

TARGET = HERE / "Module 3_NEW_draft.pptx"

# (slide_position, new_image_path, image_filename_in_pkg, x_in, y_in, w_in, h_in, attribution)
REPLACEMENTS = [
    (26, HERE / "_hassabis.jpg",     "_hassabis.jpg",
     9.7, 4.4, 3.2, 2.5,
     "Demis Hassabis (CC BY, C. Michel via Wikimedia)"),

    (51, HERE / "_chatgpt_phone.jpg", "_chatgpt_phone.jpg",
     10.1, 2.8, 2.8, 3.5,
     "ChatGPT on iPhone (CC BY-SA, Wikimedia)"),

    (72, HERE / "_googleplex.jpg",    "_googleplex.jpg",
     8.0, 4.5, 4.8, 2.4,
     "Googleplex, Mountain View (CC BY-SA, Wikimedia)"),
]

# Old logo media filenames to drop from the package
OLD_MEDIA_TO_DROP = [
    "ppt/media/_anthropic_logo.png",
    "ppt/media/_chatgpt_logo.png",
    "ppt/media/_alphabet_logo.png",
]


def remove_existing_example_image(members, slide_path):
    """Strip any <p:pic> shape named 'ExampleImage' and the caption
    text shape named 'AttrText' from the slide. Returns rIds whose
    rels can be removed from the slide rels file."""
    slide = parse(members[slide_path])
    spTree = slide.find(Q('p', 'cSld') + '/' + Q('p', 'spTree'))
    removed_rIds = []
    for shp in list(spTree):
        tag = shp.tag.split('}', 1)[1]
        if tag not in ('pic', 'sp'):
            continue
        nvPr = shp.find(Q('p', 'nvPicPr')) if tag == 'pic' else shp.find(Q('p', 'nvSpPr'))
        if nvPr is None:
            continue
        cNvPr = nvPr.find(Q('p', 'cNvPr'))
        if cNvPr is None:
            continue
        name = cNvPr.get('name', '')
        if name not in ('ExampleImage', 'AttrText'):
            continue
        # Capture image rId if this is the pic
        if tag == 'pic':
            blipFill = shp.find(Q('p', 'blipFill'))
            if blipFill is not None:
                blip = blipFill.find(Q('a', 'blip'))
                if blip is not None:
                    rid = blip.get(Q('r', 'embed'))
                    if rid:
                        removed_rIds.append(rid)
        spTree.remove(shp)
    members[slide_path] = serialize(slide)
    return removed_rIds


def drop_image_rels(members, slide_rels_path, rIds_to_drop):
    if not rIds_to_drop:
        return
    rels = parse(members[slide_rels_path])
    for rel in list(rels.findall(Q('rel', 'Relationship'))):
        if rel.get('Id') in rIds_to_drop and rel.get('Type') == REL_IMAGE:
            rels.remove(rel)
    members[slide_rels_path] = serialize(rels)


def add_picture_with_attribution(members, slide_position, image_local_path,
                                  image_dest_filename, x_in, y_in, w_in, h_in,
                                  attribution_text):
    image_bytes = image_local_path.read_bytes()
    ext = image_local_path.suffix.lstrip('.').lower()
    new_image_path = f'ppt/media/{image_dest_filename}'
    members[new_image_path] = image_bytes

    # Content_Types Default
    ct = parse(members['[Content_Types].xml'])
    if not any(d.get('Extension') == ext for d in ct.findall(Q('ct', 'Default'))):
        default = etree.Element(Q('ct', 'Default'))
        default.set('Extension', ext)
        default.set('ContentType',
                    'image/jpeg' if ext in ('jpg', 'jpeg') else f'image/{ext}')
        ct.insert(0, default)
        members['[Content_Types].xml'] = serialize(ct)

    slide_path = get_slide_path_at_position(members, slide_position)
    slide_dir, slide_file = slide_path.rsplit('/', 1)
    slide_rels_path = f'{slide_dir}/_rels/{slide_file}.rels'

    slide_rels = parse(members[slide_rels_path])
    img_rId = next_rId(slide_rels)
    img_rel = etree.SubElement(slide_rels, Q('rel', 'Relationship'))
    img_rel.set('Id', img_rId)
    img_rel.set('Type', REL_IMAGE)
    img_rel.set('Target', f'../media/{image_dest_filename}')
    members[slide_rels_path] = serialize(slide_rels)

    EMU = 914400
    x = int(x_in * EMU)
    y = int(y_in * EMU)
    w = int(w_in * EMU)
    h = int(h_in * EMU)
    pic_xml = (
        f'<p:pic xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}" xmlns:r="{NS["r"]}">'
        f'<p:nvPicPr>'
          f'<p:cNvPr id="9999" name="ExampleImage"/>'
          f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
          f'<p:nvPr/>'
        f'</p:nvPicPr>'
        f'<p:blipFill>'
          f'<a:blip r:embed="{img_rId}"/>'
          f'<a:stretch><a:fillRect/></a:stretch>'
        f'</p:blipFill>'
        f'<p:spPr>'
          f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
          f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</p:spPr>'
        f'</p:pic>'
    )
    slide = parse(members[slide_path])
    spTree = slide.find(Q('p', 'cSld') + '/' + Q('p', 'spTree'))
    spTree.append(etree.fromstring(pic_xml))

    # Attribution caption below image
    cap_xml = (
        f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">'
        f'<p:nvSpPr>'
          f'<p:cNvPr id="9998" name="AttrText"/>'
          f'<p:cNvSpPr txBox="1"/><p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
          f'<a:xfrm><a:off x="{x}" y="{y + h + int(0.05 * EMU)}"/>'
          f'<a:ext cx="{w}" cy="{int(0.3 * EMU)}"/></a:xfrm>'
          f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
          f'<a:noFill/>'
        f'</p:spPr>'
        f'<p:txBody>'
          f'<a:bodyPr wrap="square" lIns="0" rIns="0" tIns="0" bIns="0"/>'
          f'<a:lstStyle/>'
          f'<a:p><a:pPr algn="ctr"/>'
          f'<a:r><a:rPr sz="900" i="1">'
          f'<a:solidFill><a:srgbClr val="555B66"/></a:solidFill>'
          f'<a:latin typeface="Calibri"/></a:rPr>'
          f'<a:t>{attribution_text}</a:t></a:r></a:p>'
        f'</p:txBody>'
        f'</p:sp>'
    )
    spTree.append(etree.fromstring(cap_xml))
    members[slide_path] = serialize(slide)


def main():
    print(f"Reading {TARGET.name}")
    with zipfile.ZipFile(TARGET, 'r') as zf:
        members = {n: zf.read(n) for n in zf.namelist()}

    for pos, img_path, img_file, x, y, w, h, attrib in REPLACEMENTS:
        print(f"Slide {pos}:")
        slide_path = get_slide_path_at_position(members, pos)
        slide_dir, slide_file = slide_path.rsplit('/', 1)
        slide_rels_path = f'{slide_dir}/_rels/{slide_file}.rels'

        # 1) Remove existing ExampleImage + AttrText
        rIds = remove_existing_example_image(members, slide_path)
        print(f"  removed existing ExampleImage shapes (image rIds: {rIds})")
        drop_image_rels(members, slide_rels_path, rIds)

        # 2) Add new picture
        add_picture_with_attribution(
            members, pos, img_path, img_file, x, y, w, h, attrib
        )
        print(f"  added new picture {img_file} with attribution")

    # 3) Drop the orphan logo media files from the package
    for old in OLD_MEDIA_TO_DROP:
        if old in members:
            del members[old]
            print(f"  dropped old media: {old}")
        # Also drop Content_Types Default for any extension only used by these
    # (jpg/png defaults stay since real photos still use them)

    # Write atomically
    tmp = TARGET.with_suffix('.pptx.tmp')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zf:
        for n, d in members.items():
            zf.writestr(n, d)
    tmp.replace(TARGET)
    print(f"\nWrote {TARGET.name}")


if __name__ == "__main__":
    main()
