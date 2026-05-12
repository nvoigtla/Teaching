"""Normalize every character bullet in Module 3_NEW_draft.pptx to a
navy square (▪, U+25AA) rendered in Calibri.

Touches slides + slide masters. Layouts have no character bullets. Auto-
numbered bullets (a./b., 1./2., A./B.) are left as-is. Bullet colors and
sizes are preserved.

Works via direct zip + lxml surgery (python-pptx round-trip is unsafe
for this deck — it would drop the NULL image rels that several original
slides depend on).
"""

import zipfile
from pathlib import Path

from lxml import etree

HERE = Path(__file__).parent
TARGET = HERE / "Module 3_NEW_draft.pptx"

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
SQUARE = "▪"  # U+25AA BLACK SMALL SQUARE


def fix_xml(xml_bytes):
    """Update <a:buChar> chars to ▪ and their sibling <a:buFont> to Calibri.

    Returns (new_bytes, n_chars_changed, n_fonts_changed) or (None, 0, 0)
    if nothing needed updating.
    """
    root = etree.fromstring(xml_bytes)
    n_chars = 0
    n_fonts = 0
    for buChar in root.iter(f'{{{NS_A}}}buChar'):
        if buChar.get('char') != SQUARE:
            buChar.set('char', SQUARE)
            n_chars += 1
        # Sibling buFont, if any, should be Calibri so the square renders
        # in the same family as the surrounding text. (Wingdings's "l"
        # is the traditional PPT bullet — if we leave that, ▪ won't draw.)
        pPr = buChar.getparent()
        for buFont in pPr.findall(f'{{{NS_A}}}buFont'):
            if buFont.get('typeface') != 'Calibri':
                buFont.set('typeface', 'Calibri')
                n_fonts += 1
    if n_chars == 0 and n_fonts == 0:
        return None, 0, 0
    out = etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True,
    )
    return out, n_chars, n_fonts


def main():
    print(f"Reading {TARGET.name}")
    with zipfile.ZipFile(TARGET, 'r') as zf:
        members = {name: zf.read(name) for name in zf.namelist()}

    n_files_changed = 0
    total_chars = 0
    total_fonts = 0
    for name in list(members.keys()):
        is_slide = name.startswith('ppt/slides/slide') and name.endswith('.xml')
        is_master = name.startswith('ppt/slideMasters/') and name.endswith('.xml')
        if not (is_slide or is_master):
            continue
        new, nc, nf = fix_xml(members[name])
        if new is not None:
            members[name] = new
            n_files_changed += 1
            total_chars += nc
            total_fonts += nf

    print(f"  files modified: {n_files_changed}")
    print(f"  buChar changes: {total_chars}")
    print(f"  buFont changes: {total_fonts}")

    # Write back atomically via a tempfile
    tmp = TARGET.with_suffix('.pptx.tmp')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, data in members.items():
            zf.writestr(name, data)
    tmp.replace(TARGET)
    print(f"Wrote {TARGET.name}")


if __name__ == "__main__":
    main()
