"""Apply a navy bullet color + level-based size hierarchy to every
character-bullet paragraph in Module 3_NEW_draft.pptx.

  Level 0 (top):  150% (largest square)
  Level 1:        110%
  Level 2:         85%
  Level 3+:        70%

Touches:
  - Every <a:pPr> in slide XMLs that contains a <a:buChar>
  - Every <a:lvlNpPr> in slide-master XMLs that contains a <a:buChar>
    (so inherited bullets also pick up the new style)

Direct zip + lxml surgery – never round-trips through python-pptx (would
drop NULL image rels from several original slides).
"""

import re
import zipfile
from pathlib import Path

from lxml import etree

HERE = Path(__file__).parent
TARGET = HERE / "Module 3_NEW_draft.pptx"

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NAVY_HEX = "0B2B4E"

# Bullet size as percentage of the surrounding text, by paragraph level
LEVEL_SIZE_PCT = {0: 150, 1: 110, 2: 85, 3: 70}

# Schema-required ordering of <a:pPr> children. We must insert new
# elements at the right position or PowerPoint may reject the slide.
PPR_CHILD_ORDER = [
    'lnSpc', 'spcBef', 'spcAft',
    'buClr', 'buClrTx',
    'buSzTx', 'buSzPct', 'buSzPts',
    'buFontTx', 'buFont',
    'buNone', 'buAutoNum', 'buChar', 'buBlip',
    'tabLst', 'defRPr', 'extLst',
]


def _local(tag):
    return tag.split('}', 1)[1] if '}' in tag else tag


def _insert_in_order(parent, new_el):
    """Insert *new_el* among *parent*'s children at the schema-correct
    position (before the first child whose tag comes later in
    PPR_CHILD_ORDER, otherwise append)."""
    target_idx = PPR_CHILD_ORDER.index(_local(new_el.tag))
    for i, child in enumerate(parent):
        ctag = _local(child.tag)
        if ctag in PPR_CHILD_ORDER and PPR_CHILD_ORDER.index(ctag) > target_idx:
            parent.insert(i, new_el)
            return
    parent.append(new_el)


def _get_level(pPr):
    """Return the 0-based level for a pPr or lvlNpPr element."""
    tag = _local(pPr.tag)
    if tag == 'pPr':
        return int(pPr.get('lvl', '0'))
    m = re.match(r'lvl(\d+)pPr$', tag)
    if m:
        return int(m.group(1)) - 1
    return 0


def _set_bullet_color(pPr, hex_rgb):
    # Remove buClr / buClrTx (mutually exclusive)
    for t in ('buClr', 'buClrTx'):
        for el in list(pPr.findall(f'{{{NS_A}}}{t}')):
            pPr.remove(el)
    buClr = etree.Element(f'{{{NS_A}}}buClr')
    srgb = etree.SubElement(buClr, f'{{{NS_A}}}srgbClr')
    srgb.set('val', hex_rgb)
    _insert_in_order(pPr, buClr)


def _set_bullet_size_pct(pPr, pct):
    # Size variants are mutually exclusive
    for t in ('buSzTx', 'buSzPct', 'buSzPts'):
        for el in list(pPr.findall(f'{{{NS_A}}}{t}')):
            pPr.remove(el)
    buSz = etree.Element(f'{{{NS_A}}}buSzPct')
    buSz.set('val', str(pct * 1000))   # value is in thousandths of percent
    _insert_in_order(pPr, buSz)


def _fix_pPr(pPr):
    """If pPr has a buChar (square), apply navy + level-based size."""
    if pPr.find(f'{{{NS_A}}}buChar') is None:
        return False
    lvl = _get_level(pPr)
    pct = LEVEL_SIZE_PCT.get(lvl, LEVEL_SIZE_PCT[max(LEVEL_SIZE_PCT)])
    _set_bullet_color(pPr, NAVY_HEX)
    _set_bullet_size_pct(pPr, pct)
    return True


def fix_xml(xml_bytes):
    root = etree.fromstring(xml_bytes)
    changed = 0
    for el in root.iter():
        if '}' not in el.tag:
            continue
        local = _local(el.tag)
        if local == 'pPr' or re.match(r'lvl\d+pPr$', local):
            if _fix_pPr(el):
                changed += 1
    if changed == 0:
        return None, 0
    return etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True
    ), changed


def main():
    print(f'Reading {TARGET.name}')
    with zipfile.ZipFile(TARGET, 'r') as zf:
        members = {n: zf.read(n) for n in zf.namelist()}

    n_files = 0
    n_blocks = 0
    for name in list(members.keys()):
        is_slide = name.startswith('ppt/slides/slide') and name.endswith('.xml')
        is_master = name.startswith('ppt/slideMasters/') and name.endswith('.xml')
        if not (is_slide or is_master):
            continue
        new, n = fix_xml(members[name])
        if new is not None:
            members[name] = new
            n_files += 1
            n_blocks += n

    print(f'  files modified: {n_files}')
    print(f'  bullet pPr blocks updated: {n_blocks}')

    tmp = TARGET.with_suffix('.pptx.tmp')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, data in members.items():
            zf.writestr(name, data)
    tmp.replace(TARGET)
    print(f'Wrote {TARGET.name}')


if __name__ == '__main__':
    main()
