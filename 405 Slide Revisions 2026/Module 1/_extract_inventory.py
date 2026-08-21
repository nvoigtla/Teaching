# Inventory extractor for Module 1 source decks (read-only; never writes .pptx)
import zipfile, re, sys, os
from lxml import etree

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}

FOLDER = r"c:\Users\nvoigtla\Claude Code\Teaching\405 Slide Revisions 2026\Module 1"
DECKS = [
    "Module 1 - In Class.pptx",
    "Module 1 - MW.pptx",
    "Module 1 - Video 1.pptx",
    "Module 1 - Video 2.pptx",
    "Module 1 - Video 3.pptx",
    "Module 1 - Video 4.pptx",
]

def slide_order(z):
    """Return list of slide part names in display order via sldIdLst."""
    pres = etree.fromstring(z.read('ppt/presentation.xml'))
    rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    rmap = {rel.get('Id'): rel.get('Target')
            for rel in rels.iter('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')}
    order = []
    for sid in pres.iter('{%s}sldId' % NS['p']):
        rid = sid.get('{%s}id' % NS['r'])
        tgt = rmap[rid]
        order.append('ppt/' + tgt.lstrip('/').replace('../', ''))
    # slide size
    ext = pres.find('.//{%s}sldSz' % NS['p'])
    cx, cy = int(ext.get('cx')), int(ext.get('cy'))
    return order, (cx / 914400.0, cy / 914400.0)

def shape_texts(sp_tree):
    """Extract text per top-level shape (incl. group children flattened), with paragraph structure."""
    out = []
    spTree = sp_tree.find('.//{%s}cSld/{%s}spTree' % (NS['p'], NS['p']))
    for sp in spTree:
        tag = etree.QName(sp).localname
        if tag not in ('sp', 'grpSp', 'graphicFrame', 'pic'):
            continue
        entry = {'kind': tag, 'name': '', 'paras': [], 'is_title': False, 'media': None}
        nv = sp.find('.//{%s}cNvPr' % NS['p'])
        if nv is not None:
            entry['name'] = nv.get('name', '')
        ph = sp.find('.//{%s}ph' % NS['p'])
        if ph is not None and ph.get('type') in ('title', 'ctrTitle'):
            entry['is_title'] = True
        if tag == 'pic':
            entry['media'] = 'picture'
            vid = sp.find('.//{%s}videoFile' % NS['a'])
            if vid is not None:
                entry['media'] = 'VIDEO'
        if tag == 'graphicFrame':
            g = sp.find('.//{%s}graphicData' % NS['a'])
            uri = g.get('uri') if g is not None else ''
            if 'table' in uri: entry['media'] = 'TABLE'
            elif 'chart' in uri: entry['media'] = 'CHART'
            elif 'ole' in uri.lower(): entry['media'] = 'OLE'
            else: entry['media'] = uri.rsplit('/', 1)[-1]
        for para in sp.iter('{%s}p' % NS['a']):
            runs = []
            for node in para.iter():
                t = etree.QName(node).localname
                if t == 't':  # a:t and m:t both localname 't'
                    runs.append(node.text or '')
            lvl = 0
            pPr = para.find('{%s}pPr' % NS['a'])
            if pPr is not None and pPr.get('lvl'):
                lvl = int(pPr.get('lvl'))
            txt = ''.join(runs).strip()
            if txt:
                entry['paras'].append((lvl, txt))
        # table cell text comes through the a:p iteration above as well
        out.append(entry)
    return out

def notes_text(z, slide_part):
    rel_path = slide_part.replace('slides/', 'slides/_rels/') + '.rels'
    try:
        rels = etree.fromstring(z.read(rel_path))
    except KeyError:
        return None, []
    notes = None
    extras = []
    for rel in rels.iter('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        t = rel.get('Type')
        if t.endswith('/notesSlide'):
            notes = 'ppt/' + rel.get('Target').replace('../', '')
        elif t.endswith('/tags'):
            extras.append('TAGS(poll?)')
        elif t.endswith('/video') or t.endswith('/media'):
            extras.append('MEDIA:' + os.path.basename(rel.get('Target')))
        elif t.endswith('/hyperlink'):
            extras.append('LINK:' + (rel.get('Target') or '')[:80])
    if notes is None:
        return None, extras
    tree = etree.fromstring(z.read(notes))
    paras = []
    for sp in tree.iter('{%s}sp' % NS['p']):
        ph = sp.find('.//{%s}ph' % NS['p'])
        if ph is not None and ph.get('type') == 'body':
            for para in sp.iter('{%s}p' % NS['a']):
                txt = ''.join(n.text or '' for n in para.iter() if etree.QName(n).localname == 't').strip()
                if txt:
                    paras.append(txt)
    return '\n'.join(paras), extras

def has_timing(z, slide_part):
    xml = z.read(slide_part)
    return b'<p:timing>' in xml

def hidden(z, slide_part):
    root = etree.fromstring(z.read(slide_part))
    return root.get('show') == '0'

def main():
    out = []
    for deck in DECKS:
        path = os.path.join(FOLDER, deck)
        z = zipfile.ZipFile(path)
        order, (w, h) = slide_order(z)
        out.append(f"\n\n# ===== {deck} — {len(order)} slides, {w:.2f} x {h:.2f} in =====\n")
        for i, part in enumerate(order, 1):
            tree = etree.fromstring(z.read(part))
            shapes = shape_texts(tree)
            title = next((('; '.join(t for _, t in s['paras'])) for s in shapes if s['is_title']), '(no title placeholder)')
            hid = ' [HIDDEN]' if hidden(z, part) else ''
            anim = ' [ANIM]' if has_timing(z, part) else ''
            notes, extras = notes_text(z, part)
            npics = sum(1 for s in shapes if s['media'] == 'picture')
            media_bits = [s['media'] for s in shapes if s['media'] and s['media'] != 'picture']
            mflags = (f" pics:{npics}" if npics else '') + (' ' + ','.join(media_bits) if media_bits else '')
            extras_s = (' | ' + '; '.join(extras)) if extras else ''
            out.append(f"\n## S{i} ({os.path.basename(part)}){hid}{anim} — {title}{mflags}{extras_s}\n")
            for s in shapes:
                if s['is_title'] or not s['paras']:
                    continue
                body = ' / '.join(('>' * lvl) + t for lvl, t in s['paras'])
                kind = s['media'] or s['kind']
                out.append(f"- [{kind}] {body}\n")
            if notes:
                out.append(f"- NOTES: {notes}\n")
        z.close()
    dest = os.path.join(FOLDER, '_source_inventory.md')
    with open(dest, 'w', encoding='utf-8') as f:
        f.writelines(out)
    print(f"written: {dest}, {os.path.getsize(dest)} bytes")
    # compact stdout summary: deck, slides, size
    print("done")

if __name__ == '__main__':
    main()
