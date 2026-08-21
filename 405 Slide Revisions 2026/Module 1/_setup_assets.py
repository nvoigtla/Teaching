# Extract picture assets + run-format dump from the Module 1 source decks.
# Read-only on the .pptx files. Outputs:
#   _source_images/{tag}_s{NN}_{rId}.{ext}   (all embedded pictures)
#   _assets_manifest.md                      (pic positions in rendered inches)
#   _runfmt_dump.md                          (body text with **bold** / *italic*
#                                             / _underline_ / [color] markers)
import os, zipfile
from lxml import etree

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}
def q(t):
    pre, loc = t.split(':')
    return '{%s}%s' % (NS[pre], loc)

FOLDER = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(FOLDER, '_source_images')
os.makedirs(IMG_DIR, exist_ok=True)

DECKS = [
    ('ic', 'Module 1 - In Class.pptx', None),          # all slides
    ('v1', 'Module 1 - Video 1.pptx', None),
    ('v2', 'Module 1 - Video 2.pptx', None),
    ('v3', 'Module 1 - Video 3.pptx', None),
    ('v4', 'Module 1 - Video 4.pptx', None),
    ('mw', 'Module 1 - MW.pptx', {24, 25, 26, 49, 50, 51, 52, 55, 65, 67}),
]

EMU = 914400.0

def slide_order(z):
    pres = etree.fromstring(z.read('ppt/presentation.xml'))
    rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    rmap = {r_.get('Id'): r_.get('Target') for r_ in rels.iter(q('rel:Relationship'))}
    out = []
    for sid in pres.iter(q('p:sldId')):
        out.append('ppt/' + rmap[sid.get(q('r:id'))].lstrip('/').replace('../', ''))
    return out

def walk_pics(el, xform, acc):
    """Collect (pic_element, rendered_l, t, w, h) recursively through groups.
    xform: (offx, offy, scalex, scaley) mapping child coords -> slide EMU."""
    ox, oy, sx, sy = xform
    for child in el:
        tag = etree.QName(child).localname
        if tag == 'pic':
            x = child.find('.//' + q('a:xfrm'))
            if x is None:
                continue
            off = x.find(q('a:off')); ext = x.find(q('a:ext'))
            if off is None or ext is None:
                continue
            l = ox + int(off.get('x')) * sx
            t = oy + int(off.get('y')) * sy
            w = int(ext.get('cx')) * sx
            h = int(ext.get('cy')) * sy
            acc.append((child, l, t, w, h))
        elif tag == 'grpSp':
            gx = child.find(q('p:grpSpPr') + '/' + q('a:xfrm'))
            if gx is None:
                continue
            off = gx.find(q('a:off')); ext = gx.find(q('a:ext'))
            ch_off = gx.find(q('a:chOff')); ch_ext = gx.find(q('a:chExt'))
            if None in (off, ext, ch_off, ch_ext):
                continue
            cw = max(int(ch_ext.get('cx')), 1); ch = max(int(ch_ext.get('cy')), 1)
            nsx = sx * int(ext.get('cx')) / cw
            nsy = sy * int(ext.get('cy')) / ch
            nox = ox + int(off.get('x')) * sx - int(ch_off.get('x')) * nsx
            noy = oy + int(off.get('y')) * sy - int(ch_off.get('y')) * nsy
            walk_pics(child, (nox, noy, nsx, nsy), acc)

def run_marks(rPr):
    pre = suf = ''
    if rPr is not None:
        if rPr.get('b') == '1': pre += '**'; suf = '**' + suf
        if rPr.get('i') == '1': pre += '*'; suf = '*' + suf
        if rPr.get('u') not in (None, 'none'): pre += '_'; suf = '_' + suf
        clr = rPr.find(q('a:solidFill') + '/' + q('a:srgbClr'))
        if clr is not None and clr.get('val') not in ('000000',):
            suf += '[#%s]' % clr.get('val')
    return pre, suf

def dump_runs(sp):
    paras = []
    for txBody in sp.iter(q('p:txBody')):
        for para in txBody.findall(q('a:p')):
            lvl = 0
            pPr = para.find(q('a:pPr'))
            if pPr is not None and pPr.get('lvl'):
                lvl = int(pPr.get('lvl'))
            parts = []
            for node in para.iter():
                loc = etree.QName(node).localname
                if loc == 'r' and node.tag == q('a:r'):
                    t = node.find(q('a:t'))
                    txt = (t.text or '') if t is not None else ''
                    if not txt:
                        continue
                    pre, suf = run_marks(node.find(q('a:rPr')))
                    parts.append(pre + txt + suf)
                elif node.tag == q('a:fld'):
                    parts.append('[fld]')
            if parts:
                paras.append(('>' * lvl) + ''.join(parts))
    return paras

def main():
    manifest = []
    runfmt = []
    for tag, deck, only in DECKS:
        z = zipfile.ZipFile(os.path.join(FOLDER, deck))
        order = slide_order(z)
        manifest.append(f"\n# {deck}\n")
        runfmt.append(f"\n# {deck}\n")
        for i, part in enumerate(order, 1):
            if only is not None and i not in only:
                continue
            tree = etree.fromstring(z.read(part))
            # rels for this slide: rId -> media target
            relpath = part.replace('slides/', 'slides/_rels/') + '.rels'
            rmap = {}
            try:
                rels = etree.fromstring(z.read(relpath))
                for r_ in rels.iter(q('rel:Relationship')):
                    if '/image' in r_.get('Type'):
                        rmap[r_.get('Id')] = r_.get('Target')
            except KeyError:
                pass
            spTree = tree.find(q('p:cSld') + '/' + q('p:spTree'))
            pics = []
            walk_pics(spTree, (0.0, 0.0, 1.0, 1.0), pics)
            if pics:
                manifest.append(f"\n## {tag} S{i}\n")
            for pic, l, t, w, h in pics:
                blip = pic.find('.//' + q('a:blip'))
                if blip is None:
                    continue
                rid = blip.get(q('r:embed'))
                tgt = rmap.get(rid)
                if not tgt:
                    continue
                mpath = 'ppt/' + tgt.replace('../', '')
                ext = os.path.splitext(mpath)[1]
                fname = f"{tag}_s{i:02d}_{rid}{ext}"
                dest = os.path.join(IMG_DIR, fname)
                if not os.path.exists(dest):
                    with open(dest, 'wb') as f:
                        f.write(z.read(mpath))
                nm = pic.find('.//' + q('p:cNvPr'))
                nm_s = nm.get('name') if nm is not None else ''
                manifest.append(
                    f"- {fname}  pos=({l/EMU:.2f},{t/EMU:.2f}) "
                    f"size=({w/EMU:.2f}x{h/EMU:.2f})  name='{nm_s}'\n")
            # run-format dump (top-level sp text incl. group children)
            slide_lines = []
            for sp in spTree.iter(q('p:sp')):
                paras = dump_runs(sp)
                if paras:
                    slide_lines.append(' / '.join(paras))
            if slide_lines:
                runfmt.append(f"\n## {tag} S{i}\n")
                for ln in slide_lines:
                    runfmt.append(f"- {ln}\n")
        z.close()
    with open(os.path.join(FOLDER, '_assets_manifest.md'), 'w', encoding='utf-8') as f:
        f.writelines(manifest)
    with open(os.path.join(FOLDER, '_runfmt_dump.md'), 'w', encoding='utf-8') as f:
        f.writelines(runfmt)
    n = len(os.listdir(IMG_DIR))
    print(f"images: {n} in _source_images/")
    print("wrote _assets_manifest.md, _runfmt_dump.md")

if __name__ == '__main__':
    main()
