import sys, zipfile, re
from xml.etree import ElementTree as ET

NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def order(z):
    pres=ET.fromstring(z.read('ppt/presentation.xml'))
    rels=ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    rid={r.get('Id'):r.get('Target') for r in rels}
    out=[]
    for sid in pres.find('p:sldIdLst',NS):
        t=rid[sid.get('{%s}id'%NS['r'])]
        out.append('ppt/'+t.replace('../','').lstrip('/') if not t.startswith('ppt/') else t)
    return out

def texts(z,part):
    root=ET.fromstring(z.read(part))
    res=[]
    for sp in root.iter('{%s}sp'%NS['p']):
        tx=''.join(t.text or '' for t in sp.iter('{%s}t'%NS['a']))
        if tx.strip(): res.append(tx.strip())
    return res

for f in sys.argv[1:]:
    z=zipfile.ZipFile(f)
    parts=order(z)
    print('='*80); print(f, len(parts),'slides')
    for i,pt in enumerate(parts,1):
        ts=texts(z,pt)
        print(f'  [{i:3d}] {pt.split("/")[-1]:16s} | ' + ' // '.join(ts[:4])[:150])
