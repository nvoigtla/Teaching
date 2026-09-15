"""Dump raw per-shape text (including OMML math) straight from slide XML.

python-pptx hides shapes wrapped in mc:AlternateContent (which is what a
PowerPoint save does to every OMML-math textbox), so this reads the parts
directly.  Usage: python _dump_raw.py <deck.pptx> <slide numbers...>
"""
import sys, zipfile
from lxml import etree as ET
from pptx import Presentation

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
EMU = 914400.0


def sp_text(sp):
    parts = []
    for el in sp.iter():
        q = ET.QName(el)
        if q.localname == 't' and q.namespace in (A, M):
            parts.append(el.text or '')
        elif q.localname == 'br' and q.namespace == A:
            parts.append(' / ')
        elif q.localname == 'p' and q.namespace == A:
            parts.append(' ⏎ ')
    return ''.join(parts).strip()


def walk(el, depth=0):
    out = []
    for ch in el:
        q = ET.QName(ch)
        if q.localname in ('sp', 'pic', 'graphicFrame') and q.namespace == P:
            sppr = ch.find('.//{%s}xfrm' % A)
            xfrm = sppr.find('{%s}off' % A) if sppr is not None else None
            ext = sppr.find('{%s}ext' % A) if sppr is not None else None
            geo = ''
            if xfrm is not None and ext is not None and xfrm.get('x') and ext.get('cx'):
                geo = (f"[{int(xfrm.get('x'))/EMU:.2f},{int(xfrm.get('y'))/EMU:.2f} "
                       f"{int(ext.get('cx'))/EMU:.2f}x{int(ext.get('cy'))/EMU:.2f}]")
            t = sp_text(ch) if q.localname == 'sp' else ''
            lbl = q.localname.upper()
            out.append("  " * depth + f"- {lbl} {geo}" + (f": {t}" if t else ""))
        elif q.localname == 'grpSp' and q.namespace == P:
            out.append("  " * depth + "- GROUP")
            out += walk(ch, depth + 1)
        elif q.localname in ('AlternateContent',):
            ch2 = ch.find('.//{http://schemas.openxmlformats.org/markup-compatibility/2006}Choice')
            if ch2 is not None:
                out += walk(ch2, depth)
        elif q.localname == 'spTree':
            out += walk(ch, depth)
    return out


def main():
    deck = sys.argv[1]
    nums = [int(n) for n in sys.argv[2:]]
    prs = Presentation(deck)
    order = [str(s.part.partname).lstrip('/') for s in prs.slides]
    z = zipfile.ZipFile(deck)
    for n in nums:
        x = ET.fromstring(z.read(order[n - 1]))
        tree = x.find('.//{%s}cSld/{%s}spTree' % (P, P))
        print(f"===== SLIDE {n} =====")
        print("\n".join(walk(tree)))
        print()


if __name__ == '__main__':
    main()
