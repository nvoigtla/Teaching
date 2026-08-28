# -*- coding: utf-8 -*-
import sys
from lxml import etree as ET
sys.path.insert(0,'.')
import _poll_chrome_pass as base
A,P=base.A,base.P
EMU=base.EMU
class Pk(base.Pkg): pass
deck=sys.argv[1]
pkg=Pk(deck)
for i,pn in enumerate(pkg.slides(),1):
    t=pkg.xml(pn); sp=t.find(base.q(P,'cSld')).find(base.q(P,'spTree'))
    tag=''; title=''
    for c in base.shape_kids(sp):
        if ET.QName(c).localname!='sp': continue
        g=base.geom(c); s=base.txt(c)
        if not g or not s: continue
        top=g[1]/EMU; w=g[2]/EMU
        if not tag and top<0.55 and w>4: tag=s[:52]
        elif not title and 0.9<top<2.4 and w>5: title=s[:46]
    print('%3d | %-52s | %s'%(i,tag,title))
