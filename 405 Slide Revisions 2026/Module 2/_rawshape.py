import sys, zipfile
from lxml import etree as ET
import _diff_slides as D
import _dump_runs as R
deck=sys.argv[1]; disp=int(sys.argv[2]); idxs=[int(x) for x in sys.argv[3:]]
sh=R.shapes(deck, disp)
for i in idxs:
    tag,c=sh[i]
    print("=== [%d] %s"%(i,tag))
    print(ET.tostring(c, pretty_print=True).decode())
