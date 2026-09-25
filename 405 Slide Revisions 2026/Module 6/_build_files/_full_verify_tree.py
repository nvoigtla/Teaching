# -*- coding: utf-8 -*-
"""Create the full slide version, structural check (read-only): every slide of
"Module 6 - Full.pptx" against its source slide as an XML TREE --
attribute order, namespace placement, ids / rIds, cached page numbers and
default-valued attributes that a PowerPoint save drops are ignored;
anything else (geometry, fills, run formatting, text, timing) counts.
"""
import io, contextlib, re, sys
import xml.etree.ElementTree as ET
from pathlib import Path

with contextlib.redirect_stdout(io.StringIO()):
    import _full_verify as V
import _full_inventory as FI

IGN = {"id", "spid", "creationId", "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
       "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed",
       "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link",
       "smtClean", "dirty", "err", "grpId"}
DEFAULTS = {("firstRow", "0"), ("bandRow", "0"), ("lastRow", "0"),
            ("firstCol", "0"), ("lastCol", "0"), ("bandCol", "0")}
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def prep(z, part, copy):
    root = ET.fromstring(z.read(part))
    for el in root.iter():
        if el.tag == A + "fld" and el.get("type") == "slidenum":
            for t in el.iter(A + "t"):
                t.text = "#"
        if copy and el.tag == A + "t" and (el.text or "").startswith("Module 6 · "):
            el.text = re.sub(r"(Video \d+|In Class · Examples) · ", "", el.text)
    return root


def sig(el):
    at = tuple(sorted((k, v) for k, v in el.attrib.items()
                      if k not in IGN and (k, v) not in DEFAULTS))
    kids = [c for c in el if c.tag not in (A + "endParaRPr", P + "extLst")
            and not c.tag.endswith("}creationId")]
    txt = (el.text or "").strip()
    return el.tag, at, txt, kids


def diff(a, b, path, out):
    ta, aa, xa, ka = sig(a)
    tb, ab, xb, kb = sig(b)
    if a.tag == P + "grpSpPr" and (not ka or not kb):     # empty xfrm
        return
    if (ta, aa, xa) != (tb, ab, xb):
        out.append("%s: %s %s %r  !=  %s %s %r" % (path, ta.split("}")[-1], dict(ab), xb,
                                                   tb.split("}")[-1], dict(aa), xa))
        return
    if len(ka) != len(kb):
        out.append("%s/%s: %d vs %d children" % (path, ta.split("}")[-1], len(kb), len(ka)))
        return
    for i, (x, y) in enumerate(zip(ka, kb)):
        diff(x, y, path + "/" + ta.split("}")[-1] + "[%d]" % i, out)


cache, bad = {}, 0
for i, (kind, key) in enumerate(V.exp, 1):
    if kind == "divider":
        continue
    f = V.fin[i - 1]
    r = key["r"] if kind == "taped" else key
    if r in V.pairs:
        p = V.pairs[r]; vp = FI.VDIR / p["deck"]
        cache.setdefault(vp, V.raw(vp)); z, d = cache[vp]; s = d[p["vn"] - 1]
    else:
        z, s = V.rz, V.rev[r - 1]
    out = []
    diff(prep(V.fz, f["part"], kind == "copy"), prep(z, s["part"], kind == "copy"), "", out)
    if out:
        bad += 1
        print("== Full %d (%s R%d): %d differences" % (i, kind, r, len(out)))
        for l in out[:4]:
            print("    ", l[:220])
print("slides with structural differences:", bad)
