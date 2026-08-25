# Member-level geometry+text diff of selected display slides between the
# canonical deck (with Nico's hand-edits) and the fresh side-path build.
import sys
import zipfile
from lxml import etree as ET

import os
HERE = os.path.dirname(os.path.abspath(__file__))
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
EMU = 914400.0

# normalize PowerPoint math-italic codepoints to plain ASCII
MATH_OFFSETS = [(0x1D434, 0x1D44D, ord('A')), (0x1D44E, 0x1D467, ord('a'))]


def norm(t):
    out = []
    for ch in t:
        o = ord(ch)
        for lo, hi, base in MATH_OFFSETS:
            if lo <= o <= hi:
                ch = chr(base + o - lo)
                break
        out.append(ch)
    return " ".join("".join(out).split())


def q(ns, t):
    return "{%s}%s" % (ns, t)


def slide_part(z, disp):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    rmap = {r.get("Id"): r.get("Target") for r in rels}
    sld = pres.find(q(P, "sldIdLst"))[disp - 1]
    return "ppt/" + rmap[sld.get(q(R, "id"))].lstrip("/")


def decode(el, ox=0.0, oy=0.0, sx=1.0, sy=1.0, out=None, depth=0):
    """Decode shapes to rendered inches, descending into groups."""
    if out is None:
        out = []
    for c in el:
        tag = ET.QName(c).localname
        if tag == "AlternateContent":
            ch = c.find("{http://schemas.openxmlformats.org/markup-compatibility/2006}Choice")
            if ch is not None and len(ch):
                decode(ch, ox, oy, sx, sy, out, depth)
            continue
        if tag not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
            continue
        if tag == "graphicFrame":
            xf = c.find(q(P, "xfrm"))
        else:
            sppr = c.find(q(P, "spPr")) or c.find(q(P, "grpSpPr"))
            xf = sppr.find(q(A, "xfrm")) if sppr is not None else None
        x = y = w = h = 0.0
        if xf is not None:
            off = xf.find(q(A, "off")); ext = xf.find(q(A, "ext"))
            if off is not None and ext is not None:
                x = ox + (int(off.get("x"))) * sx
                y = oy + (int(off.get("y"))) * sy
                w = int(ext.get("cx")) * sx
                h = int(ext.get("cy")) * sy
        if tag == "grpSp":
            cho = xf.find(q(A, "chOff")); che = xf.find(q(A, "chExt"))
            csx = w / int(che.get("cx")) if che is not None and int(che.get("cx")) else sx
            csy = h / int(che.get("cy")) if che is not None and int(che.get("cy")) else sy
            cox = x - int(cho.get("x")) * csx if cho is not None else x
            coy = y - int(cho.get("y")) * csy if cho is not None else y
            out.append(("grp", depth, x / EMU, y / EMU, w / EMU, h / EMU, "", ""))
            decode(c, cox, coy, csx, csy, out, depth + 1)
            continue
        txt = norm("".join(t.text or "" for t in c.iter(q(A, "t"))))
        runs = []
        for r in c.iter(q(A, "rPr")):
            sz = r.get("sz"); b = r.get("b"); i = r.get("i")
            clr = r.find(".//" + q(A, "srgbClr"))
            runs.append("%s%s%s%s" % (
                sz or "-", "b" if b == "1" else "", "i" if i == "1" else "",
                ("#" + clr.get("val")) if clr is not None else ""))
        # 2026-08-25: keep the FULL text.  Truncating at 90 chars hid a
        # real rewrite (video slide 8 gained four answer options whose
        # first 90 characters were unchanged), so callers slice for
        # display instead.
        out.append((tag, depth, x / EMU, y / EMU, w / EMU, h / EMU,
                    txt, ",".join(runs[:8])))
    return out


def notes_text(z, part):
    rp = part.replace("slides/", "slides/_rels/") + ".rels"
    try:
        rels = ET.fromstring(z.read(rp))
    except KeyError:
        return ""
    for r in rels.iter(q(REL, "Relationship")):
        if r.get("Type").endswith("/notesSlide"):
            np = "ppt/" + r.get("Target").replace("../", "")
            tree = ET.fromstring(z.read(np))
            outp = []
            for sp in tree.iter(q(P, "sp")):
                ph = sp.find(".//" + q(P, "ph"))
                if ph is not None and ph.get("type") == "body":
                    for para in sp.iter(q(A, "p")):
                        t = "".join(n.text or "" for n in para.iter(q(A, "t")))
                        outp.append(t)
            return "\n".join(outp)
    return ""


def dump(deck, disp):
    z = zipfile.ZipFile(deck)
    part = slide_part(z, disp)
    tree = ET.fromstring(z.read(part))
    spTree = tree.find(".//" + q(P, "cSld") + "/" + q(P, "spTree"))
    shapes = decode(spTree)
    notes = notes_text(z, part)
    z.close()
    return shapes, notes


def main():
    disps = [int(x) for x in sys.argv[1:]] or [1, 6]
    can = HERE + r"\Module 2 - In Class Revised.pptx"
    test = HERE + r"\Module 2 - In Class Revised_test.pptx"
    for d in disps:
        s1, n1 = dump(can, d)
        s2, n2 = dump(test, d)
        print("=" * 30, "DISPLAY", d, "=" * 30)
        print("--- canonical (hand-edited): %d shapes" % len(s1))
        for s in s1:
            print("  %s%-4s (%6.2f,%5.2f) %5.2fx%4.2f | %s | %s"
                  % ("  " * s[1], s[0], s[2], s[3], s[4], s[5], s[6], s[7]))
        print("--- fresh build: %d shapes" % len(s2))
        for s in s2:
            print("  %s%-4s (%6.2f,%5.2f) %5.2fx%4.2f | %s | %s"
                  % ("  " * s[1], s[0], s[2], s[3], s[4], s[5], s[6], s[7]))
        if n1 != n2:
            print("--- NOTES DIFFER; canonical notes:")
            print(n1[:1500])
            print("--- build notes:")
            print(n2[:1500])
        else:
            print("--- notes identical")


if __name__ == "__main__":
    main()
