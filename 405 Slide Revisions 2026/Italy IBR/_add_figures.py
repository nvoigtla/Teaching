"""
Add illustrative figures to text-only content slides (High+Medium set, minus 31).
For each slide: narrow the bullets box (TextBox 6) into a left column, place a
rounded + shadowed picture in the right column, add an italic-grey caption below.
In-place OOXML surgery. Run _resize_bullets.py apply afterward to re-fit bullets.

CONFIG captions carry the attribution reported by the image-sourcing agent.
"""
import io
import os
import shutil
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape
from PIL import Image
from lxml import etree as ET

HERE = Path(__file__).parent
IMG = HERE / "Images"
DECK = HERE / "Class 1 - Revised.pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
RT_IMAGE = f"{R}/image"
EMU = 914400
GRAY = "555B66"

# profiles: text column width, image-column centre-x, max image w/h, vertical band
PROF = {
    "BIG":   dict(text_w=7.15, cx=10.40, max_w=4.85, max_h=3.85, band_top=2.00, band_h=4.15),
    "MED":   dict(text_w=7.55, cx=10.55, max_w=4.45, max_h=3.45, band_top=2.10, band_h=3.95),
    "SMALL": dict(text_w=8.25, cx=10.85, max_w=4.05, max_h=2.95, band_top=2.20, band_h=3.75),
    # slide 16 is text-heavy: give the bullets a wider column, a smaller image
    "M16":   dict(text_w=8.55, cx=10.95, max_w=3.80, max_h=3.05, band_top=2.15, band_h=3.90),
}

# slide -> (image filename, profile, caption). Captions carry real attribution.
CONFIG = {
    16: ("etruscan_sarcophagus.jpg", "M16",   "Sarcophagus of the Spouses, Villa Giulia (Sailko, CC BY-SA 4.0)"),
    # slide 20 (Roman Republic) is too text-dense for a side figure — left full-width
    22: ("augustus_primaporta.jpg",  "BIG",   "Augustus of Prima Porta (public domain)"),
    24: ("via_appia.jpg",            "MED",   "The Via Appia Antica (N. Hartmann, CC BY-SA 4.0)"),
    27: ("roman_invasions_map.png",  "BIG",   "Barbarian invasions of the Empire (MapMaster, CC BY-SA 2.5)"),
    33: ("siena_campo.jpg",          "SMALL", "Siena, Piazza del Campo (A. Otrębski, CC BY-SA 4.0)"),
    34: ("good_government.jpg",      "BIG",   "Lorenzetti, ‘Effects of Good Government’, Siena (public domain)"),
    40: ("italy_north_south_map.png", "BIG",  "Italian regions by GDP per capita, 2018 (JJLiu112, CC0)"),
    42: ("florence_duomo.jpg",       "BIG",   "Florence: Brunelleschi’s dome (Morio, CC BY-SA 3.0)"),
    47: ("plague_milan_1630.jpg",    "MED",   "Micco Spadaro, plague in Naples, 1656 (Wellcome Collection, CC BY 4.0)"),
    89: ("autostrada_del_sole.jpg",  "BIG",   "Autostrada del Sole, 1965 (public domain)"),
}

SHADOW = ('<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" rotWithShape="0">'
          '<a:srgbClr val="000000"><a:alpha val="30000"/></a:srgbClr></a:outerShdw></a:effectLst>')


def prep_image(path, ext):
    """Center-crop ultra-wide panoramas to ~1.55:1, downscale, return (bytes, w, h)."""
    im = Image.open(path)
    if im.mode not in ("RGB", "RGBA", "P", "L"):
        im = im.convert("RGB")
    w, h = im.size
    if w / h > 1.8:  # ultra-wide panorama -> crop to 1.55:1 around the centre
        new_w = int(h * 1.55)
        x0 = (w - new_w) // 2
        im = im.crop((x0, 0, x0 + new_w, h))
        w, h = im.size
    longest = max(w, h)
    if longest > 1800:
        s = 1800 / longest
        im = im.resize((int(w * s), int(h * s)), Image.LANCZOS)
        w, h = im.size
    buf = io.BytesIO()
    if ext in ("jpg", "jpeg"):
        if im.mode != "RGB":
            im = im.convert("RGB")
        im.save(buf, "JPEG", quality=88)
    else:
        im.save(buf, "PNG", optimize=True)
    return buf.getvalue(), w, h


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def emu(v):
    return int(round(v * EMU))


def pic_xml(sid, rid, x, y, w, h, name):
    return (f'<p:pic xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}">'
            f'<p:nvPicPr><p:cNvPr id="{sid}" name="{name}"/>'
            f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 6000"/></a:avLst></a:prstGeom>'
            f'{SHADOW}</p:spPr></p:pic>')


def cap_xml(sid, x, y, w, text):
    return (f'<p:sp xmlns:a="{A}" xmlns:p="{P}"><p:nvSpPr><p:cNvPr id="{sid}" name="cap{sid}"/>'
            f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'
            f'<a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(0.3)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="0" rIns="0" tIns="0" bIns="0" anchor="t"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1100" i="1">'
            f'<a:solidFill><a:srgbClr val="{GRAY}"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr>'
            f'<a:t>{escape(text)}</a:t></a:r></a:p></p:txBody></p:sp>')


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target")
             for r in ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    order = [os.path.basename(rid2t[s.get(q(R, "id"))]) for s in pres.find(q(P, "sldIdLst"))]

    ct = ET.fromstring(data["[Content_Types].xml"])
    have_ext = {d.get("Extension") for d in ct if d.tag == q(CT, "Default")}
    media_n = 500  # start well above any existing media index (image*/more*)
    while any(n.startswith(f"ppt/media/fig{media_n}.") for n in data):
        media_n += 1

    for disp, (fname, prof, caption) in CONFIG.items():
        part = order[disp - 1]
        img_path = IMG / fname
        assert img_path.exists(), f"missing {img_path}"
        ext = os.path.splitext(fname)[1].lower().lstrip(".")
        ext_ct = "jpeg" if ext in ("jpg", "jpeg") else ext
        if ext_ct not in have_ext:
            d = ET.SubElement(ct, q(CT, "Default"))
            d.set("Extension", ext_ct)
            d.set("ContentType", f"image/{ext_ct}")
            have_ext.add(ext_ct)

        blob, iw, ih = prep_image(img_path, ext)
        p = PROF[prof]
        # fit within the max box preserving aspect ratio
        w_in, h_in = iw / 96.0, ih / 96.0
        s = min(p["max_w"] / w_in, p["max_h"] / h_in)
        w_in, h_in = w_in * s, h_in * s
        cap_h = 0.30
        band = p["band_h"] - cap_h - 0.06
        img_y = p["band_top"] + (band - h_in) / 2.0
        img_x = p["cx"] - w_in / 2.0
        cap_y = img_y + h_in + 0.05

        root = ET.fromstring(data[f"ppt/slides/{part}"])
        tree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))
        # narrow the bullets box
        for sp in tree.findall(q(P, "sp")):
            if sp.find(".//" + q(P, "cNvPr")).get("name") == "TextBox 6":
                sp.find(".//" + q(A, "ext")).set("cx", str(emu(p["text_w"])))
                break
        # ids
        ids = [int(c.get("id")) for c in root.iter(q(P, "cNvPr"))]
        sid = max(ids) + 1

        # media + rel
        media_n += 1
        media_name = f"ppt/media/fig{media_n}.{ext}"
        data[media_name] = blob
        relpart = f"ppt/slides/_rels/{part}.rels"
        rels = ET.fromstring(data[relpart])
        used = {r.get("Id") for r in rels}
        k = 1
        while f"rId{k}" in used:
            k += 1
        rid = f"rId{k}"
        rr = ET.SubElement(rels, q(PKG, "Relationship"))
        rr.set("Id", rid)
        rr.set("Type", RT_IMAGE)
        rr.set("Target", f"../media/{os.path.basename(media_name)}")
        data[relpart] = ser(rels)

        # place picture before the footer/back chrome so it sits under nothing important
        pic = ET.fromstring(pic_xml(sid, rid, img_x, img_y, w_in, h_in, f"fig{disp}"))
        cap = ET.fromstring(cap_xml(sid + 1, p["cx"] - (w_in + 0.6) / 2.0, cap_y, w_in + 0.6, caption))
        tree.append(pic)
        tree.append(cap)
        data[f"ppt/slides/{part}"] = ser(root)
        print(f"slide {disp}: {fname} {iw}x{ih} -> {w_in:.2f}x{h_in:.2f}in at x{img_x:.2f} y{img_y:.2f}")

    data["[Content_Types].xml"] = ser(ct)
    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print("added figures to", ", ".join(str(d) for d in CONFIG))


if __name__ == "__main__":
    main()
