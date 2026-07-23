"""
Update slide 107 (Presentation Topics) from the latest Excel.
 - 6 debate topics in the Topic | Team A | Team B table (drop Immigration &
   'Lavazza in the U.S.'; add Family businesses & Tourism as topics 5-6).
 - Company Option becomes a single full-width block (NO A/B columns) with a
   'one team, no A/B split' note and a row of company-choice pills.
In-place OOXML surgery; back-button link (rId2 -> slide8) untouched.
"""
import zipfile
import shutil
from pathlib import Path
from xml.sax.saxutils import escape
from PIL import ImageFont
from lxml import etree as ET

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
PART = "ppt/slides/slide107.xml"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
EMU = 914400
NAVY, GOLD, GREY, CREAM = "0B2B4E", "E09F3E", "E7E9EC", "FDF6E6"
GRAY = "8A9199"  # readable mid-grey for the pending placeholder
FONT = ImageFont.truetype("C:/Windows/Fonts/calibrib.ttf", 17)  # ~12.5pt bold


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def emu(inch):
    return int(round(inch * EMU))


# ---- the 6 debate topics (faithful to the Excel, deck-concise) --------------
TOPICS = [
    ("1   Will Southern Italy catch up?", "Yes – it will converge with the North", "No – it will stay behind"),
    ("2   Italy’s future in tech & AI", "Italy is lagging far behind", "Italy is already on a good path"),
    ("3   Corruption and trust", "A serious problem, with real economic cost", "Its reputation is overblown"),
    ("4   Italy and the EU / euro", "The EU and euro are good for Italy", "Italy would be better off outside"),
    ("5   Family businesses", "A major competitive advantage", "They hold Italian companies back"),
    ("6   Tourism", "Italy should keep expanding tourism", "Italy already depends on it too much"),
]

# ---- company choices (Excel order); TBD kept as a dim placeholder -----------
COMPANIES = [
    ("1", "Campari Group", False), ("2", "Pirelli", False),
    ("3", "EssilorLuxottica", False), ("4", "Prada Group", False),
    ("5", "Kering", False), ("6", "TBD (Wed 11am)", True),
    ("7", "Inter Milan FC", False), ("8", "Lavazza", False),
    ("9", "Carioca", False), ("10", "Italdesign", False),
    ("11", "LeapFactory", False),
]


def run(text, sz, color, bold=False, italic=False):
    b = ' b="1"' if bold else ""
    i = ' i="1"' if italic else ""
    return (f'<a:r><a:rPr lang="en-US" sz="{sz}"{b}{i}>'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="Calibri"/></a:rPr><a:t>{escape(text)}</a:t></a:r>')


def txbox(sid, x, y, w, h, runs, align="l", anchor="ctr"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="t{sid}"/>'
            f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'
            f'<a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="0" rIns="0" tIns="0" bIns="0" anchor="{anchor}"/>'
            f'<a:lstStyle/><a:p><a:pPr algn="{align}"/>{runs}</a:p></p:txBody></p:sp>')


SHADOW = ('<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" rotWithShape="0">'
          '<a:srgbClr val="000000"><a:alpha val="30000"/></a:srgbClr></a:outerShdw></a:effectLst>')


def roundrect(sid, x, y, w, h, fill, line, line_w_pt, adj, runs=None,
              shadow=False, align="ctr", anchor="ctr", dash=None):
    ln = f'<a:ln w="{int(line_w_pt*12700)}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill>'
    ln += (f'<a:prstDash val="{dash}"/>' if dash else "") + '</a:ln>' if line else '<a:ln><a:noFill/></a:ln>'
    eff = SHADOW if shadow else ""
    body = ""
    if runs is not None:
        body = (f'<p:txBody><a:bodyPr wrap="square" lIns="45720" rIns="45720" tIns="0" bIns="0" anchor="{anchor}"/>'
                f'<a:lstStyle/><a:p><a:pPr algn="{align}"/>{runs}</a:p></p:txBody>')
    else:
        body = '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>'
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="rr{sid}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val {int(adj*100000)}"/></a:avLst></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{ln}{eff}</p:spPr>{body}</p:sp>')


def pill_width(num, name):
    txt = f"{num}  {name}"
    w = FONT.getbbox(txt)[2] / 96.0  # px -> inches
    return w + 0.30  # padding both sides


def main():
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    root = ET.fromstring(data[PART])
    tree = root.find(q(P, "cSld") + "/" + q(P, "spTree"))

    # --- edit / trim the table -------------------------------------------
    gf = tree.find(q(P, "graphicFrame"))
    tbl = gf.find(".//" + q(A, "tbl"))
    rows = tbl.findall(q(A, "tr"))
    for ri, (c0, c1, c2) in enumerate(TOPICS, start=1):
        for ci, val in enumerate((c0, c1, c2)):
            t = rows[ri].findall(q(A, "tc"))[ci].find(".//" + q(A, "t"))
            t.text = val
    for extra in rows[7:]:          # drop old rows 7,8,9
        tbl.remove(extra)
    n_rows = len(tbl.findall(q(A, "tr")))
    table_h = sum(int(tr.get("h")) for tr in tbl.findall(q(A, "tr")))
    gf.find(".//" + q(P, "xfrm") + "/" + q(A, "ext")).set("cy", str(table_h))

    # shrink the backing card to match the table
    for sp in tree.findall(q(P, "sp")):
        if sp.find(".//" + q(P, "cNvPr")).get("name") == "TableBacking":
            sp.find(".//" + q(A, "ext")).set("cy", str(table_h))

    tbl_top = 1.55
    tbl_bottom = tbl_top + table_h / EMU

    # --- company panel (single block, no A/B) ----------------------------
    ids = [int(cnv.get("id")) for cnv in root.iter(q(P, "cNvPr"))]
    nid = max(ids) + 1

    panel_x, panel_w = 0.33, 12.70
    panel_top = tbl_bottom + 0.18
    # lay out the pills (flow into rows)
    px0, pxmax = panel_x + 0.22, panel_x + panel_w - 0.22
    pill_h, gap, rowgap = 0.38, 0.16, 0.12
    prows, cur, curw = [], [], px0
    for num, name, dim in COMPANIES:
        w = pill_width(num, name)
        if curw + w > pxmax and cur:
            prows.append(cur)
            cur, curw = [], px0
        cur.append((num, name, dim, curw, w))
        curw += w + gap
    if cur:
        prows.append(cur)

    pills_top = panel_top + 0.62
    pills_h = len(prows) * pill_h + (len(prows) - 1) * rowgap
    panel_h = (pills_top - panel_top) + pills_h + 0.12

    shapes = []
    # panel background
    shapes.append(roundrect(nid, panel_x, panel_top, panel_w, panel_h,
                            CREAM, NAVY, 1.0, 0.05, shadow=True)); nid += 1
    # gold badge
    shapes.append(roundrect(nid, panel_x + 0.22, panel_top + 0.16, 2.55, 0.44,
                            GOLD, GOLD, 0, 0.5,
                            runs=run("Company Option", 1500, NAVY, bold=True),
                            shadow=True)); nid += 1
    # prompt / 'no A/B' note
    prompt = (run("Pick one company we’ll visit – introduce it, then argue why (not) to invest.  ", 1400, NAVY)
              + run("One team, no A/B split.", 1400, GOLD, bold=True))
    shapes.append(txbox(nid, panel_x + 3.05, panel_top + 0.14, 9.4, 0.48, prompt,
                        align="l", anchor="ctr")); nid += 1
    # pills
    for r_i, prow in enumerate(prows):
        y = pills_top + r_i * (pill_h + rowgap)
        for num, name, dim, x, w in prow:
            fill = "FFFFFF"
            border = GRAY if dim else NAVY
            tcol = GRAY if dim else NAVY
            runs = run(num + "  ", 1250, GRAY if dim else GOLD, bold=True) + run(name, 1250, tcol, italic=dim)
            shapes.append(roundrect(nid, x, y, w, pill_h, fill, border, 1.0, 0.5,
                                    runs=runs, align="ctr",
                                    dash="dash" if dim else None)); nid += 1

    # insert panel+pills just before the BackButton so the button stays clickable on top
    back = None
    for sp in tree.findall(q(P, "sp")):
        if sp.find(".//" + q(P, "cNvPr")).get("name") == "BackButton":
            back = sp
            break
    frag = ET.fromstring(f'<root xmlns:a="{A}" xmlns:p="{P}">' + "".join(shapes) + "</root>")
    for child in frag:
        back.addprevious(child)

    data[PART] = ser(root)
    tmp = DECK.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in data.items():
            out.writestr(name, blob)
    with zipfile.ZipFile(tmp) as chk:
        assert chk.testzip() is None
    shutil.move(str(tmp), str(DECK))
    print(f"updated slide 107: {n_rows-1} topics, panel_h={panel_h:.2f}, pill rows={len(prows)}, "
          f"table_bottom={tbl_bottom:.2f}, panel_bottom={panel_top+panel_h:.2f}")


if __name__ == "__main__":
    main()
