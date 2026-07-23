# -*- coding: utf-8 -*-
"""
Insert speaker notes into the deck (OOXML surgery; NOT python-pptx, to preserve
poll tags and null-target links). Regenerates the review Markdown too.

- Content slides + title/quiz/thank-you: full narration from NOTES.
- Backups 105-131: one brief line from BACKUP.
- Slide 17: narration first, then its existing 'More info on Celts' line.
- Skipped: logistics 3-9, agenda (roadmap 10 + dividers), poll slides, blanks.
Creates a new notesSlide part for slides that lack one; edits slide 17 in place.
"""
import os
import shutil
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree as ET

import _notes_data as ND

HERE = Path(__file__).parent
DECK = HERE / "Class 1 - Revised.pptx"
MD = HERE / "Speaker Notes - Class 1.md"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
RT_NOTES = f"{R}/notesSlide"
RT_NOTESMASTER = f"{R}/notesMaster"
RT_SLIDE = f"{R}/slide"
NOTES_CT = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
CHROME = {"TextBox 2", "TextBox 3", "TextBox 9", "TextBox 10", "TextBox 11"}


def q(ns, t):
    return f"{{{ns}}}{t}"


def ser(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def paras_xml(text):
    """Turn a note string into <a:p> paragraphs; bold the 'ANECDOTE —' label and
    put a blank line before an anecdote so it stands out."""
    out = []
    for para in [p for p in text.split("\n\n") if p.strip()]:
        if para.startswith("ANECDOTE"):
            out.append('<a:p/>')  # blank separator line
            # split "ANECDOTE — rest"
            marker, _, rest = para.partition(" ")
            # keep the em-dash with the marker: label = up to first space after dash
            # simpler: bold everything up to and including the dash
            if "—" in para:
                label, _, body = para.partition("—")
                label = label + "—"
                out.append(
                    f'<a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>{escape(label)}</a:t></a:r>'
                    f'<a:r><a:rPr lang="en-US"/><a:t>{escape(body)}</a:t></a:r></a:p>')
            else:
                out.append(f'<a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>{escape(para)}</a:t></a:r></a:p>')
        else:
            out.append(f'<a:p><a:r><a:rPr lang="en-US"/><a:t>{escape(para)}</a:t></a:r></a:p>')
    return "".join(out)


def notes_xml(body_paras):
    return (
        f'<p:notes xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"><p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>'
        '<p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
        '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr>'
        f'<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{body_paras}</p:txBody></p:sp>'
        '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr><p:ph type="sldNum" sz="quarter" idx="5"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
        '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>')


def main(apply):
    z = zipfile.ZipFile(DECK)
    data = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = ET.fromstring(data["ppt/presentation.xml"])
    rid2t = {r.get("Id"): r.get("Target")
             for r in ET.fromstring(data["ppt/_rels/presentation.xml.rels"])}
    order = [os.path.basename(rid2t[s.get(q(R, "id"))]) for s in pres.find(q(P, "sldIdLst"))]

    # titles/sections for the review markdown
    meta = {}
    for disp, part in enumerate(order, 1):
        root = ET.fromstring(data[f"ppt/slides/{part}"])
        sect = title = ""
        for sp in root.iter(q(P, "sp")):
            nv = sp.find(".//" + q(P, "cNvPr"))
            name = nv.get("name") if nv is not None else ""
            txt = " ".join(t.text or "" for t in sp.iter(q(A, "t"))).strip()
            if name == "TextBox 2":
                sect = txt
            elif name == "TextBox 3":
                title = txt
        meta[disp] = (sect, title)

    ct = ET.fromstring(data["[Content_Types].xml"])
    next_ns = max([int(os.path.basename(n)[len("notesSlide"):-4])
                   for n in data if n.startswith("ppt/notesSlides/notesSlide") and n.endswith(".xml")] + [0]) + 1

    def existing_notes_part(part):
        rp = f"ppt/slides/_rels/{part}.rels"
        for r in ET.fromstring(data[rp]):
            if r.get("Type") == RT_NOTES:
                return os.path.normpath("ppt/slides/" + r.get("Target")).replace(os.sep, "/")
        return None

    written = []
    targets = dict(ND.NOTES)
    targets.update(ND.BACKUP)
    for disp in sorted(targets):
        part = order[disp - 1]
        text = targets[disp]
        if not text.strip():
            continue
        if disp == 17:
            body = text + "\n\n" + ND.SLIDE17_KEEP
        else:
            body = text
        if disp in ND.ANECDOTE_SOURCES:
            body = body + "\n\nSource: " + ND.ANECDOTE_SOURCES[disp]
        body_paras = paras_xml(body)

        np = existing_notes_part(part)
        if np is not None:
            # edit in place (slide 17): replace the body placeholder txBody
            nroot = ET.fromstring(data[np])
            for sp in nroot.iter(q(P, "sp")):
                ph = sp.find(".//" + q(P, "ph"))
                if ph is not None and ph.get("type") == "body":
                    old = sp.find(q(P, "txBody"))
                    sp.remove(old)
                    sp.append(ET.fromstring(f'<p:txBody xmlns:a="{A}" xmlns:p="{P}">'
                                            f'<a:bodyPr/><a:lstStyle/>{body_paras}</p:txBody>'))
                    break
            data[np] = ser(nroot)
            written.append((disp, "edited"))
            continue

        # create a fresh notesSlide
        k = next_ns
        next_ns += 1
        nspart = f"ppt/notesSlides/notesSlide{k}.xml"
        data[nspart] = notes_xml(body_paras).encode("utf-8")
        nrels = ET.Element(q(PKG, "Relationships"))
        r1 = ET.SubElement(nrels, q(PKG, "Relationship"))
        r1.set("Id", "rId1"); r1.set("Type", RT_NOTESMASTER)
        r1.set("Target", "../notesMasters/notesMaster1.xml")
        r2 = ET.SubElement(nrels, q(PKG, "Relationship"))
        r2.set("Id", "rId2"); r2.set("Type", RT_SLIDE)
        r2.set("Target", f"../slides/{part}")
        data[f"ppt/notesSlides/_rels/notesSlide{k}.xml.rels"] = ser(nrels)
        ov = ET.SubElement(ct, q(CT, "Override"))
        ov.set("PartName", f"/ppt/notesSlides/notesSlide{k}.xml")
        ov.set("ContentType", NOTES_CT)
        # add rel from the slide to the notesSlide
        srp = f"ppt/slides/_rels/{part}.rels"
        srels = ET.fromstring(data[srp])
        used = {r.get("Id") for r in srels}
        j = 1
        while f"rId{j}" in used:
            j += 1
        rr = ET.SubElement(srels, q(PKG, "Relationship"))
        rr.set("Id", f"rId{j}"); rr.set("Type", RT_NOTES)
        rr.set("Target", f"../notesSlides/notesSlide{k}.xml")
        data[srp] = ser(srels)
        written.append((disp, f"new ns{k}"))

    data["[Content_Types].xml"] = ser(ct)

    # ---- regenerate review markdown ----
    skip_reason = {}
    for n in ND.SKIP_LOGISTICS: skip_reason[n] = "logistics — no note (per instruction)"
    for n in ND.SKIP_AGENDA: skip_reason[n] = "agenda/divider — no note (per instruction)"
    for n in ND.SKIP_POLL: skip_reason[n] = "poll slide — PollEverywhere note preserved"
    for n in ND.SKIP_BLANK: skip_reason[n] = "blank slide — no note"
    lines = ["# Speaker Notes — Class 1 (The Italian Economy)", "",
             "Teleprompter-style narration for a general reader. Anecdotes ~1 in 3, "
             "marked `ANECDOTE —`. Logistics and agenda slides carry no notes; poll "
             "slides keep their protected notes; backups get one brief line.", ""]
    for disp in range(1, len(order) + 1):
        sect, title = meta[disp]
        head = f"### Slide {disp}" + (f" — {title}" if title else (f" — {sect}" if sect else ""))
        lines.append(head)
        if disp in ND.NOTES and ND.NOTES[disp].strip():
            note = ND.NOTES[disp]
            if disp in ND.ANECDOTE_SOURCES:
                note = note + "\n\nSource: " + ND.ANECDOTE_SOURCES[disp]
            lines.append(note)
            if disp == 17:
                lines.append("")
                lines.append(f"*(existing note preserved: {ND.SLIDE17_KEEP})*")
        elif disp in ND.BACKUP:
            lines.append(ND.BACKUP[disp])
        elif disp in skip_reason:
            lines.append(f"*[{skip_reason[disp]}]*")
        else:
            lines.append("*[no note]*")
        lines.append("")
    MD.write_text("\n".join(lines), encoding="utf-8")

    print(f"notes prepared for {len(written)} slides "
          f"({sum(1 for _,s in written if s=='edited')} edited, "
          f"{sum(1 for _,s in written if s.startswith('new'))} new)")
    if apply:
        tmp = DECK.with_suffix(".pptx.tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for name, blob in data.items():
                out.writestr(name, blob if isinstance(blob, bytes) else blob)
        with zipfile.ZipFile(tmp) as chk:
            assert chk.testzip() is None
        shutil.move(str(tmp), str(DECK))
        print("APPLIED notes to deck")
    else:
        print("DRY RUN (markdown regenerated); pass 'apply' to write the deck")


if __name__ == "__main__":
    main(len(sys.argv) > 1 and sys.argv[1] == "apply")
