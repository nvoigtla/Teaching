"""Extract every media part from the PG decks, keyed to the slide that uses it.

PG builds most screenshots as picture-FILLED autoshapes (a:blipFill inside
<p:sp>), which python-pptx does not report as pictures -- so walk the raw
OOXML instead.  Build input: keep the output.
"""
import zipfile, re, os
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
out = Path("_source_images"); out.mkdir(exist_ok=True)
rows = []
for tag, deck in (("PG1", "Module 6 - PG - Part 1.pptx"),
                  ("PG2", "Module 6 - PG - Part 2.pptx"),
                  ("PG3", "Module 6 - PG - Part 3.pptx")):
    z = zipfile.ZipFile("Module 6 - PG Slides/" + deck)
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    prels = {rel.get('Id'): rel.get('Target')
             for rel in ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    order = [prels[s.get('{%s}id' % NS['r'])]
             for s in pres.find('p:sldIdLst', NS)]
    for n, target in enumerate(order, 1):
        part = "ppt/" + target.replace("../", "")
        xml = z.read(part).decode("utf-8")
        relp = part.rsplit("/", 1)[0] + "/_rels/" + part.rsplit("/", 1)[1] + ".rels"
        rels = {rel.get('Id'): rel.get('Target')
                for rel in ET.fromstring(z.read(relp))}
        for rid in sorted(set(re.findall(r'r:embed="(rId\d+)"', xml))):
            tgt = rels.get(rid, "")
            if "media" not in tgt:
                continue
            src = "ppt/" + tgt.replace("../", "")
            ext = src.rsplit(".", 1)[-1]
            name = f"{tag}_s{n:02d}_{rid}.{ext}"
            (out / name).write_bytes(z.read(src))
            rows.append((tag, n, name, len(z.read(src))))
L = ["# Assets manifest: PG decks (media pulled from picture fills too)", "",
     "| Deck | Slide | File | Bytes |", "|---|---|---|---|"]
L += ["| " + " | ".join(str(x) for x in r) + " |" for r in rows]
Path("_assets_manifest_PG.md").write_text("\n".join(L), encoding="utf-8")
print(f"{len(rows)} media placements extracted")
