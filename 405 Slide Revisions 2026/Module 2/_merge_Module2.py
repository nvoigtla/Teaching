# -*- coding: utf-8 -*-
"""Build "Module 2 - Revised.pptx" — the FULL module: the in-class deck
followed by the video part.

The module's agenda decides the order.  Items 1 (the law of demand) and
2 (elasticities) are taught in class; items 3 / 3a / 3b (demand and
revenue, elasticity and revenue, marginal revenue) are Videos 1+2 and
item 4 (demand estimation) is Video 3, all of them post-work.  So the
merged deck is:

    In-Class 1-68   the two in-class agenda items
    In-Class 69     cheat sheet
    In-Class 70-71  the post-work pointers ("next: Videos 1+2 / Video 3")
    Video    1-45   the video part itself, Videos 1, 2 and 3 in order

Both source decks are built by the same helper layer, so they share one
theme, one slide master and one layout (slideLayout7) byte for byte —
verified before this script was written.  The copied video slides are
therefore pointed at the BASE deck's layout and no master/theme parts
are duplicated.

Pure zip + ElementTree surgery, the way _splice_media.py works: a
python-pptx round-trip would strip the PollEverywhere `tags` rels and
the NULL video rels.  Everything the video deck carries comes along
verbatim — slide XML, <p:timing> builds, groups, media, notes and the
poll payload.

RED SOLUTIONS.  The standalone "Module 2 - Video Part Revised.pptx"
reproduces Nico's three recorded video decks slide for slide, so it
stays as taped.  This deck is a teaching artifact, so its video part is
rebuilt with `--red-solutions` (see _build_Module2Video.RED_SOLUTIONS),
which puts the final answer on Videos 3's solution slides in dark red —
the same rule the in-class solution slides follow.

    python _merge_Module2.py              # full run (rebuilds the video
                                          # part with red solutions)
    python _merge_Module2.py --no-rebuild # merge the video deck as-is
"""
import re
import shutil
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).parent
BASE = HERE / "Module 2 - In Class Revised.pptx"
VIDEO = HERE / "Module 2 - Video Part Revised.pptx"
# the "_red" suffix matters: _group_pass.py strips it to find the
# video deck's manual-group and spliced-slide tables
VIDEO_TMP = HERE / "Module 2 - Video Part Revised_red.pptx"
OUT = HERE / "Module 2 - Revised.pptx"

NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
R_ID = '{%s}id' % NS_R

RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
CT_SLIDE = ('application/vnd.openxmlformats-officedocument.'
            'presentationml.slide+xml')
CT_NOTES = ('application/vnd.openxmlformats-officedocument.'
            'presentationml.notesSlide+xml')
CT_TAGS = ('application/vnd.openxmlformats-officedocument.'
           'presentationml.tags+xml')
CT_CHART = ('application/vnd.openxmlformats-officedocument.'
            'drawingml.chart+xml')
CT_DRAWING = ('application/vnd.openxmlformats-officedocument.'
              'drawingml.chartshapes+xml')
CT_BY_EXT = {
    'png': 'image/png', 'jpeg': 'image/jpeg', 'jpg': 'image/jpg',
    'gif': 'image/gif', 'emf': 'image/x-emf', 'wmf': 'image/x-wmf',
    'bmp': 'image/bmp', 'tiff': 'image/tiff', 'svg': 'image/svg+xml',
    'mp4': 'video/mp4', 'm4v': 'video/mp4', 'wav': 'audio/wav',
    'xlsx': ('application/vnd.openxmlformats-officedocument.'
             'spreadsheetml.sheet'),
    'vml': 'application/vnd.openxmlformats-officedocument.vmlDrawing',
    'bin': 'application/vnd.openxmlformats-officedocument.'
           'oleObject',
}
CT_BY_FOLDER = {
    'tags': CT_TAGS, 'charts': CT_CHART, 'drawings': CT_DRAWING,
}


def display_to_part(z):
    pres = ET.fromstring(z.read('ppt/presentation.xml'))
    rels = ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    relmap = {r.get('Id'): r.get('Target') for r in rels}
    sldlst = pres.find('{%s}sldIdLst' % NS_P)
    return ['ppt/' + relmap[s.get(R_ID)].lstrip('/') for s in sldlst]


def rels_name_for(part):
    d, base = part.rsplit('/', 1)
    return '%s/_rels/%s.rels' % (d, base)


class Merger(object):
    def __init__(self, base_path, video_path):
        with zipfile.ZipFile(base_path) as z:
            self.items = {n: z.read(n) for n in z.namelist()}
        self.src = zipfile.ZipFile(video_path)
        self.src_names = set(self.src.namelist())
        self.new_parts = {}
        self.ct_overrides = []
        self.ct_defaults = {}
        self.copied = {}            # src part name -> new part name
        self.notes_master = next(
            n for n in self.items
            if re.match(r'ppt/notesMasters/notesMaster\d+\.xml$', n))
        # the layout every base slide uses (both decks share it)
        base_parts = display_to_part(_Zip(self.items))
        first = base_parts[0].split('/')[-1]
        rels = ET.fromstring(self.items['ppt/slides/_rels/%s.rels' % first])
        self.layout_target = next(
            r.get('Target') for r in rels
            if r.get('Type').endswith('slideLayout'))

    # -- part copying ----------------------------------------------------
    def copy_part(self, src_name):
        """Copy a source part (and, recursively, everything its own .rels
        points at) into the merged package under a vd_ prefix."""
        if src_name in self.copied:
            return self.copied[src_name]
        folder = src_name.rsplit('/', 2)[-2]
        leaf = src_name.split('/')[-1]
        new_name = 'ppt/%s/vd_%s' % (folder, leaf)
        i = 1
        while new_name in self.items or new_name in self.new_parts:
            new_name = 'ppt/%s/vd%d_%s' % (folder, i, leaf)
            i += 1
        self.copied[src_name] = new_name
        self.new_parts[new_name] = self.src.read(src_name)
        ext = leaf.rsplit('.', 1)[-1].lower()
        if folder in CT_BY_FOLDER:
            self.ct_overrides.append(('/' + new_name, CT_BY_FOLDER[folder]))
        elif ext in CT_BY_EXT:
            self.ct_defaults[ext] = CT_BY_EXT[ext]
        # dependent rels
        rn = rels_name_for(src_name)
        if rn in self.src_names:
            rels = ET.fromstring(self.src.read(rn))
            for r in rels:
                if r.get('TargetMode') == 'External':
                    continue
                child = 'ppt/' + r.get('Target').replace('../', '')
                child_new = self.copy_part(child)
                r.set('Target', '../%s/%s'
                      % (child_new.rsplit('/', 2)[-2],
                         child_new.split('/')[-1]))
            ET.register_namespace('', NS_REL)
            self.new_parts[rels_name_for(new_name)] = ET.tostring(
                rels, xml_declaration=True, encoding='UTF-8')
        return new_name

    # -- one video slide -------------------------------------------------
    def add_slide(self, s_part, new_disp, idx):
        s_leaf = s_part.split('/')[-1]
        slide_xml = self.src.read(s_part).decode('utf-8')
        s_rels = ET.fromstring(
            self.src.read('ppt/slides/_rels/%s.rels' % s_leaf))

        new_name = 'ppt/slides/vid%02d.xml' % idx
        assert new_name not in self.items

        out = ET.Element('{%s}Relationships' % NS_REL)
        rid_map = {}
        counter = [0]

        def add_rel(rtype, target, mode=None):
            counter[0] += 1
            rid = 'rId%d' % counter[0]
            e = ET.SubElement(out, '{%s}Relationship' % NS_REL)
            e.set('Id', rid)
            e.set('Type', rtype)
            e.set('Target', target)
            if mode:
                e.set('TargetMode', mode)
            return rid

        add_rel(RT + 'slideLayout', self.layout_target)

        for r in s_rels:
            typ, old_id = r.get('Type'), r.get('Id')
            target = r.get('Target')
            suffix = typ.split('/')[-1]
            if suffix == 'slideLayout':
                continue
            if r.get('TargetMode') == 'External':
                rid_map[old_id] = add_rel(typ, target, 'External')
                continue
            src_name = 'ppt/' + target.replace('../', '')
            if suffix == 'notesSlide':
                new_notes = 'ppt/notesSlides/vid%02d_notes.xml' % idx
                self.new_parts[new_notes] = self.src.read(src_name)
                nrels = ET.Element('{%s}Relationships' % NS_REL)
                e = ET.SubElement(nrels, '{%s}Relationship' % NS_REL)
                e.set('Id', 'rId1')
                e.set('Type', RT + 'notesMaster')
                e.set('Target',
                      '../notesMasters/' + self.notes_master.split('/')[-1])
                e = ET.SubElement(nrels, '{%s}Relationship' % NS_REL)
                e.set('Id', 'rId2')
                e.set('Type', RT + 'slide')
                e.set('Target', '../slides/' + new_name.split('/')[-1])
                ET.register_namespace('', NS_REL)
                self.new_parts[rels_name_for(new_notes)] = ET.tostring(
                    nrels, xml_declaration=True, encoding='UTF-8')
                self.ct_overrides.append(('/' + new_notes, CT_NOTES))
                rid_map[old_id] = add_rel(
                    typ, '../notesSlides/' + new_notes.split('/')[-1])
                continue
            copied = self.copy_part(src_name)
            rid_map[old_id] = add_rel(
                typ, '../%s/%s' % (copied.rsplit('/', 2)[-2],
                                   copied.split('/')[-1]))

        # SINGLE-PASS rId remap (a sequential replace clobbers itself
        # whenever the old and new id spaces overlap - see _splice_media)
        slide_xml = re.sub(
            r'"(rId\d+)"',
            lambda m: '"%s"' % rid_map.get(m.group(1), m.group(1)),
            slide_xml)
        # refresh the CACHED digit of the live slide-number field so the
        # deck looks right before PowerPoint recomputes it
        slide_xml = re.sub(
            r'(<a:fld[^>]*type="slidenum"[^>]*>.*?<a:t>)([^<]*)(</a:t>)',
            lambda m: '%s%d%s' % (m.group(1), new_disp, m.group(3)),
            slide_xml, flags=re.S)

        self.new_parts[new_name] = slide_xml.encode('utf-8')
        ET.register_namespace('', NS_REL)
        self.new_parts[rels_name_for(new_name)] = ET.tostring(
            out, xml_declaration=True, encoding='UTF-8')
        self.ct_overrides.append(('/' + new_name, CT_SLIDE))
        return new_name

    # -- package assembly ------------------------------------------------
    def run(self, out_path):
        base_parts = display_to_part(_Zip(self.items))
        src_parts = display_to_part(self.src)
        n_base = len(base_parts)

        added = []
        for i, s_part in enumerate(src_parts, start=1):
            added.append(self.add_slide(s_part, n_base + i, i))

        # presentation.xml.rels — new slide relationships
        prels = self.items['ppt/_rels/presentation.xml.rels'].decode('utf-8')
        used = set(re.findall(r'Id="(rId\d+)"', prels))
        nxt = max(int(x[3:]) for x in used) + 1
        add_xml = []
        rids = []
        for name in added:
            rid = 'rId%d' % nxt
            nxt += 1
            rids.append(rid)
            add_xml.append(
                '<Relationship Id="%s" Type="%sslide" Target="slides/%s"/>'
                % (rid, RT, name.split('/')[-1]))
        prels = prels.replace('</Relationships>',
                              ''.join(add_xml) + '</Relationships>')
        self.items['ppt/_rels/presentation.xml.rels'] = prels.encode('utf-8')

        # presentation.xml — append the sldId entries
        pres = self.items['ppt/presentation.xml'].decode('utf-8')
        ids = [int(x) for x in re.findall(r'<p:sldId id="(\d+)"', pres)]
        nid = max(ids) + 1 if ids else 256
        entries = []
        for rid in rids:
            entries.append('<p:sldId id="%d" r:id="%s"/>' % (nid, rid))
            nid += 1
        assert '</p:sldIdLst>' in pres
        pres = pres.replace('</p:sldIdLst>',
                            ''.join(entries) + '</p:sldIdLst>')
        self.items['ppt/presentation.xml'] = pres.encode('utf-8')

        # [Content_Types].xml
        ct = self.items['[Content_Types].xml'].decode('utf-8')
        have_ext = set(re.findall(r'Default Extension="([^"]+)"', ct))
        have_ov = set(re.findall(r'Override PartName="([^"]+)"', ct))
        ins = []
        for ext, mime in sorted(self.ct_defaults.items()):
            if ext not in have_ext:
                ins.append('<Default Extension="%s" ContentType="%s"/>'
                           % (ext, mime))
        for part, mime in self.ct_overrides:
            if part not in have_ov:
                ins.append('<Override PartName="%s" ContentType="%s"/>'
                           % (part, mime))
        ct = ct.replace('</Types>', ''.join(ins) + '</Types>')
        self.items['[Content_Types].xml'] = ct.encode('utf-8')

        self.items.update(self.new_parts)
        tmp = out_path.with_suffix('.merge_tmp.pptx')
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name, blob in self.items.items():
                zout.writestr(name, blob)
        self.src.close()
        shutil.move(str(tmp), str(out_path))
        print('merged %d + %d = %d slides -> %s'
              % (n_base, len(added), n_base + len(added), out_path.name))


class _Zip(object):
    """Minimal zipfile stand-in so display_to_part can read the in-memory
    part dict of the base package."""

    def __init__(self, items):
        self.items = items

    def read(self, name):
        return self.items[name]


def rebuild_video_with_red():
    """Rebuild the video part with dark-red final solutions, through the
    same four-step pipeline the standalone deck uses."""
    py = sys.executable
    steps = [
        [py, '_build_Module2Video.py', VIDEO_TMP.name, '--red-solutions'],
        [py, '_splice_video.py', VIDEO_TMP.name],
        [py, '_group_pass.py', VIDEO_TMP.name],
        [py, '_animate_video.py', VIDEO_TMP.name, 'all', 'apply'],
    ]
    for cmd in steps:
        print('  $ ' + ' '.join(cmd[1:]))
        subprocess.run(cmd, cwd=str(HERE), check=True,
                       stdout=subprocess.DEVNULL)
    return VIDEO_TMP


def main():
    rebuild = '--no-rebuild' not in sys.argv[1:]
    video = VIDEO
    if rebuild:
        print('rebuilding the video part with dark-red solutions...')
        video = rebuild_video_with_red()
    Merger(BASE, video).run(OUT)
    if rebuild and VIDEO_TMP.exists():
        VIDEO_TMP.unlink()


if __name__ == '__main__':
    main()
