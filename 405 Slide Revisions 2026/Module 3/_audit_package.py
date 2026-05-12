"""Audit a .pptx package for consistency issues."""

import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS_CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'
NS_PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'

target = sys.argv[1] if len(sys.argv) > 1 else 'Module 3_NEW_draft.pptx'

with zipfile.ZipFile(target, 'r') as zf:
    names = set(zf.namelist())
    print(f'Files in zip: {len(names)}')

    # 1) Parts declared in [Content_Types].xml as Override
    ct = ET.fromstring(zf.read('[Content_Types].xml'))
    declared = []
    for o in ct.findall(f'{NS_CT}Override'):
        declared.append(o.get('PartName', '').lstrip('/'))
    declared_set = set(declared)
    print(f'Declared in [Content_Types]: {len(declared)}')

    missing = [p for p in declared if p not in names]
    print(f'Declared but missing in zip: {len(missing)}')
    for p in missing[:20]:
        print(f'   MISSING: {p}')

    extras = [n for n in names
              if n.endswith('.xml')
              and n not in declared_set
              and not n.endswith('.rels')
              and n != '[Content_Types].xml']
    print(f'XML files in zip but not declared in [Content_Types]: {len(extras)}')
    for p in extras[:20]:
        print(f'   EXTRA: {p}')

    # 2) Check all .rels files for broken targets
    print()
    print('--- Checking .rels files for broken/unparseable targets ---')
    broken = 0
    for n in sorted(names):
        if not n.endswith('.rels'):
            continue
        try:
            root = ET.fromstring(zf.read(n))
        except Exception as e:
            print(f'   PARSE FAIL: {n}: {e}')
            continue
        for rel in root.findall(f'{NS_PKG}Relationship'):
            target_attr = rel.get('Target', '')
            mode = rel.get('TargetMode', '')
            if mode == 'External':
                continue
            if target_attr.startswith('http://') or target_attr.startswith('https://'):
                continue
            # Resolve relative path
            if n == '_rels/.rels':
                resolved = target_attr.lstrip('/')
            else:
                base = n[:n.rindex('/_rels/')]
                combined = base + '/' + target_attr
                parts = []
                for seg in combined.replace('\\', '/').split('/'):
                    if seg == '..':
                        if parts:
                            parts.pop()
                    elif seg and seg != '.':
                        parts.append(seg)
                resolved = '/'.join(parts)
            if resolved not in names:
                if broken < 20:
                    print(f'   BROKEN: in {n}: rId={rel.get("Id")} target={target_attr!r} -> resolved {resolved!r}')
                broken += 1
    print(f'Total broken rels: {broken}')
