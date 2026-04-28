#!/usr/bin/env python3
"""Extract a header from a .docx file and add it to the pool."""

import argparse
import json
import sys
import zipfile
from pathlib import Path

from lxml import etree

POOL_DIR = Path(__file__).parent / 'pool'

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL_TYPE_HEADER = f'{R}/header'
REL_TYPE_IMAGE = f'{R}/image'


def extract_header(docx_path: Path, name: str, index: int = 0):
    pool_entry = POOL_DIR / name
    if pool_entry.exists():
        print(f"Pool entry '{name}' already exists. Remove it first or choose a different name.")
        sys.exit(1)

    with zipfile.ZipFile(docx_path) as z:
        all_names = set(z.namelist())

        doc_rels_root = etree.fromstring(z.read('word/_rels/document.xml.rels'))
        header_rels = [r for r in doc_rels_root if r.get('Type') == REL_TYPE_HEADER]

        if not header_rels:
            print("No header found in document.")
            sys.exit(1)

        if index >= len(header_rels):
            print(f"Index {index} out of range — {len(header_rels)} header(s) found.")
            sys.exit(1)

        header_target = header_rels[index].get('Target')
        header_xml_bytes = z.read(f'word/{header_target}')

        # Load header's own rels (images etc.)
        header_rels_key = f'word/_rels/{header_target}.rels'
        header_own_rels_bytes = None
        image_map = {}  # rel_id -> (zip_path, filename)

        if header_rels_key in all_names:
            header_own_rels_bytes = z.read(header_rels_key)
            own_rels_root = etree.fromstring(header_own_rels_bytes)
            for rel in own_rels_root:
                if rel.get('Type') == REL_TYPE_IMAGE:
                    target = rel.get('Target', '')
                    zip_path = 'word/' + target.lstrip('../')
                    image_map[rel.get('Id')] = (zip_path, Path(target).name)

        images = {}
        for rid, (zip_path, filename) in image_map.items():
            if zip_path in all_names:
                images[filename] = z.read(zip_path)

    pool_entry.mkdir(parents=True)
    (pool_entry / 'images').mkdir()
    (pool_entry / 'header.xml').write_bytes(header_xml_bytes)

    if header_own_rels_bytes:
        (pool_entry / 'header.xml.rels').write_bytes(header_own_rels_bytes)

    for filename, data in images.items():
        (pool_entry / 'images' / filename).write_bytes(data)

    manifest = {
        'name': name,
        'description': '',
        'objects': _infer_objects(header_xml_bytes, image_map),
    }
    (pool_entry / 'manifest.json').write_text(json.dumps(manifest, indent=2))

    print(f"Extracted to pool/{name}/")
    print(f"  header.xml")
    if header_own_rels_bytes:
        print(f"  header.xml.rels")
    if images:
        print(f"  images/: {list(images.keys())}")
    print(f"  manifest.json  ← update description field")


def _infer_objects(header_xml_bytes: bytes, image_map: dict) -> list:
    objects = []
    root = etree.fromstring(header_xml_bytes)

    for para in root.iter(f'{{{W}}}p'):
        texts = [
            t.text for run in para.iter(f'{{{W}}}r')
            for t in [run.find(f'{{{W}}}t')]
            if t is not None and t.text
        ]
        if not texts:
            continue

        font, size_pt, bold, align = None, None, False, 'left'

        for run in para.iter(f'{{{W}}}r'):
            rpr = run.find(f'{{{W}}}rPr')
            if rpr is not None:
                f_el = rpr.find(f'{{{W}}}rFonts')
                if f_el is not None:
                    font = f_el.get(f'{{{W}}}ascii') or f_el.get(f'{{{W}}}hAnsi')
                sz = rpr.find(f'{{{W}}}sz')
                if sz is not None:
                    try:
                        size_pt = int(sz.get(f'{{{W}}}val', '24')) // 2
                    except ValueError:
                        pass
                bold = rpr.find(f'{{{W}}}b') is not None
            break

        ppr = para.find(f'{{{W}}}pPr')
        if ppr is not None:
            jc = ppr.find(f'{{{W}}}jc')
            if jc is not None:
                align = jc.get(f'{{{W}}}val', 'left')

        objects.append({
            'type': 'text',
            'value': ''.join(texts),
            'font': font,
            'size_pt': size_pt,
            'bold': bold,
            'align': align,
        })

    for rid, (_, filename) in image_map.items():
        objects.append({'type': 'image', 'value': filename, 'rel_id': rid})

    return objects


def main():
    p = argparse.ArgumentParser(description='Extract header from .docx into pool')
    p.add_argument('--doc', required=True, help='Source .docx file')
    p.add_argument('--name', required=True, help='Pool entry name')
    p.add_argument('--index', type=int, default=0, help='Header index if doc has multiple (default: 0)')
    args = p.parse_args()

    path = Path(args.doc)
    if not path.exists():
        print(f"Not found: {path}")
        sys.exit(1)

    extract_header(path, args.name, args.index)


if __name__ == '__main__':
    main()
