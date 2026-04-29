#!/usr/bin/env python3
"""Core header injection logic — raw XML approach."""

import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

POOL_DIR = Path(__file__).parent / 'pool'

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
REL_TYPE_HEADER = f'{R}/header'
REL_TYPE_IMAGE = f'{R}/image'
CT_HEADER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
CT_IMAGES = {
    'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
    'gif': 'image/gif', 'wmf': 'image/x-wmf', 'emf': 'image/x-emf',
}


def inject(docx_path: Path, header_name: str, out_path: Path, replace: bool = False,
           font: str = None, size_pt: int = None, separator: bool = True):
    pool_entry = POOL_DIR / header_name
    if not pool_entry.exists():
        raise FileNotFoundError(f"Pool entry not found: pool/{header_name}/")

    pool_header_xml = (pool_entry / 'header.xml').read_bytes()
    if font or size_pt:
        pool_header_xml = _apply_text_overrides(pool_header_xml, font, size_pt)
    pool_rels_path = pool_entry / 'header.xml.rels'
    pool_rels_xml = pool_rels_path.read_bytes() if pool_rels_path.exists() else None

    images_dir = pool_entry / 'images'
    pool_images = (
        {f.name: f.read_bytes() for f in images_dir.iterdir() if f.is_file()}
        if images_dir.exists() else {}
    )

    zip_contents = {}
    compress_types = {}
    with zipfile.ZipFile(docx_path) as z:
        for info in z.infolist():
            zip_contents[info.filename] = z.read(info.filename)
            compress_types[info.filename] = info.compress_type

    doc_rels_root = etree.fromstring(zip_contents['word/_rels/document.xml.rels'])
    existing_header_rel = next(
        (r for r in doc_rels_root if r.get('Type') == REL_TYPE_HEADER), None
    )

    if existing_header_rel is None:
        _add_new_header(zip_contents, doc_rels_root, pool_header_xml, pool_rels_xml, pool_images, header_name)
    else:
        header_key = f"word/{existing_header_rel.get('Target')}"
        if replace:
            _replace_header(zip_contents, header_key, pool_header_xml, pool_rels_xml, pool_images, header_name)
        else:
            if pool_images:
                print("Warning: pool header contains images and target doc already has a header.")
                if separator:
                    print("         Adding separator line between injected and existing header.")
                    print("         Use --no-separator to skip (may produce inconsistent results).")
                else:
                    print("         --no-separator set: headers merged directly. Results may be inconsistent.")
            _prepend_header(zip_contents, header_key, pool_header_xml, pool_rels_xml, pool_images,
                            header_name, separator=separator)

    _write_docx(zip_contents, compress_types, out_path)
    print(f"Saved: {out_path}")


# ── internal helpers ──────────────────────────────────────────────────────────

def _apply_text_overrides(header_xml: bytes, font: str = None, size_pt: int = None) -> bytes:
    root = etree.fromstring(header_xml)
    half_pts = str(size_pt * 2) if size_pt else None

    for run in root.iter(f'{{{W}}}r'):
        rpr = run.find(f'{{{W}}}rPr')
        if rpr is None:
            rpr = etree.Element(f'{{{W}}}rPr')
            run.insert(0, rpr)

        if font:
            f_el = rpr.find(f'{{{W}}}rFonts')
            if f_el is None:
                f_el = etree.SubElement(rpr, f'{{{W}}}rFonts')
            f_el.set(f'{{{W}}}ascii', font)
            f_el.set(f'{{{W}}}hAnsi', font)
            f_el.set(f'{{{W}}}cs', font)

        if half_pts:
            for tag in (f'{{{W}}}sz', f'{{{W}}}szCs'):
                el = rpr.find(tag)
                if el is None:
                    el = etree.SubElement(rpr, tag)
                el.set(f'{{{W}}}val', half_pts)

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def _next_rid(rels_root) -> str:
    ids = []
    for rel in rels_root:
        rid = rel.get('Id', '')
        if rid.startswith('rId'):
            try:
                ids.append(int(rid[3:]))
            except ValueError:
                pass
    return f'rId{max(ids, default=0) + 1}'


def _embed_pool_images(zip_contents, pool_images, pool_rels_xml, header_key, header_name) -> dict:
    """Copy pool images into docx media; update/create header rels. Returns old→new rId map."""
    if not pool_images or pool_rels_xml is None:
        return {}

    pool_rels_root = etree.fromstring(pool_rels_xml)
    pool_image_rels = {
        r.get('Id'): Path(r.get('Target', '')).name
        for r in pool_rels_root
        if r.get('Type') == REL_TYPE_IMAGE
    }

    header_rels_key = f"word/_rels/{Path(header_key).name}.rels"
    header_rels_root = (
        etree.fromstring(zip_contents[header_rels_key])
        if header_rels_key in zip_contents
        else etree.fromstring(
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
        )
    )

    rid_map = {}
    for old_rid, img_name in pool_image_rels.items():
        if img_name not in pool_images:
            continue

        new_img_name = f'{header_name}_{img_name}'
        zip_contents[f'word/media/{new_img_name}'] = pool_images[img_name]
        _ensure_image_ct(zip_contents, new_img_name)

        new_rid = _next_rid(header_rels_root)
        new_rel = etree.SubElement(header_rels_root, f'{{{PKG_REL}}}Relationship')
        new_rel.set('Id', new_rid)
        new_rel.set('Type', REL_TYPE_IMAGE)
        new_rel.set('Target', f'media/{new_img_name}')
        rid_map[old_rid] = new_rid

    zip_contents[header_rels_key] = etree.tostring(
        header_rels_root, xml_declaration=True, encoding='UTF-8', standalone=True
    )
    return rid_map


def _rewrite_rids(xml_bytes: bytes, rid_map: dict) -> bytes:
    """Replace rId attribute values per rid_map."""
    if not rid_map:
        return xml_bytes
    text = xml_bytes.decode('utf-8')
    for old, new in rid_map.items():
        text = text.replace(f'="{old}"', f'="{new}"').replace(f"='{old}'", f"='{new}'")
    return text.encode('utf-8')


def _ensure_image_ct(zip_contents, filename):
    ext = Path(filename).suffix.lstrip('.').lower()
    ct = CT_IMAGES.get(ext)
    if not ct:
        return
    ct_root = etree.fromstring(zip_contents['[Content_Types].xml'])
    if any(el.get('Extension') == ext for el in ct_root):
        return
    el = etree.SubElement(ct_root, f'{{{CT_NS}}}Default')
    el.set('Extension', ext)
    el.set('ContentType', ct)
    zip_contents['[Content_Types].xml'] = etree.tostring(
        ct_root, xml_declaration=True, encoding='UTF-8', standalone=True
    )


def _replace_header(zip_contents, header_key, pool_header_xml, pool_rels_xml, pool_images, header_name):
    rid_map = _embed_pool_images(zip_contents, pool_images, pool_rels_xml, header_key, header_name)
    zip_contents[header_key] = _rewrite_rids(pool_header_xml, rid_map)


def _separator_paragraph() -> etree._Element:
    p = etree.Element(f'{{{W}}}p')
    pPr = etree.SubElement(p, f'{{{W}}}pPr')
    pBdr = etree.SubElement(pPr, f'{{{W}}}pBdr')
    bottom = etree.SubElement(pBdr, f'{{{W}}}bottom')
    bottom.set(f'{{{W}}}val', 'single')
    bottom.set(f'{{{W}}}sz', '6')
    bottom.set(f'{{{W}}}space', '1')
    bottom.set(f'{{{W}}}color', 'auto')
    return p


def _prepend_header(zip_contents, header_key, pool_header_xml, pool_rels_xml, pool_images,
                    header_name, separator: bool = True):
    rid_map = _embed_pool_images(zip_contents, pool_images, pool_rels_xml, header_key, header_name)
    pool_xml = _rewrite_rids(pool_header_xml, rid_map)

    pool_root = etree.fromstring(pool_xml)
    existing_root = etree.fromstring(zip_contents[header_key])

    pool_children = list(pool_root)
    for i, child in enumerate(pool_children):
        existing_root.insert(i, deepcopy(child))

    if separator:
        existing_root.insert(len(pool_children), _separator_paragraph())

    zip_contents[header_key] = etree.tostring(
        existing_root, xml_declaration=True, encoding='UTF-8', standalone=True
    )


def _add_new_header(zip_contents, doc_rels_root, pool_header_xml, pool_rels_xml, pool_images, header_name):
    header_key = next(
        f'word/header{n}.xml' for n in range(1, 20)
        if f'word/header{n}.xml' not in zip_contents
    )

    rid_map = _embed_pool_images(zip_contents, pool_images, pool_rels_xml, header_key, header_name)
    zip_contents[header_key] = _rewrite_rids(pool_header_xml, rid_map)

    # Register content type
    ct_root = etree.fromstring(zip_contents['[Content_Types].xml'])
    override = etree.SubElement(ct_root, f'{{{CT_NS}}}Override')
    override.set('PartName', f'/{header_key}')
    override.set('ContentType', CT_HEADER)
    zip_contents['[Content_Types].xml'] = etree.tostring(
        ct_root, xml_declaration=True, encoding='UTF-8', standalone=True
    )

    # Add to document rels
    new_rid = _next_rid(doc_rels_root)
    new_rel = etree.SubElement(doc_rels_root, f'{{{PKG_REL}}}Relationship')
    new_rel.set('Id', new_rid)
    new_rel.set('Type', REL_TYPE_HEADER)
    new_rel.set('Target', Path(header_key).name)
    zip_contents['word/_rels/document.xml.rels'] = etree.tostring(
        doc_rels_root, xml_declaration=True, encoding='UTF-8', standalone=True
    )

    # Wire into sectPr in document.xml
    doc_root = etree.fromstring(zip_contents['word/document.xml'])
    body = doc_root.find(f'.//{{{W}}}body')
    sect_pr = body.find(f'{{{W}}}sectPr')
    if sect_pr is None:
        sect_pr = etree.SubElement(body, f'{{{W}}}sectPr')
    href = etree.SubElement(sect_pr, f'{{{W}}}headerReference')
    href.set(f'{{{W}}}type', 'default')
    href.set(f'{{{R}}}id', new_rid)
    zip_contents['word/document.xml'] = etree.tostring(
        doc_root, xml_declaration=True, encoding='UTF-8', standalone=True
    )


def _write_docx(zip_contents, compress_types, out_path):
    with zipfile.ZipFile(out_path, 'w') as z:
        for name, data in zip_contents.items():
            z.writestr(name, data, compress_type=compress_types.get(name, zipfile.ZIP_DEFLATED))
