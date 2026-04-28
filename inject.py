#!/usr/bin/env python3
"""CLI: inject a pool header into a document."""

import argparse
import sys
from pathlib import Path

from converter import to_docx
from header_injector import inject

POOL_DIR = Path(__file__).parent / 'pool'


def list_pool():
    entries = sorted(p.name for p in POOL_DIR.iterdir() if p.is_dir() and (p / 'header.xml').exists())
    if entries:
        print("Available headers:")
        for e in entries:
            print(f"  {e}")
    else:
        print("Pool is empty. Use extract_header.py to add headers.")


def main():
    p = argparse.ArgumentParser(description='Inject a pool header into a document')
    p.add_argument('--doc', help='Input document (.docx, .doc, .odt, .rtf)')
    p.add_argument('--header', help='Pool header name')
    p.add_argument('--out', help='Output path (default: <input>_injected.docx)')
    p.add_argument('--replace', action='store_true', help='Replace existing header (default: prepend above it)')
    p.add_argument('--font', help='Override font for all text in header (e.g. "Arial")')
    p.add_argument('--size', type=int, help='Override font size in points for all text in header (e.g. 11)')
    p.add_argument('--list', action='store_true', help='List available pool headers')
    args = p.parse_args()

    if args.list:
        list_pool()
        return

    if not args.doc or not args.header:
        p.print_help()
        sys.exit(1)

    input_path = Path(args.doc)
    if not input_path.exists():
        print(f"Not found: {input_path}")
        sys.exit(1)

    docx_path, _ = to_docx(input_path)

    out_path = Path(args.out) if args.out else input_path.with_name(input_path.stem + '_injected.docx')

    inject(docx_path, args.header, out_path, replace=args.replace,
           font=args.font, size_pt=args.size)


if __name__ == '__main__':
    main()
