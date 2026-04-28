#!/usr/bin/env python3
"""Convert non-docx formats to .docx via LibreOffice headless."""

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

_LO_CANDIDATES = [
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    '/usr/bin/libreoffice',
    '/usr/bin/soffice',
]


def _find_lo() -> str | None:
    for p in _LO_CANDIDATES:
        if Path(p).exists():
            return p
    return shutil.which('libreoffice') or shutil.which('soffice')


def to_docx(input_path: Path) -> tuple[Path, bool]:
    """Return (docx_path, was_converted). Pass-through for .docx inputs."""
    if input_path.suffix.lower() == '.docx':
        return input_path, False

    lo = _find_lo()
    if not lo:
        print("LibreOffice not found. Install: brew install --cask libreoffice", file=sys.stderr)
        sys.exit(1)

    tmp = Path(tempfile.mkdtemp())
    r = subprocess.run(
        [lo, '--headless', '--convert-to', 'docx', '--outdir', str(tmp), str(input_path)],
        capture_output=True, text=True
    )
    if r.returncode != 0:
        print(f"Conversion failed:\n{r.stderr}", file=sys.stderr)
        sys.exit(1)

    candidates = list(tmp.glob('*.docx'))
    if not candidates:
        print("Conversion produced no .docx output.", file=sys.stderr)
        sys.exit(1)

    return candidates[0], True
