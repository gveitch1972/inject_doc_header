# inject_doc_header

CLI tool to stamp a named header from a pool onto any Word document. Headers can be plain text or images (logos, letterheads). Output is always `.docx`.

## Requirements

- Python 3.8+
- LibreOffice — only needed for `.doc` / `.odt` / `.rtf` inputs (not required for `.docx`)

## Setup

```bash
git clone https://github.com/gveitch1972/inject_doc_header.git
cd inject_doc_header
pip install -r requirements.txt
```

For non-docx inputs (macOS):

```bash
brew install --cask libreoffice
```

## Quick start

Example letterhead templates and a sample target document are included.

```bash
# 1. Extract a header from one of the included templates
python3 extract_header.py --doc "example_hdrs/California_Letterhead.doc" --name california

# 2. Check the pool
python3 inject.py --list

# 3. Inject into the sample document
python3 inject.py --doc sample_target.docx --header california
# → outputs sample_target_injected.docx
```

## Usage

```bash
# Prepend header above existing content (default)
python3 inject.py --doc TARGET.docx --header NAME

# Replace existing header
python3 inject.py --doc TARGET.docx --header NAME --replace

# Custom output path
python3 inject.py --doc TARGET.docx --header NAME --out OUTPUT.docx

# List available headers in pool
python3 inject.py --list

# Extract header from an existing Word doc into the pool
python3 extract_header.py --doc SOURCE.docx --name NAME
```

## Included examples

| Path | Purpose |
|------|---------|
| `example_hdrs/` | 13 generic `.doc` letterhead templates — use with `extract_header.py` to build pool entries |
| `sample_target.docx` | Neutral target document for testing injection |
| `pool/acme_test/` | Pre-built text-only pool entry (no images needed) |
| `pool/example_header/` | Minimal manifest-only pool entry |

## Full documentation

See [HOWTO.md](HOWTO.md) for pool folder structure, manifest schema, and troubleshooting.
