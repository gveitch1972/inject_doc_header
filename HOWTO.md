# inject_doc_header — How To Use

## What it does

Takes a Word document and stamps a header from a pool onto it. Headers can be text or images (company logos etc). Output is always .docx.

Included for testing: `example_hdrs/` (13 generic letterhead `.doc` files) and `sample_target.docx` (neutral target document).

---

## 1. Install dependencies

```bash
pip install -r requirements.txt
```

For `.doc`, `.odt`, `.rtf` inputs (not needed for `.docx`):
```bash
brew install --cask libreoffice
```

---

## 2. Add a header to the pool

### From an existing Word document (recommended)

```bash
python3 extract_header.py --doc "Company Letterhead.docx" --name mycompany
```

This creates `pool/mycompany/` with the header XML and any images extracted.

### Check what's in the pool

```bash
python3 inject.py --list
```

---

## 3. Inject a header into a document

### Add header above existing content (default)

```bash
python3 inject.py --doc "my_document.docx" --header mycompany
```

Output: `my_document_injected.docx` in same folder.

### Replace existing header

```bash
python3 inject.py --doc "my_document.docx" --header mycompany --replace
```

### Specify output path

```bash
python3 inject.py --doc "my_document.docx" --header mycompany --out "output/final.docx"
```

### Use a .doc file as input

```bash
python3 inject.py --doc "old_document.doc" --header mycompany
```

(Requires LibreOffice installed — see step 1.)

---

## Pool folder structure

```
pool/
  mycompany/
    header.xml        ← Word header XML (do not edit manually)
    header.xml.rels   ← image relationships (auto-generated)
    images/           ← any images used in the header
    manifest.json     ← metadata: edit description here
```

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `Pool entry not found` | Run `--list` to see what's available |
| `No header found in document` | Source doc has no Word header — use a doc with a proper header section |
| Image shows as broken | Re-extract the header; the source doc may have embedded the image differently |
| `.doc` conversion fails | Check LibreOffice is installed: `which soffice` |

---

## Quick reference

```bash
# List pool
python3 inject.py --list

# Extract header from doc
python3 extract_header.py --doc SOURCE.docx --name NAME

# Inject (prepend)
python3 inject.py --doc TARGET.docx --header NAME

# Inject (replace)
python3 inject.py --doc TARGET.docx --header NAME --replace

# Custom output path
python3 inject.py --doc TARGET.docx --header NAME --out OUTPUT.docx
```
