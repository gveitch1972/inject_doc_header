# Header Pool

Each subdirectory is one named header.

## Structure

```
pool/
  <header_name>/
    header.xml      # required — w:hdr XML element
    manifest.json   # required — metadata and object config
    images/         # optional — any images referenced in header.xml
```

## Adding a header from an existing .docx

Extract `word/header1.xml` from the .docx zip, save as `header.xml`.
Copy any referenced images from `word/media/` into `images/`.
Update `manifest.json` with name and object descriptions.

## manifest.json schema

```json
{
  "name": "string",
  "description": "string",
  "objects": [
    {
      "type": "text|image",
      "value": "text string or image filename",
      "font": "font name (text only)",
      "size_pt": 12,
      "bold": false,
      "align": "left|center|right"
    }
  ]
}
```
