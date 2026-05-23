# How to Read the Output (for AI)

This file explains the output structure produced by **akira-sheet-extractor** from `.xlsx` files.

## Directory Structure

```
output/
└── <excel_filename>/
    ├── markdown/               # present when run with md format
    │   ├── sheets/*.md
    │   ├── images/
    │   └── manifest.json
    ├── html/                   # present when run with html format
    │   ├── sheets/*.html
    │   ├── images/
    │   └── manifest.json
    └── (both if both formats selected)
```

## Recommended Reading Order

1. Read `manifest.json` to get the sheet list, row/column counts, and image positions.
2. Read each file under `sheets/` in order, or by the sheet name you need.
3. Images are stored in `images/`; their cell coordinates are recorded in the manifest.

---

## Markdown Conventions

### Tables
- Each sheet is one `.md` file with a `# Sheet Name` header followed by a table.
- Completely empty rows and columns are omitted to reduce noise.

### Merged Cells
The top-left cell holds the value; covered cells are left empty. The span is annotated right after the value:

```
|Title[2r×3c]|...|
```

`[2r×3c]` means the cell spans 2 rows × 3 columns.

### Strikethrough
```
~~old content~~ new content
```
`~~...~~` marks content that has been **deleted or replaced** — read it together with the text that follows to understand the change.

### Embedded Images
```
<!-- image cell=D10 range=D10:F14 --> ![](../images/sheet_img_1.png)
```
Images are injected near their visual cell when possible. Because Excel images are floating objects and Markdown tables cannot perfectly preserve layout, each sheet also ends with an explicit image layer:

```md
## Images
- `D10:F14`, placed at `D10`: ![image at D10](../images/sheet_img_1.png)
```
Use this section and `manifest.json` as the reliable source for image-to-cell mapping.

---

## HTML Conventions

### Merged Cells
Standard HTML attributes — read as normal:
```html
<td rowspan="2" colspan="3">Title</td>
```

### Strikethrough
```html
<s>old content</s> new content
```

### Embedded Images
```html
<img src="../images/sheet_img_1.png">
```

---

## manifest.json Structure

```json
{
  "source": "filename.xlsx",
  "format": "md",
  "sheets": [
    {
      "name": "Original Sheet Name",
      "sheet_file": "sheets/sheet_name.md",
      "rows": 50,
      "cols": 10,
      "images": [
        {
          "file": "images/sheet_name_img_1.png",
          "row": 3,
          "col": 2,
          "cell": "B3",
          "anchor_cell": "B3",
          "visual_cell": "B3",
          "range": "B3:D6",
          "anchor_type": "twoCellAnchor",
          "confidence": "high",
          "placement": "ooxml"
        }
      ]
    }
  ]
}
```

- `rows` / `cols`: data area dimensions of the sheet.
- `images[].cell`: backward-compatible placement cell, usually the visual cell.
- `images[].anchor_cell`: original top-left anchor cell from Excel.
- `images[].visual_cell`: best cell for Markdown injection.
- `images[].range`: Excel cell range covered by the image when available.
- `images[].confidence`: `high` for two-cell OOXML anchors, `medium` for one-cell/fallback mappings.
