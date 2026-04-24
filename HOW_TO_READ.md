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
![](../images/sheet_img_1.png)
```
Images are injected into the cell that matches their anchor position in the original Excel file.

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
        {"file": "images/sheet_name_img_1.png", "row": 3, "col": 2, "cell": "B3"}
      ]
    }
  ]
}
```

- `rows` / `cols`: data area dimensions of the sheet.
- `images[].cell`: Excel coordinate (e.g. `B3`) where the image is anchored in the original file.
