# akira-sheet-extractor

Extract `.xlsx` files into compact markdown tables, images, and a manifest — optimized for AI consumption.

## Output

For each sheet, generates:
- `sheets/<sheet>.md` — markdown table (compact, no padding)
- `images/<sheet>_img_N.png` — embedded images with anchor position
- `manifest.json` — sheet list, row/col counts, image metadata

## Usage

```bash
# auto-detect first .xlsx in current directory
python extract.py

# specify file
python extract.py path/to/file.xlsx

# specify file + output directory
python extract.py path/to/file.xlsx output_dir
```

Output goes to `output/<filename>/` by default.

## Install

```bash
pip install openpyxl pillow
```

## Markdown format

- Merged cells are annotated: `value[2r×3c]`
- Strikethrough text is preserved as `~~text~~`
- Empty rows and trailing empty cells are trimmed
- Newlines within cells are collapsed to a space

## Warnings

openpyxl may emit warnings for:
- **DrawingML** — shapes/drawings are not supported, only images and charts
- **Invalid date serials** — cells with out-of-range date values are treated as errors
