"""
xlsx extractor — outputs markdown tables + PNG images + manifest.json
"""
import io
import json
import re
import sys
from io import BytesIO
from pathlib import Path

# Fix Windows console encoding for Vietnamese text
if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def slugify(name: str) -> str:
    return re.sub(r"[^\w]+", "_", name).strip("_")


def get_merged_map(sheet):
    """Return dict {(row, col): (value, rowspan, colspan)} for top-left cells of merged ranges."""
    merged = {}
    covered = set()
    for rng in sheet.merged_cells.ranges:
        min_row, min_col = rng.min_row, rng.min_col
        max_row, max_col = rng.max_row, rng.max_col
        value = sheet.cell(min_row, min_col).value
        rowspan = max_row - min_row + 1
        colspan = max_col - min_col + 1
        merged[(min_row, min_col)] = (value, rowspan, colspan)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                if not (r == min_row and c == min_col):
                    covered.add((r, c))
    return merged, covered


def cell_text(sheet, r, c) -> str:
    cell = sheet.cell(r, c)
    val = str(cell.value).strip() if cell.value is not None else ""
    if val and cell.font and cell.font.strike:
        val = f"~~{val}~~"
    return val


def sheet_to_markdown(sheet) -> str:
    if sheet.max_row is None or sheet.max_column is None:
        return "_empty sheet_\n"

    merged_map, covered = get_merged_map(sheet)

    rows = []
    for r in range(1, sheet.max_row + 1):
        row = []
        for c in range(1, sheet.max_column + 1):
            if (r, c) in covered:
                row.append("")
                continue
            if (r, c) in merged_map:
                val, rowspan, colspan = merged_map[(r, c)]
                cell_str = cell_text(sheet, r, c)
                if rowspan > 1 or colspan > 1:
                    cell_str = f"{cell_str}[{rowspan}r×{colspan}c]" if cell_str else f"[{rowspan}r×{colspan}c]"
                row.append(cell_str)
            else:
                row.append(cell_text(sheet, r, c))
        # trim trailing empty cells
        while row and row[-1] == "":
            row.pop()
        rows.append(row)

    # drop fully empty rows
    rows = [r for r in rows if any(c != "" for c in r)]

    if not rows:
        return "_empty sheet_\n"

    col_count = max(len(r) for r in rows)
    rows = [r + [""] * (col_count - len(r)) for r in rows]

    def fmt_row(row):
        return "|" + "|".join(row) + "|"

    lines = []
    lines.append(fmt_row(rows[0]))
    lines.append("|" + "|".join("---" for _ in range(col_count)) + "|")
    for row in rows[1:]:
        lines.append(fmt_row(row))

    return "\n".join(lines) + "\n"


def extract_images(sheet, images_dir: Path, sheet_slug: str):
    records = []
    for i, img in enumerate(sheet._images):
        try:
            data = img._data()
            pil_img = __import__("PIL.Image", fromlist=["Image"]).open(BytesIO(data))
            ext = pil_img.format.lower() if pil_img.format else "png"
            filename = f"{sheet_slug}_img_{i + 1}.{ext}"
            out_path = images_dir / filename
            pil_img.save(out_path)

            anchor = img.anchor
            if hasattr(anchor, "_from"):
                row = anchor._from.row + 1
                col = anchor._from.col + 1
                col_letter = get_column_letter(col)
            else:
                row, col, col_letter = None, None, None

            records.append({
                "file": f"images/{filename}",
                "row": row,
                "col": col,
                "cell": f"{col_letter}{row}" if col_letter else None,
            })
        except Exception as e:
            records.append({"error": str(e), "index": i})
    return records


def extract(xlsx_path: str, output_dir: str = "output"):
    src = Path(xlsx_path)
    out = Path(output_dir)
    sheets_dir = out / "sheets"
    images_dir = out / "images"
    sheets_dir.mkdir(parents=True, exist_ok=True)
    images_dir.mkdir(parents=True, exist_ok=True)

    print(f"Loading {src.name} ...")
    wb = load_workbook(src, data_only=True)

    manifest = {"source": src.name, "sheets": []}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        slug = slugify(sheet_name)
        print(f"  Sheet: {sheet_name}")

        md = sheet_to_markdown(ws)
        md_file = sheets_dir / f"{slug}.md"
        md_file.write_text(f"# {sheet_name}\n\n{md}", encoding="utf-8")

        images = extract_images(ws, images_dir, slug)

        manifest["sheets"].append({
            "name": sheet_name,
            "markdown_file": f"sheets/{slug}.md",
            "rows": ws.max_row,
            "cols": ws.max_column,
            "images": images,
        })

    manifest_path = out / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    total_images = sum(len(s["images"]) for s in manifest["sheets"])
    print(f"\nDone! Output in '{out}/'")
    print(f"  Sheets : {len(manifest['sheets'])}")
    print(f"  Images : {total_images}")
    print(f"  Manifest: {manifest_path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        # default to first xlsx in current dir
        files = list(Path(".").glob("*.xlsx"))
        if not files:
            print("Usage: python extract.py <file.xlsx> [output_dir]")
            sys.exit(1)
        xlsx = str(files[0])
    else:
        xlsx = sys.argv[1]

    base_out = sys.argv[2] if len(sys.argv) > 2 else "output"
    file_slug = slugify(Path(xlsx).stem)
    out = str(Path(base_out) / file_slug)
    extract(xlsx, out)
