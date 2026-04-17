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
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    cell = sheet.cell(r, c)
    if cell.value is None:
        return ""
    if isinstance(cell.value, CellRichText):
        cell_strike = bool(cell.font and cell.font.strike)
        parts = []
        for run in cell.value:
            if isinstance(run, TextBlock):
                text = run.text or ""
                strike = bool(run.font.strike) if run.font else cell_strike
                if strike:
                    text = f"~~{text}~~"
                parts.append(text)
            else:
                # plain str inherits cell-level font
                text = str(run)
                if cell_strike:
                    text = f"~~{text}~~"
                parts.append(text)
        result = "".join(parts).strip()
    else:
        result = str(cell.value).strip()
        if result and cell.font and cell.font.strike:
            result = f"~~{result}~~"
    # newline inside a table cell breaks markdown — collapse to space
    result = result.replace("\n", " ").replace("\r", "")
    return result


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
    wb = load_workbook(src, data_only=True, rich_text=True)

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


CONFIG_PATH = Path(__file__).parent / "config.json"


def load_config() -> dict:
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(data: dict):
    CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def run_gui():
    import queue
    import threading
    import os
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    class QueueWriter:
        def __init__(self, q):
            self.q = q
        def write(self, text):
            if text:
                self.q.put(text)
        def flush(self):
            pass

    class ExtractorApp:
        def __init__(self, root):
            self.root = root
            self.root.title("akira-sheet-extractor")
            self.root.resizable(True, True)
            self.log_queue = queue.Queue()
            self._build_ui()
            self._load_config()
            self.root.after(100, self._poll_log)

        def _build_ui(self):
            pad = {"padx": 8, "pady": 4}
            frm = ttk.Frame(self.root, padding=10)
            frm.grid(sticky="nsew")
            self.root.columnconfigure(0, weight=1)
            self.root.rowconfigure(0, weight=1)
            frm.columnconfigure(1, weight=1)

            # --- Input file ---
            ttk.Label(frm, text="Input file (.xlsx)").grid(row=0, column=0, columnspan=3, sticky="w", **pad)
            self.xlsx_var = tk.StringVar()
            self.xlsx_entry = ttk.Entry(frm, textvariable=self.xlsx_var, width=55)
            self.xlsx_entry.grid(row=1, column=0, columnspan=2, sticky="ew", **pad)
            ttk.Button(frm, text="Browse…", command=self._browse_xlsx).grid(row=1, column=2, **pad)
            self.open_input_btn = ttk.Button(frm, text="Open Folder", command=self._open_input_folder)
            self.open_input_btn.grid(row=2, column=2, sticky="e", **pad)

            # --- Output folder ---
            ttk.Label(frm, text="Output folder").grid(row=3, column=0, columnspan=3, sticky="w", **pad)
            self.out_var = tk.StringVar()
            self.out_entry = ttk.Entry(frm, textvariable=self.out_var, width=55)
            self.out_entry.grid(row=4, column=0, columnspan=2, sticky="ew", **pad)
            ttk.Button(frm, text="Browse…", command=self._browse_output).grid(row=4, column=2, **pad)
            self.open_output_btn = ttk.Button(frm, text="Open Folder", command=self._open_output_folder)
            self.open_output_btn.grid(row=5, column=2, sticky="e", **pad)

            # --- Extract button ---
            self.extract_btn = ttk.Button(frm, text="Extract", command=self._run_extract)
            self.extract_btn.grid(row=6, column=0, columnspan=3, pady=10)

            # --- Log ---
            self.log = tk.Text(frm, height=14, state="disabled", wrap="word", bg="#1e1e1e", fg="#d4d4d4",
                               font=("Consolas", 9))
            self.log.grid(row=7, column=0, columnspan=3, sticky="nsew", **pad)
            frm.rowconfigure(7, weight=1)
            sb = ttk.Scrollbar(frm, command=self.log.yview)
            sb.grid(row=7, column=3, sticky="ns")
            self.log["yscrollcommand"] = sb.set

        def _load_config(self):
            cfg = load_config()
            xlsx = cfg.get("last_xlsx", "")
            out = cfg.get("last_output", "")
            if xlsx:
                self.xlsx_var.set(xlsx)
            if out:
                self.out_var.set(out)
            else:
                self.out_var.set(str(Path(__file__).parent / "output"))

        def _browse_xlsx(self):
            path = filedialog.askopenfilename(
                title="Chọn file Excel",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if path:
                self.xlsx_var.set(path)
                if not self.out_var.get().strip():
                    self.out_var.set(str(Path(path).parent / "output"))

        def _browse_output(self):
            path = filedialog.askdirectory(title="Chọn output folder")
            if path:
                self.out_var.set(path)

        def _open_input_folder(self):
            xlsx = self.xlsx_var.get().strip()
            folder = str(Path(xlsx).parent) if xlsx else ""
            if folder and Path(folder).exists():
                os.startfile(folder)

        def _open_output_folder(self):
            out = self.out_var.get().strip()
            if out and Path(out).exists():
                os.startfile(out)
            else:
                messagebox.showinfo("Thông báo", "Folder chưa tồn tại. Hãy chạy Extract trước.")

        def _log_write(self, text):
            self.log.configure(state="normal")
            self.log.insert("end", text)
            self.log.see("end")
            self.log.configure(state="disabled")

        def _poll_log(self):
            while True:
                try:
                    self._log_write(self.log_queue.get_nowait())
                except Exception:
                    break
            self.root.after(100, self._poll_log)

        def _run_extract(self):
            xlsx = self.xlsx_var.get().strip()
            out = self.out_var.get().strip() or str(Path(__file__).parent / "output")

            if not xlsx:
                messagebox.showerror("Lỗi", "Chưa chọn file xlsx.")
                return
            if not Path(xlsx).exists():
                messagebox.showerror("Lỗi", f"File không tồn tại:\n{xlsx}")
                return

            self.extract_btn.configure(state="disabled", text="Đang xử lý…")
            self._log_write(f"\n--- Extract: {Path(xlsx).name} ---\n")

            def worker():
                old_stdout, old_stderr = sys.stdout, sys.stderr
                writer = QueueWriter(self.log_queue)
                sys.stdout = writer
                sys.stderr = writer
                try:
                    file_slug = slugify(Path(xlsx).stem)
                    extract(xlsx, str(Path(out) / file_slug))
                    save_config({"last_xlsx": xlsx, "last_output": out})
                except Exception as e:
                    self.log_queue.put(f"\n[ERROR] {e}\n")
                finally:
                    sys.stdout, sys.stderr = old_stdout, old_stderr
                    self.root.after(0, lambda: self.extract_btn.configure(state="normal", text="Extract"))

            threading.Thread(target=worker, daemon=True).start()

    root = tk.Tk()
    app = ExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        try:
            import tkinter  # noqa: F401
            run_gui()
        except ImportError:
            print("Usage: python extract.py <file.xlsx> [output_dir]")
            sys.exit(1)
    else:
        xlsx = sys.argv[1]
        base_out = sys.argv[2] if len(sys.argv) > 2 else "output"
        file_slug = slugify(Path(xlsx).stem)
        out = str(Path(base_out) / file_slug)
        extract(xlsx, out)
