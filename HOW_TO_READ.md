# Hướng dẫn đọc output cho AI

File này giải thích cấu trúc output được tạo bởi **akira-sheet-extractor** từ file `.xlsx`.

## Cấu trúc thư mục

```
output/
└── <tên_file_excel>/
    ├── markdown/               # có nếu chạy với format md
    │   ├── sheets/*.md
    │   ├── images/
    │   └── manifest.json
    ├── html/                   # có nếu chạy với format html
    │   ├── sheets/*.html
    │   ├── images/
    │   └── manifest.json
    └── (cả hai nếu chọn cùng lúc)
```

## Bước đọc đề xuất

1. Đọc `manifest.json` để nắm danh sách sheet, số dòng/cột, và vị trí ảnh
2. Đọc từng file trong `sheets/` theo thứ tự hoặc theo tên sheet cần thiết
3. Ảnh lưu trong `images/`, tọa độ cell được ghi trong manifest

---

## Quy ước Markdown

### Bảng
- Mỗi sheet là 1 file `.md` với header `# Tên Sheet` và bảng phía dưới
- Hàng và cột rỗng hoàn toàn bị lược bỏ để giảm noise

### Ô merge
Ô merge chỉ có nội dung ở **ô đầu tiên (top-left)**, các ô bị bao phủ bị bỏ trống. Kích thước được ghi chú ngay sau giá trị:

```
|Tiêu đề[2r×3c]|...|
```

`[2r×3c]` = ô này span 2 hàng × 3 cột.

### Strikethrough
```
~~nội dung cũ~~ nội dung mới
```
`~~...~~` là nội dung đã bị **xóa hoặc thay thế** — đọc kèm phần text liền sau để hiểu sự thay đổi.

### Ảnh nhúng
```
![](../images/sheet_img_1.png)
```
Ảnh được inject vào đúng cell có anchor trong file Excel gốc.

---

## Quy ước HTML

### Ô merge
Dùng thuộc tính chuẩn HTML — đọc như bình thường:
```html
<td rowspan="2" colspan="3">Tiêu đề</td>
```

### Strikethrough
```html
<s>nội dung cũ</s> nội dung mới
```

### Ảnh nhúng
```html
<img src="../images/sheet_img_1.png">
```

---

## Cấu trúc manifest.json

```json
{
  "source": "tên_file.xlsx",
  "format": "md",
  "sheets": [
    {
      "name": "Tên sheet gốc",
      "sheet_file": "sheets/ten_sheet.md",
      "rows": 50,
      "cols": 10,
      "images": [
        {"file": "images/ten_sheet_img_1.png", "row": 3, "col": 2, "cell": "B3"}
      ]
    }
  ]
}
```

- `rows` / `cols`: kích thước vùng dữ liệu của sheet
- `images[].cell`: tọa độ Excel (ví dụ `B3`) nơi ảnh được neo trong file gốc
