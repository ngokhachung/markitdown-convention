# MarkItDown Office Convention — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Bộ 4 tài liệu convention tiếng Việt (README + docx/xlsx/pptx) kèm thư mục `samples/` chứng minh từng quy tắc bằng output MarkItDown thật.

**Architecture:** Làm samples trước (Task 1–4) để có bằng chứng thực nghiệm, viết tài liệu sau (Task 5–8) đối chiếu với output thật. Hai script Python: `generate_samples.py` sinh cặp file good/bad cho mỗi định dạng, `convert_and_verify.py` convert bằng MarkItDown và assert các khác biệt mà quy tắc mô tả — đây chính là "test" của dự án.

**Tech Stack:** Python 3.10+, markitdown (extras docx/xlsx/pptx), python-docx, openpyxl, python-pptx. Tài liệu là Markdown thuần.

**Spec:** `docs/superpowers/specs/2026-07-03-markitdown-convention-design.md`

## Global Constraints

- Ngôn ngữ tài liệu: tiếng Việt; thuật ngữ kỹ thuật (Style, alt text, placeholder, merge cell…) giữ tiếng Anh.
- Mã quy tắc: `DOCX-NN` / `XLSX-NN` / `PPTX-NN`; đúng 3 mức: `[BẮT BUỘC]`, `[NÊN]`, `[TRÁNH]`.
- Danh mục quy tắc lấy VERBATIM từ spec §3.3: 12 DOCX, 10 XLSX, 10 PPTX. Không thêm/bớt/đổi ID.
- 4 file tài liệu đặt ở repo root: `README.md`, `convention-docx.md`, `convention-xlsx.md`, `convention-pptx.md`.
- `samples/output/*.md` PHẢI được commit (team xem không cần chạy lại).
- Nguyên tắc "thực nghiệm thắng tài liệu": nếu output thật mâu thuẫn quy tắc nào → sửa quy tắc theo thực tế, ghi chú thay đổi trong commit message.
- Mọi file text ghi bằng UTF-8 (script Python phải truyền `encoding="utf-8"` khi write).
- Python chạy qua venv `.venv` (không commit); trên Git Bash gọi `.venv/Scripts/python`.
- Commit sau mỗi task, message tiếng Anh, kết thúc bằng `Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>`.

---

### Task 1: Môi trường Python + scaffolding repo

**Files:**
- Create: `.gitignore`
- Create: `.venv/` (không commit)

**Interfaces:**
- Produces: lệnh `.venv/Scripts/python` chạy được với các package `markitdown`, `docx`, `openpyxl`, `pptx` import được. Task 2–4 và 8 dùng interpreter này.

- [ ] **Step 1: Kiểm tra Python có sẵn**

Run: `python --version`
Expected: `Python 3.1x.x` (≥3.10). Nếu lệnh không tồn tại, thử `py -3 --version`. Nếu cả hai fail → BLOCKED, báo user cài Python.

- [ ] **Step 2: Tạo `.gitignore`**

```gitignore
.venv/
__pycache__/
*.pyc
```

- [ ] **Step 3: Tạo venv và cài package**

Run (từ repo root `D:\project\MD convention`):
```bash
python -m venv .venv
.venv/Scripts/python -m pip install --quiet "markitdown[docx,xlsx,pptx]" python-docx openpyxl python-pptx
```
Expected: exit 0, không có ERROR trong output (warning về pip version bỏ qua được).

- [ ] **Step 4: Xác minh import**

Run:
```bash
.venv/Scripts/python -c "from markitdown import MarkItDown; import docx, openpyxl, pptx; print('env ok')"
```
Expected: `env ok`

- [ ] **Step 5: Commit**

```bash
git add .gitignore
git commit -m "chore: add gitignore for python venv"
```

---

### Task 2: Samples DOCX — sinh, convert, verify

**Files:**
- Create: `samples/generate_samples.py`
- Create: `samples/convert_and_verify.py`
- Create (generated, commit): `samples/docx-good.docx`, `samples/docx-bad.docx`, `samples/output/docx-good.md`, `samples/output/docx-bad.md`

**Interfaces:**
- Consumes: `.venv/Scripts/python` từ Task 1.
- Produces (Task 3–4 sẽ thêm hàm vào đúng 2 file này):
  - `generate_samples.py`: hằng `SAMPLES_DIR: Path`, `PNG_1PX: bytes`; hàm `gen_docx() -> None`; `main()` gọi lần lượt các hàm `gen_*`.
  - `convert_and_verify.py`: hàm `convert(name: str) -> str` (convert `samples/<name>` → `samples/output/<stem>.md`, trả về nội dung); hàm `verify_docx() -> None` (assert, in `PASS docx`); `main()` gọi lần lượt các hàm `verify_*` rồi in `ALL PASS`.

- [ ] **Step 1: Viết `samples/generate_samples.py`**

```python
"""Sinh file Office mẫu good/bad để kiểm chứng convention MarkItDown."""
import base64
import io
from pathlib import Path

SAMPLES_DIR = Path(__file__).parent

# PNG 1x1 hợp lệ — đủ để python-docx/python-pptx nhúng làm ảnh
PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ"
    "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def gen_docx() -> None:
    from docx import Document
    from docx.shared import Inches, Pt

    # good: heading dùng Style, bảng sạch có header, ảnh có alt text
    doc = Document()
    doc.add_heading("Báo cáo doanh thu 2026", level=1)
    doc.add_heading("Tình hình quý 1", level=2)
    doc.add_paragraph("Doanh thu quý 1 đạt 600 triệu, tăng 20% so với cùng kỳ.")
    table = doc.add_table(rows=3, cols=2)
    data = [("Tháng", "Doanh thu"), ("01", "100"), ("02", "200")]
    for row, (c1, c2) in zip(table.rows, data):
        row.cells[0].text = c1
        row.cells[1].text = c2
    pic = doc.add_picture(io.BytesIO(PNG_1PX), width=Inches(1))
    pic._inline.docPr.set(
        "descr", "Biểu đồ cột doanh thu theo tháng: tháng 1 đạt 100, tháng 2 đạt 200"
    )
    pic._inline.docPr.set("title", "Biểu đồ doanh thu")
    doc.save(SAMPLES_DIR / "docx-good.docx")

    # bad: "tiêu đề" bằng bold + cỡ chữ, bảng merge header, ảnh không alt text
    doc = Document()
    p = doc.add_paragraph()
    run = p.add_run("Báo cáo doanh thu 2026")
    run.bold = True
    run.font.size = Pt(16)
    table = doc.add_table(rows=3, cols=2)
    table.rows[0].cells[0].merge(table.rows[0].cells[1])
    table.rows[0].cells[0].text = "Doanh thu theo tháng"
    table.rows[1].cells[0].text = "01"
    table.rows[1].cells[1].text = "100"
    table.rows[2].cells[0].text = "02"
    table.rows[2].cells[1].text = "200"
    doc.add_picture(io.BytesIO(PNG_1PX), width=Inches(1))
    doc.save(SAMPLES_DIR / "docx-bad.docx")


def main() -> None:
    (SAMPLES_DIR / "output").mkdir(exist_ok=True)
    gen_docx()
    print("generated: docx")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Viết `samples/convert_and_verify.py`**

```python
"""Convert file mẫu bằng MarkItDown và assert đúng các khác biệt convention mô tả."""
import sys
from pathlib import Path

from markitdown import MarkItDown

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

SAMPLES_DIR = Path(__file__).parent
OUTPUT_DIR = SAMPLES_DIR / "output"

_md = MarkItDown()


def convert(name: str) -> str:
    """Convert samples/<name> -> samples/output/<stem>.md, trả về nội dung Markdown."""
    text = _md.convert(str(SAMPLES_DIR / name)).text_content
    (OUTPUT_DIR / (Path(name).stem + ".md")).write_text(text, encoding="utf-8")
    return text


def verify_docx() -> None:
    good = convert("docx-good.docx")
    bad = convert("docx-bad.docx")
    # DOCX-01: Heading Style -> '#'; bold thủ công -> văn bản thường
    assert "# Báo cáo doanh thu 2026" in good, "good: thiếu heading H1"
    assert "## Tình hình quý 1" in good, "good: thiếu heading H2"
    assert "# " not in bad, "bad: không được có heading nào"
    assert "Báo cáo doanh thu 2026" in bad, "bad: text tiêu đề vẫn phải còn (dạng thường)"
    # DOCX-03: alt text là thứ duy nhất của ảnh sống sót
    assert "Biểu đồ" in good, "good: alt text phải xuất hiện trong output"
    assert "Biểu đồ" not in bad, "bad: không alt text thì không có mô tả ảnh"
    print("PASS docx")


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    verify_docx()
    print("ALL PASS")


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Chạy sinh file rồi verify**

Run:
```bash
.venv/Scripts/python samples/generate_samples.py
.venv/Scripts/python samples/convert_and_verify.py
```
Expected: `generated: docx` rồi `PASS docx` + `ALL PASS`.
Nếu một assert fail: đọc `samples/output/docx-*.md`, xác định hành vi thật của MarkItDown, sửa assert theo thực tế và GHI LẠI phát hiện đó (sẽ đưa vào `convention-docx.md` ở Task 5) — thực nghiệm thắng tài liệu.

- [ ] **Step 4: Xem output bằng mắt, ghi nhận hành vi bảng merge**

Run: đọc `samples/output/docx-bad.md`, quan sát bảng bị merge header render lệch/thiếu cột thế nào. Ghi chú 1–2 dòng mô tả hiện tượng thật (dùng cho ví dụ ❌ của DOCX-02 ở Task 5).

- [ ] **Step 5: Commit**

```bash
git add samples/
git commit -m "feat: add docx good/bad samples with verified markitdown output"
```

---

### Task 3: Samples XLSX — sinh, convert, verify

**Files:**
- Modify: `samples/generate_samples.py` (thêm `gen_xlsx()`, gọi trong `main()`)
- Modify: `samples/convert_and_verify.py` (thêm `verify_xlsx()`, gọi trong `main()`)
- Create (generated, commit): `samples/xlsx-good.xlsx`, `samples/xlsx-bad.xlsx`, `samples/output/xlsx-good.md`, `samples/output/xlsx-bad.md`

**Interfaces:**
- Consumes: `SAMPLES_DIR`, `PNG_1PX` (không cần cho xlsx), `convert(name)` như định nghĩa Task 2.
- Produces: `gen_xlsx() -> None`, `verify_xlsx() -> None` (in `PASS xlsx`).

- [ ] **Step 1: Thêm `gen_xlsx()` vào `generate_samples.py`**

```python
def gen_xlsx() -> None:
    from openpyxl import Workbook

    # good: 1 sheet = 1 bảng từ A1, dòng 1 là header, không merge,
    # dòng Tổng ghi GIÁ TRỊ chứ không ghi công thức
    wb = Workbook()
    ws = wb.active
    ws.title = "DoanhThu2026"
    ws.append(["Tháng", "Doanh thu", "Trạng thái"])
    ws.append(["01", 100, "Đạt"])
    ws.append(["02", 200, "Đạt"])
    ws.append(["03", 300, "Vượt"])
    ws.append(["Tổng", 600, "N/A"])
    wb.save(SAMPLES_DIR / "xlsx-good.xlsx")

    # bad: tiêu đề trang trí + merge, bảng lệch khỏi A1, ô bỏ trống,
    # công thức chưa từng được Excel tính (không cached value), sheet ẩn
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "BÁO CÁO DOANH THU NĂM 2026"
    ws.merge_cells("A1:C1")
    ws["A3"] = "Tháng"
    ws["B3"] = "Doanh thu"
    ws["A4"] = "01"
    ws["B4"] = 110
    ws["A5"] = "02"          # B5 cố tình bỏ trống
    ws["A6"] = "Gấp ba"
    ws["B6"] = "=B4*3"       # openpyxl không tính -> không có cached value
    hidden = wb.create_sheet("NhapLieuTam")
    hidden["A1"] = "dữ liệu nháp không nên lộ ra"
    hidden.sheet_state = "hidden"
    wb.save(SAMPLES_DIR / "xlsx-bad.xlsx")
```

Trong `main()` của `generate_samples.py`, thêm sau `gen_docx()`:

```python
    gen_xlsx()
    print("generated: xlsx")
```

- [ ] **Step 2: Thêm `verify_xlsx()` vào `convert_and_verify.py`**

```python
def verify_xlsx() -> None:
    good = convert("xlsx-good.xlsx")
    bad = convert("xlsx-bad.xlsx")
    # XLSX-07: tên sheet thành heading ##
    assert "## DoanhThu2026" in good, "good: tên sheet phải thành heading"
    # XLSX-08: bảng sạch không có NaN
    assert "NaN" not in good, "good: không được xuất hiện NaN"
    assert "Tổng" in good and "600" in good, "good: dòng tổng ghi giá trị phải giữ nguyên"
    # XLSX-03/08: merge + ô trống sinh chữ NaN trong output
    assert "NaN" in bad, "bad: merge cell và ô trống phải sinh NaN"
    # XLSX-01/02: tiêu đề trang trí chiếm dòng header -> cột Unnamed
    assert "Unnamed" in bad, "bad: header giả phải sinh cột Unnamed"
    # XLSX-04: công thức không cached value -> kết quả biến mất
    assert "330" not in bad, "bad: 110*3=330 không được xuất hiện (không cached value)"
    assert "=B4" not in bad, "bad: chuỗi công thức cũng không được xuất hiện"
    # XLSX-06: sheet ẩn vẫn bị convert
    assert "nháp" in bad, "bad: nội dung sheet ẩn vẫn lộ ra output"
    print("PASS xlsx")
```

Trong `main()` của `convert_and_verify.py`, thêm sau `verify_docx()`:

```python
    verify_xlsx()
```

- [ ] **Step 3: Chạy lại toàn bộ**

Run:
```bash
.venv/Scripts/python samples/generate_samples.py
.venv/Scripts/python samples/convert_and_verify.py
```
Expected: `generated: docx`, `generated: xlsx`, rồi `PASS docx`, `PASS xlsx`, `ALL PASS`.
Assert fail → đọc `samples/output/xlsx-*.md`, sửa assert theo hành vi thật, ghi lại phát hiện cho Task 6.

- [ ] **Step 4: Commit**

```bash
git add samples/
git commit -m "feat: add xlsx good/bad samples with verified markitdown output"
```

---

### Task 4: Samples PPTX — sinh, convert, verify

**Files:**
- Modify: `samples/generate_samples.py` (thêm `gen_pptx()`, gọi trong `main()`)
- Modify: `samples/convert_and_verify.py` (thêm `verify_pptx()`, gọi trong `main()`)
- Create (generated, commit): `samples/pptx-good.pptx`, `samples/pptx-bad.pptx`, `samples/output/pptx-good.md`, `samples/output/pptx-bad.md`

**Interfaces:**
- Consumes: `SAMPLES_DIR`, `PNG_1PX`, `convert(name)` như Task 2.
- Produces: `gen_pptx() -> None`, `verify_pptx() -> None` (in `PASS pptx`).

- [ ] **Step 1: Thêm `gen_pptx()` vào `generate_samples.py`**

```python
def gen_pptx() -> None:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches

    # good: layout có Title placeholder, bullet trong body, ảnh có alt text,
    # speaker notes, chart cột cơ bản
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
    slide.shapes.title.text = "Kế hoạch quý 3"
    body = slide.placeholders[1].text_frame
    body.text = "Bước 1: khảo sát khách hàng"
    body.add_paragraph().text = "Bước 2: chốt tính năng"
    pic = slide.shapes.add_picture(
        io.BytesIO(PNG_1PX), Inches(6), Inches(1), width=Inches(1)
    )
    pic._element._nvXxPr.cNvPr.set(
        "descr", "Sơ đồ timeline quý 3: tháng 7 khảo sát, tháng 8 chốt tính năng"
    )
    slide.notes_slide.notes_text_frame.text = (
        "Chi tiết: khảo sát 50 khách hàng nhóm A trong tháng 7."
    )
    slide2 = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only
    slide2.shapes.title.text = "Doanh thu theo tháng"
    chart_data = CategoryChartData()
    chart_data.categories = ["Tháng 7", "Tháng 8"]
    chart_data.add_series("Doanh thu", (100, 200))
    slide2.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1), Inches(2), Inches(6), Inches(4), chart_data,
    )
    prs.save(SAMPLES_DIR / "pptx-good.pptx")

    # bad: slide Blank, "tiêu đề" là textbox nằm ĐÁY slide, các bước đặt
    # đảo vị trí (Bước 2 ở trên, Bước 1 ở dưới), ảnh không alt text
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    box2 = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    box2.text_frame.text = "Bước 2: chốt tính năng"
    box1 = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(6), Inches(1))
    box1.text_frame.text = "Bước 1: khảo sát khách hàng"
    title_box = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(6), Inches(1))
    title_box.text_frame.text = "Kế hoạch quý 3"
    slide.shapes.add_picture(io.BytesIO(PNG_1PX), Inches(6), Inches(5), width=Inches(1))
    prs.save(SAMPLES_DIR / "pptx-bad.pptx")
```

Trong `main()`, thêm sau `gen_xlsx()`:

```python
    gen_pptx()
    print("generated: pptx")
```

- [ ] **Step 2: Thêm `verify_pptx()` vào `convert_and_verify.py`**

```python
def verify_pptx() -> None:
    good = convert("pptx-good.pptx")
    bad = convert("pptx-bad.pptx")
    # PPTX-01: Title placeholder -> H1
    assert "# Kế hoạch quý 3" in good, "good: title placeholder phải thành H1"
    # PPTX-05: nội dung đúng thứ tự đọc
    assert good.index("Bước 1") < good.index("Bước 2"), "good: thứ tự đọc phải đúng"
    # PPTX-02: alt text sống sót
    assert "Sơ đồ timeline" in good, "good: alt text phải xuất hiện"
    # PPTX-07: speaker notes được convert
    assert "Notes:" in good and "50 khách hàng" in good, "good: notes phải được convert"
    # PPTX-06: chart cơ bản thành bảng dữ liệu
    assert "Tháng 7" in good and "100" in good, "good: chart phải thành bảng dữ liệu"
    # bad: textbox không thành heading, thứ tự đảo theo tọa độ, không mô tả ảnh
    assert "# Kế hoạch quý 3" not in bad, "bad: textbox không được thành heading"
    assert bad.index("Bước 2") < bad.index("Bước 1"), "bad: shape sort theo tọa độ nên đảo thứ tự"
    assert "Sơ đồ timeline" not in bad, "bad: ảnh không alt thì không có mô tả"
    print("PASS pptx")
```

Trong `main()`, thêm sau `verify_xlsx()`:

```python
    verify_pptx()
```

- [ ] **Step 3: Chạy lại toàn bộ**

Run:
```bash
.venv/Scripts/python samples/generate_samples.py
.venv/Scripts/python samples/convert_and_verify.py
```
Expected: 3 dòng `generated:` + `PASS docx`, `PASS xlsx`, `PASS pptx`, `ALL PASS`.
Assert fail → đọc `samples/output/pptx-*.md`, sửa assert theo hành vi thật, ghi lại phát hiện cho Task 7.

- [ ] **Step 4: Commit**

```bash
git add samples/
git commit -m "feat: add pptx good/bad samples with verified markitdown output"
```

---

### Task 5: Viết `convention-docx.md`

**Files:**
- Create: `convention-docx.md`

**Interfaces:**
- Consumes: `samples/output/docx-good.md`, `samples/output/docx-bad.md` (Task 2) và ghi chú hành vi bảng merge từ Task 2 Step 4.
- Produces: file `convention-docx.md` mà `README.md` (Task 8) sẽ link tới.

- [ ] **Step 1: Viết file với nội dung sau** (nếu Task 2 phát hiện hành vi khác mô tả — sửa phần "Tại sao"/ví dụ tương ứng theo output thật):

````markdown
# Convention: Word (.docx)

Quy tắc tạo file Word để MarkItDown convert sang Markdown không mất thông tin.

**Cơ chế:** MarkItDown dùng thư viện `mammoth` chuyển docx → HTML → Markdown, kèm bước
tiền xử lý chuyển công thức OMML → LaTeX. Mọi thứ mammoth không hiểu sẽ bị **bỏ qua
lặng lẽ, không có cảnh báo**.

**Mức độ:** `[BẮT BUỘC]` vi phạm là mất/hỏng dữ liệu · `[NÊN]` cải thiện chất lượng
output · `[TRÁNH]` tính năng Office bị MarkItDown bỏ qua hoàn toàn.

**Bằng chứng:** xem cặp file `samples/docx-good.docx` / `samples/docx-bad.docx`
và output thật tại `samples/output/docx-good.md` / `samples/output/docx-bad.md`.

---

### DOCX-01 — Dùng Heading Style cho mọi tiêu đề `[BẮT BUỘC]`

**Tại sao:** mammoth chỉ map Style `Heading 1..6` thành heading Markdown (`#`..`######`).
Tiêu đề tạo bằng bold + cỡ chữ to là văn bản thường trong output — mất toàn bộ cây cấu
trúc, pipeline RAG không chunk được theo section.

✅ **Đúng:** Bôi đen tiêu đề → Home → Styles → chọn "Heading 1/2/3"
❌ **Sai:** Bôi đen → tăng cỡ chữ 16pt → Bold

### DOCX-02 — Bảng không merge cell, đúng 1 dòng header `[BẮT BUỘC]`

**Tại sao:** bảng Markdown không biểu diễn được ô gộp (colspan/rowspan). Ô merge làm
các dòng lệch cột — dữ liệu bị gán nhầm cột mà không có cảnh báo.

✅ **Đúng:** mỗi ô một giá trị; dòng đầu là tên cột; lặp lại giá trị thay vì merge dọc.
❌ **Sai:** merge ô "Q1" ngang 3 cột; merge dọc ô "Miền Bắc" cho 4 dòng.

### DOCX-03 — Mọi hình ảnh phải có alt text `[BẮT BUỘC]`

**Tại sao:** ảnh nhúng bị cắt khỏi output (chỉ còn placeholder rỗng). **Alt text là thứ
duy nhất của ảnh sống sót.** Pipeline không bật LLM caption nên không có mô tả tự động.

✅ **Đúng:** Right-click ảnh → View Alt Text → mô tả đầy đủ nội dung và ý nghĩa
(vd: "Biểu đồ cột doanh thu theo tháng: tháng 1 đạt 100, tháng 2 đạt 200").
❌ **Sai:** chèn ảnh không có alt text, hoặc alt text vô nghĩa kiểu "image1", "ảnh".

### DOCX-04 — Thông tin quan trọng phải có bản chữ, không chỉ nằm trong hình `[BẮT BUỘC]`

**Tại sao:** không có OCR. Ảnh chụp bảng số liệu, đoạn code, sơ đồ = mất toàn bộ nội dung.

✅ **Đúng:** sơ đồ kèm đoạn văn mô tả; bảng số liệu gõ thành bảng Word thật; code dán
dạng text.
❌ **Sai:** screenshot bảng Excel dán vào Word.

### DOCX-05 — Công thức toán dùng Insert > Equation (OMML) `[BẮT BUỘC]`

**Tại sao:** MarkItDown chuyển equation OMML thành LaTeX (`$..$` inline, `$$..$$` block)
— giữ nguyên ngữ nghĩa toán học. MathType/Equation 3.0 là OLE object và ảnh chụp công
thức là ảnh → mất hoàn toàn.

✅ **Đúng:** Insert → Equation, gõ công thức trực tiếp.
❌ **Sai:** dùng MathType, Equation 3.0, hoặc dán ảnh chụp công thức.

### DOCX-06 — Resolve mọi tracked change trước khi nộp `[BẮT BUỘC]`

**Tại sao:** file còn tracked changes cho output không dự đoán được (nội dung lấy ra
phụ thuộc cách mammoth xử lý phần thêm/xóa đang treo).

✅ **Đúng:** Review → Accept All Changes → tắt Track Changes → save.

### DOCX-07 — Không dùng text box `[TRÁNH]`

**Tại sao:** mammoth bỏ qua nội dung trong text box — chữ biến mất không cảnh báo.

✅ **Thay bằng:** đoạn văn thường; nếu cần đóng khung nhấn mạnh, dùng bảng 1 ô.

### DOCX-08 — Không dùng SmartArt, WordArt, shape có chữ `[TRÁNH]`

**Tại sao:** chữ bên trong các đối tượng đồ họa không được trích xuất.

✅ **Thay bằng:** bullet list (quy trình), bảng (so sánh), heading (phân cấp).

### DOCX-09 — Không đặt nội dung trong header/footer `[TRÁNH]`

**Tại sao:** MarkItDown chỉ xử lý thân tài liệu, footnote và endnote. Header/footer bị
bỏ qua (kể cả công thức toán đặt trong đó).

✅ **Thay bằng:** thông tin định danh tài liệu (mã dự án, phiên bản…) đặt ở đầu thân bài.

### DOCX-10 — Không đặt nội dung trong comment `[TRÁNH]`

**Tại sao:** comment không được convert.

✅ **Thay bằng:** ghi chú cần giữ lại → chuyển thành đoạn "Lưu ý:" trong thân bài.

### DOCX-11 — Dùng bullet/numbering chuẩn của Word `[NÊN]`

**Tại sao:** list style chuẩn → list Markdown đúng cấu trúc, giữ được cấp lồng nhau.
Gõ tay "- ", "1)" chỉ là văn bản thường.

✅ **Đúng:** Home → Bullets / Numbering.
❌ **Sai:** gõ tay dấu gạch đầu dòng.

### DOCX-12 — Link chèn bằng Insert > Link; footnote dùng được `[NÊN]`

**Tại sao:** hyperlink thật → `[chữ](url)` trong Markdown. Footnote/endnote được hỗ
trợ (kể cả công thức trong đó), nhưng nội dung chính vẫn nên nằm ở thân bài.

✅ **Đúng:** bôi đen chữ → Insert → Link → dán URL.

---

## Checklist trước khi nộp file Word

| ID | Quy tắc | Mức |
|---|---|---|
| DOCX-01 | Tiêu đề dùng Heading Style | BẮT BUỘC |
| DOCX-02 | Bảng không merge cell, 1 dòng header | BẮT BUỘC |
| DOCX-03 | Mọi ảnh có alt text mô tả đủ | BẮT BUỘC |
| DOCX-04 | Thông tin quan trọng có bản chữ | BẮT BUỘC |
| DOCX-05 | Công thức dùng Insert > Equation | BẮT BUỘC |
| DOCX-06 | Đã resolve hết tracked changes | BẮT BUỘC |
| DOCX-07 | Không text box | TRÁNH |
| DOCX-08 | Không SmartArt/WordArt/shape có chữ | TRÁNH |
| DOCX-09 | Không nội dung trong header/footer | TRÁNH |
| DOCX-10 | Không nội dung trong comment | TRÁNH |
| DOCX-11 | List dùng bullet/numbering chuẩn | NÊN |
| DOCX-12 | Link qua Insert > Link | NÊN |
````

- [ ] **Step 2: Đối chiếu với output thật**

Đọc lại `samples/output/docx-good.md` và `docx-bad.md`; xác nhận mô tả trong DOCX-01,
DOCX-02, DOCX-03 khớp hiện tượng thật. Bổ sung ví dụ ❌ của DOCX-02 bằng ghi chú hành
vi merge từ Task 2 Step 4 nếu hiện tượng đủ rõ để trích dẫn.

- [ ] **Step 3: Commit**

```bash
git add convention-docx.md
git commit -m "docs: add Word convention (12 rules) backed by sample evidence"
```

---

### Task 6: Viết `convention-xlsx.md`

**Files:**
- Create: `convention-xlsx.md`

**Interfaces:**
- Consumes: `samples/output/xlsx-good.md`, `samples/output/xlsx-bad.md` (Task 3).
- Produces: file `convention-xlsx.md` mà `README.md` (Task 8) sẽ link tới.

- [ ] **Step 1: Viết file với nội dung sau** (sửa theo output thật nếu Task 3 phát hiện khác):

````markdown
# Convention: Excel (.xlsx)

Quy tắc tạo file Excel để MarkItDown convert sang Markdown không mất thông tin.

**Cơ chế:** MarkItDown đọc bằng `pandas`: **mọi sheet (kể cả sheet ẩn)** trở thành
`## <tên sheet>` + một bảng Markdown. Nó giả định mỗi sheet là **một bảng duy nhất bắt
đầu tại A1, dòng 1 là header**. Đây là converter đơn giản nhất trong ba định dạng —
mọi thứ nằm ngoài giả định đó đều vỡ.

**Mức độ:** `[BẮT BUỘC]` vi phạm là mất/hỏng dữ liệu · `[NÊN]` cải thiện chất lượng
output · `[TRÁNH]` cấu trúc mà converter không xử lý được.

**Bằng chứng:** `samples/xlsx-good.xlsx` / `samples/xlsx-bad.xlsx` và output thật tại
`samples/output/xlsx-good.md` / `samples/output/xlsx-bad.md`.

---

### XLSX-01 — 1 sheet = 1 bảng duy nhất, bắt đầu tại A1 `[BẮT BUỘC]`

**Tại sao:** pandas đọc cả sheet thành đúng một bảng. Bảng đặt lệch vị trí sinh cột
`Unnamed` và dòng `NaN` rác.

✅ **Đúng:** ô A1 là ô đầu tiên của dòng header.
❌ **Sai:** chừa vài dòng trống trang trí, bảng bắt đầu tại B3.

### XLSX-02 — Dòng 1 là header thật `[BẮT BUỘC]`

**Tại sao:** dòng đầu tiên có dữ liệu trở thành tên cột cho cả bảng. Dòng tiêu đề trang
trí ("BÁO CÁO QUÝ 1") sẽ chiếm chỗ header — toàn bộ cột thành `Unnamed: 1`, `Unnamed: 2`…

✅ **Đúng:** dòng 1 = tên cột; tên báo cáo đặt vào **tên sheet** hoặc tên file.
❌ **Sai:** dòng 1 là tên báo cáo, dòng 3 mới là header (xem `samples/output/xlsx-bad.md`).

### XLSX-03 — Cấm merge cell `[BẮT BUỘC]`

**Tại sao:** giá trị chỉ nằm ở ô trên-trái của vùng merge; các ô còn lại rỗng và render
thành chữ `NaN` trong output — dòng dữ liệu lệch nghĩa.

✅ **Đúng:** lặp lại giá trị ở từng dòng (vd cột "Miền" ghi "Bắc" cho cả 4 dòng).
❌ **Sai:** merge dọc ô "Bắc" cho 4 dòng.

### XLSX-04 — Công thức phải có cached value `[BẮT BUỘC]`

**Tại sao:** MarkItDown chỉ đọc **giá trị đã tính** mà Excel lưu trong file ở lần save
cuối, không đọc công thức. File sinh bằng code (openpyxl, script xuất báo cáo…) chưa
từng được Excel mở/tính → ô công thức **rỗng hoàn toàn** trong output.

✅ **Đúng:** trước khi nộp, mở file bằng Excel và Ctrl+S. File sinh bằng code: ghi
giá trị đã tính sẵn, không ghi chuỗi công thức.
❌ **Sai:** script xuất file có `=SUM(...)` rồi nộp thẳng (xem ô "Gấp ba" biến mất
trong `samples/output/xlsx-bad.md`).

### XLSX-05 — Chart/pivot/comment/màu sắc không được là nơi duy nhất chứa thông tin `[BẮT BUỘC]`

**Tại sao:** chart, pivot table, comment, màu nền, conditional formatting bị bỏ qua
hoàn toàn. Quy ước "tô đỏ = quá hạn" mất sạch sau convert.

✅ **Đúng:** thêm cột "Trạng thái" bằng chữ; bảng dữ liệu nguồn của chart phải hiện
diện trong sheet.
❌ **Sai:** thông tin chỉ tồn tại trong biểu đồ hoặc màu ô.

### XLSX-06 — Xóa sheet nháp và sheet ẩn trước khi nộp `[BẮT BUỘC]`

**Tại sao:** mọi sheet đều được convert, **kể cả sheet ẩn** — dữ liệu nháp/nội bộ sẽ
lọt vào nguồn tri thức của AI (xem sheet `NhapLieuTam` lộ ra trong
`samples/output/xlsx-bad.md`).

✅ **Đúng:** xóa hẳn sheet nháp; đừng chỉ Hide.

### XLSX-07 — Tên sheet ngắn, có nghĩa `[NÊN]`

**Tại sao:** tên sheet trở thành heading `##` — là tiêu đề section trong nguồn tri thức.

✅ **Đúng:** `DoanhThu2026`, `NhanSu-Q3`.
❌ **Sai:** `Sheet1`, `Copy of Final (2)`.

### XLSX-08 — Không để ô trống giữa bảng `[NÊN]`

**Tại sao:** ô trống render thành chữ `NaN` — gây nhiễu và mơ hồ cho LLM.

✅ **Đúng:** điền `0`, `N/A`, hoặc "Chưa có" tùy ngữ nghĩa.

### XLSX-09 — Hiểu rằng output là giá trị thô, không phải giá trị hiển thị `[NÊN]`

**Tại sao:** converter lấy giá trị gốc, bỏ định dạng hiển thị: ngày ra
`2026-07-03 00:00:00`, ô 15% ra `0.15`.

✅ **Đúng:** chấp nhận dạng thô; nếu bắt buộc cần dạng hiển thị đẹp, nhập dạng text.

### XLSX-10 — Không nhiều bảng con hoặc ghi chú xen kẽ trong 1 sheet `[TRÁNH]`

**Tại sao:** pandas không nhận biết ranh giới bảng — các bảng con và dòng ghi chú bị
trộn thành một bảng sai cấu trúc.

✅ **Thay bằng:** mỗi bảng một sheet riêng; ghi chú đặt thành cột "Ghi chú".

---

## Checklist trước khi nộp file Excel

| ID | Quy tắc | Mức |
|---|---|---|
| XLSX-01 | 1 sheet = 1 bảng từ A1 | BẮT BUỘC |
| XLSX-02 | Dòng 1 là header thật | BẮT BUỘC |
| XLSX-03 | Không merge cell | BẮT BUỘC |
| XLSX-04 | Công thức có cached value (save bằng Excel) | BẮT BUỘC |
| XLSX-05 | Không giấu thông tin trong chart/màu/comment | BẮT BUỘC |
| XLSX-06 | Đã xóa sheet nháp/ẩn | BẮT BUỘC |
| XLSX-07 | Tên sheet có nghĩa | NÊN |
| XLSX-08 | Không ô trống giữa bảng | NÊN |
| XLSX-09 | Ý thức về giá trị thô (ngày, %) | NÊN |
| XLSX-10 | Không nhiều bảng con trong 1 sheet | TRÁNH |
````

- [ ] **Step 2: Đối chiếu với output thật**

Đọc `samples/output/xlsx-bad.md`: xác nhận có `Unnamed`, `NaN`, sheet ẩn lộ ra, và ô
công thức trống — đúng như XLSX-01/02/03/04/06 mô tả. Sửa câu chữ nếu hiện tượng khác.

- [ ] **Step 3: Commit**

```bash
git add convention-xlsx.md
git commit -m "docs: add Excel convention (10 rules) backed by sample evidence"
```

---

### Task 7: Viết `convention-pptx.md`

**Files:**
- Create: `convention-pptx.md`

**Interfaces:**
- Consumes: `samples/output/pptx-good.md`, `samples/output/pptx-bad.md` (Task 4).
- Produces: file `convention-pptx.md` mà `README.md` (Task 8) sẽ link tới.

- [ ] **Step 1: Viết file với nội dung sau** (sửa theo output thật nếu Task 4 phát hiện khác):

````markdown
# Convention: PowerPoint (.pptx)

Quy tắc tạo file PowerPoint để MarkItDown convert sang Markdown không mất thông tin.

**Cơ chế:** MarkItDown dùng `python-pptx`, duyệt shape theo **tọa độ (trên→dưới,
trái→phải)**. Title placeholder → `#`; speaker notes → `### Notes:`; chart cơ bản →
bảng dữ liệu Markdown. **Không xử lý SmartArt, không OCR chữ trong ảnh.**

**Mức độ:** `[BẮT BUỘC]` vi phạm là mất/hỏng dữ liệu · `[NÊN]` cải thiện chất lượng
output · `[TRÁNH]` tính năng bị bỏ qua hoàn toàn.

**Bằng chứng:** `samples/pptx-good.pptx` / `samples/pptx-bad.pptx` và output thật tại
`samples/output/pptx-good.md` / `samples/output/pptx-bad.md`.

---

### PPTX-01 — Dùng layout có Title placeholder, mỗi slide đúng 1 tiêu đề `[BẮT BUỘC]`

**Tại sao:** chỉ Title placeholder mới thành heading `#`. Textbox tự vẽ làm "tiêu đề"
chỉ là văn bản thường — mất ranh giới nội dung giữa các slide, RAG không chunk được.

✅ **Đúng:** New Slide → chọn layout "Title and Content" / "Title Only", gõ vào ô title.
❌ **Sai:** slide Blank + textbox chữ to làm tiêu đề.

### PPTX-02 — Alt text cho mọi hình ảnh `[BẮT BUỘC]`

**Tại sao:** ảnh chỉ còn `![alt](tên-file)` trong output — alt text (hoặc tên shape)
là thứ duy nhất sống sót. Pipeline không bật LLM caption.

✅ **Đúng:** Right-click ảnh → Edit Alt Text → mô tả đầy đủ nội dung/ý nghĩa.
❌ **Sai:** dán ảnh không alt text (output chỉ còn tên mặc định kiểu "Picture 4").

### PPTX-03 — Cấm SmartArt `[BẮT BUỘC]`

**Tại sao:** converter không có code xử lý SmartArt — **toàn bộ chữ bên trong biến
mất**. Đây là nguồn mất dữ liệu lớn nhất của PowerPoint.

✅ **Thay bằng:** bullet list (quy trình, chu trình), bảng (so sánh, ma trận).

### PPTX-04 — Chữ quan trọng không nằm trong screenshot/ảnh `[BẮT BUỘC]`

**Tại sao:** không có OCR. Slide chụp màn hình số liệu = slide rỗng sau convert.

✅ **Đúng:** gõ lại số liệu chính thành bullet/bảng, ảnh chỉ để minh họa (kèm alt text).

### PPTX-05 — Bố trí shape theo trục đọc trên→dưới, trái→phải `[NÊN]`

**Tại sao:** shape được sort theo tọa độ (top, left). Đặt lệch thứ tự thị giác làm văn
bản đảo lộn (xem "Bước 2" đứng trước "Bước 1" trong `samples/output/pptx-bad.md`).

✅ **Đúng:** nội dung đọc trước đặt cao hơn/trái hơn; tránh chồng lấn shape.

### PPTX-06 — Chart dùng loại cơ bản `[NÊN]`

**Tại sao:** chart column/bar/line/pie được chuyển thành bảng dữ liệu Markdown (giữ
được số liệu!). Loại phức tạp không hỗ trợ chỉ còn `[unsupported chart]`.

✅ **Đúng:** dùng chart cơ bản, đặt tên series/category rõ nghĩa.
❌ **Sai:** chart combo/3D lạ — mất toàn bộ số liệu.

### PPTX-07 — Dùng speaker notes cho nội dung chi tiết `[NÊN]`

**Tại sao:** notes được convert (mục `Notes:` cuối mỗi slide) — là nơi tốt nhất để
viết đoạn văn đầy đủ ngữ cảnh cho RAG mà không làm rối slide.

✅ **Đúng:** slide để ý chính ngắn gọn; diễn giải đầy đủ viết vào Notes.

### PPTX-08 — Bảng đơn giản, không merge cell `[NÊN]`

**Tại sao:** như mọi định dạng khác, bảng Markdown không biểu diễn được ô gộp.

### PPTX-09 — Mỗi slide 1 chủ đề `[NÊN]`

**Tại sao:** mỗi slide là một khối nội dung có tiêu đề riêng trong output — slide đơn
chủ đề giúp chunking RAG sạch hơn.

### PPTX-10 — Không WordArt trang trí, không icon mang nghĩa mà thiếu chữ `[TRÁNH]`

**Tại sao:** icon/hình trang trí không có chữ đi kèm thì không để lại dấu vết ngữ
nghĩa nào trong output.

✅ **Thay bằng:** icon + nhãn chữ bên cạnh.

---

## Checklist trước khi nộp file PowerPoint

| ID | Quy tắc | Mức |
|---|---|---|
| PPTX-01 | Layout có Title placeholder | BẮT BUỘC |
| PPTX-02 | Mọi ảnh có alt text | BẮT BUỘC |
| PPTX-03 | Không SmartArt | BẮT BUỘC |
| PPTX-04 | Chữ quan trọng không nằm trong ảnh | BẮT BUỘC |
| PPTX-05 | Shape xếp theo trục đọc | NÊN |
| PPTX-06 | Chart loại cơ bản | NÊN |
| PPTX-07 | Nội dung chi tiết vào speaker notes | NÊN |
| PPTX-08 | Bảng không merge cell | NÊN |
| PPTX-09 | Mỗi slide 1 chủ đề | NÊN |
| PPTX-10 | Không WordArt/icon thiếu nhãn chữ | TRÁNH |
````

- [ ] **Step 2: Đối chiếu với output thật**

Đọc `samples/output/pptx-good.md` / `pptx-bad.md`: xác nhận title thành `#`, notes xuất
hiện, chart thành bảng, thứ tự đảo ở bản bad. Đặc biệt kiểm tra format thật của mục
notes (`### Notes:` hay khác) và sửa PPTX-07 cho khớp từng ký tự.

- [ ] **Step 3: Commit**

```bash
git add convention-pptx.md
git commit -m "docs: add PowerPoint convention (10 rules) backed by sample evidence"
```

---

### Task 8: Viết `README.md` + kiểm tra chéo toàn bộ

**Files:**
- Create: `README.md`
- Verify: `convention-docx.md`, `convention-xlsx.md`, `convention-pptx.md`, `samples/`

**Interfaces:**
- Consumes: 3 file convention (Task 5–7), samples (Task 2–4).

- [ ] **Step 1: Viết `README.md` với nội dung sau:**

````markdown
# Convention tạo file Office cho pipeline MarkItDown → Markdown

Tài liệu Office (.docx, .xlsx, .pptx) của team được convert tự động sang Markdown bằng
[MarkItDown](https://github.com/microsoft/markitdown) để làm nguồn tri thức cho LLM/RAG.
Bộ convention này giúp file bạn tạo ra **không mất thông tin** khi đi qua bước convert.

## Ba nguyên tắc gốc

Mọi quy tắc chi tiết đều bắt nguồn từ ba nguyên tắc này:

1. **Thông tin quan trọng phải tồn tại ở dạng chữ** — không chỉ ở hình ảnh, màu sắc
   hay vị trí trình bày. Converter không có OCR và bỏ qua mọi định dạng thị giác.
2. **Dùng cấu trúc ngữ nghĩa, không định dạng thủ công** — Heading Style thay vì chữ
   to in đậm, Title placeholder thay vì textbox, bullet chuẩn thay vì gõ tay gạch đầu dòng.
3. **File phải "tự đứng" khi bị bóc hết trình bày** — tự hỏi: nếu file này mất toàn bộ
   hình, màu, vị trí, nội dung còn lại có đọc hiểu được không?

## Mức độ quy tắc

| Mức | Ý nghĩa |
|---|---|
| `[BẮT BUỘC]` | Vi phạm là mất hoặc sai lệch dữ liệu trong output |
| `[NÊN]` | Cải thiện đáng kể chất lượng output |
| `[TRÁNH]` | Tính năng Office bị MarkItDown bỏ qua hoàn toàn |

## Quy tắc theo định dạng

| Định dạng | File | Số quy tắc |
|---|---|---|
| Word | [convention-docx.md](convention-docx.md) | 12 |
| Excel | [convention-xlsx.md](convention-xlsx.md) | 10 |
| PowerPoint | [convention-pptx.md](convention-pptx.md) | 10 |

Cuối mỗi file có bảng checklist tóm tắt để rà nhanh trước khi nộp. Khi review, viện
dẫn quy tắc bằng mã: "file này vi phạm DOCX-01".

## Pipeline

Pipeline chạy MarkItDown **mặc định, không bật LLM caption** — vì vậy alt text cho
hình ảnh là bắt buộc (DOCX-03, PPTX-02). Nếu sau này pipeline bật `llm_client` để AI
tự mô tả ảnh, hai quy tắc đó có thể nới lỏng cho pptx (docx vẫn cần alt text).

## Kiểm chứng

Mọi quy tắc chính được chứng minh bằng cặp file mẫu good/bad trong [`samples/`](samples/)
kèm output MarkItDown thật trong [`samples/output/`](samples/output/). Tái tạo:

```bash
python -m venv .venv
.venv/Scripts/python -m pip install "markitdown[docx,xlsx,pptx]" python-docx openpyxl python-pptx
.venv/Scripts/python samples/generate_samples.py
.venv/Scripts/python samples/convert_and_verify.py   # kỳ vọng: ALL PASS
```

Nếu nâng version MarkItDown, chạy lại lệnh verify — assert nào fail nghĩa là hành vi
converter đã đổi và convention cần cập nhật (thực nghiệm thắng tài liệu).
````

- [ ] **Step 2: Kiểm tra chéo toàn bộ**

Checklist (làm thủ công, từng mục):
- 4 link trong README trỏ tới file tồn tại.
- Số quy tắc thực tế trong 3 file convention = 12/10/10, khớp bảng README và spec §3.3.
- Mọi ID quy tắc duy nhất, đúng format, mức độ chỉ dùng 3 giá trị chuẩn.
- Chạy lại lần cuối: `.venv/Scripts/python samples/convert_and_verify.py` → `ALL PASS`.

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: add README with core principles, rule index and verification guide"
```

---

## Self-Review Notes (đã chạy khi viết plan)

- **Spec coverage:** 4 file tài liệu (Task 5–8) ✔; samples + output commit (Task 2–4) ✔; 32 quy tắc verbatim từ spec §3.3 ✔; nguyên tắc "thực nghiệm thắng tài liệu" nhúng vào Step verify của Task 2–4 và Step đối chiếu của Task 5–7 ✔. Bonus: XLSX-06 (sheet ẩn) demo được bằng sample dù spec xếp vào nhóm "không demo" — plan demo luôn.
- **Placeholder scan:** không còn TBD/TODO; mọi step có code hoặc lệnh cụ thể.
- **Type consistency:** `convert(name: str) -> str`, `gen_*() -> None`, `verify_*() -> None`, `SAMPLES_DIR`, `PNG_1PX`, `OUTPUT_DIR` dùng thống nhất Task 2→4.
- **Rủi ro đã biết:** một số assert dựa trên format output cụ thể (vd `### Notes:`, `Unnamed`) — nếu version MarkItDown mới đổi format, sửa assert + câu chữ doc theo output thật, đúng nguyên tắc thực nghiệm.
