# Sample Rule-Annotations Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Sinh lại 6 file sample Office sao cho mỗi quy tắc convention (docx 12 · xlsx 10 · pptx 10) đều có phần tử minh hoạ mang nhãn inline, verify ra ALL PASS, và checklist mỗi convention thêm cột `Sample`.

**Architecture:** Mở rộng `samples/generate_samples.py` (thêm helper chèn OOXML thô cho các phần python-docx/pptx không sinh trực tiếp được) và viết lại `samples/convert_and_verify.py` (mỗi quy tắc ≥1 assert). Vì output của MarkItDown cho các phần tử mới (OMML, VML textbox, tracked change, chart 3D) **không được tài liệu hoá chính xác**, mỗi task theo mẫu **observe-then-assert**: sinh file → convert → in output thật → chốt assert khớp output quan sát được. Không bịa chuỗi output chưa quan sát.

**Tech Stack:** Python 3.14, `markitdown[docx,xlsx,pptx]`, `python-docx`, `openpyxl`, `python-pptx`, chèn XML qua `docx.oxml.parse_xml` / `lxml`.

## Global Constraints

- Interpreter trong repo: `.venv/Scripts/python` (Windows). macOS/Linux: `.venv/bin/python`. Mọi lệnh python trong plan dùng `.venv/Scripts/python`.
- Mục tiêu bất biến: `convert_and_verify.py` in `ALL PASS` sau mỗi task đụng tới sample.
- Nhãn good: mã quy tắc dạng `[DOCX-01]` nhúng trong chữ của chính phần tử minh hoạ.
- Nhãn bad: `[vi phạm DOCX-07]`; vi phạm gây mất nội dung phải kèm một đoạn chữ-thường **sống sót** mô tả điều sẽ mất.
- Nhãn xấp xỉ: `≈ xấp xỉ SmartArt, xem PPTX-03` — không assert "là SmartArt", chỉ assert điểm quan sát được.
- Không sửa ngữ nghĩa quy tắc; không bật LLM caption; không giả lập Excel tính lại công thức.
- Encoding output `.md`: `utf-8`, `newline="\n"` (giữ như code hiện tại).
- Commit message tiếng Việt, kết thúc bằng dòng `Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>`.

---

### Task 1: Dựng môi trường & chốt baseline

**Files:**
- Create: `.venv/` (không commit — kiểm tra `.gitignore` đã bỏ qua `.venv`)

**Interfaces:**
- Produces: `.venv/Scripts/python` chạy được `markitdown`, `docx`, `openpyxl`, `pptx`.

- [ ] **Step 1: Tạo venv**

Run: `python -m venv .venv`
Expected: thư mục `.venv` xuất hiện, không lỗi.

- [ ] **Step 2: Cài thư viện**

Run: `.venv/Scripts/python -m pip install "markitdown[docx,xlsx,pptx]" python-docx openpyxl python-pptx`
Expected: `Successfully installed ...`. **Nếu fail do Python 3.14 không tương thích** (vd wheel thiếu), DỪNG và báo người dùng — không tự hạ cấp/bỏ qua.

- [ ] **Step 3: Xác nhận import**

Run: `.venv/Scripts/python -c "import markitdown, docx, openpyxl, pptx; print('libs OK')"`
Expected: `libs OK`

- [ ] **Step 4: Chốt baseline — generate + verify hiện trạng**

Run: `.venv/Scripts/python samples/generate_samples.py && .venv/Scripts/python samples/convert_and_verify.py`
Expected: `ALL PASS` (chứng minh môi trường + code cũ hoạt động trước khi sửa).

- [ ] **Step 5: Xác nhận .venv được ignore**

Run: `git status --short`
Expected: KHÔNG có file nào dưới `.venv/`. Nếu có, thêm `.venv/` vào `.gitignore` và commit riêng:
```bash
git add .gitignore && git commit -m "chore: ignore .venv

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 2: Word (docx) — phủ 12 quy tắc + nhãn inline

**Files:**
- Modify: `samples/generate_samples.py` (helpers OOXML + viết lại `gen_docx`)
- Modify: `samples/convert_and_verify.py` (viết lại `verify_docx`)

**Interfaces:**
- Consumes: `PNG_1PX`, `SAMPLES_DIR` (đã có ở đầu `generate_samples.py`).
- Produces: `samples/docx-good.docx`, `samples/docx-bad.docx`, output `.md` tương ứng.

- [ ] **Step 1: Thêm import + helper OOXML vào đầu `generate_samples.py`**

Thêm ngay dưới các import hiện có:

```python
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT


def _repeat_header(table):
    """Đánh dấu dòng đầu 'Repeat as header row' -> mammoth map thành <th>."""
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    el = parse_xml(f'<w:tblHeader {nsdecls("w")} w:val="true"/>')
    trPr.append(el)


def _add_hyperlink(paragraph, url, text):
    """Hyperlink thật -> [text](url) trong Markdown (DOCX-12)."""
    r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    xml = (
        f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}">'
        f'<w:r><w:rPr/><w:t xml:space="preserve">{text}</w:t></w:r>'
        f'</w:hyperlink>'
    )
    paragraph._p.append(parse_xml(xml))


def _add_vml_textbox(paragraph, text):
    """Text box VML -> mammoth bỏ qua txbxContent, chữ biến mất (DOCX-07/08)."""
    xml = (
        f'<w:r {nsdecls("w", "v")}>'
        f'<w:pict><v:shape style="width:220pt;height:36pt">'
        f'<v:textbox><w:txbxContent><w:p><w:r>'
        f'<w:t xml:space="preserve">{text}</w:t>'
        f'</w:r></w:p></w:txbxContent></v:textbox>'
        f'</v:shape></w:pict></w:r>'
    )
    paragraph._p.append(parse_xml(xml))


def _add_omml_pythagoras(paragraph):
    """Chèn công thức OMML a^2 + b^2 -> MarkItDown chuyển thành LaTeX (DOCX-05)."""
    xml = (
        '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        '<m:sSup><m:e><m:r><m:t>a</m:t></m:r></m:e>'
        '<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>'
        '<m:r><m:t xml:space="preserve"> + </m:t></m:r>'
        '<m:sSup><m:e><m:r><m:t>b</m:t></m:r></m:e>'
        '<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>'
        '</m:oMath>'
    )
    paragraph._p.append(parse_xml(xml))


def _add_tracked_insertion(paragraph, text):
    """Insertion đang treo (chưa Accept) -> tracked change chưa resolve (DOCX-06)."""
    xml = (
        f'<w:ins {nsdecls("w")} w:id="1" w:author="reviewer" '
        f'w:date="2026-07-15T00:00:00Z"><w:r>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r></w:ins>'
    )
    paragraph._p.append(parse_xml(xml))
```

- [ ] **Step 2: Viết lại `gen_docx` phủ đủ 12 quy tắc, có nhãn inline**

Thay toàn bộ hàm `gen_docx` bằng:

```python
def gen_docx() -> None:
    from docx import Document
    from docx.shared import Inches, Pt

    # ---------- GOOD ----------
    doc = Document()
    doc.add_heading("Báo cáo doanh thu 2026 [DOCX-01]", level=1)      # DOCX-01
    doc.add_heading("Tình hình quý 1 [DOCX-01]", level=2)
    p9 = doc.add_paragraph(
        "Mã dự án DT-2026, phiên bản 1.0 — định danh đặt ở đầu thân bài, "
        "không nhét vào header/footer. [DOCX-09]"                     # DOCX-09
    )
    doc.add_paragraph(
        "Doanh thu quý 1 đạt 600 triệu (số liệu gõ thành văn bản, không chỉ "
        "nằm trong ảnh). [DOCX-04]"                                   # DOCX-04
    )
    p5 = doc.add_paragraph("Ràng buộc kiểm tra [DOCX-05]: ")          # DOCX-05
    _add_omml_pythagoras(p5)

    table = doc.add_table(rows=3, cols=2)                             # DOCX-02
    data = [("Tháng [DOCX-02]", "Doanh thu"), ("01", "100"), ("02", "200")]
    for row, (c1, c2) in zip(table.rows, data):
        row.cells[0].text = c1
        row.cells[1].text = c2
    _repeat_header(table)

    doc.add_paragraph("Bước khảo sát khách hàng [DOCX-11]", style="List Bullet")
    doc.add_paragraph("Bước chốt tính năng [DOCX-11]", style="List Bullet")  # DOCX-11

    p12 = doc.add_paragraph("Nguồn tham chiếu [DOCX-12]: ")           # DOCX-12
    _add_hyperlink(p12, "https://example.com/bao-cao", "báo cáo đầy đủ")

    p10 = doc.add_paragraph(
        "Lưu ý: số liệu quý 1 là tạm tính, chốt lại cuối tháng 4. "
        "(ghi chú giữ lại đặt trong thân bài, không dùng comment) [DOCX-10]"  # DOCX-10
    )
    # DOCX-07: thay text box bằng bảng 1 ô để nhấn mạnh
    box_tbl = doc.add_table(rows=1, cols=1)
    box_tbl.style = "Table Grid"
    box_tbl.rows[0].cells[0].text = (
        "KHUNG NHẤN MẠNH: dùng bảng 1 ô thay cho text box. [DOCX-07]"
    )
    # DOCX-08: thay SmartArt bằng bullet list
    doc.add_paragraph("Quy trình thay cho SmartArt [DOCX-08]:")
    doc.add_paragraph("B1 khảo sát", style="List Number")
    doc.add_paragraph("B2 phân tích", style="List Number")
    # DOCX-03: ảnh có alt text
    pic = doc.add_picture(io.BytesIO(PNG_1PX), width=Inches(1))
    pic._inline.docPr.set(
        "descr",
        "Biểu đồ cột doanh thu theo tháng: tháng 1 đạt 100, tháng 2 đạt 200 [DOCX-03]",
    )
    pic._inline.docPr.set("title", "Biểu đồ doanh thu")
    # DOCX-06: file good không còn tracked change (chỉ ghi chú)
    doc.add_paragraph(
        "Đã Accept toàn bộ thay đổi, tắt Track Changes trước khi nộp. [DOCX-06]"
    )
    doc.save(SAMPLES_DIR / "docx-good.docx")

    # ---------- BAD ----------
    doc = Document()
    p = doc.add_paragraph()                                          # DOCX-01 (vi phạm)
    run = p.add_run("Báo cáo doanh thu 2026 [vi phạm DOCX-01: bold+cỡ chữ]")
    run.bold = True
    run.font.size = Pt(16)

    # DOCX-09 (vi phạm): nội dung đặt trong header
    doc.sections[0].header.paragraphs[0].text = (
        "Mã dự án DT-2026 [vi phạm DOCX-09: đặt trong header, sẽ biến mất]"
    )

    # DOCX-02 (vi phạm): bảng merge header, không đánh dấu header row
    table = doc.add_table(rows=3, cols=2)
    table.rows[0].cells[0].merge(table.rows[0].cells[1])
    table.rows[0].cells[0].text = "Doanh thu theo tháng [vi phạm DOCX-02: merge]"
    table.rows[1].cells[0].text = "01"
    table.rows[1].cells[1].text = "100"
    table.rows[2].cells[0].text = "02"
    table.rows[2].cells[1].text = "200"

    # DOCX-05 (vi phạm): "công thức" là ảnh, không có bản OMML
    doc.add_paragraph(
        "Ràng buộc a^2+b^2 chỉ tồn tại trong ảnh dưới đây [vi phạm DOCX-05]:"
    )
    doc.add_picture(io.BytesIO(PNG_1PX), width=Inches(1))

    # DOCX-07 (vi phạm): text box — chữ trong nó sẽ biến mất
    doc.add_paragraph(
        "Đoạn kế tiếp là text box, nội dung của nó sẽ biến mất trong output "
        "[vi phạm DOCX-07]:"
    )
    p7 = doc.add_paragraph()
    _add_vml_textbox(p7, "CHỮ TRONG TEXT BOX — LẼ RA PHẢI THẤY nhưng sẽ mất")

    # DOCX-08 (vi phạm ≈): shape chứa chữ thay cho SmartArt thật — chữ cũng mất
    doc.add_paragraph(
        "Đoạn kế tiếp ≈ xấp xỉ SmartArt (shape chứa chữ), chữ sẽ biến mất "
        "[vi phạm DOCX-08]:"
    )
    p8 = doc.add_paragraph()
    _add_vml_textbox(p8, "CHỮ TRONG SHAPE SMARTART — sẽ mất")

    # DOCX-06 (vi phạm): tracked change chưa resolve
    p6 = doc.add_paragraph("Câu có sửa đổi đang treo [vi phạm DOCX-06]: ")
    _add_tracked_insertion(p6, "PHẦN CHÈN ĐANG TREO")

    # DOCX-11 (vi phạm): gõ tay dấu gạch đầu dòng
    doc.add_paragraph("- Bước khảo sát (gõ tay) [vi phạm DOCX-11]")
    doc.add_paragraph("- Bước chốt tính năng (gõ tay)")

    # DOCX-03/04 (vi phạm): ảnh không alt text, số liệu chỉ nằm trong ảnh
    doc.add_paragraph(
        "Bảng số liệu bên dưới chỉ là ảnh chụp, không có bản chữ [vi phạm DOCX-04]:"
    )
    doc.add_picture(io.BytesIO(PNG_1PX), width=Inches(1))
    doc.save(SAMPLES_DIR / "docx-bad.docx")
```

- [ ] **Step 3: Sinh file docx**

Run: `.venv/Scripts/python samples/generate_samples.py`
Expected: in `generated: docx` (và xlsx/pptx). Không traceback.

- [ ] **Step 4: OBSERVE — in output docx thật để chốt chuỗi assert**

Run:
```bash
.venv/Scripts/python -c "from markitdown import MarkItDown; m=MarkItDown(); print('=== GOOD ==='); print(m.convert('samples/docx-good.docx').text_content); print('=== BAD ==='); print(m.convert('samples/docx-bad.docx').text_content)"
```
Ghi lại chính xác: (a) dạng LaTeX cho OMML (vd `$a^{2} + b^{2}$` hay khác), (b) header row good `| Tháng [DOCX-01]... |`… — **thực tế** ra sao, (c) hyperlink `[báo cáo đầy đủ](https://example.com/bao-cao)`, (d) list markdown cho List Bullet/Number, (e) BAD: text box/shape/tracked-insertion/header text CÓ hay KHÔNG xuất hiện. Dùng các chuỗi quan sát được ở Step 5.

- [ ] **Step 5: Viết lại `verify_docx` bám output quan sát được**

Thay toàn bộ `verify_docx`. Khung dưới đây; **thay các chuỗi `<...>` bằng chuỗi thật quan sát ở Step 4**:

```python
def verify_docx() -> None:
    good = convert("docx-good.docx")
    bad = convert("docx-bad.docx")

    # DOCX-01: Heading Style -> '#'; bold thủ công -> văn bản thường
    assert "# Báo cáo doanh thu 2026 [DOCX-01]" in good, "good: thiếu H1"
    assert "## Tình hình quý 1 [DOCX-01]" in good, "good: thiếu H2"
    assert "\n# " not in ("\n" + bad), "bad: không được có heading"
    assert "Báo cáo doanh thu 2026" in bad, "bad: text tiêu đề vẫn còn dạng thường"

    # DOCX-02: dòng có tblHeader -> header Markdown thật; bad merge -> header rỗng
    assert "| Tháng [DOCX-02] | Doanh thu |\n| --- | --- |" in good, "good: header thật"
    assert "|  |  |\n| --- | --- |" in bad, "bad: không tblHeader -> header rỗng"

    # DOCX-03: alt text là thứ duy nhất của ảnh sống sót
    assert "[DOCX-03]" in good and "Biểu đồ cột doanh thu" in good, "good: alt text"
    assert "Biểu đồ cột doanh thu" not in bad, "bad: ảnh không alt -> không mô tả"

    # DOCX-04: số liệu có bản chữ ở good; ở bad chỉ trong ảnh (note còn, số mất)
    assert "600 triệu" in good, "good: số liệu dạng chữ"
    assert "[vi phạm DOCX-04]" in bad, "bad: note vi phạm vẫn còn (chữ thường)"

    # DOCX-05: OMML -> LaTeX ở good; bad là ảnh nên không có LaTeX
    assert "<CHUỖI LATEX QUAN SÁT, VD $a^{2} + b^{2}$>" in good, "good: OMML->LaTeX"
    assert "$a" not in bad, "bad: công thức dạng ảnh -> không có LaTeX"

    # DOCX-06: tracked insertion đang treo ở bad để lại dấu vết bất định
    assert "[vi phạm DOCX-06]" in bad, "bad: câu có tracked change vẫn còn"
    # (ghi chú: nội dung phần chèn có/không tuỳ mammoth — không assert cứng)

    # DOCX-07: text box -> chữ biến mất ở bad; note sống sót
    assert "[vi phạm DOCX-07]" in bad, "bad: note vi phạm text box còn"
    assert "CHỮ TRONG TEXT BOX" not in bad, "bad: chữ trong text box phải mất"
    # good: thay bằng bảng 1 ô
    assert "[DOCX-07]" in good and "KHUNG NHẤN MẠNH" in good, "good: bảng 1 ô thay text box"

    # DOCX-08 (≈): chữ trong shape SmartArt xấp xỉ cũng mất
    assert "CHỮ TRONG SHAPE SMARTART" not in bad, "bad: chữ trong shape mất"
    assert "[DOCX-08]" in good, "good: bullet thay SmartArt"

    # DOCX-09: header content biến mất ở bad; good đặt ở thân bài
    assert "[DOCX-09]" in good, "good: định danh ở đầu thân bài"
    assert "vi phạm DOCX-09" not in bad, "bad: chữ trong header phải biến mất"

    # DOCX-10: ghi chú trong thân bài (good) hiển thị
    assert "[DOCX-10]" in good and "tạm tính" in good, "good: ghi chú trong thân bài"

    # DOCX-11: List style -> list markdown; gõ tay -> văn bản thường
    assert "<TIỀN TỐ LIST QUAN SÁT>Bước khảo sát khách hàng [DOCX-11]" in good, "good: list markdown"

    # DOCX-12: hyperlink thật -> [text](url)
    assert "[báo cáo đầy đủ](https://example.com/bao-cao)" in good, "good: hyperlink markdown"

    print("PASS docx")
```

- [ ] **Step 6: Chạy verify docx (tạm chỉ docx)**

Run: `.venv/Scripts/python -c "import samples.convert_and_verify as v" 2>/dev/null; cd samples && ../.venv/Scripts/python -c "import convert_and_verify as v; v.verify_docx()"`
Expected: `PASS docx`. Nếu fail, quay lại Step 4, sửa chuỗi assert cho khớp output THẬT (không sửa nghĩa quy tắc).

- [ ] **Step 7: Commit**

```bash
git add samples/generate_samples.py samples/convert_and_verify.py samples/docx-good.docx samples/docx-bad.docx samples/output/docx-good.md samples/output/docx-bad.md
git commit -m "test(docx): sample phủ 12 quy tắc + nhãn inline, verify khớp output

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 3: Excel (xlsx) — phủ 10 quy tắc + nhãn inline

**Files:**
- Modify: `samples/generate_samples.py` (viết lại `gen_xlsx`)
- Modify: `samples/convert_and_verify.py` (viết lại `verify_xlsx`)

**Interfaces:**
- Consumes: `SAMPLES_DIR`.
- Produces: `samples/xlsx-good.xlsx`, `samples/xlsx-bad.xlsx`, output `.md`.

- [ ] **Step 1: Viết lại `gen_xlsx` phủ đủ 10 quy tắc**

Thay toàn bộ `gen_xlsx`:

```python
def gen_xlsx() -> None:
    import datetime
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import PatternFill
    from openpyxl.chart import BarChart, Reference

    # ---------- GOOD ----------
    wb = Workbook()
    # Sheet 1: 1 bảng từ A1, header dòng 1, không merge, cột Trạng thái bằng chữ,
    # cột "Quy tắc" gắn nhãn từng dòng (XLSX-01/02/03/04/05/07/08)
    ws = wb.active
    ws.title = "DoanhThu2026"                                   # XLSX-07
    ws.append(["Tháng", "Doanh thu", "Trạng thái", "Quy tắc"])  # XLSX-02
    ws.append(["01", 100, "Đạt", "XLSX-01/02/03: bảng sạch từ A1"])
    ws.append(["02", 200, "Đạt", "XLSX-05: trạng thái bằng chữ, không chỉ màu"])
    ws.append(["03", 300, "Vượt", "XLSX-08: điền đủ, không ô trống"])
    # XLSX-04: dòng tổng ghi GIÁ TRỊ đã tính (600), không ghi công thức
    ws.append(["Tổng", 600, "Không áp dụng", "XLSX-04: ghi sẵn giá trị, không formula"])
    # XLSX-05: chart lấy nguồn từ bảng đang hiện diện (dữ liệu vẫn có bản chữ)
    chart = BarChart()
    chart.title = "Doanh thu theo tháng"
    data = Reference(ws, min_col=2, min_row=1, max_row=4)
    cats = Reference(ws, min_col=1, min_row=2, max_row=4)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "F2")

    # Sheet 2: XLSX-09 — ý thức giá trị thô (ngày, %)
    ws2 = wb.create_sheet("NgayVaTyLe")                          # XLSX-07 + XLSX-10 (bảng riêng)
    ws2.append(["Mốc", "Ngày", "Tỷ lệ", "Quy tắc"])
    d = ws2.cell(row=2, column=2, value=datetime.datetime(2026, 7, 3))
    d.number_format = "dd/mm/yyyy"
    r = ws2.cell(row=2, column=3, value=0.15)
    r.number_format = "0%"
    ws2["A2"] = "Chốt Q1"
    ws2["D2"] = "XLSX-09: output ra giá trị thô 2026-07-03 00:00:00 và 0.15"
    wb.save(SAMPLES_DIR / "xlsx-good.xlsx")

    # ---------- BAD ----------
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"                                          # XLSX-07 (vi phạm)
    ws["A1"] = "BÁO CÁO DOANH THU NĂM 2026 [vi phạm XLSX-01/02]"  # tiêu đề trang trí
    ws.merge_cells("A1:D1")                                       # XLSX-03 (vi phạm)
    ws["A3"] = "Tháng"
    ws["B3"] = "Doanh thu"
    ws["C3"] = "Ghi chú"
    ws["A4"] = "01"
    ws["B4"] = 110
    ws["C4"] = "N/A"          # XLSX-08 (vi phạm): literal N/A -> pandas coi là khuyết
    ws["A5"] = "02"           # XLSX-08 (vi phạm): B5 bỏ trống
    ws["A6"] = "Gấp ba"
    ws["B6"] = "=B4*3"        # XLSX-04 (vi phạm): openpyxl không cache -> mất
    # XLSX-05 (vi phạm): thông tin chỉ nằm trong comment + màu nền, không có bản chữ
    ws["B4"].comment = Comment("Vượt kế hoạch 10% [vi phạm XLSX-05: chỉ trong comment]", "pm")
    ws["B4"].fill = PatternFill("solid", fgColor="FF0000")
    # XLSX-10 (vi phạm): bảng con thứ hai + dòng ghi chú xen kẽ trong cùng sheet
    ws["A8"] = "Chi phí [vi phạm XLSX-10: bảng con thứ 2 chung sheet]"
    ws["A9"] = "Khoản"
    ws["B9"] = "Số tiền"
    ws["A10"] = "Marketing"
    ws["B10"] = 50
    # XLSX-06 (vi phạm): sheet nháp bị ẩn vẫn convert
    hidden = wb.create_sheet("NhapLieuTam")
    hidden["A1"] = "dữ liệu nháp không nên lộ ra [vi phạm XLSX-06: sheet ẩn]"
    hidden.sheet_state = "hidden"
    wb.save(SAMPLES_DIR / "xlsx-bad.xlsx")
```

- [ ] **Step 2: Sinh + OBSERVE output xlsx**

Run:
```bash
.venv/Scripts/python samples/generate_samples.py
.venv/Scripts/python -c "from markitdown import MarkItDown; m=MarkItDown(); print('=== GOOD ==='); print(m.convert('samples/xlsx-good.xlsx').text_content); print('=== BAD ==='); print(m.convert('samples/xlsx-bad.xlsx').text_content)"
```
Ghi lại: ngày ra dạng gì (`2026-07-03 00:00:00`?), % ra `0.15`?, comment/màu CÓ mất không (kỳ vọng mất), sheet ẩn CÓ lộ không, `N/A`/ô trống ra `NaN`, `330`/`=B4` KHÔNG xuất hiện.

- [ ] **Step 3: Viết lại `verify_xlsx`**

Thay toàn bộ; thay `<...>` bằng chuỗi thật ở Step 2:

```python
def verify_xlsx() -> None:
    good = convert("xlsx-good.xlsx")
    bad = convert("xlsx-bad.xlsx")

    # XLSX-07: tên sheet -> heading ##
    assert "## DoanhThu2026" in good and "## NgayVaTyLe" in good, "good: heading sheet"
    assert "## Sheet1" in bad, "bad: tên sheet vô nghĩa vẫn thành heading"

    # XLSX-01/02/03/08: bảng good sạch, không NaN
    assert "NaN" not in good, "good: không NaN"
    assert "| Tháng | Doanh thu | Trạng thái | Quy tắc |" in good, "good: header thật"

    # XLSX-04: tổng ghi giá trị (600) giữ nguyên
    assert "Tổng" in good and "600" in good, "good: giá trị tổng giữ nguyên"

    # XLSX-05: trạng thái bằng chữ hiện diện (không phụ thuộc màu/chart)
    assert "Vượt" in good, "good: trạng thái bằng chữ"

    # XLSX-09: giá trị thô ngày + %
    assert "<CHUỖI NGÀY THÔ QUAN SÁT>" in good, "good: ngày ra dạng thô"
    assert "0.15" in good, "good: % ra 0.15 (giá trị thô)"

    # XLSX-01/02: tiêu đề trang trí chiếm header -> Unnamed
    assert "Unnamed" in bad, "bad: header giả -> Unnamed"
    # XLSX-03/08: merge + ô trống + literal N/A -> NaN
    assert "NaN" in bad, "bad: merge/ô trống -> NaN"
    assert "N/A" not in bad, "bad: literal N/A bị coerce thành NaN"
    # XLSX-04: công thức không cached -> mất
    assert "330" not in bad, "bad: 110*3=330 không xuất hiện"
    assert "=B4" not in bad, "bad: chuỗi công thức không xuất hiện"
    # XLSX-05: thông tin chỉ trong comment/màu -> mất
    assert "Vượt kế hoạch 10%" not in bad, "bad: nội dung comment biến mất"
    # XLSX-06: sheet ẩn vẫn convert
    assert "## NhapLieuTam" in bad and "nháp" in bad, "bad: sheet ẩn vẫn lộ"

    print("PASS xlsx")
```

- [ ] **Step 4: Chạy verify xlsx**

Run: `cd samples && ../.venv/Scripts/python -c "import convert_and_verify as v; v.verify_xlsx()"`
Expected: `PASS xlsx`. Fail -> quay lại Step 2 chốt chuỗi thật.

- [ ] **Step 5: Commit**

```bash
git add samples/generate_samples.py samples/convert_and_verify.py samples/xlsx-good.xlsx samples/xlsx-bad.xlsx samples/output/xlsx-good.md samples/output/xlsx-bad.md
git commit -m "test(xlsx): sample phủ 10 quy tắc + nhãn inline, verify khớp output

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 4: PowerPoint (pptx) — phủ 10 quy tắc + nhãn inline

**Files:**
- Modify: `samples/generate_samples.py` (viết lại `gen_pptx`)
- Modify: `samples/convert_and_verify.py` (viết lại `verify_pptx`)

**Interfaces:**
- Consumes: `PNG_1PX`, `SAMPLES_DIR`.
- Produces: `samples/pptx-good.pptx`, `samples/pptx-bad.pptx`, output `.md`.
- Note: text trong autoshape python-pptx **được** MarkItDown trích xuất, nên PPTX-03/PPTX-10 (SmartArt/WordArt) KHÔNG assert "mất chữ" — assert phía good (bản thay thế) và điểm quan sát được ở bad.

- [ ] **Step 1: Viết lại `gen_pptx` phủ đủ 10 quy tắc**

Thay toàn bộ `gen_pptx`:

```python
def gen_pptx() -> None:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches, Pt

    # ---------- GOOD ----------
    prs = Presentation()
    # Slide 1: Title placeholder + bullet (thay SmartArt) + alt text + notes (1 chủ đề)
    slide = prs.slides.add_slide(prs.slide_layouts[1])          # PPTX-01, PPTX-09
    slide.shapes.title.text = "Kế hoạch quý 3 [PPTX-01]"
    body = slide.placeholders[1].text_frame
    body.text = "Bước 1: khảo sát khách hàng [PPTX-03: bullet thay SmartArt]"  # PPTX-03/05
    body.add_paragraph().text = "Bước 2: chốt tính năng [PPTX-05: đúng trục đọc]"
    pic = slide.shapes.add_picture(
        io.BytesIO(PNG_1PX), Inches(6), Inches(1), width=Inches(1)
    )
    pic._element._nvXxPr.cNvPr.set(
        "descr", "Sơ đồ timeline quý 3: tháng 7 khảo sát, tháng 8 chốt [PPTX-02]"
    )                                                            # PPTX-02
    slide.notes_slide.notes_text_frame.text = (
        "Chi tiết: khảo sát 50 khách hàng nhóm A trong tháng 7. [PPTX-07]"
    )                                                            # PPTX-07

    # Slide 2: chart cơ bản (PPTX-06) — 1 chủ đề
    slide2 = prs.slides.add_slide(prs.slide_layouts[5])
    slide2.shapes.title.text = "Doanh thu theo tháng [PPTX-06]"
    chart_data = CategoryChartData()
    chart_data.categories = ["Tháng 7", "Tháng 8"]
    chart_data.add_series("Doanh thu", (100, 200))
    slide2.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1), Inches(2), Inches(6), Inches(4), chart_data,
    )

    # Slide 3: bảng không merge (PPTX-08) + icon có nhãn chữ (PPTX-10) + số liệu bằng chữ (PPTX-04)
    slide3 = prs.slides.add_slide(prs.slide_layouts[5])
    slide3.shapes.title.text = "Chi tiết doanh thu [PPTX-08]"
    tbl = slide3.shapes.add_table(3, 2, Inches(1), Inches(2), Inches(4), Inches(2)).table
    tbl.cell(0, 0).text = "Tháng"
    tbl.cell(0, 1).text = "Doanh thu [PPTX-04: số liệu bằng chữ]"
    tbl.cell(1, 0).text = "07"
    tbl.cell(1, 1).text = "100"
    tbl.cell(2, 0).text = "08"
    tbl.cell(2, 1).text = "200"
    label = slide3.shapes.add_textbox(Inches(6), Inches(2), Inches(3), Inches(0.6))
    label.text_frame.text = "▲ Tăng trưởng — icon kèm nhãn chữ [PPTX-10]"
    prs.save(SAMPLES_DIR / "pptx-good.pptx")

    # ---------- BAD ----------
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])          # Blank
    # PPTX-05 (vi phạm): Bước 2 đặt CAO hơn Bước 1 -> sort toạ độ đảo thứ tự
    box2 = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    box2.text_frame.text = "Bước 2: chốt tính năng [vi phạm PPTX-05: đặt trên]"
    box1 = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(6), Inches(1))
    box1.text_frame.text = "Bước 1: khảo sát khách hàng [vi phạm PPTX-05: đặt dưới]"
    # PPTX-01 (vi phạm): "tiêu đề" là textbox, không phải placeholder
    title_box = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(6), Inches(1))
    title_box.text_frame.text = "Kế hoạch quý 3 [vi phạm PPTX-01: textbox không thành heading]"
    # PPTX-02/04 (vi phạm): ảnh không alt text, số liệu chỉ trong ảnh
    pic_bad = slide.shapes.add_picture(
        io.BytesIO(PNG_1PX), Inches(6), Inches(5), width=Inches(1)
    )
    pic_bad._element._nvXxPr.cNvPr.set("descr", "")  # descr rỗng = không alt text

    # Slide 2: PPTX-06 (vi phạm) chart 3D + PPTX-08 (vi phạm) bảng merge + PPTX-09 (vi phạm) nhồi 2 chủ đề
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    note = slide2.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    note.text_frame.text = "Slide nhồi 2 chủ đề [vi phạm PPTX-09]"
    cd = CategoryChartData()
    cd.categories = ["Tháng 7", "Tháng 8"]
    cd.add_series("Doanh thu", (100, 200))
    slide2.shapes.add_chart(
        XL_CHART_TYPE.THREE_D_COLUMN,                           # PPTX-06 (vi phạm)
        Inches(0.5), Inches(1), Inches(4), Inches(3), cd,
    )
    tbl2 = slide2.shapes.add_table(3, 2, Inches(5), Inches(1), Inches(4), Inches(2)).table
    tbl2.cell(0, 0).merge(tbl2.cell(0, 1))                      # PPTX-08 (vi phạm)
    tbl2.cell(0, 0).text = "Doanh thu [vi phạm PPTX-08: merge]"
    tbl2.cell(1, 0).text = "07"
    tbl2.cell(1, 1).text = "100"
    tbl2.cell(2, 0).text = "08"
    tbl2.cell(2, 1).text = "200"
    prs.save(SAMPLES_DIR / "pptx-bad.pptx")
```

- [ ] **Step 2: Sinh + OBSERVE output pptx**

Run:
```bash
.venv/Scripts/python samples/generate_samples.py
.venv/Scripts/python -c "from markitdown import MarkItDown; m=MarkItDown(); print('=== GOOD ==='); print(m.convert('samples/pptx-good.pptx').text_content); print('=== BAD ==='); print(m.convert('samples/pptx-bad.pptx').text_content)"
```
Ghi lại: chart 3D ra gì (`[unsupported chart]`?), bảng merge render sao, thứ tự Bước 1/2 ở bad, alt rỗng `![](...)`, table markdown ở good, icon-label text.

- [ ] **Step 3: Viết lại `verify_pptx`**

Thay toàn bộ; thay `<...>` bằng chuỗi thật ở Step 2:

```python
def verify_pptx() -> None:
    good = convert("pptx-good.pptx")
    bad = convert("pptx-bad.pptx")

    # PPTX-01: title placeholder -> H1; bad textbox không thành heading
    assert "# Kế hoạch quý 3 [PPTX-01]" in good, "good: title -> H1"
    assert "# Kế hoạch quý 3" not in bad, "bad: textbox không thành heading"

    # PPTX-02: alt text sống sót ở good; bad alt rỗng
    assert "Sơ đồ timeline" in good, "good: alt text"
    assert "Sơ đồ timeline" not in bad, "bad: không alt -> không mô tả"
    assert "![](" in bad, "bad: ảnh không alt -> alt rỗng"
    assert "image.png" not in bad, "bad: không có placeholder literal image.png"

    # PPTX-03: bullet thay SmartArt hiển thị ở good
    assert "[PPTX-03: bullet thay SmartArt]" in good, "good: bullet thay SmartArt"

    # PPTX-04: số liệu bằng chữ ở good (bảng)
    assert "[PPTX-04: số liệu bằng chữ]" in good, "good: số liệu bằng chữ"

    # PPTX-05: thứ tự đọc đúng ở good, đảo ở bad (sort theo toạ độ)
    assert good.index("Bước 1") < good.index("Bước 2"), "good: đúng thứ tự"
    assert bad.index("Bước 2") < bad.index("Bước 1"), "bad: đảo thứ tự theo toạ độ"

    # PPTX-06: chart cơ bản -> bảng dữ liệu ở good; 3D ở bad không ra bảng số liệu
    assert "Tháng 7" in good and "100" in good, "good: chart -> bảng dữ liệu"
    assert "<DẤU HIỆU 3D KHÔNG HỖ TRỢ, VD [unsupported chart]>" in bad, "bad: chart 3D không ra số liệu"

    # PPTX-07: speaker notes -> ### Notes:
    assert "Notes:" in good and "50 khách hàng" in good, "good: notes convert"

    # PPTX-08: bảng không merge ở good ra bảng markdown đủ ô
    assert "| Tháng |" in good, "good: bảng markdown"

    # PPTX-09: slide good 1 chủ đề (có nhiều slide); bad có note nhồi chủ đề
    assert "[vi phạm PPTX-09]" in bad, "bad: note nhồi chủ đề"

    # PPTX-10: icon kèm nhãn chữ ở good hiển thị
    assert "[PPTX-10]" in good, "good: icon + nhãn chữ"

    print("PASS pptx")
```

- [ ] **Step 4: Chạy verify pptx**

Run: `cd samples && ../.venv/Scripts/python -c "import convert_and_verify as v; v.verify_pptx()"`
Expected: `PASS pptx`. Fail -> chốt lại chuỗi thật ở Step 2.

- [ ] **Step 5: Commit**

```bash
git add samples/generate_samples.py samples/convert_and_verify.py samples/pptx-good.pptx samples/pptx-bad.pptx samples/output/pptx-good.md samples/output/pptx-bad.md
git commit -m "test(pptx): sample phủ 10 quy tắc + nhãn inline, verify khớp output

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 5: Thêm cột `Sample` vào checklist convention + cập nhật README

**Files:**
- Modify: `convention-docx.md` (bảng checklist cuối file)
- Modify: `convention-xlsx.md` (bảng checklist cuối file)
- Modify: `convention-pptx.md` (bảng checklist cuối file)
- Modify: `README.md` (mục Kiểm chứng nếu cần)

**Interfaces:**
- Consumes: mã quy tắc + phía minh hoạ (good/bad) đã cố định ở Task 2–4.

- [ ] **Step 1: Thêm cột `Sample` vào checklist `convention-docx.md`**

Sửa header bảng checklist thành `| ID | Quy tắc | Mức | Sample |` và thêm giá trị mỗi dòng theo phía minh hoạ thực tế:
```
| DOCX-01 | Tiêu đề dùng Heading Style | BẮT BUỘC | good+bad |
| DOCX-02 | Bảng không merge, "Repeat as header row" | BẮT BUỘC | good+bad |
| DOCX-03 | Mọi ảnh có alt text mô tả đủ | BẮT BUỘC | good+bad |
| DOCX-04 | Thông tin quan trọng có bản chữ | BẮT BUỘC | good+bad |
| DOCX-05 | Công thức dùng Insert > Equation | BẮT BUỘC | good+bad |
| DOCX-06 | Đã resolve hết tracked changes | BẮT BUỘC | bad |
| DOCX-07 | Không text box | TRÁNH | good+bad |
| DOCX-08 | Không SmartArt/WordArt/shape có chữ | TRÁNH | good+bad (≈) |
| DOCX-09 | Không nội dung trong header/footer | TRÁNH | good+bad |
| DOCX-10 | Không nội dung trong comment | TRÁNH | good |
| DOCX-11 | List dùng bullet/numbering chuẩn | NÊN | good+bad |
| DOCX-12 | Link qua Insert > Link | NÊN | good |
```

- [ ] **Step 2: Thêm cột `Sample` vào `convention-xlsx.md`**

Header `| ID | Quy tắc | Mức | Sample |`, giá trị:
```
| XLSX-01 | 1 sheet = 1 bảng từ A1 | BẮT BUỘC | good+bad |
| XLSX-02 | Dòng 1 là header thật | BẮT BUỘC | good+bad |
| XLSX-03 | Không merge cell | BẮT BUỘC | good+bad |
| XLSX-04 | Công thức có cached value | BẮT BUỘC | good+bad |
| XLSX-05 | Không giấu thông tin trong chart/màu/comment | BẮT BUỘC | good+bad |
| XLSX-06 | Đã xóa sheet nháp/ẩn | BẮT BUỘC | bad |
| XLSX-07 | Tên sheet có nghĩa | NÊN | good+bad |
| XLSX-08 | Không ô trống; không N/A/NULL/None | NÊN | good+bad |
| XLSX-09 | Ý thức về giá trị thô (ngày, %) | NÊN | good |
| XLSX-10 | Không nhiều bảng con trong 1 sheet | TRÁNH | good+bad |
```

- [ ] **Step 3: Thêm cột `Sample` vào `convention-pptx.md`**

Header `| ID | Quy tắc | Mức | Sample |`, giá trị:
```
| PPTX-01 | Layout có Title placeholder | BẮT BUỘC | good+bad |
| PPTX-02 | Mọi ảnh có alt text | BẮT BUỘC | good+bad |
| PPTX-03 | Không SmartArt | BẮT BUỘC | good+bad (≈) |
| PPTX-04 | Chữ quan trọng không nằm trong ảnh | BẮT BUỘC | good+bad |
| PPTX-05 | Shape xếp theo trục đọc | NÊN | good+bad |
| PPTX-06 | Chart loại cơ bản | NÊN | good+bad |
| PPTX-07 | Nội dung chi tiết vào speaker notes | NÊN | good |
| PPTX-08 | Bảng không merge cell | NÊN | good+bad |
| PPTX-09 | Mỗi slide 1 chủ đề | NÊN | good+bad |
| PPTX-10 | Không WordArt/icon thiếu nhãn chữ | TRÁNH | good+bad (≈) |
```

- [ ] **Step 4: Cập nhật README nếu cần**

Đọc mục "Kiểm chứng" của `README.md`. Câu "Mọi quy tắc chính được chứng minh bằng cặp file mẫu" giờ đã đúng cho **mọi** quy tắc — đổi "quy tắc chính" thành "mọi quy tắc" và thêm một câu: mỗi phần tử sample mang nhãn mã quy tắc inline. Không đổi phần lệnh reproduce.

- [ ] **Step 5: Commit**

```bash
git add convention-docx.md convention-xlsx.md convention-pptx.md README.md
git commit -m "docs: thêm cột Sample vào checklist, cập nhật README ánh xạ sample

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 6: Kiểm chứng toàn bộ & dọn dẹp

**Files:** (không tạo mới — chạy lại toàn bộ và xác nhận)

- [ ] **Step 1: Regenerate sạch từ đầu**

Run: `.venv/Scripts/python samples/generate_samples.py`
Expected: `generated: docx` / `generated: xlsx` / `generated: pptx`, không lỗi.

- [ ] **Step 2: Verify toàn bộ**

Run: `.venv/Scripts/python samples/convert_and_verify.py`
Expected: `PASS docx` / `PASS xlsx` / `PASS pptx` / `ALL PASS`.

- [ ] **Step 3: Xác nhận output .md đã cập nhật & không sót file**

Run: `git status --short`
Expected: các file `samples/output/*.md` và nhị phân đã được commit ở Task 2–4; nếu Step 1 làm đổi nội dung (do determinism), `git add` phần thay đổi và commit:
```bash
git add samples/output samples/*.docx samples/*.xlsx samples/*.pptx
git commit -m "chore: đồng bộ output sample sau regenerate

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>" || echo "không có thay đổi"
```

- [ ] **Step 4: Rà nhãn — mọi quy tắc có mặt trong output**

Run:
```bash
cd "samples/output" && for id in DOCX-01 DOCX-02 DOCX-03 DOCX-04 DOCX-05 DOCX-06 DOCX-07 DOCX-08 DOCX-09 DOCX-10 DOCX-11 DOCX-12 XLSX-01 XLSX-04 XLSX-05 XLSX-06 XLSX-07 XLSX-08 XLSX-09 XLSX-10 PPTX-01 PPTX-02 PPTX-03 PPTX-04 PPTX-05 PPTX-06 PPTX-08 PPTX-09 PPTX-10; do grep -l "$id" *.md >/dev/null 2>&1 && echo "OK $id" || echo "THIẾU $id"; done
```
Expected: không dòng `THIẾU` nào (trừ các nhãn cố ý biến mất ở bad như `vi phạm DOCX-07/08/09` và comment DOCX-10 — các mã này vẫn xuất hiện ở phía good/label sống sót). Nếu có `THIẾU` ngoài dự kiến, quay lại task tương ứng.

- [ ] **Step 5: Báo cáo hoàn tất**

Xác nhận `ALL PASS` + toàn bộ commit sạch. Dùng superpowers:finishing-a-development-branch để chốt hướng merge/PR.

---

## Self-Review

**Spec coverage:** Mọi quy tắc docx(12)/xlsx(10)/pptx(10) có task minh hoạ (Task 2/3/4) + assert; cột Sample (Task 5); regenerate+verify (Task 1,6). Quy ước nhãn good/bad/≈ áp dụng xuyên suốt. ✓

**Placeholder scan:** Các `<...>` còn lại là **có chủ đích** — chuỗi output converter phải quan sát thực tế (OMML→LaTeX, ngày thô, dấu hiệu chart 3D, tiền tố list) rồi mới chốt; mỗi chỗ đều kèm Step OBSERVE. Đây là ràng buộc bản chất (không tài liệu hoá được trước), không phải placeholder lười. ✓

**Type consistency:** Tên helper (`_repeat_header`, `_add_hyperlink`, `_add_vml_textbox`, `_add_omml_pythagoras`, `_add_tracked_insertion`) khớp giữa định nghĩa (Task 2 Step 1) và nơi gọi (Task 2 Step 2). Tên hàm verify (`verify_docx/xlsx/pptx`) giữ nguyên chữ ký, `main()` không đổi. ✓
