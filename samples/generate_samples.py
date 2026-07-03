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
    # "Không áp dụng" thay vì "N/A": pandas coi literal "N/A" là giá trị khuyết -> NaN
    ws.append(["Tổng", 600, "Không áp dụng"])
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


def main() -> None:
    (SAMPLES_DIR / "output").mkdir(exist_ok=True)
    gen_docx()
    print("generated: docx")
    gen_xlsx()
    print("generated: xlsx")


if __name__ == "__main__":
    main()
