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
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

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
    # Đánh dấu dòng đầu là "Repeat as header row at the top of each page"
    # (Table Properties -> Row) -> mammoth map dòng này thành <th>, ra header
    # Markdown thật.
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    trPr.append(tbl_header)
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
    ws["C3"] = "Ghi chú"
    ws["A4"] = "01"
    ws["B4"] = 110
    ws["C4"] = "N/A"         # literal "N/A" -> pandas coi là giá trị khuyết -> NaN
    ws["A5"] = "02"          # B5 cố tình bỏ trống
    ws["A6"] = "Gấp ba"
    ws["B6"] = "=B4*3"       # openpyxl không tính -> không có cached value
    hidden = wb.create_sheet("NhapLieuTam")
    hidden["A1"] = "dữ liệu nháp không nên lộ ra"
    hidden.sheet_state = "hidden"
    wb.save(SAMPLES_DIR / "xlsx-bad.xlsx")


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
    pic_bad = slide.shapes.add_picture(
        io.BytesIO(PNG_1PX), Inches(6), Inches(5), width=Inches(1)
    )
    # python-pptx tự viết descr="image.png" khi add_picture từ stream — đó là
    # artifact của python-pptx, KHÔNG phải hành vi MarkItDown. Xóa để mẫu phản
    # ánh đúng trường hợp "ảnh không có alt text" (descr rỗng) mà convention mô tả.
    pic_bad._element._nvXxPr.cNvPr.set("descr", "")
    prs.save(SAMPLES_DIR / "pptx-bad.pptx")


def main() -> None:
    (SAMPLES_DIR / "output").mkdir(exist_ok=True)
    gen_docx()
    print("generated: docx")
    gen_xlsx()
    print("generated: xlsx")
    gen_pptx()
    print("generated: pptx")


if __name__ == "__main__":
    main()
