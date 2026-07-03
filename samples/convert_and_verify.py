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
    (OUTPUT_DIR / (Path(name).stem + ".md")).write_text(
        text, encoding="utf-8", newline="\n"
    )
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
    # DOCX-02: dòng có w:tblHeader (Table Properties -> Row -> "Repeat as header
    # row") được mammoth map thành <th> -> ra header Markdown thật, không rỗng.
    assert "| Tháng | Doanh thu |\n| --- | --- |" in good, (
        "good: dòng có tblHeader phải ra header Markdown thật"
    )
    # bad: không set tblHeader -> dòng đầu vẫn bị đẩy xuống thân bảng, header rỗng
    assert "|  |  |\n| --- | --- |" in bad, (
        "bad: không tblHeader thì header Markdown vẫn rỗng"
    )
    print("PASS docx")


def verify_xlsx() -> None:
    good = convert("xlsx-good.xlsx")
    bad = convert("xlsx-bad.xlsx")
    # XLSX-07: tên sheet thành heading ##
    assert "## DoanhThu2026" in good, "good: tên sheet phải thành heading"
    # XLSX-08: bảng sạch không có NaN. Lưu ý (phát hiện thực nghiệm): pandas coi các
    # literal "N/A"/"NA"/"NULL"/"None"/"n/a"/"nan"... là giá trị khuyết -> NaN dù ô
    # KHÔNG trống, nên file good phải dùng "Không áp dụng" thay vì "N/A" ở dòng Tổng.
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
    # XLSX-08: literal "N/A" (cột "Ghi chú") bị pandas coi là giá trị khuyết ->
    # biến mất hoàn toàn, chỉ còn lại NaN, dù ô đó không hề trống.
    assert "N/A" not in bad, "bad: literal N/A phải bị coerce thành NaN"
    print("PASS xlsx")


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
    # PPTX-02: ảnh không alt text (descr rỗng) ra alt RỖNG, không phải placeholder
    # literal "image.png" — link target là tên shape tự sinh, không phải tên file gốc.
    assert "![](" in bad, "bad: ảnh không alt phải ra alt rỗng"
    assert "image.png" not in bad, "bad: không được có placeholder literal image.png"
    print("PASS pptx")


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    verify_docx()
    verify_xlsx()
    verify_pptx()
    print("ALL PASS")


if __name__ == "__main__":
    main()
