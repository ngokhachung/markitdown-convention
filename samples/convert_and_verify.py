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
    print("PASS xlsx")


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    verify_docx()
    verify_xlsx()
    print("ALL PASS")


if __name__ == "__main__":
    main()
