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
