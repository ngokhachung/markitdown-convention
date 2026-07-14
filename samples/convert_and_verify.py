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
    # Quan sát thực tế (Step 4): MarkItDown chuyển oMath a^2 + b^2 thành đúng
    # "$a^{2} + b^{2}$" (LaTeX inline, có dấu cách quanh dấu +, mũ dùng {}).
    assert "$a^{2} + b^{2}$" in good, "good: OMML->LaTeX"
    assert "$a" not in bad, "bad: công thức dạng ảnh -> không có LaTeX"

    # DOCX-06: tracked insertion đang treo ở bad để lại dấu vết bất định
    assert "[vi phạm DOCX-06]" in bad, "bad: câu có tracked change vẫn còn"
    # (ghi chú: nội dung phần chèn có/không tuỳ mammoth — không assert cứng)

    # DOCX-07: text box -> chữ sống sót nhưng bị dồn ra cuối đoạn (rối thứ tự đọc)
    assert "[vi phạm DOCX-07]" in bad, "bad: note vi phạm text box còn"
    assert "ĐẠT 600 TRIỆU" in bad, "bad: chữ trong text box vẫn sống sót"
    assert bad.index("đã kiểm toán") < bad.index("ĐẠT 600 TRIỆU"), \
        "bad: nội dung text box bị dồn ra sau -> sai thứ tự đọc"
    assert "[DOCX-07]" in good and "KHUNG NHẤN MẠNH" in good, "good: bảng 1 ô thay text box"

    # DOCX-08: SmartArt/đồ hoạ DrawingML không-fallback -> chữ MẤT
    assert "[vi phạm DOCX-08]" in bad, "bad: note vi phạm còn"
    assert "KHẢO SÁT" not in bad, "bad: chữ trong SmartArt/đồ hoạ phải mất"
    assert "[DOCX-08]" in good, "good: bullet thay SmartArt"

    # DOCX-09: header content biến mất ở bad; good đặt ở thân bài
    assert "[DOCX-09]" in good, "good: định danh ở đầu thân bài"
    assert "vi phạm DOCX-09" not in bad, "bad: chữ trong header phải biến mất"

    # DOCX-10: ghi chú trong thân bài (good) hiển thị
    assert "[DOCX-10]" in good and "tạm tính" in good, "good: ghi chú trong thân bài"

    # DOCX-11: List style -> list markdown; gõ tay -> văn bản thường
    # Quan sát thực tế (Step 4): List Bullet -> tiền tố "* " (dấu sao + cách).
    assert "* Bước khảo sát khách hàng [DOCX-11]" in good, "good: list markdown"

    # DOCX-12: hyperlink thật -> [text](url)
    assert "[báo cáo đầy đủ](https://example.com/bao-cao)" in good, "good: hyperlink markdown"

    print("PASS docx")


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

    # XLSX-09: giá trị thô ngày + % — pin đúng Ô Ngày/Tỷ lệ bằng assert row-scoped,
    # không để lọt qua chữ trong ô nhãn.
    # Quan sát thực tế: pandas.read_excel trả về Timestamp('2026-07-03 00:00:00'),
    # nhưng MarkItDown render qua DataFrame.to_html() — pandas rút gọn cột datetime
    # toàn giờ-phút-giây = 00:00:00 thành CHỈ ngày, không có "00:00:00".
    # -> raw output thực tế là "2026-07-03", KHÔNG PHẢI "2026-07-03 00:00:00".
    assert "| Chốt Q1 | 2026-07-03 | 0.15 |" in good, "good: ngày & % ra giá trị thô"

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
    # XLSX-10: bảng con thứ 2 xen kẽ trong 1 sheet vẫn lọt vào output (cấu trúc sai)
    assert "Chi phí [vi phạm XLSX-10" in bad, "bad: bảng con thứ 2 chung sheet"

    print("PASS xlsx")


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
    assert "[unsupported chart]" in bad, "bad: chart 3D không ra số liệu"

    # PPTX-07: speaker notes -> ### Notes:
    assert "Notes:" in good and "50 khách hàng" in good, "good: notes convert"

    # PPTX-08: bảng không merge ở good ra bảng markdown đủ ô
    assert "| Tháng |" in good, "good: bảng markdown"

    # PPTX-09: slide good 1 chủ đề (có nhiều slide); bad có note nhồi chủ đề
    assert "[vi phạm PPTX-09]" in bad, "bad: note nhồi chủ đề"

    # PPTX-10: icon kèm nhãn chữ ở good hiển thị
    assert "[PPTX-10]" in good, "good: icon + nhãn chữ"

    print("PASS pptx")


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    verify_docx()
    verify_xlsx()
    verify_pptx()
    print("ALL PASS")


if __name__ == "__main__":
    main()
