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

Mọi quy tắc (32/32) được chứng minh bằng cặp file mẫu với nhãn mã quy tắc inline trong chính nội dung good/bad trong [`samples/`](samples/)
kèm output MarkItDown thật trong [`samples/output/`](samples/output/).  
Mỗi phần tử trong sample Office mang nhãn inline (vd `[DOCX-01]`, `[vi phạm PPTX-01]`) để tra chéo nhanh giữa convention và bằng chứng.
Tái tạo:

```bash
python -m venv .venv
.venv/Scripts/python -m pip install "markitdown[docx,xlsx,pptx]" python-docx openpyxl python-pptx
.venv/Scripts/python samples/generate_samples.py
.venv/Scripts/python samples/convert_and_verify.py   # kỳ vọng: ALL PASS
```

Trên macOS/Linux, đường dẫn interpreter trong `.venv` là `.venv/bin/python` thay vì
`.venv/Scripts/python`.

Nếu nâng version MarkItDown, chạy lại lệnh verify — assert nào fail nghĩa là hành vi
converter đã đổi và convention cần cập nhật (thực nghiệm thắng tài liệu).
