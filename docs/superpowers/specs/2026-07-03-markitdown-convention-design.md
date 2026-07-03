# Design: Bộ convention tạo file Office cho MarkItDown

**Ngày:** 2026-07-03
**Trạng thái:** Đã duyệt thiết kế, chờ triển khai

## 1. Mục tiêu

Định nghĩa bộ tài liệu convention (tiếng Việt) hướng dẫn team member tạo file `.docx`, `.xlsx`, `.pptx` sao cho MarkItDown convert sang Markdown đạt chất lượng cao nhất, phục vụ pipeline **LLM/RAG**.

## 2. Bối cảnh & ràng buộc

- Pipeline chạy **MarkItDown mặc định, không bật LLM caption** (`llm_client` không được cấu hình) → hình ảnh chỉ còn lại alt text sau convert.
- Markdown output dùng làm nguồn tri thức cho AI (RAG/chatbot), không ưu tiên trình bày cho người đọc.
- Toàn bộ giới hạn kỹ thuật đã được xác minh trực tiếp từ mã nguồn MarkItDown (nhánh `main`, tháng 07/2026):
  - **DOCX**: dùng `mammoth` (docx → HTML → Markdown). Có bước tiền xử lý chuyển công thức OMML → LaTeX (`$...$` / `$$...$$`), chỉ áp dụng cho `document.xml`, `footnotes.xml`, `endnotes.xml`. Ảnh nhúng thành data URI bị cắt cụt (chỉ còn alt text). Text box, SmartArt, header/footer, comment bị bỏ qua. Heading chỉ nhận qua Style.
  - **XLSX**: dùng `pandas` + `openpyxl`/`xlrd`, đọc mọi sheet (kể cả sheet ẩn), mỗi sheet → `## tên sheet` + 1 bảng Markdown. Công thức chỉ lấy cached value; không có cached value → ô rỗng. Ô trống/merged cell render thành chữ "NaN". Chart, pivot, comment, màu sắc, conditional formatting mất hoàn toàn. Giả định: 1 bảng duy nhất từ A1, dòng đầu là header.
  - **PPTX**: dùng `python-pptx`. Shape được sort theo tọa độ (top, left); group xử lý đệ quy. Ảnh → `![alt](tên-file)`, alt lấy từ descr nhúng hoặc tên shape. Chart cơ bản → bảng dữ liệu Markdown, loại không hỗ trợ → `[unsupported chart]`. Speaker notes → `### Notes:`. Title placeholder → `#`. SmartArt không được xử lý (mất toàn bộ chữ). Không OCR chữ trong ảnh.

## 3. Sản phẩm

Đặt tại thư mục gốc `D:\project\MD convention`:

| Đường dẫn | Nội dung |
|---|---|
| `README.md` | Mục đích, 3 nguyên tắc gốc, quy ước mức độ, cách tra mã quy tắc, link các file convention, ghi chú pipeline |
| `convention-docx.md` | ~12 quy tắc cho Word |
| `convention-xlsx.md` | ~10 quy tắc cho Excel |
| `convention-pptx.md` | ~10 quy tắc cho PowerPoint |
| `samples/` | File mẫu đúng/sai + script sinh file + output Markdown thực tế để kiểm chứng |

### 3.1. Ba nguyên tắc gốc (trong README)

1. **Thông tin quan trọng phải tồn tại ở dạng chữ** — không chỉ ở hình, màu, layout.
2. **Dùng cấu trúc ngữ nghĩa** (Style, placeholder, list/table chuẩn) — không định dạng thủ công.
3. **File phải "tự đứng" sau khi mất toàn bộ yếu tố trình bày** — tưởng tượng file bị bóc hết hình/màu/vị trí, nội dung còn đọc được không.

### 3.2. Template mỗi quy tắc

```markdown
### DOCX-01 — Dùng Heading Style cho mọi tiêu đề  `[BẮT BUỘC]`

**Tại sao:** <cơ chế kỹ thuật của MarkItDown gây ra hạn chế này>

✅ **Đúng:** <thao tác cụ thể trong Office>
❌ **Sai:** <thao tác sai thường gặp>
```

Ba mức độ:
- `[BẮT BUỘC]` — vi phạm là mất hoặc hỏng dữ liệu trong output.
- `[NÊN]` — cải thiện đáng kể chất lượng output.
- `[TRÁNH]` — tính năng Office mà MarkItDown bỏ qua hoàn toàn.

Cuối mỗi file convention: **bảng checklist tóm tắt** (ID — quy tắc — mức độ).

### 3.3. Danh mục quy tắc

**DOCX (12):**

| ID | Mức | Quy tắc |
|---|---|---|
| DOCX-01 | BẮT BUỘC | Dùng Heading Style (Heading 1–6) cho mọi tiêu đề; cấm tiêu đề bằng bold/cỡ chữ thủ công |
| DOCX-02 | BẮT BUỘC | Bảng không merge cell; đúng 1 dòng header |
| DOCX-03 | BẮT BUỘC | Mọi hình ảnh phải có alt text mô tả đủ nội dung |
| DOCX-04 | BẮT BUỘC | Thông tin quan trọng không được chỉ tồn tại trong hình (ảnh chụp bảng, code, sơ đồ phải kèm bản chữ) |
| DOCX-05 | BẮT BUỘC | Công thức toán dùng Insert > Equation (OMML); cấm MathType/Equation 3.0/ảnh chụp công thức |
| DOCX-06 | BẮT BUỘC | Resolve mọi tracked change trước khi nộp |
| DOCX-07 | TRÁNH | Text box — nội dung bên trong bị bỏ qua |
| DOCX-08 | TRÁNH | SmartArt và shape có chữ — chữ bên trong mất |
| DOCX-09 | TRÁNH | Đặt nội dung (kể cả công thức) trong header/footer |
| DOCX-10 | TRÁNH | Đặt nội dung trong comment — chuyển vào thân bài |
| DOCX-11 | NÊN | Dùng bullet/numbering chuẩn của Word thay vì gõ tay "- ", "1)" |
| DOCX-12 | NÊN | Dùng hyperlink thật (Insert > Link); footnote/endnote dùng được nhưng nội dung chính đặt ở thân bài |

**XLSX (10):**

| ID | Mức | Quy tắc |
|---|---|---|
| XLSX-01 | BẮT BUỘC | 1 sheet = 1 bảng duy nhất, bắt đầu tại ô A1 |
| XLSX-02 | BẮT BUỘC | Dòng 1 là header; không dòng tiêu đề trang trí phía trên bảng |
| XLSX-03 | BẮT BUỘC | Cấm merge cell |
| XLSX-04 | BẮT BUỘC | File có công thức phải được mở & save bằng Excel trước khi nộp (để có cached value); file sinh bằng code phải ghi giá trị thay vì công thức |
| XLSX-05 | BẮT BUỘC | Chart/pivot/comment/màu nền/conditional formatting không được là nơi duy nhất chứa thông tin |
| XLSX-06 | BẮT BUỘC | Xóa sheet nháp và sheet ẩn trước khi nộp (mọi sheet đều bị convert) |
| XLSX-07 | NÊN | Tên sheet ngắn, có nghĩa (trở thành heading `##`) |
| XLSX-08 | NÊN | Không để ô trống giữa bảng — điền "0"/"N/A" (ô trống thành chữ "NaN") |
| XLSX-09 | NÊN | Hiểu cách xuất giá trị thô: ngày → `2026-07-03 00:00:00`, 15% → `0.15`; nếu cần dạng hiển thị, lưu dạng text |
| XLSX-10 | TRÁNH | Nhiều bảng con hoặc ghi chú xen kẽ trong 1 sheet — tách sheet riêng |

**PPTX (10):**

| ID | Mức | Quy tắc |
|---|---|---|
| PPTX-01 | BẮT BUỘC | Dùng layout chuẩn với Title placeholder (title → `#`); mỗi slide có đúng 1 tiêu đề |
| PPTX-02 | BẮT BUỘC | Alt text cho mọi hình ảnh |
| PPTX-03 | BẮT BUỘC | Cấm SmartArt — thay bằng bullet list hoặc bảng |
| PPTX-04 | BẮT BUỘC | Chữ quan trọng không nằm trong screenshot/ảnh (không có OCR) |
| PPTX-05 | NÊN | Bố trí shape theo trục đọc trên→dưới, trái→phải; tránh chồng lấn (converter sort theo tọa độ) |
| PPTX-06 | NÊN | Chart dùng loại cơ bản (column/bar/line/pie); loại phức tạp thành `[unsupported chart]` |
| PPTX-07 | NÊN | Dùng speaker notes cho nội dung chi tiết/ngữ cảnh (được convert thành `### Notes:`) |
| PPTX-08 | NÊN | Bảng đơn giản, không merge cell |
| PPTX-09 | NÊN | Mỗi slide 1 chủ đề (giúp chunking RAG) |
| PPTX-10 | TRÁNH | WordArt trang trí, icon mang nghĩa mà không có chữ đi kèm |

## 4. Kiểm chứng thực nghiệm (`samples/`)

- `samples/generate_samples.py` — script Python dùng `python-docx`, `openpyxl`, `python-pptx` sinh cặp file `good.*` / `bad.*` cho mỗi định dạng.
- `samples/output/` — kết quả chạy `markitdown` trên từng file mẫu, commit kèm để team xem không cần chạy lại.
- Yêu cầu môi trường: Python 3.10+, `pip install "markitdown[docx,xlsx,pptx]" python-docx openpyxl python-pptx`.

**Phạm vi demo được bằng script** (không phải quy tắc nào cũng sinh được bằng code):

| Demo được | Không demo được (chỉ giải thích trong tài liệu) |
|---|---|
| DOCX-01 (heading style vs bold thủ công), DOCX-02 (merge cell), DOCX-03 (có/không alt text) | DOCX-05 (MathType), DOCX-06 (tracked changes), DOCX-07/08 (text box, SmartArt — python-docx không hỗ trợ tạo) |
| XLSX-01/02 (bảng lệch A1, tiêu đề trang trí), XLSX-03 (merge), XLSX-04 (công thức không cached value — openpyxl ghi công thức thuần), XLSX-08 (ô trống → NaN) | XLSX-05 (chart/pivot) |
| PPTX-01 (có/không title placeholder), PPTX-02 (alt text), PPTX-05 (thứ tự shape), PPTX-07 (notes), PPTX-06 (chart cơ bản) | PPTX-03 (SmartArt — python-pptx không tạo được) |

Tiêu chí đạt: output Markdown của từng cặp good/bad thể hiện đúng sự khác biệt mà quy tắc mô tả. Nếu output thực tế mâu thuẫn với quy tắc nào → sửa lại quy tắc theo thực tế (thực nghiệm thắng tài liệu).

## 5. Ngoài phạm vi

- Hướng dẫn cài đặt/vận hành pipeline MarkItDown cho production.
- Convention cho PDF và các định dạng khác.
- Quy trình review/enforce trong team (chỉ cung cấp checklist, không quy định quy trình).
- Bật LLM caption (`llm_client`) — nếu sau này bật, việc nới lỏng chỉ áp dụng cho
  PPTX-02 (converter pptx của MarkItDown có gọi hook `llm_client`/`llm_caption` khi
  ảnh không có alt text). DOCX-03 vẫn bắt buộc: pipeline docx đi qua `mammoth`, không
  có hook LLM caption nào trong đường xử lý đó, nên ảnh docx không alt text vẫn mất
  hoàn toàn bất kể pipeline có bật `llm_client` hay không. Đã có ghi chú trong README.

## 6. Quyết định đã chốt

| Câu hỏi | Quyết định |
|---|---|
| Mục đích Markdown | LLM/RAG |
| Ngôn ngữ tài liệu | Tiếng Việt (thuật ngữ kỹ thuật giữ tiếng Anh) |
| Cấu trúc | Tách file: README + 3 file theo định dạng |
| Pipeline | MarkItDown mặc định, không LLM caption |
| Kiểu trình bày | Hướng A: quy tắc có ID + mức độ + lý do + ví dụ đúng/sai + checklist |
| Kiểm chứng | Có — thư mục `samples/` với script sinh file và output thực tế |
