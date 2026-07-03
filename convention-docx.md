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
các dòng lệch cột — dữ liệu bị gán nhầm cột mà không có cảnh báo. Quan sát thực tế:
bảng docx luôn render với một dòng header Markdown **rỗng** (mọi dòng của bảng docx,
kể cả dòng định làm header, đều bị đẩy xuống phần thân bảng) — nhưng dòng đầu tiên
vẫn nên là header ngữ nghĩa vì người đọc và LLM đều dựa vào vị trí đó để hiểu ý nghĩa
các cột.

✅ **Đúng:** mỗi ô một giá trị; dòng đầu là tên cột; lặp lại giá trị thay vì merge dọc.
❌ **Sai:** merge ô "Q1" ngang 3 cột; merge dọc ô "Miền Bắc" cho 4 dòng. Hậu quả quan
sát được: ô merge ngang chỉ còn text ở cột đầu tiên, cột bị merge biến thành ô rỗng
(Markdown không có colspan) — dòng bị lệch so với các dòng dữ liệu bên dưới, dễ khiến
người đọc hiểu sai nội dung thuộc cột nào.

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
