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

### DOCX-02 — Bảng không merge cell, đánh dấu dòng đầu là header `[BẮT BUỘC]`

**Tại sao:** bảng Markdown không biểu diễn được ô gộp (colspan/rowspan). Ô merge làm
các dòng lệch cột — dữ liệu bị gán nhầm cột mà không có cảnh báo. Quan sát thực tế về
header: mammoth chỉ map dòng ra `<th>` (→ dòng header Markdown thật) nếu dòng đó được
đánh dấu **"Repeat as header row at the top of each page"** (Word: bôi đen dòng 1 →
Table Properties → Row → tick ô đó). **Không** đánh dấu, dòng đầu bị đẩy xuống thân
bảng như mọi dòng khác và header Markdown ra **rỗng** (`|  |  |`), dù về mặt hình thức
dòng đó vẫn trông giống header trong Word. Có đánh dấu, output có header Markdown thật
(`| Tháng | Doanh thu |`).

✅ **Đúng:** mỗi ô một giá trị; dòng đầu là tên cột; lặp lại giá trị thay vì merge dọc;
tick "Repeat as header row at the top of each page" cho dòng header (Table Properties
→ Row).
❌ **Sai:** merge ô "Q1" ngang 3 cột; merge dọc ô "Miền Bắc" cho 4 dòng; hoặc bỏ qua
việc đánh dấu header row (dòng đầu vẫn ra được nhưng header Markdown thành rỗng). Hậu
quả merge quan sát được: ô merge ngang chỉ còn text ở cột đầu tiên, cột bị merge biến
thành ô rỗng (Markdown không có colspan) — dòng bị lệch so với các dòng dữ liệu bên
dưới, dễ khiến người đọc hiểu sai nội dung thuộc cột nào.

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

### DOCX-07 — Không đặt nội dung trong text box `[NÊN]`

**Tại sao:** mammoth KHÔNG bỏ chữ trong text box, nhưng cũng không giữ đúng chỗ. Text
box Word thật được lưu dạng `mc:AlternateContent` (nhánh DrawingML + nhánh VML dự phòng);
mammoth đọc nhánh VML nên chữ **sống sót**, nhưng bị **dồn ra cuối đoạn văn chứa nó** —
thứ tự đọc bị xáo trộn âm thầm. (Text box DrawingML thuần, không có nhánh VML, thì mất
hẳn chữ.) Dù trường hợp nào, output cũng không đáng tin.

✅ **Thay bằng:** đoạn văn thường; nếu cần đóng khung nhấn mạnh, dùng bảng 1 ô.
❌ **Sai:** đặt số liệu vào text box — "ĐẠT 600 TRIỆU" bị dồn ra sau "đã kiểm toán"
trong `samples/output/docx-bad.md`.

### DOCX-08 — Không đặt chữ mang nghĩa trong SmartArt/đồ hoạ `[TRÁNH]`

**Tại sao:** chữ trong SmartArt thật nằm ở một part riêng (`diagrams/data*.xml`) mà
mammoth không bao giờ mở — **mất toàn bộ, không dấu vết**. Đồ hoạ DrawingML không có bản
VML dự phòng cũng mất chữ y hệt. (Shape/WordArt có nhánh VML thì chữ sống sót nhưng bị
dồn vị trí như text box — xem DOCX-07.)

✅ **Thay bằng:** bullet list (quy trình), bảng (so sánh), heading (phân cấp).
❌ **Sai:** để nội dung chỉ trong SmartArt/đồ hoạ — chữ "KHẢO SÁT → PHÂN TÍCH → CHỐT"
biến mất trong `samples/output/docx-bad.md`. SmartArt thật khó sinh bằng code nên sample
dùng đồ hoạ DrawingML tương đương về hành vi mất chữ.

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

| ID | Quy tắc | Mức | Sample |
|---|---|---|---|---|
| DOCX-01 | Tiêu đề dùng Heading Style | BẮT BUỘC | good+bad |
| DOCX-02 | Bảng không merge cell, đánh dấu "Repeat as header row" | BẮT BUỘC | good+bad |
| DOCX-03 | Mọi ảnh có alt text mô tả đủ | BẮT BUỘC | good+bad |
| DOCX-04 | Thông tin quan trọng có bản chữ | BẮT BUỘC | good+bad |
| DOCX-05 | Công thức dùng Insert > Equation | BẮT BUỘC | good+bad |
| DOCX-06 | Đã resolve hết tracked changes | BẮT BUỘC | bad |
| DOCX-07 | Không nội dung trong text box (sống sót nhưng sai thứ tự) | NÊN | good+bad |
| DOCX-08 | Không chữ mang nghĩa trong SmartArt/đồ hoạ | TRÁNH | good+bad |
| DOCX-09 | Không nội dung trong header/footer | TRÁNH | good+bad |
| DOCX-10 | Không nội dung trong comment | TRÁNH | good |
| DOCX-11 | List dùng bullet/numbering chuẩn | NÊN | good+bad |
| DOCX-12 | Link qua Insert > Link | NÊN | good |
