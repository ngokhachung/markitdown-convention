# Convention: PowerPoint (.pptx)

Quy tắc tạo file PowerPoint để MarkItDown convert sang Markdown không mất thông tin.

**Cơ chế:** MarkItDown dùng `python-pptx`, duyệt shape theo **tọa độ (trên→dưới,
trái→phải)**. Mỗi slide được ngăn cách trong output bằng comment
`<!-- Slide number: N -->` (đánh số từ 1). Title placeholder → `#`; speaker notes →
heading cố định `### Notes:`; chart cơ bản → heading `### Chart` kèm bảng dữ liệu
Markdown. **Không xử lý SmartArt, không OCR chữ trong ảnh.**

**Mức độ:** `[BẮT BUỘC]` vi phạm là mất/hỏng dữ liệu · `[NÊN]` cải thiện chất lượng
output · `[TRÁNH]` tính năng bị bỏ qua hoàn toàn.

**Bằng chứng:** `samples/pptx-good.pptx` / `samples/pptx-bad.pptx` và output thật tại
`samples/output/pptx-good.md` / `samples/output/pptx-bad.md`.

---

### PPTX-01 — Dùng layout có Title placeholder, mỗi slide đúng 1 tiêu đề `[BẮT BUỘC]`

**Tại sao:** chỉ Title placeholder mới thành heading `#`. Textbox tự vẽ làm "tiêu đề"
chỉ là văn bản thường — mất ranh giới nội dung giữa các slide, RAG không chunk được.

✅ **Đúng:** New Slide → chọn layout "Title and Content" / "Title Only", gõ vào ô title.
❌ **Sai:** slide Blank + textbox chữ to làm tiêu đề.

### PPTX-02 — Alt text cho mọi hình ảnh `[BẮT BUỘC]`

**Tại sao:** ảnh có alt text sẽ ra `![mô tả](tên-file)` — alt text là thứ duy nhất
của ảnh sống sót, pipeline không bật LLM caption. Ảnh **không có** alt text cũng
không biến mất lặng lẽ, nhưng cũng không sinh ra tên file thật: MarkItDown đọc
thuộc tính `descr` của shape; không có descr, alt ra **rỗng** — `![](tên-file)` —
và `tên-file` không phải tên file ảnh gốc mà là tên shape trong PowerPoint (vd
`Picture 4` → `Picture4.jpg`) được PowerPoint tự đặt, không mang ý nghĩa gì. Vì vậy
ảnh không alt text để lại đúng một dấu vết trong output: cú pháp ảnh rỗng, không mô
tả nội dung.

✅ **Đúng:** Right-click ảnh → Edit Alt Text → mô tả đầy đủ nội dung/ý nghĩa.
❌ **Sai:** dán ảnh không alt text — output ra `![](Picture4.jpg)`, alt rỗng không mô
tả gì nội dung ảnh (xem `samples/output/pptx-bad.md`).

### PPTX-03 — Cấm SmartArt `[BẮT BUỘC]`

**Tại sao:** converter không có code xử lý SmartArt — **toàn bộ chữ bên trong biến
mất**. Đây là nguồn mất dữ liệu lớn nhất của PowerPoint.

✅ **Thay bằng:** bullet list (quy trình, chu trình), bảng (so sánh, ma trận).

### PPTX-04 — Chữ quan trọng không nằm trong screenshot/ảnh `[BẮT BUỘC]`

**Tại sao:** không có OCR. Slide chụp màn hình số liệu = slide rỗng sau convert.

✅ **Đúng:** gõ lại số liệu chính thành bullet/bảng, ảnh chỉ để minh họa (kèm alt text).

### PPTX-05 — Bố trí shape theo trục đọc trên→dưới, trái→phải `[NÊN]`

**Tại sao:** shape được sort theo tọa độ (top, left). Đặt lệch thứ tự thị giác làm văn
bản đảo lộn (xem "Bước 2" đứng trước "Bước 1" trong `samples/output/pptx-bad.md`).

✅ **Đúng:** nội dung đọc trước đặt cao hơn/trái hơn; tránh chồng lấn shape.

### PPTX-06 — Chart dùng loại cơ bản `[NÊN]`

**Tại sao:** chart column/bar/line/pie được chuyển thành bảng dữ liệu Markdown (giữ
được số liệu!). Loại phức tạp không hỗ trợ chỉ còn `[unsupported chart]`. Quan sát
thực tế: chart hiện dưới heading `### Chart` kèm bảng Markdown chuẩn, và giá trị số
bị ép về kiểu float — `100`/`200` nhập vào chart ra `100.0`/`200.0` trong bảng, không
giữ định dạng số nguyên. Nếu chart có đặt tiêu đề (chart title), heading sẽ thành
`### Chart: <tiêu đề>` thay vì chỉ `### Chart` trơn — vì vậy đặt tên chart rõ nghĩa
là đáng làm, nó cho RAG một mốc ngữ cảnh ngay trong heading.

✅ **Đúng:** dùng chart cơ bản, đặt tên series/category rõ nghĩa; đặt tiêu đề chart để
heading output có ngữ cảnh (`### Chart: <tiêu đề>`); chấp nhận số liệu ra dạng thập
phân có đuôi `.0`.
❌ **Sai:** chart combo/3D lạ — mất toàn bộ số liệu.

### PPTX-07 — Dùng speaker notes cho nội dung chi tiết `[NÊN]`

**Tại sao:** khi slide có notes, notes được convert thành heading cố định
`### Notes:` ở cuối slide đó — chuỗi này do MarkItDown hardcode, luôn xuất hiện y hệt
bất kể nội dung notes là gì (slide không có notes thì không có heading này). Đây là
nơi tốt nhất để viết đoạn văn đầy đủ ngữ cảnh cho RAG mà không làm rối slide.

✅ **Đúng:** slide để ý chính ngắn gọn; diễn giải đầy đủ viết vào Notes; không cần tự
gõ thêm chữ "Notes" trong nội dung notes vì heading đã tự sinh.

### PPTX-08 — Bảng đơn giản, không merge cell `[NÊN]`

**Tại sao:** như mọi định dạng khác, bảng Markdown không biểu diễn được ô gộp.

### PPTX-09 — Mỗi slide 1 chủ đề `[NÊN]`

**Tại sao:** mỗi slide là một khối nội dung có tiêu đề riêng trong output — slide đơn
chủ đề giúp chunking RAG sạch hơn.

### PPTX-10 — Không WordArt trang trí, không icon mang nghĩa mà thiếu chữ `[TRÁNH]`

**Tại sao:** icon/hình trang trí không có chữ đi kèm thì không để lại dấu vết ngữ
nghĩa nào trong output.

✅ **Thay bằng:** icon + nhãn chữ bên cạnh.

---

## Checklist trước khi nộp file PowerPoint

| ID | Quy tắc | Mức | Sample |
|---|---|---|---|---|
| PPTX-01 | Layout có Title placeholder | BẮT BUỘC | good+bad |
| PPTX-02 | Mọi ảnh có alt text | BẮT BUỘC | good+bad |
| PPTX-03 | Không SmartArt | BẮT BUỘC | good+bad |
| PPTX-04 | Chữ quan trọng không nằm trong ảnh | BẮT BUỘC | good+bad |
| PPTX-05 | Shape xếp theo trục đọc | NÊN | good+bad |
| PPTX-06 | Chart loại cơ bản | NÊN | good+bad |
| PPTX-07 | Nội dung chi tiết vào speaker notes | NÊN | good |
| PPTX-08 | Bảng không merge cell | NÊN | good+bad |
| PPTX-09 | Mỗi slide 1 chủ đề | NÊN | good+bad |
| PPTX-10 | Không WordArt/icon thiếu nhãn chữ | TRÁNH | good+bad |
