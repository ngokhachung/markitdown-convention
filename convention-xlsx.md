# Convention: Excel (.xlsx)

Quy tắc tạo file Excel để MarkItDown convert sang Markdown không mất thông tin.

**Cơ chế:** MarkItDown đọc bằng `pandas`: **mọi sheet (kể cả sheet ẩn)** trở thành
`## <tên sheet>` + một bảng Markdown. Nó giả định mỗi sheet là **một bảng duy nhất bắt
đầu tại A1, dòng 1 là header**. Đây là converter đơn giản nhất trong ba định dạng —
mọi thứ nằm ngoài giả định đó đều vỡ.

**Mức độ:** `[BẮT BUỘC]` vi phạm là mất/hỏng dữ liệu · `[NÊN]` cải thiện chất lượng
output · `[TRÁNH]` cấu trúc mà converter không xử lý được.

**Bằng chứng:** `samples/xlsx-good.xlsx` / `samples/xlsx-bad.xlsx` và output thật tại
`samples/output/xlsx-good.md` / `samples/output/xlsx-bad.md`.

---

### XLSX-01 — 1 sheet = 1 bảng duy nhất, bắt đầu tại A1 `[BẮT BUỘC]`

**Tại sao:** pandas đọc cả sheet thành đúng một bảng. Bảng đặt lệch vị trí sinh cột
`Unnamed` và dòng `NaN` rác.

✅ **Đúng:** ô A1 là ô đầu tiên của dòng header.
❌ **Sai:** chừa vài dòng trống trang trí, bảng bắt đầu tại B3.

### XLSX-02 — Dòng 1 là header thật `[BẮT BUỘC]`

**Tại sao:** dòng đầu tiên có dữ liệu trở thành tên cột cho cả bảng. Dòng tiêu đề trang
trí ("BÁO CÁO QUÝ 1") sẽ chiếm chỗ header — toàn bộ cột thành `Unnamed: 1`, `Unnamed: 2`…

✅ **Đúng:** dòng 1 = tên cột; tên báo cáo đặt vào **tên sheet** hoặc tên file.
❌ **Sai:** dòng 1 là tên báo cáo, dòng 3 mới là header (xem `samples/output/xlsx-bad.md`).

### XLSX-03 — Cấm merge cell `[BẮT BUỘC]`

**Tại sao:** giá trị chỉ nằm ở ô trên-trái của vùng merge; các ô còn lại rỗng và render
thành chữ `NaN` trong output — dòng dữ liệu lệch nghĩa.

✅ **Đúng:** lặp lại giá trị ở từng dòng (vd cột "Miền" ghi "Bắc" cho cả 4 dòng).
❌ **Sai:** merge dọc ô "Bắc" cho 4 dòng.

### XLSX-04 — Công thức phải có cached value `[BẮT BUỘC]`

**Tại sao:** MarkItDown chỉ đọc **giá trị đã tính** mà Excel lưu trong file ở lần save
cuối, không đọc công thức. File sinh bằng code (openpyxl, script xuất báo cáo…) chưa
từng được Excel mở/tính → ô công thức **rỗng hoàn toàn** trong output.

✅ **Đúng:** trước khi nộp, mở file bằng Excel và Ctrl+S. File sinh bằng code: ghi
giá trị đã tính sẵn, không ghi chuỗi công thức.
❌ **Sai:** script xuất file có `=SUM(...)` rồi nộp thẳng (xem ô "Gấp ba" biến mất
trong `samples/output/xlsx-bad.md`).

### XLSX-05 — Chart/pivot/comment/màu sắc không được là nơi duy nhất chứa thông tin `[BẮT BUỘC]`

**Tại sao:** chart, pivot table, comment, màu nền, conditional formatting bị bỏ qua
hoàn toàn. Quy ước "tô đỏ = quá hạn" mất sạch sau convert.

✅ **Đúng:** thêm cột "Trạng thái" bằng chữ; bảng dữ liệu nguồn của chart phải hiện
diện trong sheet.
❌ **Sai:** thông tin chỉ tồn tại trong biểu đồ hoặc màu ô.

### XLSX-06 — Xóa sheet nháp và sheet ẩn trước khi nộp `[BẮT BUỘC]`

**Tại sao:** mọi sheet đều được convert, **kể cả sheet ẩn** — dữ liệu nháp/nội bộ sẽ
lọt vào nguồn tri thức của AI (xem sheet `NhapLieuTam` lộ ra trong
`samples/output/xlsx-bad.md`).

✅ **Đúng:** xóa hẳn sheet nháp; đừng chỉ Hide.

### XLSX-07 — Tên sheet ngắn, có nghĩa `[NÊN]`

**Tại sao:** tên sheet trở thành heading `##` — là tiêu đề section trong nguồn tri thức.

✅ **Đúng:** `DoanhThu2026`, `NhanSu-Q3`.
❌ **Sai:** `Sheet1`, `Copy of Final (2)`.

### XLSX-08 — Không để ô trống, và không điền bằng chuỗi giả-NaN (`N/A`, `NULL`…) `[NÊN]`

**Tại sao:** ô trống render thành chữ `NaN` — gây nhiễu và mơ hồ cho LLM. Nhưng điền
bằng đúng những chuỗi tưởng chừng an toàn như `N/A`, `NA`, `NULL`, `None`, `nan` **cũng
không giải quyết được gì**: MarkItDown đọc xlsx qua `pandas`, và pandas mặc định coi các
chuỗi literal này (cùng nhóm với `n/a`, `NaN`, `#N/A`, `null`...) là giá trị khuyết —
tự động biến ô thành `NaN` y hệt một ô trống, **dù ô đó thực sự có nội dung**. Phát hiện
thực nghiệm khi làm sample cho convention này: cột "Trạng thái" ghi literal `"N/A"` ở
dòng "Tổng" vẫn ra `NaN` trong output dù ô không hề rỗng.

✅ **Đúng:** điền `0`, "Chưa có", hoặc "Không áp dụng" tùy ngữ nghĩa — xem
`samples/output/xlsx-good.md`, dòng "Tổng" dùng "Không áp dụng" và bảng ra 0 `NaN`.
❌ **Sai:** điền ô bằng `N/A`, `NA`, `NULL`, `None`, hoặc `nan` để "cho khỏi trống" —
pandas vẫn coi đây là giá trị khuyết và xuất ra `NaN` trong Markdown, y như để trống.

### XLSX-09 — Hiểu rằng output là giá trị thô, không phải giá trị hiển thị `[NÊN]`

**Tại sao:** converter lấy giá trị gốc, bỏ định dạng hiển thị: ngày ra
`2026-07-03 00:00:00`, ô 15% ra `0.15`.

✅ **Đúng:** chấp nhận dạng thô; nếu bắt buộc cần dạng hiển thị đẹp, nhập dạng text.

### XLSX-10 — Không nhiều bảng con hoặc ghi chú xen kẽ trong 1 sheet `[TRÁNH]`

**Tại sao:** pandas không nhận biết ranh giới bảng — các bảng con và dòng ghi chú bị
trộn thành một bảng sai cấu trúc.

✅ **Thay bằng:** mỗi bảng một sheet riêng; ghi chú đặt thành cột "Ghi chú".

---

## Checklist trước khi nộp file Excel

| ID | Quy tắc | Mức |
|---|---|---|
| XLSX-01 | 1 sheet = 1 bảng từ A1 | BẮT BUỘC |
| XLSX-02 | Dòng 1 là header thật | BẮT BUỘC |
| XLSX-03 | Không merge cell | BẮT BUỘC |
| XLSX-04 | Công thức có cached value (save bằng Excel) | BẮT BUỘC |
| XLSX-05 | Không giấu thông tin trong chart/màu/comment | BẮT BUỘC |
| XLSX-06 | Đã xóa sheet nháp/ẩn | BẮT BUỘC |
| XLSX-07 | Tên sheet có nghĩa | NÊN |
| XLSX-08 | Không ô trống; không dùng N/A/NULL/None làm giá trị | NÊN |
| XLSX-09 | Ý thức về giá trị thô (ngày, %) | NÊN |
| XLSX-10 | Không nhiều bảng con trong 1 sheet | TRÁNH |
