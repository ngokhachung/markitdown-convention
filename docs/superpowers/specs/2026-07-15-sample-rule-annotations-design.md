# Thiết kế: Sinh lại sample, minh hoạ + gắn nhãn inline mọi quy tắc

Ngày: 2026-07-15

## Mục tiêu

Tạo lại 6 file sample Office (good/bad × docx/xlsx/pptx) sao cho **mỗi quy tắc**
trong ba convention (docx 12 · xlsx 10 · pptx 10) đều có phần tử minh hoạ cụ thể,
và **mỗi phần tử mang nhãn inline** ghi rõ mã quy tắc nó minh hoạ. Khi mở file Office
hoặc đọc file `.md` sau convert, người xem thấy ngay quy tắc nào gắn ở đâu.

Kết quả cuối: `generate_samples.py` + `convert_and_verify.py` chạy ra **ALL PASS**,
các file nhị phân và output `.md` được sinh lại và commit; checklist trong mỗi
convention thêm cột `Sample`.

## Quy ước gắn nhãn

- **File good** — phần tử minh hoạ mang mã quy tắc ngay trong nội dung chữ của nó,
  dạng `[DOCX-01]`, đặt tự nhiên (đuôi heading, đầu đoạn, một cột "Quy tắc"…). Nhãn
  là chữ nên đi kèm ra output `.md`, khiến cả file Office lẫn `.md` tự chú giải. Với
  quy tắc `[TRÁNH]`, file good minh hoạ **cách thay thế** ("Thay bằng").
- **File bad** — phần tử vi phạm mang nhãn `[vi phạm DOCX-07]`. Với vi phạm gây *mất
  nội dung* (text box, SmartArt, comment, header/footer…), đặt kèm một đoạn chữ-thường
  **sống sót** mô tả: "Đoạn dưới vi phạm DOCX-07, nội dung sẽ biến mất". Nhờ vậy output
  chứng minh: nhãn/mô tả còn, nội dung vi phạm biến mất.
- **Nhãn xấp xỉ** — SmartArt/WordArt không sinh trung thực bằng code: dùng shape chứa
  chữ (qua XML) có hành vi bị-bỏ-qua giống hệt, nhãn ghi rõ `≈ xấp xỉ SmartArt, xem
  DOCX-08`. Không giả lập sai hành vi converter.

## Phủ quy tắc — Word (docx)

| Quy tắc | good minh hoạ | bad minh hoạ | Cách sinh |
|---|---|---|---|
| DOCX-01 Heading Style | H1/H2/H3 thật, chữ mang `[DOCX-01]` | tiêu đề giả bold+16pt | python-docx |
| DOCX-02 Bảng no-merge + header row | bảng sạch, row0 có `w:tblHeader` | ô header merge ngang | python-docx + XML |
| DOCX-03 Alt text | ảnh có `descr` chứa `[DOCX-03]` | ảnh không `descr` | python-docx |
| DOCX-04 Bản chữ, không chỉ trong ảnh | số liệu gõ thành đoạn `[DOCX-04]` + ảnh minh hoạ | ảnh "chụp bảng số liệu" không bản chữ + đoạn note sống sót | python-docx |
| DOCX-05 Equation OMML | equation OMML → LaTeX `$…$`, đoạn dẫn `[DOCX-05]` | công thức dạng ảnh | chèn XML OMML |
| DOCX-06 Resolve tracked change | không còn tracked change, note `[DOCX-06]` | chèn `w:ins`/`w:del` treo | chèn XML |
| DOCX-07 Không text box | bảng 1 ô thay thế, `[DOCX-07]` | text box (chữ biến mất) + note sống sót | chèn XML txbx |
| DOCX-08 Không SmartArt/WordArt/shape-chữ | bullet/bảng thay thế `[DOCX-08]` | shape chứa chữ (≈ SmartArt) + note | chèn XML shape |
| DOCX-09 Không content ở header/footer | định danh tài liệu ở đầu thân `[DOCX-09]` | đặt chữ vào header/footer | python-docx section |
| DOCX-10 Không content trong comment | đoạn "Lưu ý:" trong thân `[DOCX-10]` | comment (biến mất) + note | python-docx/XML comment |
| DOCX-11 Bullet/numbering chuẩn | style "List Bullet"/"List Number" `[DOCX-11]` | gõ tay `- ` | python-docx style |
| DOCX-12 Hyperlink | hyperlink thật `[text](url)` `[DOCX-12]` | URL gõ trơn (không link) | chèn XML hyperlink |

## Phủ quy tắc — Excel (xlsx)

Ràng buộc "1 sheet = 1 bảng từ A1": file **good** dùng nhiều sheet sạch, mỗi sheet
một bảng; nhãn quy tắc đặt trong **một cột "Quy tắc"** cuối bảng và/hoặc trong tên sheet.

| Quy tắc | good minh hoạ | bad minh hoạ | Cách sinh |
|---|---|---|---|
| XLSX-01 1 bảng từ A1 | bảng bắt đầu A1 | bảng lệch xuống B3/A3 | openpyxl |
| XLSX-02 Dòng 1 là header | dòng 1 = tên cột | dòng 1 = tiêu đề trang trí | openpyxl |
| XLSX-03 Không merge | không merge; cột "Quy tắc" ghi `XLSX-03` | merge dọc/ngang | openpyxl |
| XLSX-04 Công thức có cached value | ghi sẵn giá trị (600) | `=B4*3` không cached | openpyxl |
| XLSX-05 Không giấu info trong chart/màu/comment | cột "Trạng thái" bằng chữ `XLSX-05`; chart lấy nguồn từ bảng hiện diện | info chỉ ở chart + comment + fill màu | openpyxl chart/comment/fill |
| XLSX-06 Xoá sheet nháp/ẩn | không có sheet ẩn; note trong ô | sheet `NhapLieuTam` ẩn | openpyxl hidden |
| XLSX-07 Tên sheet có nghĩa | `DoanhThu2026`, `NgayVaTyLe` | `Sheet1`, `Copy of Final (2)` | openpyxl title |
| XLSX-08 Không ô trống / không N/A giả | điền "Không áp dụng", `0` | ô trống + literal `N/A` | openpyxl |
| XLSX-09 Ý thức giá trị thô | sheet có cột ngày + % (ra `2026-…`, `0.15`) `XLSX-09` | (không cần bad riêng) | openpyxl number_format |
| XLSX-10 Không nhiều bảng con/1 sheet | mỗi bảng một sheet | 2 bảng con + dòng ghi chú xen kẽ trong 1 sheet | openpyxl |

Nhãn vi phạm trong file bad đặt ở cột "Ghi chú"/dòng note khi cấu trúc còn cho phép;
chỗ merge/lệch làm nhãn xô lệch — chính sự xô lệch đó là bằng chứng.

## Phủ quy tắc — PowerPoint (pptx)

| Quy tắc | good minh hoạ | bad minh hoạ | Cách sinh |
|---|---|---|---|
| PPTX-01 Title placeholder | layout Title, `[PPTX-01]` trong body | textbox chữ to làm tiêu đề | python-pptx |
| PPTX-02 Alt text | ảnh có `descr` `[PPTX-02]` | ảnh `descr` rỗng | python-pptx |
| PPTX-03 Không SmartArt | bullet list thay thế `[PPTX-03]` | shape-nhóm chứa chữ (≈ SmartArt) + note | python-pptx/XML |
| PPTX-04 Chữ không nằm trong ảnh | số liệu gõ thành bullet + ảnh minh hoạ | slide "ảnh chụp số liệu" không bản chữ + note | python-pptx |
| PPTX-05 Bố trí theo trục đọc | shape xếp trên→dưới đúng | Bước 2 trên, Bước 1 dưới | python-pptx |
| PPTX-06 Chart cơ bản | column clustered → bảng `[PPTX-06]` | chart 3D/combo → `[unsupported chart]` | python-pptx chart |
| PPTX-07 Speaker notes | notes có nội dung → `### Notes:` | (không cần bad riêng) | python-pptx notes |
| PPTX-08 Bảng không merge | bảng sạch `[PPTX-08]` | bảng có ô merge | python-pptx table |
| PPTX-09 Mỗi slide 1 chủ đề | 2 slide, mỗi slide 1 chủ đề, note `[PPTX-09]` | 1 slide nhồi 2 chủ đề + note | python-pptx |
| PPTX-10 Không WordArt/icon thiếu chữ | icon + nhãn chữ `[PPTX-10]` | shape trang trí không chữ (≈ WordArt) + note | python-pptx/XML |

## Script

- `generate_samples.py`: mở rộng 6 hàm sinh; thêm helper chèn XML (OMML, txbx, w:ins,
  hyperlink, comment, shape-chữ). Mỗi phần tử gắn nhãn theo quy ước trên.
- `convert_and_verify.py`: viết lại assertions để **mỗi quy tắc có ≥1 assert** (good
  hoặc bad, hoặc cả hai), bám output mới; mục tiêu **ALL PASS**. Với quy tắc xấp xỉ,
  assert điểm quan sát được (chữ trong shape biến mất) chứ không assert "là SmartArt".
- Convention `.md` × 3: thêm cột `Sample` vào bảng checklist cuối file (giá trị `good`,
  `bad`, `good+bad`, `good+bad (≈)`). Cập nhật câu "Bằng chứng" nếu ánh xạ thay đổi.
- `README.md`: cập nhật nếu số quy tắc có sample thay đổi cách mô tả (mục Kiểm chứng).

## Sinh lại & kiểm chứng

1. Tạo `.venv` bằng Python hệ thống; `pip install "markitdown[docx,xlsx,pptx]"
   python-docx openpyxl python-pptx`.
2. Chạy `generate_samples.py` rồi `convert_and_verify.py` tới khi **ALL PASS**.
3. Commit: 6 file nhị phân + 6 output `.md` + 2 script + 3 convention + README.

**Rủi ro:** Python hệ thống là 3.14.4 (rất mới) và chưa có venv. Nếu cài đặt vướng
tương thích, dừng và báo người dùng thay vì loay hoay — không tự hạ cấp/bỏ qua verify.

## Đính chính hành vi phát hiện khi làm sample (2026-07-15)

Khi dựng sample DOCX-07/08, thực nghiệm (`probe_docx.py`) cho thấy convention mô tả
SAI hành vi MarkItDown — người dùng quyết định **sửa convention theo thực nghiệm**:

- **DOCX-07 (text box):** mammoth đọc VML text box (và nhánh VML Fallback trong
  `mc:AlternateContent` — cách Word thật lưu text box) → chữ **KHÔNG mất** mà bị **dồn
  ra cuối đoạn chứa nó** (rối thứ tự đọc). Chỉ text box DrawingML thuần (không có VML
  fallback) mới mất hẳn. → sửa DOCX-07: "chữ sống sót nhưng sai thứ tự", đề xuất đổi
  mức `[TRÁNH]`→`[NÊN]`.
- **DOCX-08 (SmartArt/WordArt/shape):** shape/WordArt (có VML fallback) → chữ sống sót
  như text box. Chỉ **SmartArt thật** (`dgm:`, chữ nằm trong part `diagrams/data*.xml`
  mammoth không mở) mới mất toàn bộ chữ → giữ `[TRÁNH]`, thu hẹp về đúng SmartArt thật.
  Demo mất chữ bằng đồ hoạ DrawingML không-fallback; SmartArt thật khó sinh bằng code
  nên ghi chú trung thực.

## Không làm (non-goals)

- Không sửa ngữ nghĩa quy tắc **ngoài** DOCX-07/08 đã đính chính ở trên (các quy tắc
  khác chỉ thêm cột Sample + minh hoạ).
- Không bật LLM caption trong pipeline.
- Không giả lập Excel tính lại công thức (openpyxl không cache được — đó chính là
  điều XLSX-04 dạy); file good ghi sẵn giá trị đã tính.
- Không sinh SmartArt/WordArt "thật" 100% — dùng xấp xỉ có nhãn.
