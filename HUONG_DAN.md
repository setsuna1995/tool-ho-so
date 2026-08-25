# Hướng dẫn sử dụng công cụ tạo hồ sơ dự án

Công cụ này giúp bạn tự động tạo một bộ hồ sơ đạo đức và khoa học hoàn chỉnh cho dự án nghiên cứu, dựa trên dữ liệu nhập từ một file Excel.

## 1. Yêu cầu máy chạy

Để sử dụng công cụ này, máy tính của bạn cần có:

- **Hệ điều hành:** Windows (Windows 10 trở lên)
- **Python:** Phiên bản 3.9 hoặc cao hơn (cần được cài đặt)
- **Microsoft Word:** Tùy chọn
  - Nếu có Word: Kết quả chính xác nhất, định dạng tuyệt vời
  - Nếu không có Word: Công cụ sẽ tự động chuyển sang chế độ fallback (python-docx). Chế độ này có một số giới hạn đã biết:
    - Một vài dòng trống trong file "00. QĐ Giao đề tài" có thể không được xóa hoàn toàn — script sẽ in cảnh báo cụ thể nếu gặp tình huống này.
    - Với những câu chữ vốn bị chia thành nhiều đoạn định dạng nhỏ (runs) trong file Word gốc, sau khi thay thế nội dung có thể mất định dạng chi tiết (ví dụ: in đậm một phần trong câu) — chữ vẫn đúng, chỉ định dạng ở cấp ký tự trong câu đã thay có thể bị đơn giản hóa. Tình trạng này không xảy ra nếu máy có cài Word.

## 2. Cài đặt lần đầu

Cài đặt chỉ cần làm **một lần duy nhất**. Các bước sau:

1. Mở thư mục chứa công cụ trên máy tính
2. Tìm file `setup.bat` (là file batch)
3. **Double-click** vào `setup.bat`
   - Một cửa sổ Terminal sẽ mở ra
   - Script sẽ kiểm tra xem máy bạn có Python không
   - Nếu không có, script sẽ yêu cầu bạn tải Python từ https://www.python.org/downloads/ (nhớ ticked vào checkbox "Add python.exe to PATH" khi cài)
   - Sau khi có Python, script sẽ cài đặt các thư viện cần thiết
4. Đợi đến khi thấy dòng chữ **"Cai dat xong"** (setup complete) xuất hiện
5. Đóng cửa sổ Terminal

Từ bây giờ, công cụ đã sẵn sàng sử dụng.

## 3. Trước khi tạo hồ sơ cho dự án mới

Mỗi lần bạn muốn tạo hồ sơ cho một dự án mới, bạn cần chuẩn bị dữ liệu trong file Excel. Các bước:

### 3.1 Mở file Excel
- Tìm file tên `Form checklist hồ sơ dự án.xlsx` trong thư mục công cụ
- Mở file bằng Excel (double-click nó)

### 3.2 Tạo một sheet mới
Bạn sẽ thấy file Excel có 2 sheet hiện tại:
- `Đề tài - Bánh ăn dặm VIAM 2027` (đây là ví dụ từ dự án mẫu)
- `Đề tài - Mẫu trắng dự án mới` (đây là mẫu trắng)

Để tạo hồ sơ cho dự án mới:
1. **Click chuột phải** vào sheet `Đề tài - Mẫu trắng dự án mới`
2. Chọn "Copy" (hoặc "Move or Copy" → chọn "Create a copy")
3. Đặt tên mới cho sheet theo **tên dự án của bạn** (ví dụ: "Đề tài - Nghiên cứu về dinh dưỡng 2026")

### 3.3 Điền dữ liệu bắt buộc
Sheet mới sẽ có các ô được tô nền **vàng** — đây là những ô **BẮT BUỘC** phải điền. Bạn cần nhập:

- **Tên đề tài** (A01)
- **Kiểu nghiên cứu** (A02) — ví dụ: "Nghiên cứu can thiệp"
- **Năm thực hiện** (A03) — ví dụ: 2026
- **Đơn vị chủ trì** (A04)
- **Năm thực hiện, dự kiến hoàn thành** (A05) — ví dụ: "2026–2027"
- **Chủ nhiệm đề tài** (B01) — tên, học vị (ví dụ: PGS.TS.), và đơn vị công tác
- **Thư ký hội đồng đạo đức** (C09, C10) — **bắt buộc** phải điền đủ cả 2 người (tên, học vị, đơn vị)
- **Thư ký hội đồng khoa học** (D09, D10) — **bắt buộc**
- **Thư ký hội đồng nghiệm thu** (E09, E10) — **bắt buộc**
- **Các thành viên hội đồng** (Chủ tịch, phản biện, ủy viên của 3 hội đồng)

**Lưu ý:** Bạn có thể điền thêm các ô không tô vàng (như đơn vị đối tác, cộng tác viên), nhưng nó không bắt buộc.

### 3.4 Lưu file Excel
- Nhấn **Ctrl + S** để lưu

## 4. Chạy tạo hồ sơ

Sau khi đã chuẩn bị đủ dữ liệu trong Excel, các bước tiếp theo:

### 4.1 Mở file Python
1. Tìm file `tao_ho_so_moi.py` trong thư mục công cụ
2. Mở file bằng một text editor (ví dụ: Notepad)

### 4.2 Sửa tên sheet
Trong file, bạn sẽ thấy dòng:
```
SHEET_NAME = "Đề tài - Bánh ăn dặm VIAM 2027"
```

Đổi tên này thành **tên sheet bạn vừa tạo** (cũng là tên dự án). Ví dụ:
```
SHEET_NAME = "Đề tài - Nghiên cứu về dinh dưỡng 2026"
```

**Lưu ý:** Tên phải khớp **chính xác** với tên sheet trong Excel (có khoảng trắng, dấu, chữ hoa-thường đều phải đúng)

### 4.3 Lưu file Python
- Nhấn **Ctrl + S** để lưu

### 4.4 Chạy script
1. Mở **Command Prompt** hoặc **PowerShell** trong thư mục công cụ
   - Cách dễ nhất: Trong File Explorer, click phải trong thư mục công cụ → chọn "Open in Terminal"
2. Gõ lệnh:
   ```
   python tao_ho_so_moi.py
   ```
3. Nhấn **Enter** và đợi script chạy xong

Script sẽ in ra các bước đang xử lý:
- "Dang doc du lieu tu Excel checklist..."
- "Dang sao chep file mau..."
- "Dang sinh ho so dao duc..."
- "Dang sinh ho so khoa hoc de cuong..."
- "Dang sinh cong van moi chuyen gia..."
- "Dang sinh ho so nghiem thu..."
- "XONG. Bo ho so da tao tai: ..."

Khi thấy **"XONG"**, hồ sơ của bạn đã được tạo thành công!

## 5. Kiểm tra sau khi chạy

Sau khi script chạy xong, một thư mục mới sẽ được tạo. Tên của nó là:
```
Hồ sơ - <tên đề tài> (<năm>)
```

Ví dụ: `Hồ sơ - Nghiên cứu về dinh dưỡng 2026 (2026)`

Bạn nên kiểm tra lại một số file quan trọng:

1. Mở thư mục vừa tạo
2. Vào thư mục con `01. Hồ sơ đạo đức đề cương`
3. Mở file `00. QĐ Giao đề tài.docx`
   - Kiểm tra **tên đề tài** có chính xác không
   - Kiểm tra **năm** có đúng không
4. Vào thư mục con `01. Hồ sơ đạo đức đề cương` → mở file `01. QĐTLHĐ đạo đức đề cương.docx`
   - Kiểm tra **tên chủ tịch hội đồng** có chính xác không
   - Kiểm tra **đơn vị** của các thành viên có đúng không

Nếu tất cả đều chính xác, hồ sơ của bạn đã sẵn sàng gửi đi!

## Những điều cần tự kiểm tra/sửa tay cho dự án mới

Công cụ tự động hóa được phần lớn công việc, nhưng có một số chỗ **chưa** tự động — bạn cần tự kiểm tra và sửa tay cho từng dự án mới:

- **Danh sách thành viên nhóm nghiên cứu trong file "00. QĐ Giao đề tài.docx":** File này vẫn giữ nguyên danh sách tên thành viên nhóm nghiên cứu từ mẫu cũ — công cụ **chưa** tự động cập nhật danh sách này theo các trường thành viên nhóm nghiên cứu bạn điền trong checklist Excel. Hãy mở file này và đối chiếu kỹ xem danh sách có khớp với nhóm nghiên cứu thực tế của dự án bạn không, trước khi gửi hồ sơ đi.

- **Phiếu chấm điểm nghiệm thu:** Công cụ luôn tạo ra bản "TVCT_ĐGHQ". Nếu kiểu nghiên cứu của dự án bạn (mã mục A02) là **"TNLS"** thay vì "TVCT_ĐGHQ", bạn cần tự thay thủ công đúng mẫu phiếu chấm điểm phù hợp.

- **File CV chuyên gia đính kèm hồ sơ đạo đức:** File CV hiện đang cố định là CV của một người cụ thể (theo mẫu COLOSTRUM cũ). Nếu dự án mới của bạn có chủ nhiệm đề tài khác, bạn cần tự thay file CV này bằng CV đúng người.

- **Mốc thời gian nghiên cứu:** Các mốc thời gian dạng "01/2027 đến 12/2027" xuất hiện trong nhiều tài liệu được công cụ tự suy ra **chỉ từ trường năm thực hiện (A03)**, chứ không dựa vào nội dung mốc thời gian chi tiết hơn mà bạn có thể đã điền ở checklist (A05). Hãy kiểm tra kỹ xem tháng/ngày trong các tài liệu có đúng với thực tế dự án của bạn không.

- **Các dấu "……" hoặc "20xx"/"20XX" còn sót lại:** Một số file có thể vẫn còn các chỗ đánh dấu chưa điền (ví dụ: ngày họp cụ thể, số quyết định...) cần bạn tự điền tay sau khi hồ sơ được tạo ra. Đây là tình trạng đã có từ quy trình cũ (script PowerShell) và chưa thay đổi trong công cụ này.

## 6. Về các script chạy một lần duy nhất

Trong thư mục công cụ, có 2 file script sau:
- `migrate_add_partner_org.py`
- `convert_doc_templates.py`

**Những script này chỉ cần chạy MỘT LẦN duy nhất** khi công cụ được thiết lập lần đầu. Bạn **KHÔNG** cần chạy lại chúng cho mỗi dự án mới.

Bạn chỉ cần chạy lại chúng nếu:
- Bạn thêm các file mẫu `.doc` mới vào thư mục công cụ
- Quản trị viên yêu cầu cập nhật cấu hình

Nếu không chắc, hãy liên hệ với người tạo công cụ.

### 6.1 Khi cần thêm một mẫu tài liệu hoàn toàn mới

**Lưu ý:** Phần này dành cho người quản trị/lập trình viên của công cụ, không dành cho người dùng thông thường. Nếu bạn không rành Python, hãy liên hệ người tạo công cụ thay vì tự làm theo phần này.

Khác với việc chỉ đổi tên sheet Excel (mục 4), **thêm một loại tài liệu mẫu mới** (ví dụ thêm một file quyết định/biên bản mới chưa từng có trong bộ hồ sơ) không thể làm được chỉ bằng cách bỏ file `.doc`/`.docx` vào thư mục — công cụ không tự động dò tìm file mới. Cần sửa code theo đúng thứ tự sau:

1. **Đặt file mẫu mới** vào đúng thư mục mẫu tương ứng (ví dụ mẫu cho hồ sơ nghiệm thu thì đặt vào thư mục `04. Hồ sơ nghiệm thu\04. Hồ sơ nghiệm thu\`).
2. **Nếu file gốc là `.doc`:** mở `convert_doc_templates.py`, thêm đường dẫn file mới vào danh sách `DOC_FILES`, rồi chạy lại script này (máy cần có cài Word) để tự động tạo ra bản `.docx` tương ứng bên cạnh file `.doc` gốc.
3. **Đăng ký file mẫu để được copy vào hồ sơ đầu ra:** mở `tao_ho_so_moi.py`, thêm một dòng vào danh sách `COPIES` gồm cặp (đường dẫn file mẫu `.docx`, đường dẫn file đích trong bộ hồ sơ sẽ tạo ra).
4. **Viết hàm điền dữ liệu cho mẫu mới:** mở file `section_*.py` tương ứng với phần hồ sơ đó (ví dụ `section_nghiem_thu.py` cho hồ sơ nghiệm thu), viết thêm một hàm `_ten_ham(session, dest_dir, info)` mở file `.docx` vừa copy và thay thế các chỗ giữ chỗ (tên đề tài, năm, tên chủ nhiệm, danh sách hội đồng...) bằng dữ liệu lấy từ `info` (đối tượng `ProjectInfo`). Sau đó gọi hàm này trong hàm `generate()` của file đó.
5. **Nếu mẫu cần một trường dữ liệu chưa có trong checklist Excel:** phải thêm cột/ô mới vào file `Form checklist hồ sơ dự án.xlsx` (cả 2 sheet) và cập nhật `excel_reader.py` để đọc trường đó vào `ProjectInfo`.
6. **Viết/cập nhật test:** nên thêm test tương ứng vào file `test_section_*.py` để đảm bảo thay đổi không làm hỏng các mẫu khác đang chạy tốt.
7. **Chạy thử toàn bộ** `python tao_ho_so_moi.py` với một sheet Excel thử nghiệm để kiểm tra mẫu mới được điền đúng, trước khi dùng cho dự án thật.

## 7. Xử lý lỗi thường gặp

Nếu gặp lỗi, bạn có thể tham khảo danh sách các lỗi phổ biến dưới đây:

### 7.1 Lỗi: Python không được tìm thấy
**Dấu hiệu:**
- Khi chạy `setup.bat`, lệnh dừng và in "[LOI] Khong tim thay Python tren may nay"
- Hoặc khi chạy `python tao_ho_so_moi.py` trong Command Prompt, máy báo "python is not recognized"

**Nguyên nhân:** Python chưa được cài đặt hoặc chưa được thêm vào PATH

**Cách khắc phục:**
1. Tải Python từ: https://www.python.org/downloads/
2. Chạy trình cài đặt
3. **Rất quan trọng:** Tick vào checkbox "Add python.exe to PATH" (checkbox này ở dưới cùng của cửa sổ cài đặt)
4. Hoàn tất cài đặt
5. Chạy lại `setup.bat`

### 7.2 Lỗi: Thiếu trường bắt buộc trong Excel
**Dấu hiệu:**
- Script báo lỗi kiểu: "Tên đề tài (A01) đang trống trong checklist" hoặc "Chủ nhiệm đề tài (B01) là bắt buộc nhưng đang trống"
- Script cũng có thể báo: "Thư ký hội đồng (C09) là bắt buộc nhưng đang trống"

**Nguyên nhân:** Bạn chưa điền đầy đủ các ô bắt buộc (tô nền vàng) trong sheet Excel

**Cách khắc phục:**
1. Quay lại file Excel
2. Tìm ô có mã nêu trong lỗi (ví dụ: A01, B01, C09)
3. Điền dữ liệu đầy đủ
4. Lưu Excel
5. Chạy lại script

### 7.3 Lỗi: File đang mở trong Word
**Dấu hiệu:**
- Script in ra lỗi kiểu: "File is locked" hoặc "File is open in another application"
- Script dừng giữa chừng, các file chưa được tạo hoàn toàn

**Nguyên nhân:** Bạn đang mở một trong các file `.docx` trong thư mục `Hồ sơ - ...` bằng Word trong khi script chạy

**Cách khắc phục:**
1. Đóng tất cả các file Word đang mở
2. Chạy lại script

### 7.4 Lỗi: Quá nhiều phản biện hoặc ủy viên
**Dấu hiệu:**
- Script báo lỗi kiểu: "Số thành viên hội đồng (7) không khớp số vai trò truyền vào (5)"
- Hoặc: `ValueError`

**Nguyên nhân:** Bảng trong file Word chỉ có 5 chỗ cố định (1 chủ tịch + 2 phản biện + 2 ủy viên), nhưng bạn đã điền quá nhiều người

**Cách khắc phục:**
1. Quay lại file Excel
2. Kiểm tra số lượng phản biện và ủy viên của mỗi hội đồng
   - Hội đồng đạo đức: tối đa 2 phản biện (C02, C03) + tối đa 2 ủy viên (C04, C05) = tối đa 4 người + 1 chủ tịch = 5 người
   - Hội đồng khoa học: tối đa 2 phản biện + tối đa 2 ủy viên
   - Hội đồng nghiệm thu: tối đa 2 phản biện + tối đa 2 ủy viên
3. Xóa bớt các thành viên vừa đủ 5 người (1 chủ tịch + 4 người khác)
4. Lưu Excel và chạy lại script

**Lưu ý:** Nếu bạn cần thêm nhiều thành viên hơn, cần phải mở rộng bảng trong file Word mẫu — vấn đề này cần liên hệ người tạo công cụ.

---

## Cần trợ giúp?

Nếu gặp vấn đề không được liệt kê ở trên, hãy:
1. Ghi lại **toàn bộ thông báo lỗi** mà script in ra
2. Ghi lại **tên file Excel sheet** bạn sử dụng
3. Ghi lại **hệ điều hành** (Windows 10 hay 11) và **phiên bản Python** (chạy `python --version` để kiểm tra)
4. Liên hệ người quản trị công cụ cùng thông tin trên
