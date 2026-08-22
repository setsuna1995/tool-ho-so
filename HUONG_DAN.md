# Hướng dẫn sử dụng công cụ tạo hồ sơ dự án

Công cụ này giúp bạn tự động tạo một bộ hồ sơ đạo đức và khoa học hoàn chỉnh cho dự án nghiên cứu, dựa trên dữ liệu nhập từ một file Excel.

## 1. Yêu cầu máy chạy

Để sử dụng công cụ này, máy tính của bạn cần có:

- **Hệ điều hành:** Windows (Windows 10 trở lên)
- **Python:** Phiên bản 3.9 hoặc cao hơn (cần được cài đặt)
- **Microsoft Word:** Tùy chọn
  - Nếu có Word: Kết quả chính xác nhất, định dạng tuyệt vời
  - Nếu không có Word: Công cụ sẽ tự động chuyển sang chế độ fallback (python-docx). Lúc này, một vài dòng trống trong file "00. QĐ Giao đề tài" có thể không được xóa hoàn toàn — script sẽ in cảnh báo cụ thể nếu gặp tình huống này.

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
- **Thư ký hội đồng đạo đức** (C09, C10) — **bắt buộc** phải có ít nhất 1-2 người (tên, học vị, đơn vị)
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

## 6. Về các script chạy một lần duy nhất

Trong thư mục công cụ, có 2 file script sau:
- `migrate_add_partner_org.py`
- `convert_doc_templates.py`

**Những script này chỉ cần chạy MỘT LẦN duy nhất** khi công cụ được thiết lập lần đầu. Bạn **KHÔNG** cần chạy lại chúng cho mỗi dự án mới.

Bạn chỉ cần chạy lại chúng nếu:
- Bạn thêm các file mẫu `.doc` mới vào thư mục công cụ
- Quản trị viên yêu cầu cập nhật cấu hình

Nếu không chắc, hãy liên hệ với người tạo công cụ.

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
