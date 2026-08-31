# Hướng dẫn sử dụng công cụ tạo hồ sơ dự án

Công cụ này giúp bạn tự động tạo một bộ hồ sơ đạo đức, hồ sơ khoa học đề cương, công văn mời chuyên gia và hồ sơ nghiệm thu hoàn chỉnh cho dự án nghiên cứu, dựa trên dữ liệu nhập từ một file Excel.

---

## 1. Yêu cầu máy chạy

Để sử dụng công cụ này, máy tính của bạn cần có:

- **Hệ điều hành:** Windows (Windows 10 trở lên)
- **Python:** Phiên bản 3.9 hoặc cao hơn (cần được cài đặt)
- **Microsoft Word:** Tùy chọn
  - **Nếu có Word (khuyến nghị):** Kết quả chính xác nhất, giữ nguyên toàn bộ định dạng đẹp của file mẫu Word.
  - **Nếu không có Word:** Công cụ tự động chuyển sang chế độ fallback (`python-docx`). Chế độ này vẫn tạo đủ file nhưng có thể đơn giản hóa một số định dạng ký tự phức tạp trong câu văn.

---

## 2. Cài đặt lần đầu

Cài đặt chỉ cần làm **một lần duy nhất**:

1. Mở thư mục chứa công cụ trên máy tính.
2. Tìm file **`setup.bat`**.
3. **Double-click** vào `setup.bat`:
   - Script sẽ kiểm tra xem máy bạn đã cài Python chưa.
   - Nếu chưa có, script sẽ hướng dẫn bạn tải Python từ https://www.python.org/downloads/ (nhớ tích chọn **"Add python.exe to PATH"** khi cài đặt).
   - Tự động cài đặt các thư viện cần thiết (`openpyxl`, `python-docx`, `pywin32`, `pytest`...).
4. Khi thấy thông báo **"Cai dat xong"**, bạn có thể đóng cửa sổ Terminal.

---

## 3. Chuẩn bị dữ liệu cho dự án mới

### 3.1 Mở file Excel
- Mở file **`Form checklist hồ sơ dự án.xlsx`** trong thư mục công cụ.

### 3.2 Tạo sheet cho dự án mới
File Excel có sẵn các sheet:
- `Đề tài - Bánh ăn dặm VIAM 2027`: Ví dụ dự án mẫu đã điền đầy đủ.
- `Đề tài - Mẫu trắng dự án mới`: Mẫu trắng chuẩn.

**Cách tạo sheet mới:**
1. Click chuột phải vào tab sheet **`Đề tài - Mẫu trắng dự án mới`**.
2. Chọn **Move or Copy...** → tích vào ô **Create a copy** → bấm **OK**.
3. Đổi tên sheet mới theo cú pháp: `Đề tài - <Tên dự án của bạn>` (ví dụ: `Đề tài - Nghiên cứu sữa non 2027`).

### 3.3 Điền dữ liệu vào sheet dự án
Các ô có nền **màu vàng** là các ô **bắt buộc**:

- **Tên đề tài** (`A01`): Tên đầy đủ của đề tài nghiên cứu.
- **Kiểu nghiên cứu** (`A02`): Chỉ chấp nhận 1 trong 2 giá trị chính xác:
  - `TVCT_ĐGHQ` (Tư vấn công thức / Đánh giá hiệu quả)
  - `TNLS` (Thử nghiệm lâm sàng)
  *(Công cụ dựa vào ô này để tự động chọn đúng mẫu Phiếu chấm điểm nghiệm thu).*
- **Năm thực hiện** (`A03`): Ví dụ `2027`.
- **Đơn vị chủ trì** (`A04`): Ví dụ `Viện Y học ứng dụng Việt Nam`.
- **Thời gian triển khai NC** (`A05`): Bắt buộc chứa 2 mốc `MM/YYYY`, ví dụ `Tháng 01/2027 đến tháng 12/2027` hoặc `Từ 03/2027 đến 09/2028`.
- **Đơn vị đối tác** (`A06`): Đơn vị phối hợp (nếu có, không bắt buộc).
- **Địa điểm triển khai** (`A07`): Tỉnh/thành phố triển khai nghiên cứu (nếu để trống, mẫu sẽ giữ dấu `……` để điền tay).
- **Đầu mối liên hệ** (`A08`): Thông tin cán bộ đầu mối liên hệ trong thư mời chuyên gia (nếu để trống sẽ hiện `……`).
- **Chủ nhiệm đề tài** (`B01`): Họ và tên, học vị (PGS.TS., TS.BS...), đơn vị công tác.
- **Nghiên cứu viên** (`B04–B20`): Danh sách các thành viên thực hiện đề tài (tự động đánh số thứ tự `1. `, `2. ` trong Quyết định giao đề tài).
- **Thư ký các hội đồng**:
  - Thư ký HĐ Đạo đức (`C09`, `C10`): Bắt buộc đủ 2 người.
  - Thư ký HĐ Khoa học (`D09`, `D10`): Bắt buộc đủ 2 người.
  - Thư ký HĐ Nghiệm thu (`E09`, `E10`): Bắt buộc đủ 2 người.
- **Thành viên các hội đồng**: Điền Chủ tịch, phản biện 1, phản biện 2, các ủy viên của 3 hội đồng (Đạo đức, Khoa học, Nghiệm thu).

### 3.4 Chuẩn bị file Lý lịch khoa học
Công cụ tự động tìm và đính kèm file Lý lịch khoa học (CV) vào hồ sơ theo tên:

1. Đặt file Lý lịch khoa học (`.docx` hoặc `.pdf`) vào thư mục **`Lý lịch khoa học/`** ở thư mục gốc.
2. Tên file chỉ cần chứa họ tên của người đó (không phân biệt hoa/thường, có dấu hay không dấu) — ví dụ: `Lý lịch khoa học - Trương Hồng Sơn.docx`.
3. **Chủ nhiệm đề tài:** Chỉ cần tên ở `B01` đã khai, công cụ tự động tìm file lý lịch khoa học khớp tên trong thư mục `Lý lịch khoa học/` và copy vào `01. Hồ sơ đạo đức đề cương/`.
4. **Chuyên gia khác (nếu có CV cần kèm theo):** Khai tên ở `F02–F10` trong checklist; công cụ sẽ tự tìm file tương ứng trong `Lý lịch khoa học/`.

### 3.5 Lưu file Excel
- Nhấn **Ctrl + S** để lưu lại các thay đổi.

---

## 4. Chạy tạo hồ sơ

1. Chạy file **`chay_tao_ho_so.bat`** (hoặc mở Terminal gõ: `python tao_ho_so_moi.py`).
2. Nếu trong file Excel có nhiều sheet dự án, công cụ sẽ liệt kê danh sách để bạn chọn số thứ tự dự án cần tạo.
3. Chờ script xử lý:
   - Tự động copy file mẫu từ các thư mục `- MẪU`.
   - Tự động đính kèm file Lý lịch khoa học của Chủ nhiệm.
   - Tự động điền tất cả các token, thông tin hội đồng, danh sách nghiên cứu viên.
   - Tự động tạo công văn mời chuyên gia ngoài cho từng hội đồng.
4. Khi hoàn thành, công cụ thông báo đường dẫn thư mục hồ sơ đầu ra dạng: `Hồ sơ - <Tên đề tài> (<Năm>)`.

---

## 5. Cấu trúc bộ hồ sơ đầu ra

Thư mục kết quả chứa 4 phần hoàn chỉnh:
```
Hồ sơ - <Tên đề tài> (<Năm>)/
├── 01. Hồ sơ đạo đức đề cương/
│   ├── 00. QĐ Giao đề tài.docx
│   ├── 01. QĐTLHĐ đạo đức đề cương.docx
│   ├── 02. BB họp HĐ đạo đức.docx
│   ├── 03. BB kiểm phiếu HĐ đạo đức.docx
│   ├── 04. QĐ chấp nhận đạo đức.docx
│   ├── Bảng kiểm đánh giá đạo đức.docx
│   └── Lý lịch khoa học - <Tên Chủ nhiệm>.docx
├── 02. Hồ sơ khoa học đề cương/
│   ├── 05. QĐ TLHĐ khoa học xét đề cương.docx
│   ├── 06. BB họp thông qua đề cương.docx
│   ├── 07. BB kiểm phiếu thông qua đề cương.docx
│   ├── 08. QĐ phê duyệt đề tài.docx
│   ├── Phiếu chấm điểm HĐ đề cương.docx
│   └── Phiếu nhận xét đánh giá hồ sơ.docx
├── 03. Công văn mời chuyên gia/
│   ├── Công văn mời chuyên gia.docx (mời chuyên gia ngoài HĐ Đạo đức & Khoa học)
│   └── Công văn mời chuyên gia nghiệm thu.docx (mời chuyên gia ngoài HĐ Nghiệm thu)
└── 04. Hồ sơ nghiệm thu/
    ├── 9. Quyết định thành lập HĐ nghiệm thu.docx
    ├── 10. Biên bản họp HĐ nghiệm thu.docx
    ├── 11. Biên bản kiểm phiếu nghiệm thu.docx
    ├── 12. Quyết định công nhận kết quả đề tài.docx
    ├── Phiếu chấm điểm nghiệm thu (<TVCT_ĐGHQ hoặc TNLS>).docx
    ├── Phiếu ký nhận tiền.docx
    └── Phiếu nhận xét nghiệm thu.docx
```

---

---

## 6. Xử lý các lỗi thường gặp

### 6.1 Lỗi: Thiếu thông tin bắt buộc trong Excel
- **Dấu hiệu:** Script báo lỗi `ValueError: ... là bắt buộc nhưng đang trống` (ví dụ: Tên đề tài A01, Chủ nhiệm B01, Thư ký hội đồng...).
- **Cách khắc phục:** Mở Excel checklist, điền đầy đủ các ô màu vàng bị thiếu, lưu file rồi chạy lại.

### 6.2 Lỗi: Định dạng thời gian (A05) không hợp lệ
- **Dấu hiệu:** Script báo lỗi `ValueError: Không thể phân tích mốc thời gian...`.
- **Cách khắc phục:** Sửa ô A05 theo đúng định dạng có 2 mốc `MM/YYYY`, ví dụ `Tháng 01/2027 đến tháng 12/2027`.

### 6.3 Lỗi: Không tìm thấy file Lý lịch khoa học
- **Dấu hiệu:** Báo lỗi `FileNotFoundError: Không tìm thấy file CV nào khớp tên '...' trong thư mục 'Lý lịch khoa học/'`.
- **Cách khắc phục:**
  1. Mở thư mục `Lý lịch khoa học/`.
  2. Đảm bảo có file `.docx` hoặc `.pdf` chứa đúng họ và tên của Chủ nhiệm đề tài (ví dụ: `Lý lịch khoa học - Trương Hồng Sơn.docx`).
  3. Kiểm tra xem tên trong file và tên khai ở ô B01 có khớp nhau không.

### 6.4 Lỗi: File đang mở trong Microsoft Word
- **Dấu hiệu:** Báo lỗi file bị khóa (File is locked / Permission denied).
- **Cách khắc phục:** Đóng tất cả các cửa sổ Word đang mở file trong thư mục hồ sơ đầu ra rồi chạy lại.

---

## 7. Tài liệu Kỹ thuật dành cho Quản trị viên (Cấu hình Token & Mẫu mới)

Nếu bạn là **người quản trị hệ thống hoặc lập trình viên** muốn:
- Xem bảng tra cứu toàn bộ 19 token có sẵn trong hệ thống (`{{TEN_DE_TAI}}`, `{{CHU_NHIEM_HO_TEN}}`, `{{DANH_SACH_NGHIEN_CUU_VIEN}}`...)
- Chỉnh sửa hoặc thêm các file mẫu `.docx` mới vào 4 thư mục mẫu
- Thêm các trường token mới qua sheet ẩn `_Tokens` trong Excel checklist

👉 Vui lòng xem tài liệu kỹ thuật chi tiết tại: **[CAU_HINH_TOKEN_VA_FILE_MAU.md](file:///f:/tool-ho-so/CAU_HINH_TOKEN_VA_FILE_MAU.md)**.
