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
- **Kiểu nghiên cứu** (A02) — **chỉ chấp nhận đúng 2 giá trị:** `TVCT_ĐGHQ` hoặc `TNLS`. Công cụ dùng đúng giá trị này để **tự chọn đúng mẫu phiếu chấm điểm nghiệm thu tương ứng** — gõ sai chính tả (thừa/thiếu dấu, khoảng trắng...) sẽ khiến script báo lỗi khi chạy.
- **Năm thực hiện** (A03) — ví dụ: 2026
- **Đơn vị chủ trì** (A04)
- **Thời gian triển khai NC** (A05) — **bắt buộc** phải có đủ 2 mốc dạng `MM/YYYY`, ví dụ: "Tháng 01/2027 đến tháng 12/2027" hoặc "Từ 03/2027 đến 09/2028". Công cụ tự tách mốc bắt đầu/kết thúc từ đây để điền vào các văn bản — sai định dạng sẽ khiến script báo lỗi rõ ràng khi chạy.
- **Chủ nhiệm đề tài** (B01) — tên, học vị (ví dụ: PGS.TS.), và đơn vị công tác
- **Nghiên cứu viên** (B04–B20) — điền tên các thành viên nhóm nghiên cứu; danh sách này sẽ được **tự động điền** vào bảng "danh sách đơn vị triển khai" trong file "00. QĐ Giao đề tài.docx"
- **Thư ký hội đồng đạo đức** (C09, C10) — **bắt buộc** phải điền đủ cả 2 người (tên, học vị, đơn vị)
- **Thư ký hội đồng khoa học** (D09, D10) — **bắt buộc**
- **Thư ký hội đồng nghiệm thu** (E09, E10) — **bắt buộc**
- **Các thành viên hội đồng** (Chủ tịch, phản biện, ủy viên của 3 hội đồng)
- **Hồ sơ Chủ nhiệm đề tài** (F01) — **bắt buộc**. Cột "TÊN FILE CV" phải ghi **đúng tên file** (kể cả hoa/thường, khoảng trắng) của file CV chủ nhiệm đề tài — xem chi tiết ở mục 3.5.

**Lưu ý:** Bạn có thể điền thêm các ô không tô vàng (như đơn vị đối tác, cộng tác viên), nhưng nó không bắt buộc.

### 3.4 Chuẩn bị file CV chủ nhiệm đề tài

Công cụ tự động đính kèm CV chủ nhiệm đề tài vào hồ sơ đạo đức, lấy theo đúng tên file bạn khai ở mã mục **F01** (cột "TÊN FILE CV"). Để việc này chạy đúng:

1. Đặt file CV (`.docx`) của chủ nhiệm đề tài vào thư mục **`CV chuyên gia/`** ở thư mục gốc công cụ.
2. Mở Excel, vào ô F01 (cột thứ 3 trong bảng — "TÊN FILE CV"), gõ **đúng tên file** bạn vừa đặt vào thư mục đó, kể cả hoa/thường và khoảng trắng.
3. Nếu chủ nhiệm đề tài dự án mới khác với dự án trước, chỉ cần đổi tên file trong ô F01 cho khớp file CV mới — **không cần sửa code**.

Nếu tên file trong F01 không khớp file nào trong thư mục `CV chuyên gia/`, script sẽ báo lỗi rõ ràng ngay khi chạy, kèm hướng dẫn khắc phục.

### 3.5 Lưu file Excel
- Nhấn **Ctrl + S** để lưu

## 4. Chạy tạo hồ sơ

Sau khi đã chuẩn bị đủ dữ liệu trong Excel, các bước tiếp theo:

### 4.1 Mở Terminal trong thư mục công cụ
- Cách dễ nhất: Trong File Explorer, click phải trong thư mục công cụ → chọn "Open in Terminal" (hoặc "Open PowerShell window here")

### 4.2 Chạy script và chọn dự án
1. Gõ lệnh:
   ```
   python tao_ho_so_moi.py
   ```
2. Nhấn **Enter**. Script sẽ hiện danh sách các sheet dự án đang có trong file Excel, đánh số thứ tự, ví dụ:
   ```
   Chon sheet du an muon tao ho so:
     1. Đề tài - Bánh ăn dặm VIAM 2027
     2. Đề tài - Nghiên cứu về dinh dưỡng 2026
   Nhap so thu tu (1-2):
   ```
3. Gõ **số thứ tự** đúng với sheet dự án bạn vừa tạo (ví dụ: `2`), rồi nhấn **Enter**.

**Không cần sửa file `.py` nào nữa** — chỉ cần chọn đúng số trong danh sách.

**Mẹo cho người dùng thành thạo:** có thể bỏ qua bước chọn bằng cách truyền thẳng tên sheet làm đối số dòng lệnh:
```
python tao_ho_so_moi.py "Đề tài - Nghiên cứu về dinh dưỡng 2026"
```
(tên phải khớp **chính xác** với tên sheet trong Excel — khoảng trắng, dấu, chữ hoa-thường đều phải đúng)

Script sẽ in ra các bước đang xử lý:
- "Dang doc du lieu tu Excel checklist..."
- "Dang sao chep file mau..."
- "Dang sao chep CV chu nhiem de tai..."
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
   - Kiểm tra **danh sách chủ nhiệm + thành viên** trong bảng "danh sách đơn vị triển khai" có khớp với nhóm nghiên cứu thực tế không
4. Kiểm tra file CV (`Lý lịch khoa học ...docx` hoặc tên bạn khai ở F01) đã có mặt trong cùng thư mục và đúng là CV của chủ nhiệm đề tài
5. Vào thư mục con `01. Hồ sơ đạo đức đề cương` → mở file `01. QĐTLHĐ đạo đức đề cương.docx`
   - Kiểm tra **tên chủ tịch hội đồng** có chính xác không
   - Kiểm tra **đơn vị** của các thành viên có đúng không
6. Vào thư mục con `04. Hồ sơ nghiệm thu`, xác nhận chỉ có **đúng một** file "Phiếu chấm điểm nghiệm thu" khớp với kiểu nghiên cứu (A02) của dự án

Nếu tất cả đều chính xác, hồ sơ của bạn đã sẵn sàng gửi đi!

## Những điều cần tự kiểm tra/sửa tay cho dự án mới

Công cụ tự động hóa được phần lớn công việc. Các mục sau **đã được tự động hóa** (không cần sửa tay nữa, chỉ cần điền đúng dữ liệu Excel):

- **Danh sách thành viên nhóm nghiên cứu** trong "00. QĐ Giao đề tài.docx" — tự điền từ B01 (chủ nhiệm) và B04–B20 (nghiên cứu viên).
- **Phiếu chấm điểm nghiệm thu** — tự chọn đúng mẫu "TVCT_ĐGHQ" hoặc "TNLS" theo giá trị bạn điền ở A02.
- **File CV chuyên gia đính kèm hồ sơ đạo đức** — tự copy đúng file theo tên khai ở F01 (xem mục 3.5).
- **Mốc thời gian nghiên cứu** — tự tách mốc bắt đầu/kết thúc thật từ A05, không còn suy đại từ tháng 01 đến tháng 12 của năm thực hiện nữa.

Vẫn còn một chỗ **chưa** tự động, cần tự kiểm tra/điền tay:

- **Các dấu "……" còn sót lại:** Một số file có thể vẫn còn các chỗ đánh dấu chưa điền (ví dụ: ngày họp cụ thể, số quyết định...) cần bạn tự điền tay sau khi hồ sơ được tạo ra. Đây là tình trạng đã có từ quy trình cũ (script PowerShell) và chưa thay đổi trong công cụ này.

## 6. Về các script chạy một lần duy nhất

Trong thư mục công cụ, có các file script chạy một lần sau:
- `migrate_add_partner_org.py`
- `migrate_remove_template_config_sheet.py`
- `migrate_fix_f01_cv_filename.py`
- `convert_doc_templates.py`

**Những script này chỉ cần chạy MỘT LẦN duy nhất** khi công cụ được thiết lập lần đầu (hoặc khi có hướng dẫn nâng cấp). Bạn **KHÔNG** cần chạy lại chúng cho mỗi dự án mới.

Bạn chỉ cần chạy lại `convert_doc_templates.py` nếu bạn thêm một file mẫu `.doc` mới vào một trong 4 thư mục "- MẪU" (xem mục 6.1) — công cụ sẽ **tự nhận ra** file `.doc` nào chưa có bản `.docx` song song, không cần khai báo ở đâu cả.

Nếu không chắc, hãy liên hệ với người tạo công cụ.

### 6.1 Quy ước đặt tên và tổ chức thư mục/file mẫu

Toàn bộ 4 thư mục mẫu gốc ở thư mục gốc dự án (`01. Hồ sơ đạo đức đề cương - MẪU`, `02. Hồ sơ khoa học đề cương - MẪU`, `03. Công văn mời chuyên gia - MẪU`, `04. Hồ sơ nghiệm thu - MẪU`) đều theo **đúng một quy ước duy nhất**:

- **Tên thư mục** = `<STT>. <Tên phần hồ sơ đầu ra> - MẪU` (luôn có hậu tố `" - MẪU"` ở cuối).
- **Tên file bên trong** phải **giống hệt 100%** tên file sẽ xuất hiện trong hồ sơ đầu ra (đúng từng chữ hoa/thường, không có tiền tố/hậu tố thừa như tên người, tên dự án cũ...).
- **Mỗi thư mục "- MẪU" chỉ được chứa đúng những file sẽ copy vào hồ sơ đầu ra** — công cụ **tự quét** toàn bộ file `.docx` trong 4 thư mục này, không đọc danh sách khai báo từ Excel nữa (sheet "Cấu hình mẫu" đã bị xoá). File tham khảo/slide/PDF/biến thể không thuộc luồng chuẩn phải để ở thư mục riêng **`Tài liệu tham khảo (không dùng tạo hồ sơ)/`**, không được để lẫn trong 4 thư mục "- MẪU".

Nhờ quy ước này, đường dẫn **đích** trong bộ hồ sơ đầu ra luôn tự suy ra được từ đường dẫn **nguồn** — chỉ cần bỏ hậu tố `" - MẪU"` khỏi tên thư mục gốc, giữ nguyên phần còn lại. Ví dụ:

```
Nguồn: 01. Hồ sơ đạo đức đề cương - MẪU/00. QĐ Giao đề tài.docx
Đích:  01. Hồ sơ đạo đức đề cương/00. QĐ Giao đề tài.docx
```

### 6.2 Khi cần thêm một mẫu tài liệu hoàn toàn mới

Xem hướng dẫn chi tiết tại [`HUONG_DAN_LAM_MAU_MOI.md`](HUONG_DAN_LAM_MAU_MOI.md)
— quy trình đã đổi sang dùng token `{{TEN_BIEN}}` thay vì tìm-thay theo câu
chữ mẫu cũ.

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

### 7.2b Lỗi: Kiểu nghiên cứu (A02) không hợp lệ
**Dấu hiệu:** Script báo lỗi kiểu: "Loại hình nghiên cứu 'xxx' (mã A02) không hợp lệ - chỉ chấp nhận ['TNLS', 'TVCT_ĐGHQ']"

**Nguyên nhân:** Ô A02 bị gõ sai chính tả/khoảng trắng — công cụ dùng đúng giá trị này để tự chọn mẫu phiếu chấm điểm nghiệm thu.

**Cách khắc phục:** Sửa ô A02 thành đúng một trong hai giá trị `TVCT_ĐGHQ` hoặc `TNLS` (không thêm khoảng trắng, đúng dấu), lưu Excel rồi chạy lại.

### 7.2c Lỗi: Không đọc được mốc thời gian nghiên cứu (A05)
**Dấu hiệu:** Script báo lỗi kiểu: "Không đọc được mốc thời gian nghiên cứu (A05) từ nội dung '...'"

**Nguyên nhân:** Ô A05 không chứa đủ 2 mốc dạng `MM/YYYY`.

**Cách khắc phục:** Sửa ô A05 theo đúng mẫu có 2 mốc `MM/YYYY`, ví dụ "Tháng 01/2027 đến tháng 12/2027", lưu Excel rồi chạy lại.

### 7.2d Lỗi: Không tìm thấy file CV chuyên gia
**Dấu hiệu:** Script báo lỗi kiểu: "Không tìm thấy file CV '...' (khai báo ở mã mục F01) trong thư mục 'CV chuyên gia/'"

**Nguyên nhân:** Tên file khai ở ô F01 không khớp (kể cả sai một ký tự, hoa/thường, khoảng trắng) với tên file thật trong thư mục `CV chuyên gia/`, hoặc bạn quên đặt file CV vào đó.

**Cách khắc phục:**
1. Mở thư mục `CV chuyên gia/`, kiểm tra tên file CV thật
2. Sửa ô F01 cho khớp chính xác tên file đó (hoặc đổi tên file cho khớp ô F01)
3. Lưu Excel rồi chạy lại

### 7.3 Lỗi: File đang mở trong Word
**Dấu hiệu:**
- Script in ra lỗi kiểu: "File is locked" hoặc "File is open in another application"
- Script dừng giữa chừng, các file chưa được tạo hoàn toàn

**Nguyên nhân:** Bạn đang mở một trong các file `.docx` trong thư mục `Hồ sơ - ...` bằng Word trong khi script chạy

**Cách khắc phục:**
1. Đóng tất cả các file Word đang mở
2. Chạy lại script

### 7.3b Lỗi COM chung chung khi công cụ nằm trong OneDrive/SharePoint đồng bộ
**Dấu hiệu:** Script báo lỗi kiểu `(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'Command failed', ...))`, thường ngay khi đang mở hoặc ghi vào một file `.docx` nào đó (thông báo lỗi mới sẽ cho biết rõ tên file và ô đang xử lý).

**Nguyên nhân:** Nếu thư mục công cụ nằm trong một thư mục **đồng bộ OneDrive/SharePoint** (ví dụ đường dẫn có chữ "OneDrive"), tính năng AutoSave, chính sách Protected View của tài khoản Office 365 tổ chức, hoặc việc OneDrive đang khoá file để đồng bộ lên cloud có thể làm gián đoạn Word đang chạy ngầm (COM), gây ra lỗi chung chung này.

**Công cụ đã tự xử lý phần lớn vấn đề này:** từ phiên bản hiện tại, toàn bộ việc mở/ghi/lưu bằng Word được thực hiện ở một **thư mục tạm hoàn toàn local** (không nằm trong OneDrive) — chỉ khi xong xuôi mới copy nguyên bộ hồ sơ vào đúng thư mục đích (thao tác copy file thuần, không qua Word). Nếu vẫn gặp lỗi này:
1. Đọc kỹ thông báo lỗi mới — nó cho biết chính xác file/ô nào đang xử lý khi lỗi xảy ra, gửi lại thông tin này để được hỗ trợ tiếp
2. Nếu lỗi xảy ra, các file dang dở sẽ được **giữ lại** ở một thư mục tạm (đường dẫn được in ra trong thông báo lỗi, dạng `...\AppData\Local\Temp\tao_ho_so_...`) để kiểm tra — thư mục đích trong OneDrive sẽ **không** bị tạo dở dang
3. Thử tắt tạm AutoSave (thanh công cụ trên cùng cửa sổ Word, nút gạt "Lưu tự động") trước khi chạy script, rồi bật lại sau
4. Nếu vẫn không được, thử tạm dừng đồng bộ OneDrive (click phải biểu tượng OneDrive ở khay hệ thống → "Pause syncing") trong lúc chạy script, rồi bật lại sau khi xong

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
