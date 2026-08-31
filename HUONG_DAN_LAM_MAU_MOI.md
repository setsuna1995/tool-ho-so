# Hướng dẫn làm mẫu tài liệu mới (dùng token)

Công cụ điền dữ liệu vào file mẫu `.docx` bằng cách tìm và thay các token
dạng `{{TEN_BIEN}}` được gõ sẵn trong nội dung file mẫu — không còn tìm-thay
theo câu chữ dự án mẫu cũ nữa.

## Bảng token dùng chung

| Token | Lấy từ (ProjectInfo) | Mã checklist |
|---|---|---|
| `{{TEN_DE_TAI}}` | `info.title` | A01 |
| `{{NAM}}` | `info.year` | A03 |
| `{{DON_VI_CHU_TRI}}` | `info.host_org` | A04 |
| `{{DON_VI_DOI_TAC}}` | `info.partner_org` (rỗng nếu không khai) | A06 |
| `{{CHU_NHIEM_HO_TEN}}` | `"<học vị> <tên>"` của chủ nhiệm | B01 |
| `{{CHU_NHIEM_TEN}}` | chỉ tên chủ nhiệm, không kèm học vị | B01 |
| `{{DONG_CHU_NHIEM_TEN}}` | tên đồng chủ nhiệm (rỗng nếu không có) | B02 |
| `{{THU_KY_DE_TAI}}` | `"<học vị> <tên>"` của thư ký đề tài (rỗng nếu không có) | B03 |
| `{{THOI_GIAN_BAT_DAU}}` / `{{THOI_GIAN_KET_THUC}}` | tách từ mốc thời gian | A05 |
| `{{DIA_DIEM_TRIEN_KHAI}}` | địa điểm triển khai (dấu `……` nếu không khai) | A07 |
| `{{DAU_MOI_LIEN_HE}}` | `info.common_tokens` | A08 |

## Token riêng theo trang (chỉ dùng trong thư mời chuyên gia)

Khác với bảng token dùng chung ở trên (1 giá trị/dự án, khai qua sheet
`_Tokens`), thư mời chuyên gia (`expert_invitation.py`) còn có 2 token
**riêng theo từng trang**, tính động lúc sinh hồ sơ theo từng người nhận,
không khai báo qua Excel:

| Token | Ý nghĩa |
|---|---|
| `{{CHUYEN_GIA_HO_TEN}}` | `"<học vị> <tên>"` của người nhận trang đó |
| `{{CHUYEN_GIA_DON_VI}}` | Đơn vị công tác của người nhận trang đó |

## Cách 1 — Mẫu chỉ cần token dùng chung (không cần logic riêng, nhưng vẫn cần một hàm 3 dòng)

**Quan trọng:** công cụ chỉ **tự copy** file `.docx` từ thư mục `- MẪU` vào hồ
sơ đầu ra (xem `copy_templates`/`discover_copies` trong `tao_ho_so_moi.py`) —
nó **không** tự mở file ra để điền token. Nếu bỏ qua bước 3 dưới đây, file mẫu
sẽ được copy nguyên với các `{{TOKEN}}` chưa điền vào hồ sơ đầu ra, và không
có cảnh báo gì báo cho bạn biết (`fill_tokens` không tự la lên khi không có
ai gọi nó).

1. Đặt file `.docx` vào đúng thư mục `- MẪU` tương ứng, đặt tên file **giống
   hệt** tên sẽ xuất hiện trong hồ sơ đầu ra (không có hậu tố `" - MẪU"`).
2. Gõ trực tiếp các token cần dùng (ví dụ `{{TEN_DE_TAI}}`, `{{NAM}}`) vào
   đúng chỗ trong nội dung file Word.
3. Mở file `section_*.py` tương ứng với phần hồ sơ đó, thêm một hàm nhỏ chỉ
   gồm 3 dòng (mở file, điền token, lưu — không có logic riêng gì thêm), rồi
   gọi hàm đó trong `generate()` của file đó:
   ```python
   def _ten_ham(session, dest_dir, info, common_tokens):
       doc = session.open(dest_dir / "Tên file.docx")
       session.fill_tokens(doc, common_tokens)
       session.save_close(doc)
   ```
4. Viết/cập nhật test tương ứng trong `test_section_*.py`, xác nhận token đã
   được điền và không còn `{{...}}` nào sót lại trong file.
5. Chạy thử toàn bộ `python tao_ho_so_moi.py` với một sheet Excel thử
   nghiệm trước khi dùng cho dự án thật.

## Cách 2 — Mẫu cần dữ liệu riêng (bảng hội đồng, chọn file theo điều kiện...)

Giống hệt Cách 1, chỉ khác ở bước 3: hàm `_ten_ham` có thêm logic riêng của
mẫu đó, chèn vào giữa `fill_tokens` và `save_close` — ví dụ ghi bảng hội đồng:

```python
def _ten_ham(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Tên file.docx")
    session.fill_tokens(doc, common_tokens)
    # phần logic riêng của mẫu này, ví dụ bảng hội đồng:
    committee_writer.write_committee_roster(session, doc, 2, info.ethics_committee, ...)
    session.save_close(doc)
```

Nếu mẫu cần một trường dữ liệu chưa có trong checklist Excel: thêm mã mục
mới vào `Form checklist hồ sơ dự án.xlsx` (theo mẫu
`migrate_add_research_location.py`), đọc trường đó trong `excel_reader.py`,
rồi quyết định: nếu trường đó dùng chung cho nhiều mẫu → thêm vào
`build_common_tokens()` trong `tokens.py`; nếu chỉ dùng riêng cho 1 mẫu →
truyền trực tiếp trong hàm `_ten_ham`.

Các bước còn lại (viết test, chạy thử toàn bộ) giống hệt bước 4-5 của Cách 1.

## Trường không thể biết trước lúc sinh hồ sơ (ngày họp, số quyết định...)

Không cần token — gõ thẳng dấu `……` (chấm chấm) vào đúng chỗ trong file
mẫu, để người dùng tự điền tay sau khi hồ sơ được tạo ra.

## Thêm token dùng chung mới

Đặt tên `{{VIET_HOA_CO_GACH_DUOI}}`, có ý nghĩa rõ ràng bằng tiếng Việt.

Danh sách token dùng chung **được khai báo bằng dữ liệu** trong sheet ẩn
`_Tokens` của file `Form checklist hồ sơ dự án.xlsx` — không phải bằng code.
`excel_reader.load_project_data` đọc sheet đó qua `token_rules.resolve_tokens`
và tính sẵn toàn bộ token một lần (`info.common_tokens`);
`tokens.build_common_tokens(info)` chỉ là hàm trả thẳng lại giá trị đó, **không
còn logic riêng cho từng token để thêm vào nữa**.

Vì vậy, trong đa số trường hợp thêm token mới **không cần sửa code Python** —
chỉ cần thêm một dòng vào sheet `_Tokens`.

### Cấu trúc sheet `_Tokens` (5 cột)

| Cột | Tên cột | Ý nghĩa |
|---|---|---|
| A | `token_name` | Tên token VIẾT_HOA_GẠCH_DƯỚI, **không** kèm dấu `{{ }}`. Trong file `.docx` bạn gõ `{{TEN_NAY}}`. |
| B | `code` | Mã mục trong checklist mà token này đọc (ví dụ `A08`) |
| C | `kind` | Kiểu xử lý giá trị — một trong 6 giá trị hợp lệ ở bảng dưới |
| D | `param` | Tham số phụ, **chỉ** `raw_or_placeholder` dùng đến |
| E | `note` | Mô tả tự do bằng tiếng Việt, để người sau đọc hiểu |

### 6 giá trị `kind` hợp lệ

| `kind` | Làm gì |
|---|---|
| `raw` | Đọc thẳng nội dung ô (cột C của dòng mã mục đó) |
| `raw_or_placeholder` | Như `raw`, nhưng nếu ô trống — hoặc mã mục chưa có trong checklist — thì dùng giá trị ở cột `param` làm chỗ đánh dấu (ví dụ `……`) |
| `person_ho_ten` | Dòng nhân sự → `"<học hàm/học vị> <tên>"` (rỗng nếu chưa khai tên) |
| `person_ten` | Dòng nhân sự → chỉ tên, không kèm học hàm/học vị |
| `timeline_start` | Ô mốc thời gian → mốc bắt đầu dạng `MM/YYYY` |
| `timeline_end` | Ô mốc thời gian → mốc kết thúc dạng `MM/YYYY` |

Danh sách này là bản sao của dict `TRANSFORMS` trong `token_rules.py` — nếu gõ
một `kind` không nằm trong 6 giá trị trên, công cụ sẽ **báo lỗi rõ ràng** ngay
khi chạy chứ không âm thầm bỏ qua.

### Các bước

1. Mở file `Form checklist hồ sơ dự án.xlsx`.
2. Nếu token cần một trường dữ liệu chưa có trong checklist: thêm mã mục mới
   vào 2 sheet dự án trước (theo mẫu `migrate_add_research_location.py`).
3. Vào sheet `_Tokens` (sheet này bình thường bị ẩn — click chuột phải vào tab
   sheet bất kỳ, chọn "Unhide", rồi chọn `_Tokens`).
4. Thêm một dòng mới, điền đủ 5 cột theo bảng trên. Ví dụ, thêm token
   `{{DIA_BAN_TINH}}` đọc mã mục `A08`, nếu bỏ trống thì hiện dấu chấm chấm:

   | token_name | code | kind | param | note |
   |---|---|---|---|---|
   | `DIA_BAN_TINH` | `A08` | `raw_or_placeholder` | `……` | Tỉnh/thành triển khai |

5. Lưu file Excel (Ctrl + S).
6. **Chỉ khi** không có `kind` nào trong 6 kiểu trên phù hợp: mở `token_rules.py`,
   viết thêm một hàm `_resolve_<ten_kind>(ws, index, spec)` và đăng ký nó vào
   dict `TRANSFORMS`. Đây là trường hợp **duy nhất** phải sửa code xử lý token.
7. Thêm token vào bảng ở đầu file này (`HUONG_DAN_LAM_MAU_MOI.md`) để tài liệu
   luôn khớp với sheet `_Tokens` thật.
8. Viết/cập nhật test trong `test_token_rules.py` cho token mới.
9. Chạy `pytest` để xác nhận test đã pass.
10. Bây giờ bạn có thể gõ `{{DIA_BAN_TINH}}` trong file mẫu `.docx` — nó sẽ được
    điền tự động khi chạy công cụ (miễn là mẫu đó đã có hàm `_ten_ham` gọi
    `fill_tokens` như ở Cách 1 bước 3).
