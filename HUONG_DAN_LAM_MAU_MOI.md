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
Thêm vào bảng ở đầu tài liệu này và vào `build_common_tokens()` trong
`tokens.py` cùng lúc, để tài liệu này luôn khớp với code thật.
