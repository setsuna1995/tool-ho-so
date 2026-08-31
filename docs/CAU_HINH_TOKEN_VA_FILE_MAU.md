# Hướng dẫn Kỹ thuật: Cấu hình Token & Tạo Mẫu Tài Liệu Mới

Tài liệu này dành cho **người quản trị / kỹ thuật viên** phụ trách thiết kế, chỉnh sửa file mẫu `.docx` và mở rộng hệ thống token trong công cụ tạo hồ sơ.

> [!NOTE]
> Đối với người sử dụng thông thường (chỉ cần nhập liệu Excel và chạy sinh hồ sơ), vui lòng xem [HUONG_DAN.md](file:///f:/tool-ho-so/HUONG_DAN.md).

---

## 1. Cơ chế Token hóa

Công cụ điền dữ liệu vào file mẫu `.docx` bằng cách tìm và thay các token dạng `{{TEN_TOKEN}}` được gõ sẵn trong nội dung văn bản Word. Dữ liệu thay thế được lấy tự động từ file `Form checklist hồ sơ dự án.xlsx` và file cấu hình độc lập `config_tokens.json`.

---

## 2. Bảng tra cứu Token dùng chung (19 Token mặc định)

Tất cả các token dưới đây được quản lý tập trung trong file cấu hình **[`config_tokens.json`](file:///f:/tool-ho-so/config_tokens.json)** ở thư mục gốc công cụ và tự động áp dụng cho tất cả các file mẫu:

| Token | Ý nghĩa | Lấy từ mã mục | Kiểu xử lý (`kind`) |
|---|---|---|---|
| `{{TEN_DE_TAI}}` | Tên đầy đủ của đề tài nghiên cứu | `A01` | `raw` |
| `{{NAM}}` | Năm thực hiện hồ sơ | `A03` | `raw` |
| `{{DON_VI_CHU_TRI}}` | Tên cơ quan chủ trì đề tài | `A04` | `raw` |
| `{{DON_VI_DOI_TAC}}` | Tên cơ quan phối hợp / đối tác (nếu có) | `A06` | `raw_or_placeholder` |
| `{{THOI_GIAN_BAT_DAU}}` | Mốc thời gian bắt đầu (dạng `MM/YYYY`) | `A05` | `timeline_start` |
| `{{THOI_GIAN_KET_THUC}}` | Mốc thời gian kết thúc (dạng `MM/YYYY`) | `A05` | `timeline_end` |
| `{{DIA_DIEM_TRIEN_KHAI}}` | Địa điểm triển khai nghiên cứu | `A07` | `raw_or_placeholder` |
| `{{DAU_MOI_LIEN_HE}}` | Thông tin đầu mối liên hệ thư mời | `A08` | `raw_or_placeholder` |
| `{{CHU_NHIEM_HO_TEN}}` | Họ tên có học hàm/học vị của Chủ nhiệm | `B01` | `person_ho_ten` |
| `{{CHU_NHIEM_TEN}}` | Chỉ tên Chủ nhiệm đề tài (không kèm học vị) | `B01` | `person_ten` |
| `{{CHU_NHIEM_DON_VI}}` | Đơn vị công tác của Chủ nhiệm đề tài | `B01` | `person_org` |
| `{{DONG_CHU_NHIEM_HO_TEN}}` | Họ tên có học vị Đồng chủ nhiệm (nếu có) | `B02` | `person_ho_ten` |
| `{{DONG_CHU_NHIEM_TEN}}` | Chỉ tên Đồng chủ nhiệm đề tài (nếu có) | `B02` | `person_ten` |
| `{{THU_KY_DE_TAI}}` | Họ tên có học vị Thư ký đề tài (nếu có) | `B03` | `person_ho_ten` |
| `{{DANH_SACH_NGHIEN_CUU_VIEN}}` | Danh sách nghiên cứu viên có đánh số `1. ...
2. ...` | `B04–B20` | `numbered_researchers` |
| `{{CHU_TICH_HD_DAO_DUC}}` | Họ tên có học vị Chủ tịch HĐ Đạo đức | `C01` | `person_ho_ten` |
| `{{CHU_TICH_HD_KHOA_HOC}}` | Họ tên có học vị Chủ tịch HĐ Khoa học | `D01` | `person_ho_ten` |
| `{{CHU_TICH_HD_NGHIEM_THU}}` | Họ tên có học vị Chủ tịch HĐ Nghiệm thu | `E01` | `person_ho_ten` |
| `{{CHU_TICH_HD_NGHIEM_THU_TEN}}` | Chỉ tên Chủ tịch HĐ Nghiệm thu | `E01` | `person_ten` |

---

## 3. Token riêng theo từng trang (Thư mời chuyên gia)

Thư mời chuyên gia đa trang ([`section_moi_chuyen_gia.py`](file:///f:/tool-ho-so/section_moi_chuyen_gia.py)) sinh ra từng trang thư mời riêng cho các chuyên gia ngoài thuộc 3 hội đồng. Trong file mẫu thư mời, sử dụng 2 token riêng được tính toán động lúc tạo trang:

| Token | Ý nghĩa |
|---|---|
| `{{CHUYEN_GIA_HO_TEN}}` | Họ tên kèm học hàm/học vị của chuyên gia nhận thư ở trang đó |
| `{{CHUYEN_GIA_DON_VI}}` | Đơn vị công tác của chuyên gia nhận thư ở trang đó |

---

## 4. Hướng dẫn thêm / sửa file mẫu `.docx`

File mẫu được lưu trữ tập trung trong thư mục **`templates/`**:
- `templates/01. Hồ sơ đạo đức đề cương - MẪU/`
- `templates/02. Hồ sơ khoa học đề cương - MẪU/`
- `templates/03. Công văn mời chuyên gia - MẪU/`
- `templates/04. Hồ sơ nghiệm thu - MẪU/`

### Cách 1: Mẫu dùng Jinja2 Template Engine (docxtpl) - Khuyến nghị cho mẫu mới
Bạn có thể viết trực tiếp biểu thức Jinja2 vào nội dung file Word:
- **Biến thông tin đề tài:** `{{ project.title }}`, `{{ project.year }}`, `{{ project.host_org }}`
- **Bộ lọc in hoa:** `{{ project.title | upper }}`
- **Họ tên chủ nhiệm/thư ký:** `{{ project.head.full_name }}`, `{{ project.project_secretary.full_name }}`
- **Điều kiện ẩn hiện đoạn văn:** `{% if project.co_head %}Đồng chủ nhiệm: {{ project.co_head.full_name }}{% endif %}`
- **Lặp danh sách trong bảng:** Đặt `{% tr for r in researchers %}` ở ô đầu, `{{ r.name }}` / `{{ r.org }}` ở các ô tiếp theo và `{% tr endfor %}` ở ô cuối. `docxtpl` sẽ tự động nhân bản dòng bảng!

### Cách 2: Mẫu dùng các Token phẳng truyền thống
1. Đặt file `.docx` mẫu vào thư mục `templates/* - MẪU` tương ứng. Tên file mẫu phải **giống hệt** tên file muốn sinh ra trong hồ sơ đầu ra.
2. Gõ trực tiếp các token như `{{TEN_DE_TAI}}`, `{{NAM}}`, `{{CHU_NHIEM_HO_TEN}}`... vào đúng vị trí trong nội dung file Word.
3. Mở file module Python tương ứng trong `src/` (`section_dao_duc.py`, `section_khoa_hoc.py`, hoặc `section_nghiem_thu.py`), thêm hàm sinh file:
   ```python
   def _ten_mau_moi(session, dest_dir, info, common_tokens):
       doc = session.open(dest_dir / "Tên file mẫu mới.docx")
       session.fill_tokens(doc, common_tokens)
       session.save_close(doc)
   ```
4. Gọi hàm `_ten_mau_moi(...)` trong hàm `generate()` của module đó.

### Cách 2: Mẫu cần ghi bảng danh sách hội đồng
Nếu mẫu có bảng danh sách thành viên hội đồng, dùng `committee_writer` để ghi tự động vào các dòng của bảng:
```python
def _ten_mau_hoi_dong(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Tên file.docx")
    session.fill_tokens(doc, common_tokens)
    committee_writer.write_committee_roster(
        session, doc, table_index=2, committee=info.ethics_committee,
        roles=ROLES, name_col=1, org_col=2, role_col=3
    )
    committee_writer.write_committee_secretaries(
        session, doc, table_index=3, committee=info.ethics_committee,
        name_col=1, org_col=2
    )
    session.save_close(doc)
```

---

## 5. Thêm Token mới qua File `config_tokens.json` (Không cần sửa Code)

Toàn bộ token dùng chung được khai báo độc lập trong file **[`config_tokens.json`](file:///f:/tool-ho-so/config_tokens.json)**. File `Form checklist hồ sơ dự án.xlsx` hoàn toàn sạch sẽ, chỉ chứa các sheet dữ liệu dự án.

### 5.1 Cấu trúc một mục Token trong `config_tokens.json`

```json
{
  "token_name": "TEN_TOKEN_MOI",
  "code": "A09",
  "kind": "raw_or_placeholder",
  "param": "……………………",
  "note": "Ghi chú ý nghĩa token"
}
```

- **`token_name`**: Tên token viết hoa, có gạch dưới, **không** kèm dấu `{{ }}` (ví dụ `DON_VI_TAI_TRO`).
- **`code`**: Mã ô trong checklist Excel mà token này đọc giá trị (ví dụ `A09`).
- **`kind`**: Kiểu chuyển đổi dữ liệu — một trong các kiểu ở mục 5.2.
- **`param`**: Giá trị mặc định khi ô trong checklist để trống (dành riêng cho `raw_or_placeholder`, ví dụ `……`).
- **`note`**: Ghi chú mô tả bằng tiếng Việt.

### 5.2 Các giá trị `kind` được hỗ trợ trong `token_rules.py`

| `kind` | Cách thức xử lý |
|---|---|
| `raw` | Đọc thẳng chuỗi văn bản từ ô (cột C của dòng mã mục tương ứng) |
| `raw_or_placeholder` | Đọc văn bản; nếu ô trống hoặc chưa có mã mục trong checklist thì trả về giá trị ở `param` (ví dụ `……`) |
| `person_ho_ten` | Dòng nhân sự → `"<học hàm/học vị> <họ và tên>"` |
| `person_ten` | Dòng nhân sự → chỉ lấy `<họ và tên>`, không kèm học hàm/học vị |
| `person_org` | Dòng nhân sự → lấy đơn vị công tác |
| `numbered_researchers` | Đọc danh sách nghiên cứu viên B04–B20 → nối thành chuỗi có đánh số `1. ...
2. ...` |
| `timeline_start` | Ô mốc thời gian → tách lấy mốc bắt đầu `MM/YYYY` |
| `timeline_end` | Ô mốc thời gian → tách lấy mốc kết thúc `MM/YYYY` |

### 5.3 Các bước thêm Token mới (2 Cách)

#### Cách 1: Thêm trực tiếp trên Excel (Cực nhanh - Dynamic Token Discovery)
1. Mở file **`Form checklist hồ sơ dự án.xlsx`**.
2. Thêm một dòng mới trong sheet dự án:
   - Cột A: Gõ tên token dạng `{{TEN_TOKEN_MOI}}` (hoặc chọn từ Dropdown).
   - Cột C: Nhập giá trị tương ứng.
3. Trong file Word: Gõ `{{TEN_TOKEN_MOI}}` vào vị trí cần điền.
4. Chạy tạo hồ sơ: Tool tự động nhận diện và thay thế giá trị mà **không cần sửa code hay cấu hình gì thêm**!

#### Cách 2: Khai báo vào `config_tokens.json`
1. Mở file **`config_tokens.json`**.
2. Thêm một khối JSON mới:
   ```json
   {
     "token_name": "CO_QUAN_TAI_TRO",
     "code": "A09",
     "kind": "raw_or_placeholder",
     "param": "……………………",
     "note": "Đơn vị tài trợ kinh phí"
   }
   ```
3. Chạy file `capnhat_nhan_su.bat` để cập nhật danh mục Token vào file Excel.

---

## 6. Hệ thống Bộ hồ sơ (Dossier Package System)

Công cụ cho phép quản lý và xuất tài liệu theo từng **Bộ hồ sơ** độc lập hoặc trọn bộ:

### 6.1 Các Bộ hồ sơ mặc định trong `src/dossier_packages.py`
1. **`full` (Trọn bộ đầy đủ):** Sinh toàn bộ 4 phần hồ sơ (`01.`, `02.`, `03.`, `04.`).
2. **`dao_duc` (Bộ hồ sơ Đạo đức đề cương):** Chỉ sinh `01. Hồ sơ đạo đức đề cương` và `03. Công văn mời chuyên gia` (kèm CV Chủ nhiệm).
3. **`khoa_hoc` (Bộ hồ sơ Khoa học đề cương):** Chỉ sinh `02. Hồ sơ khoa học đề cương` và `03. Công văn mời chuyên gia`.
4. **`nghiem_thu` (Bộ hồ sơ Nghiệm thu đề tài):** Chỉ sinh `04. Hồ sơ nghiệm thu` và `03. Công văn mời chuyên gia`.

### 6.2 Cách thêm một Bộ hồ sơ mới trong tương lai
Mở file **[`src/dossier_packages.py`](file:///f:/tool-ho-so/src/dossier_packages.py)** và thêm vào danh sách `_DEFAULT_PACKAGES`:
```python
DossierPackage(
    id="thu_nghiem_lam_sang",
    name="Bộ hồ sơ Thử nghiệm lâm sàng pha 3",
    description="Hồ sơ phục vụ nghiệm thu đề tài thử nghiệm lâm sàng",
    sections=["dao_duc", "khoa_hoc", "nghiem_thu"],
    template_dirs=[
        "01. Hồ sơ đạo đức đề cương - MẪU",
        "04. Hồ sơ nghiệm thu - MẪU",
    ],
    copy_head_cv=True,
    copy_expert_cvs=True,
)
```
Khi chạy `tao_ho_so_moi.py`, bộ hồ sơ mới sẽ tự động xuất hiện trong menu lựa chọn!

## 7. Quy trình Kiểm thử & Xác minh

Mỗi khi chỉnh sửa file mẫu hoặc thêm token/code mới, chạy bộ test tự động để đảm bảo tính toàn vẹn:

```bash
# Chạy toàn bộ test suite
pytest -v

# Chạy riêng kiểm tra token và checklist
pytest test_token_rules.py test_excel_reader.py -v
```
