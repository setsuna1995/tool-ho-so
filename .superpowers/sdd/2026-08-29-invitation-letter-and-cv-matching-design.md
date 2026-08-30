# Multi-Page Expert Invitation Letters & Name-Based CV Matching — Design Spec

*(Thư mời chuyên gia nhiều trang & tự động khớp CV theo tên — Đặc tả thiết kế)*

- **Date / Ngày:** 2026-08-29
- **Status / Trạng thái:** Approved by user in chat, written up for record — *Đã được người dùng duyệt qua chat, viết lại để lưu hồ sơ*
- **Classification / Phân loại:** Architectural — *Kiến trúc*

## 1. Background & Problem

*Bối cảnh & vấn đề*

Two related gaps were found while auditing the "công văn mời chuyên gia" (expert invitation letter) flow and the PHẦN F (expert CV) checklist section:

1. **Invitation letters are single-page and manually addressed.** The only invitation letter in the system (`03. Công văn mời chuyên gia - MẪU/Công văn mời chuyên gia.docx`, covering the combined Hội đồng khoa học + Hội đồng đạo đức meeting) has three dotted blanks (`Kính gửi: …`, `trân trọng kính mời …`, `gửi hồ sơ … để đọc`) that `section_moi_chuyen_gia.py` never fills — every generated project ships this letter with the recipient name left for a human to type by hand, once per invitee. There is no invitation letter at all for Hội đồng nghiệm thu (acceptance council), even though PHẦN F/checklist requires that committee's data just like the other two.
2. **CV attachment requires typing an exact filename.** PHẦN F (F01–F10) requires a "TÊN FILE CV" column typed by hand, matched byte-for-byte (including case/whitespace) against files in `CV chuyên gia/`. This is exactly the class of typo-prone manual step already flagged in `HUONG_DAN.md` §7.2d, and a dedicated one-off script (`capnhat_danh_sach_cv.py`) exists solely to keep its dropdown list in sync with the folder.

*Hai lỗ hổng: (1) thư mời chuyên gia hiện chỉ có 1 trang, chỗ ghi tên người nhận vẫn là dấu chấm chấm phải gõ tay, và chưa có thư mời riêng cho Hội đồng nghiệm thu; (2) khai báo CV chuyên gia (PHẦN F) bắt buộc gõ đúng tên file, dễ sai chính tả, cần cả một script riêng chỉ để đồng bộ dropdown tên file.*

A hardcode audit (grepping all 20 `- MẪU` template masters for the old sample project's real names) additionally found two genuinely-unwired spots, bundled into this same pass because they touch the same "committee data → template" pipeline:

- `04. Hồ sơ nghiệm thu - MẪU/Phiếu ký nhận tiền.docx`: of 6 signer rows, only the secretary's row (row 7 in Word's 1-indexed table) is filled by code (`section_nghiem_thu._phieu_ky_nhan_tien`); the other 5 rows (chair + 2 reviewers + 2 members) still show the old sample project's committee names verbatim.
- The invitation letter's contact line ("Ông Lê Minh Khánh, Trung tâm nghiên cứu - Viện Y học ứng dụng Việt Nam (Email: … - Điện thoại: …)") is static text, never wired to checklist data.

*(Hai chỗ hardcode thật đã xác minh bằng cách đối chiếu code với nội dung file mẫu thật — khác với bảng "ĐƠN VỊ TRIỂN KHAI"/hội đồng ở `00. QĐ Giao đề tài.docx`/`01. QĐTLHĐ đạo đức đề cương.docx`, vốn đã được `set_cell`/`committee_writer` ghi đè động, không phải bug.)*

## 2. Goals

*Mục tiêu*

1. Generate one Word file per applicable council, with one page per external (non-host-org) committee member, each page addressed by name — no more hand-typed dotted blanks. *Sinh 1 file/hội đồng, mỗi trang 1 chuyên gia ngoài đơn vị chủ trì, tự điền tên.*
2. Add the missing Hội đồng nghiệm thu invitation letter. *Bổ sung thư mời còn thiếu cho hội đồng nghiệm thu.*
3. Fix the two confirmed hardcode gaps (`Phiếu ký nhận tiền.docx`, contact line). *Vá 2 lỗ hổng hardcode đã xác minh.*
4. Replace the PHẦN F "TÊN FILE CV" manual column with automatic name-based matching against `CV chuyên gia/`, for both the head's CV and other declared experts' CVs. Retire the now-purposeless dropdown-refresh script. *Bỏ cột tên file CV thủ công, tự khớp theo tên; gỡ bỏ script đồng bộ dropdown không còn cần thiết.*

### Non-Goals

*Ngoài phạm vi*

- No macro/VBA Excel automation — matching happens in the Python generation step, not live in Excel (consistent with the project's existing all-Python-scripts convention).
- No change to how committee roster tables themselves are populated (`committee_writer.write_committee_roster`/`write_committee_secretaries` already work correctly) — only the invitation-letter and CV-attachment layers change.

## 3. Multi-Page Letter Generation Mechanism

*Cơ chế sinh thư nhiều trang*

The existing `word_writer.Session` abstraction (COM-preferred, `python-docx` fallback) has no operation for cloning a block of content N times with page breaks — COM Range copy/paste across pages is unreliable, and this is a structural edit, not a search-replace. **New module `expert_invitation.py` bypasses `Session` entirely for this one file type and works directly with `python-docx`**, regardless of which backend the rest of the run is using. This is a deliberate, scoped exception to the Session-abstraction pattern used everywhere else in the codebase; it is safe because duplication copies the source page's own XML verbatim (no reformatting risk), and text substitution reuses the exact same paragraph-level replace helper already used by the fallback backend (`word_writer._docx_replace_in_paragraph`), so no new run-splitting risk is introduced beyond what already exists and is documented in `HUONG_DAN.md`.

```python
def generate_multi_page_letter(path: Path, recipients: list[Person], common_tokens: dict) -> bool:
    """Mo file .docx tai `path` (da duoc copy san tu thu muc MAU), nhan ban
    toan bo noi dung than file (tru sectPr) thanh N trang - moi trang 1
    nguoi trong `recipients` - roi ghi de luu lai dung `path`.

    Tra ve False (khong dong/luu gi ca, giu nguyen file mau chua dien) neu
    `recipients` rong.
    """
```

Per-page tokens `{{CHUYEN_GIA_HO_TEN}}` (`"<degree> <name>"`) and `{{CHUYEN_GIA_DON_VI}}` (org) are typed directly into the template's three dotted-blank spots, replacing the dots. Page breaks are inserted between clones (not before the first). Implementation clones each top-level body element via `copy.deepcopy`, inserts before the document's `sectPr` (which must remain the last body child per OOXML), and applies both `common_tokens` and the page's own recipient tokens to only that page's cloned elements.

*Token riêng theo trang (`{{CHUYEN_GIA_HO_TEN}}`, `{{CHUYEN_GIA_DON_VI}}`) khác với token dùng chung (`_Tokens` sheet) — token riêng chỉ áp dụng cho đúng 1 trang, tính động lúc sinh hồ sơ, không khai báo qua Excel.*

## 4. Recipient Filtering & Deduplication

*Lọc & khử trùng lặp người nhận*

For each committee (`CommitteeData`), candidates = `chair + reviewers + members` (secretaries excluded — they are host-org staff in every observed case and are not the ones being invited). A candidate qualifies if `person.org.strip().lower() != info.host_org.strip().lower()`.

- **Thư mời đề cương** (`Công văn mời chuyên gia.docx`, stays in `03. Công văn mời chuyên gia`): recipients = qualifying members of `ethics_committee` **∪** `proposal_committee`, deduplicated by `(name, org)` case-insensitive (both councils typically share the same external reviewers meeting the same day).
- **Thư mời nghiệm thu** (new file `Công văn mời chuyên gia nghiệm thu.docx`, same folder): recipients = qualifying members of `acceptance_committee` only.

If a letter's recipient list is empty, `generate_multi_page_letter` returns `False` and the caller prints an informational (not error) message; the file is left as the unfilled template copy already staged by `copy_templates`.

## 5. New Checklist Field: Đầu mối liên hệ

*Trường Excel mới: Đầu mối liên hệ*

New optional checklist code **A08** "Đầu mối liên hệ (họ tên, đơn vị, email, điện thoại)" — one freeform text cell, e.g. `"Ông Lê Minh Khánh, Trung tâm nghiên cứu - Viện Y học ứng dụng Việt Nam (Email: ... - Điện thoại: ...)"`. Wired as a new common token `{{DAU_MOI_LIEN_HE}}` (`kind=raw_or_placeholder`, `param="……"`) purely through the existing `_Tokens`-sheet mechanism — **no new Python token-resolution code needed**, per the established pattern in `HUONG_DAN_LAM_MAU_MOI.md`. The invitation letter master's static contact line is replaced with `{{DAU_MOI_LIEN_HE}}`.

## 6. Bug Fixes Bundled In

*Các lỗi được vá kèm*

- **`committee_writer.write_committee_roster`**: `org_col` becomes `Optional[int] = 2` (matching the existing `role_col: Optional[int]` pattern) — org is only written when `org_col is not None`. Required because `Phiếu ký nhận tiền.docx`'s table has no "Đơn vị" column (`TT | Họ và tên | Số tiền | Ký nhận`); writing org there today would silently corrupt the "Số tiền" column if the roster call were added naively.
- **`section_nghiem_thu._phieu_ky_nhan_tien`**: now also calls `write_committee_roster(..., name_col=2, org_col=None, role_col=None, start_row=2)` for the 5 committee rows before writing the secretary's row 7 (unchanged).
- **Contact line**: fixed by §5's new token, not by bespoke code.

## 7. Name-Based CV Matching (replaces "TÊN FILE CV" column)

*Tự động khớp file CV theo tên (thay cột "TÊN FILE CV")*

**New module `cv_matching.py`, two-tier matching:**

```python
def find_cv_file(cv_dir: Path, person_name: str, context: str = "") -> Path:
    """Tim file trong cv_dir co ten (khong ke phan mo rong) chua cum
    `person_name` LIEN NHAU, dung thu tu. Uu tien khop DUNG DAU truoc (chi
    chuan hoa hoa/thuong + khoang trang, giu nguyen dau tieng Viet) - chi
    khi khong file nao khop dung dau moi thu lai sau khi bo dau (de van
    khop duoc file dat ten khong dau nhu "TM-Gs. Nguyen Cong Khan.pdf").
    Nem FileNotFoundError neu 0 hoac >1 file khop o vong duoc dung."""
```

Rationale: stripping diacritics unconditionally would let two genuinely different Vietnamese names collide (e.g. "Đặng Thị Bình" and "Đăng Thị Bình" both reduce to `"dang thi binh"`). Trying an exact-diacritics pass first resolves such cases correctly whenever the CV filename was typed with correct diacritics; the diacritics-stripped pass only kicks in when the exact pass finds **zero** matches (i.e. the filename likely has no Vietnamese diacritics at all, as with the `TM-*.pdf` files). If the exact pass finds more than one match, that is a real ambiguity and is reported immediately without falling back — falling back would not resolve a genuine duplicate anyway.

Example: `"Nguyễn Công Khẩn"` finds zero exact-diacritics matches against `"TM-Gs. Nguyen Cong Khan.pdf"`, falls back to the diacritics-stripped pass, and matches there. Zero matches (at the final tier used) or more than one match raises a clear, actionable error (same tone as existing `copy_head_cv`/`copy_expert_cvs` errors), naming the declared person and, on ambiguity, listing every matching filename.

**Applied uniformly to both CV attachment points:**
- Head's CV (previously `F01`'s filename column, previously a required field): now resolved via `cv_matching.find_cv_file(cv_dir, info.head.name, " (chủ nhiệm đề tài)")`, using the already-existing `info.head.name` — the F01 row's filename column stops being read entirely.
- F02–F10 declared experts: `ExpertCvEntry` drops its `filename` field (keeps `code`, `name`, `role`); `excel_reader.read_expert_cvs` now includes a row when its **name** (not filename) is non-blank; `tao_ho_so_moi.copy_expert_cvs` resolves each entry's file via `cv_matching.find_cv_file` before copying (fail-fast: resolve all entries first, copy only if every entry resolved).

**Checklist/spreadsheet changes** (new one-off migration script `migrate_remove_cv_filename_column.py`, run once):
- Removes the CV-filename `DataValidation` (list dropdown, `E68:E77`-style ranges) from both project sheets.
- Rewrites each F01–F10 row's auto-status formula (column F) to drop the `ISBLANK(E<row>)` half of its `OR(...)` condition — today it requires both role (D) and filename (E) non-blank to show "✅"; after this change, filename is no longer part of the workflow, so only the name (C) presence matters. (F01 keeps its distinct "CHƯA ĐIỀN THÔNG TIN CNĐT" wording for a blank name; other rows keep their "Tùy chọn (Trống)" wording.)
- Clears/removes the `_Lists` hidden sheet (was populated solely for this dropdown).
- Column E's existing typed-in filename values in the live VIAM 2027 sample sheet are left untouched (harmless dead data, not read by any code path) — no destructive edit needed there.

**Retired:** `capnhat_danh_sach_cv.py`, `capnhat_danh_sach_cv.bat`, and their `HUONG_DAN.md` §3.4/§6.1 references — the dropdown they maintained no longer exists.

**Known behavior change:** the live checklist's **F03** row (Lưu Liên Hương, Thư ký Đề tài) currently has a name but a *blank* filename cell, so today it is silently skipped. After this change, a non-blank name is enough to require a resolvable CV — since `CV chuyên gia/` currently has no file matching "Lưu Liên Hương", running the tool against the unmodified sample data will raise a clear `FileNotFoundError` for F03. This must be resolved (by adding her CV file, or clearing the F03 name) before the sample project can regenerate cleanly; this is intentional (fail loud instead of silently dropping a declared attachment) and is called out explicitly here so it isn't mistaken for a regression during testing.

## 8. New/Modified Files Summary

*Tổng hợp file mới/sửa*

| File | Change |
|---|---|
| `expert_invitation.py` | **New.** `generate_multi_page_letter`. |
| `cv_matching.py` | **New.** `find_cv_file`. |
| `section_moi_chuyen_gia.py` | Rewritten: filtering/dedup + calls into `expert_invitation`; drops direct `Session`/token-fill use. |
| `03. Công văn mời chuyên gia - MẪU/Công văn mời chuyên gia.docx` | Edit: dotted blanks → `{{CHUYEN_GIA_HO_TEN}}`/`{{CHUYEN_GIA_DON_VI}}`; contact line → `{{DAU_MOI_LIEN_HE}}`. |
| `03. Công văn mời chuyên gia - MẪU/Công văn mời chuyên gia nghiệm thu.docx` | **New template file**, modeled on the existing letter, nghiệm thu wording. |
| `committee_writer.py` | `write_committee_roster`: `org_col` becomes `Optional[int]`. |
| `section_nghiem_thu.py` | `_phieu_ky_nhan_tien`: wire the 5 missing roster rows. |
| `excel_reader.py` | Drop `head_cv_filename` field/read; `ExpertCvEntry` drops `filename`; `read_expert_cvs` keys off name not filename. |
| `tao_ho_so_moi.py` | `copy_head_cv`/`copy_expert_cvs`: resolve via `cv_matching.find_cv_file` instead of reading a filename field. |
| `migrate_add_contact_person.py` | **New one-off migration**, inserts A08 into both sheets (pattern of `migrate_add_research_location.py`). |
| `migrate_add_tokens_sheet.py` | Add `DAU_MOI_LIEN_HE` entry to `TOKEN_SPECS`. |
| `migrate_remove_cv_filename_column.py` | **New one-off migration**: strip E-column CV validations, rewrite F0x status formulas, clear `_Lists`. |
| `capnhat_danh_sach_cv.py`, `capnhat_danh_sach_cv.bat` | **Deleted.** |
| `HUONG_DAN.md`, `HUONG_DAN_LAM_MAU_MOI.md` | Updated: remove CV-filename-dropdown instructions, document A08/new tokens, document name-based CV matching. |

## 9. Testing Strategy

*Chiến lược kiểm thử*

- `expert_invitation.py`: page count matches qualifying-recipient count; each page's tokens match only that recipient (no cross-page bleed); empty-recipients returns `False` without modifying the file; common tokens fill on every page.
- `cv_matching.py`: exact-diacritics match wins when available (including the case where a diacritics-stripped match would have been ambiguous); diacritics-stripped fallback succeeds when no exact match exists; zero-match and ambiguous-match (at whichever tier is used) both raise `FileNotFoundError` with the person's name in the message.
- `section_moi_chuyen_gia.py`: filtering excludes host-org members and secretaries; dedup collapses a person present on both ethics and proposal committees; nghiệm thu letter only pulls from `acceptance_committee`.
- `committee_writer.py`: new `org_col=None` case does not touch the org column (regression test against the "Số tiền" corruption risk described in §6).
- `excel_reader.py`: `ExpertCvEntry`/`read_expert_cvs` keyed on name; rewritten fixtures in `test_tao_ho_so_moi.py`/`test_excel_reader.py` drop `filename`.
- `migrate_add_contact_person.py`/`migrate_remove_cv_filename_column.py`: idempotency (safe to rerun), correct row insertion, correct formula rewrite — following the exact fixture-based test pattern already used by `test_migrate_add_research_location.py`.
- Full run: `python tao_ho_so_moi.py` against the VIAM 2027 sample sheet (after fixing the F03 CV gap per §7), confirm both invitation letters render with the right number of correctly-addressed pages, `Phiếu ký nhận tiền.docx` shows all 6 real names, and both migrated `.docx` templates render clean (no fragmented `{{...}}`).

## 10. Rollout & Risks

*Triển khai & rủi ro*

- **Manual verification gate**: after editing the two invitation-letter template masters and creating the new nghiệm thu one, open all three in Word and visually confirm tokens are intact, unfragmented text (same run-splitting risk as any prior template edit, per `HUONG_DAN.md`).
- **Sample data must be fixed before a clean full-pipeline test** (§7's F03 gap) — this is expected, not a bug to chase.
- **Backward compatibility**: old checklist copies still containing values in the (now-unused) CV-filename column are unaffected — the column is simply never read again, not physically deleted from data cells.
