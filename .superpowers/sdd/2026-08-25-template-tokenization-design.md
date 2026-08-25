# Template Tokenization & Document Generation Improvements — Design Spec

*(Token hóa file mẫu & cải tiến quy trình sinh hồ sơ — Đặc tả thiết kế)*

- **Date / Ngày:** 2026-08-25
- **Status / Trạng thái:** Draft, pending user review — *Bản nháp, chờ người dùng duyệt*
- **Classification / Phân loại:** Architectural (upgraded mid-brainstorm from a "hardcode audit" spike) — *Kiến trúc (được nâng cấp giữa chừng từ một spike audit hardcode)*

## 1. Background & Problem

*Bối cảnh & vấn đề*

The dossier generator (`tao_ho_so_moi.py`) fills ~20 Word template files (in four `- MẪU` folders) by running `Session.replace_text(doc, "<old literal sentence>", "<new value>")` calls hardcoded in `section_*.py`. The "old literal sentence" is verbatim text from the original sample project (KUN DOCTOR COLOSTRUM / Bánh ăn dặm VIAM). This search-and-replace-by-old-text approach was audited and found to have two classes of problems:

1. **Silent breakage risk.** `word_writer.Session.replace_text` does not raise when the search text isn't found — it only prints a `[CANH BAO]` warning and continues (`word_writer.py:129-131`). Template files (`- MẪU`) are living documents that do get edited over time, so any wording change silently desyncs the corresponding hardcoded search string, and the generated dossier ships with stale sample text, easy to miss among other console output.
2. **Confirmed live bugs from the audit.** By inspecting the actual `.docx` content (not just the code), we found the sample project's head name ("Ts. Bs. Trương Hồng Sơn") and/or host organization ("Viện Y học ứng dụng Việt Nam") appear as literal, never-replaced text in **11 of the 20 template files**, plus the project secretary name is never wired in the invitation letter, plus a "địa điểm triển khai nghiên cứu" line is replaced with a static "……" placeholder instead of real Excel data. These didn't surface before because the sample data in the repo happens to coincidentally match.

*Bộ sinh hồ sơ hiện dùng cơ chế tìm-thay dựa trên nguyên văn câu chữ của dự án mẫu cũ. Cơ chế này (a) thất bại âm thầm (chỉ in cảnh báo, không dừng chương trình) khi ai đó sửa file mẫu, và (b) qua audit thực tế đã lộ ra ít nhất 11 file mẫu đang có tên chủ nhiệm/đơn vị chủ trì của dự án mẫu cũ bị hardcode và chưa từng được thay thế — một bug thật đang tồn tại, chỉ không lộ ra vì dữ liệu mẫu trong repo tình cờ trùng khớp.*

Separately, the audit surfaced two related but independent asks that this spec bundles in because they touch the same files and the same "how do templates/checklist get authored" workflow:

- Committee-member CVs (for invited experts) have no attachment mechanism at all — only the head's CV is handled (`copy_head_cv`, via checklist code F01). Five PDF CVs for the sample project's invited experts currently sit orphaned in `Tài liệu tham khảo (không dùng tạo hồ sơ)/03. Công văn mời chuyên gia/`, a folder that is explicitly excluded from generation.
- Checklist fields that must currently be typed by hand (research type, CV filenames) are error-prone (HUONG_DAN.md §7.2b/7.2d document exactly this class of typo bug) and should become dropdowns.

## 2. Goals

*Mục tiêu*

1. Replace literal-old-text search keys with explicit `{{TOKEN}}` placeholders embedded in the template `.docx` masters, filled generically from a shared token map. *Thay tìm-thay bằng câu chữ cũ bằng token tường minh, điền tự động từ một bảng token dùng chung.*
2. While migrating, wire up the ~15 confirmed-broken spots (head name, host org, project secretary, research location) using real `ProjectInfo` data instead of leftover sample text or static dots. *Vá luôn các chỗ hardcode/bug đã phát hiện trong lúc migrate.*
3. Finish wiring up the checklist's existing PHẦN F (F02-F10) so every declared expert CV — not just the head's (F01) — actually gets attached to the output. *Hoàn thiện việc nối dây PHẦN F sẵn có trong checklist (F02-F10), để mọi CV chuyên gia đã khai đều được đính kèm vào hồ sơ đầu ra, không chỉ CV chủ nhiệm (F01).*
4. Make CV-filename and research-type checklist fields dropdowns instead of free-typed text; CV-filename dropdowns must reflect the actual contents of `CV chuyên gia/` at the time the user last refreshed them. *Biến các trường CV filename và kiểu nghiên cứu thành dropdown; dropdown CV phải phản ánh đúng nội dung folder `CV chuyên gia/`.*
5. Clean up the reference-document folder structure (move misplaced CV PDFs, delete redundant superseded `.doc` sources) and document the new template-authoring workflow so it's followed correctly next time. *Dọn dẹp cấu trúc thư mục tham khảo, viết hướng dẫn quy trình làm mẫu mới.*

### Non-Goals

*Ngoài phạm vi*

- Wiring `partner_org` (đơn vị đối tác) into any template — no template currently mentions it; the token is reserved but not used by any file in this pass.
- Any macro/VBA-based live Excel dropdown refresh (`.xlsm`) — out of scope; see §6 for why a Python refresh script was chosen instead.
- Changing how committee roster tables are populated (`committee_writer.py`, `set_cell`-based) — these are already structured/data-driven, not text-search hardcode, and are unaffected.

## 3. Token Mechanism

*Cơ chế token*

**Syntax:** `{{UPPER_SNAKE_CASE}}`, e.g. `{{TEN_DE_TAI}}`. Chosen over `[[...]]` / `«...»` for familiarity (Jinja/Handlebars-style) and near-zero collision risk with real Vietnamese administrative text.

**Common token table** — one function, single source of truth:

```python
# New module: tokens.py (or added to excel_reader.py)
def build_common_tokens(info: ProjectInfo) -> dict[str, str]:
    start, end = parse_timeline(info.timeline)
    secretary = info.project_secretary
    return {
        "{{TEN_DE_TAI}}": info.title,
        "{{NAM}}": str(info.year),
        "{{DON_VI_CHU_TRI}}": info.host_org,
        "{{DON_VI_DOI_TAC}}": info.partner_org or "",  # reserved, unused by any template today
        "{{CHU_NHIEM_HO_TEN}}": f"{info.head.degree} {info.head.name}".strip(),
        "{{THU_KY_DE_TAI}}": f"{secretary.degree} {secretary.name}".strip() if secretary else "",
        "{{THOI_GIAN_BAT_DAU}}": start,
        "{{THOI_GIAN_KET_THUC}}": end,
        "{{DIA_DIEM_TRIEN_KHAI}}": info.research_location or "……………………………….",
    }
```

**Generic filler** — new `Session` method in `word_writer.py`:

```python
def fill_tokens(self, doc: OpenDoc, tokens: dict[str, str]) -> set[str]:
    """Apply every token in `tokens` to `doc`, silently skipping ones not present.

    Not every template contains every common token, so this always passes
    warn_if_missing=False — a mangled/orphaned token instead shows up as
    literal '{{...}}' text in the generated .docx, which is self-evidently
    wrong on visual review (unlike stale-but-plausible old sentences).
    """
    return {t for t, v in tokens.items() if self.replace_text(doc, t, v, warn_if_missing=False)}
```

**Per-section usage:** each `_ten_ham(session, dest_dir, info)` function calls `session.fill_tokens(doc, common_tokens)` as its first step (computed once in `generate()`/`generate_all()` and threaded down, replacing the current `title_old` parameter). Bespoke logic that remains:

- Committee roster/secretary tables — unchanged, still `committee_writer.write_committee_roster` / `set_cell`.
- Fields with no data source (meeting date, decision number assigned after a real meeting happens) — these stop being *runtime* replacements entirely. The "……" placeholder is baked directly into the template master, since there's nothing to substitute at generation time. This removes several `replace_text` calls outright.
- File-selection logic unique to one section (e.g. `_phieu_cham_diem_nghiem_thu`'s `SCORING_FORM_FILENAMES` lookup by `research_type`) is untouched — it's not text-replacement hardcode, it's legitimate control flow.

## 4. Template Migration Script

*Script migrate file mẫu*

One-off script `migrate_templates_to_tokens.py` (same throwaway-script convention as `migrate_add_partner_org.py` etc.), run once against the master `.docx` files directly inside the four `- MẪU` folders (not against generated output):

- For every *existing* `replace_text(doc, "<old text>", ...)` call currently in `section_*.py`, the script performs the equivalent substitution directly on the template master, swapping the old literal text for the corresponding `{{TOKEN}}`.
- For the ~15 *newly discovered* gaps (head name / host org / project secretary appearing as plain unreplaced text; see §1), the script additionally substitutes those occurrences with the matching token.
- Implementation uses `python-docx` directly (this is a one-time, offline content edit of the masters — not part of the runtime generation path, so COM is not required).
- **Known risk:** `python-docx`'s run-splitting behavior (already documented as a fallback-backend limitation in `HUONG_DAN.md`) could leave a token fragmented across runs during this edit. Mitigation: after running the script, manually open and visually spot-check all ~20 migrated files before committing (see §10).
- Output of this script is a one-time content commit to the `- MẪU` template masters, not a tool users run repeatedly.

## 5. Excel Checklist Changes

*Thay đổi checklist Excel*

**Revision note (post-spec discovery):** the live checklist already has a full **"PHẦN F: LÝ LỊCH KHOA HỌC CHUYÊN GIA"** section (rows F01-F10, `SEC_F` header at row 66 in the current workbook) that `excel_reader.py` only partially reads today (F01 only, via the ad-hoc `head_cv_filename` read). Each F-row already has the exact layout needed: column 3 = name, column 4 = role/degree label, column 5 = CV filename — confirmed directly from the live workbook, including a populated example at **F03** ("Lưu Liên Hương / Thư ký Đề tài / Lý lịch khoa học_Liên Hương-2023.docx" — the exact project-secretary CV that was never wired anywhere). This removes the need to add any new columns to the C/D/E committee tables — see the corrected §7 below.

`excel_reader.py`:

- New optional field, code mục **A07** "Địa điểm triển khai nghiên cứu" → `ProjectInfo.research_location: Optional[str]`. Blank is valid (falls back to the static "……" placeholder per §3). Row insertion follows the exact pattern already established in `migrate_add_partner_org.py` (`ws.insert_rows`, copy style from a neighboring row, set code/label/status-formula cells, then `_fix_status_formula_row_refs` to repair the `[CDE]<row>` regex references in the auto-status column) — applied to both `Đề tài - Bánh ăn dặm VIAM 2027` and `Đề tài - Mẫu trắng dự án mới` sheets, inserted directly after A06 (before the `SEC_B` row).
- New function `read_expert_cvs(ws, index) -> list[ExpertCvEntry]` where `ExpertCvEntry = dataclass(code: str, name: str, role: str, filename: str)`, reading **F02 through F10** (F01 stays as the existing dedicated `head_cv_filename` field — it's required and already has its own code path), skipping any row whose filename cell (column 5) is blank. `ProjectInfo` gains `expert_cvs: List[ExpertCvEntry]`.

`Form checklist hồ sơ dự án.xlsx` (the physical template workbook) needs matching edits:

- Insert row for A07 with the same visual convention as other optional (non-yellow) fields, per the pattern above.
- **No new columns needed** for committee CVs — PHẦN F (F01-F10) already exists and already has the right shape.
- Add Data Validation dropdowns:
  - **A02 (research type)** — static list `["TVCT_ĐGHQ", "TNLS"]`, set once, never needs refreshing.
  - **F01-F10 CV filename column (column 5 of each F-row)** — dynamic list sourced from a hidden sheet (see §6).
- This must be applied to the **`Đề tài - Mẫu trắng dự án mới`** template sheet specifically, since users duplicate that sheet per HUONG_DAN.md §3.2 — validation rules copy along with the sheet duplication.

## 6. Dynamic CV Dropdown

*Dropdown CV động theo folder*

Plain (non-macro) Excel cannot scan the filesystem when opened. Since this project uses no VBA anywhere (only Python + `openpyxl` throughout), staying consistent: a new script, `capnhat_danh_sach_cv.py`, run manually by the user (same pattern as `convert_doc_templates.py` — run it after changing the `CV chuyên gia/` folder contents, before opening Excel to fill in a checklist):

1. Scans `CV chuyên gia/` for all files.
2. Writes the filename list into a hidden sheet (e.g. `_Lists`, column A), overwriting the previous contents.
3. Ensures Data Validation on the column-5 filename cell of every PHẦN F row (F01-F10) in the template sheet has `formula1="='_Lists'!$A$1:$A$<n>"` (list type), creating the validation objects if they don't exist yet, updating the range if `n` changed.

A `.bat` wrapper (matching `setup.bat`'s convention) should be provided so non-technical users can double-click it rather than use a terminal.

**Explicitly rejected alternative:** converting the workbook to `.xlsm` with VBA to auto-refresh on open — adds macro-security friction (Office trust warnings) and breaks from the project's all-Python-scripts convention for a marginal UX gain (one manual script run vs. zero).

## 7. Expert CV Attachment (PHẦN F)

*Đính kèm CV chuyên gia (PHẦN F trong checklist)*

**Revised** after discovering PHẦN F already exists in the checklist (see §5) — this is no longer tied to the C/D/E committee `Person` records at all; it's a standalone list already designed for exactly this purpose.

- `tao_ho_so_moi.py`: add `copy_expert_cvs(root: Path, dest_root: Path, info: ProjectInfo) -> None`, called alongside the existing `copy_head_cv` (F01 keeps its current behavior unchanged, → `01. Hồ sơ đạo đức đề cương/`):
  - Iterates `info.expert_cvs` (built from F02-F10 per §5).
  - For each `ExpertCvEntry`, copies `CV chuyên gia/<entry.filename>` into the output's `03. Công văn mời chuyên gia/` folder.
  - Raises the same clear, actionable `FileNotFoundError` pattern as `copy_head_cv` on a missing file, including the F-code (e.g. F04) and declared name in the message so the user knows exactly which checklist row to fix.
  - No dedup logic needed — each F-row is a distinct declared entry; if the same person's file is listed twice under different codes, both entries copy the same file harmlessly (`write_bytes` is idempotent).

## 8. Folder Structure Cleanup

*Dọn dẹp cấu trúc thư mục*

Audited every file under the four `- MẪU` folders and `Tài liệu tham khảo (không dùng tạo hồ sơ)/`:

- **`- MẪU` folders:** all 20 files are confirmed live and used by `section_*.py` — no removals. `Tài liệu tham khảo`'s existence doesn't duplicate or replace this folder's role; the two are complementary (live templates vs. non-generation reference material), both still needed.
- **Move:** `Tài liệu tham khảo (không dùng tạo hồ sơ)/03. Công văn mời chuyên gia/TM-*.pdf` (5 files) → `CV chuyên gia/`.
- **Delete** (superseded by an already-converted `.docx` counterpart living in the matching `- MẪU` folder — confirmed redundant, user approved deletion):
  - `Tài liệu tham khảo (không dùng tạo hồ sơ)/01. Hồ sơ đạo đức đề cương/Bảng kiểm đánh giá đạo đức.doc`
  - `Tài liệu tham khảo (không dùng tạo hồ sơ)/03. Công văn mời chuyên gia/Công văn mời chuyên gia.doc`
  - `Tài liệu tham khảo (không dùng tạo hồ sơ)/04. Hồ sơ nghiệm thu/10. Biên bản họp HĐ nghiệm thu.doc`
  - `Tài liệu tham khảo (không dùng tạo hồ sơ)/04. Hồ sơ nghiệm thu/11. Biên bản kiểm phiếu nghiệm thu.doc`
  - `Tài liệu tham khảo (không dùng tạo hồ sơ)/04. Hồ sơ nghiệm thu/12. Quyết định công nhận kết quả đề tài.doc`
  - `Tài liệu tham khảo (không dùng tạo hồ sơ)/04. Hồ sơ nghiệm thu/9. Quyết định thành lập HĐ nghiệm thu.doc`
  - `Tài liệu tham khảo (không dùng tạo hồ sơ)/04. Hồ sơ nghiệm thu/Phiếu nhận xét nghiệm thu.doc`
- **Keep as-is** (genuine reference material, correctly placed): `Hồ sơ đạo đức COLOSTRUM .docx`, `Slide đạo đức.pptx`, `7.5-11.5. Quyết định Thay đổi tên đề tài.docx`, `Quy trình NC.pptx`, `Slide đề cương.pptx`, `11.5. Quyết định Thay đổi tên đề tài.docx`.

## 9. New Template-Authoring Guide

*Hướng dẫn làm mẫu mới*

New file `HUONG_DAN_LAM_MAU_MOI.md` (or a new §6 replacing the current one in `HUONG_DAN.md`) documenting, for future template authoring (including for an AI assistant picking this up cold):

- The full common-token table from §3, with the `ProjectInfo` field each maps to.
- Naming convention for any *new* token: `{{UPPER_SNAKE_CASE}}`, Vietnamese-meaningful name, no diacritics-stripping ambiguity.
- **Path A — template only needs common tokens:** drop the `.docx` into the right `- MẪU` folder with the exact output filename (existing §6.1 naming convention unchanged), embed the needed `{{TOKENS}}` directly in the Word text. *Zero code required* — the generic `fill_tokens` pass picks it up automatically the next time `discover_copies` finds it.
- **Path B — template needs bespoke data** (a committee table, conditional file selection, a field with no common-token equivalent): still requires writing a `_ten_ham(session, dest_dir, info)` function, but it only needs to handle the bespoke part — call `session.fill_tokens(doc, common_tokens)` first, then the bespoke logic.
- Note on fields that can never be known at generation time (meeting date, decision number): write the static "……" placeholder directly into the template text; no token, no code.

## 10. Testing Strategy

*Chiến lược kiểm thử*

- Rewrite existing `test_section_*.py` assertions that currently check old-literal-text replacement to instead check token replacement.
- New unit tests: `build_common_tokens()` field mapping correctness (including `None` secretary / blank `research_location` fallback); `Session.fill_tokens()` (fills present tokens, silently skips absent ones, no warning printed); `copy_expert_cvs()` (found/missing-file error, cross-committee dedup); `excel_reader` for A07 optional blank handling and per-person `cv_filename` optional blank handling.
- After running `migrate_templates_to_tokens.py`: manual visual spot-check of all ~20 migrated `- MẪU` files (open each, confirm tokens render as clean, unfragmented `{{...}}` text) before committing, per the run-splitting risk noted in §4.
- Full end-to-end run: `python tao_ho_so_moi.py` against a test checklist sheet, confirm every previously-broken field (head name, host org, project secretary, research location) now renders correctly across all affected files, and confirm committee-member CVs land in `03. Công văn mời chuyên gia/`.

## 11. Rollout & Risks

*Triển khai & rủi ro*

- **Backward compatibility:** existing checklist sheets created before this change lack the new A07/CV-filename columns. `excel_reader.py` must treat them as optional/blank-safe so old project sheets don't break when re-run.
- **Scope size:** this single spec touches ~20 template masters, 6 Python modules (`word_writer.py`, `tokens.py` (new), `excel_reader.py`, `tao_ho_so_moi.py`, all four `section_*.py`), the Excel checklist template, one new migration script, one new dropdown-refresh script, and one new guide doc. The user confirmed doing this as one full pass rather than a piecemeal rollout (§3 scope decision) — the implementation plan (via `writing-plans`) should still sequence it into clearly separated, independently-testable steps internally, even though delivery is a single pass.
- **Manual verification gate:** because of the run-splitting risk (§4/§10), this change should not be considered done until the manual spot-check of all migrated template masters has been performed.
