# Template Tokenization & Document Generation Overhaul Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace literal-old-text search-and-replace in the dossier generator with explicit `{{TOKEN}}` placeholders baked into the template masters, fixing ~15 confirmed hardcode bugs in the process, add a full "attach any declared expert CV" mechanism using the checklist's already-existing but under-used PHẦN F section, add Excel dropdowns for CV filenames and research type, and clean up the reference-document folder structure.

**Architecture:** A single shared `build_common_tokens(info)` function produces a `{token: value}` dict from `ProjectInfo`; a new `Session.fill_tokens()` method applies it generically to any opened `.docx`. A one-off script edits the 20 template masters in place to swap literal old sample text for `{{TOKEN}}` placeholders (reusing the exact old search strings already in the current codebase, plus newly-audited gaps). `section_*.py` functions call `fill_tokens` first, then keep only genuinely bespoke `replace_text`/`set_cell` calls (committee tables, per-document unique fields). Checklist reading gains a `research_location` field and a `read_expert_cvs()` reader for the pre-existing but unused F02-F10 rows; `tao_ho_so_moi.py` gains `copy_expert_cvs()` to attach every declared file.

**Tech Stack:** Python 3.9+, `openpyxl` (Excel), `python-docx` (fallback Word backend + one-off template edits), `pywin32`/Word COM (primary Word backend), `pytest`.

**Spec:** `.superpowers/sdd/2026-08-25-template-tokenization-design.md` — this plan implements that spec section-by-section; read both together.

## Global Constraints

- Token syntax is exactly `{{UPPER_SNAKE_CASE}}` — no other bracket style. (Spec §3)
- `Session.fill_tokens()` always calls `replace_text(..., warn_if_missing=False)` — a missing common token in a given template is expected/normal, not an error. (Spec §3)
- No VBA/macros anywhere in this codebase — the CV dropdown refresh is a manually-run Python script, matching every other `migrate_*.py`/`convert_doc_templates.py` script in the repo. (Spec §6)
- `research_location` and every `expert_cvs` entry are optional — blank checklist cells must not raise. Existing checklist sheets created before this change lack the new fields entirely and must still load without error. (Spec §5, §11)
- F01 keeps its current, unchanged behavior (`head_cv_filename`, required, copied to `01. Hồ sơ đạo đức đề cương/`) — only F02-F10 are new. (Spec §7, revision note)
- Every new/changed Python file must have passing tests before its task is considered done — this repo has 100% file-level test coverage on its Python modules (`test_*.py` per module) and that convention continues here.

---

## File Structure

**New files:**
- `tokens.py` — `build_common_tokens(info: ProjectInfo) -> dict[str, str]`. Single source of truth for the common token vocabulary.
- `test_tokens.py`
- `migrate_add_research_location.py` — one-off script, adds checklist code A07. Same pattern as `migrate_add_partner_org.py`.
- `test_migrate_add_research_location.py`
- `migrate_templates_to_tokens.py` — one-off script, edits the 20 `- MẪU` masters in place.
- `test_migrate_templates_to_tokens.py`
- `capnhat_danh_sach_cv.py` — refreshes the hidden `_Lists` sheet + Data Validation dropdowns (CV filenames dynamic, research type static).
- `test_capnhat_danh_sach_cv.py`
- `capnhat_danh_sach_cv.bat` — double-click wrapper, same pattern as `setup.bat`.
- `HUONG_DAN_LAM_MAU_MOI.md` — new template-authoring guide.

**Modified files:**
- `word_writer.py` — add `Session.fill_tokens()`.
- `test_word_writer.py` — add tests for it.
- `excel_reader.py` — add `research_location` to `ProjectInfo`; add `ExpertCvEntry` dataclass, `read_expert_cvs()`, `expert_cvs` field.
- `test_excel_reader.py` — add tests for both.
- `tao_ho_so_moi.py` — add `copy_expert_cvs()`; wire it and `tokens.build_common_tokens()` into `generate_all()`/`main()`; drop `TITLE_OLD`/`title_old` threading.
- `test_tao_ho_so_moi.py` — add/update tests.
- `section_dao_duc.py`, `section_khoa_hoc.py`, `section_moi_chuyen_gia.py`, `section_nghiem_thu.py` — swap `title_old` parameter for `common_tokens: dict[str, str]`; call `session.fill_tokens(doc, common_tokens)` first in every function; delete now-redundant `replace_text` calls; fix the `_phieu_ky_nhan_tien` secretary bug.
- `test_section_dao_duc.py`, `test_section_khoa_hoc.py`, `test_section_moi_chuyen_gia.py`, `test_section_nghiem_thu.py` — update assertions to token-era behavior.
- The 20 `.docx` masters across the four `- MẪU` folders — content edited by `migrate_templates_to_tokens.py`, then committed as data files (not hand-edited).
- `Form checklist hồ sơ dự án.xlsx` — A07 row inserted (script); F01-F10 filename cells + A02 get Data Validation (script); `Đề tài - Bánh ăn dặm VIAM 2027` sheet's F04-F08 rows populated with the 5 relocated expert CVs (manual cell edit as part of Task 6).
- `HUONG_DAN.md` — §6 replaced with a pointer to the new guide.

**Filesystem moves/deletes (Task 6, no code):**
- Move: `Tài liệu tham khảo (không dùng tạo hồ sơ)/03. Công văn mời chuyên gia/TM-*.pdf` (5 files) → `CV chuyên gia/`.
- Delete: 7 redundant `.doc` masters already superseded by `.docx` counterparts in `- MẪU` (see spec §8 for the exact list).

---

### Task 1: `Session.fill_tokens()`

**Files:**
- Modify: `word_writer.py`
- Test: `test_word_writer.py`

**Interfaces:**
- Consumes: existing `Session.replace_text(doc, find, replace, wildcards=False, warn_if_missing=True) -> bool`.
- Produces: `Session.fill_tokens(self, doc: OpenDoc, tokens: dict[str, str]) -> set[str]` — every later task that opens a doc and needs common tokens filled calls this.

- [ ] **Step 1: Write the failing tests**

```python
# test_word_writer.py — add near the other Session tests

def test_fill_tokens_replaces_every_present_token(tmp_path):
    src = _make_paragraph_fixture(tmp_path, "De tai {{TEN_DE_TAI}}, nam {{NAM}}.")

    session = word_writer.Session(force_backend="docx")
    try:
        doc = session.open(src)
        filled = session.fill_tokens(doc, {"{{TEN_DE_TAI}}": "ABC", "{{NAM}}": "2027"})
        session.save_close(doc)
    finally:
        session.quit()

    assert filled == {"{{TEN_DE_TAI}}", "{{NAM}}"}
    check = docx.Document(str(src))
    text = "\n".join(p.text for p in check.paragraphs)
    assert text == "De tai ABC, nam 2027."


def test_fill_tokens_silently_skips_absent_tokens_without_warning(tmp_path, capsys):
    src = _make_paragraph_fixture(tmp_path, "Chi co {{NAM}} o day.")

    session = word_writer.Session(force_backend="docx")
    try:
        doc = session.open(src)
        filled = session.fill_tokens(doc, {"{{NAM}}": "2027", "{{TEN_DE_TAI}}": "ABC"})
        session.save_close(doc)
    finally:
        session.quit()

    assert filled == {"{{NAM}}"}
    captured = capsys.readouterr()
    assert "CANH BAO" not in captured.out
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_word_writer.py -k fill_tokens -v`
Expected: FAIL with `AttributeError: 'Session' object has no attribute 'fill_tokens'`

- [ ] **Step 3: Implement `fill_tokens`**

```python
# word_writer.py — add as a new Session method, after replace_text_any

    def fill_tokens(self, doc: OpenDoc, tokens: dict[str, str]) -> set[str]:
        """Ap dung moi token trong `tokens` vao `doc`, bo qua lang le token nao khong co mat.

        Khong phai template nao cung chua moi common token, nen luon goi voi
        warn_if_missing=False - mot token bi go sai/xoa nham se hien nguyen van
        '{{...}}' trong file .docx sinh ra, tu no da la dau hieu ro rang sai
        khi xem lai bang mat, khong can dua vao canh bao console.
        """
        return {token for token, value in tokens.items() if self.replace_text(doc, token, value, warn_if_missing=False)}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_word_writer.py -k fill_tokens -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Run the full word_writer test suite to confirm no regressions**

Run: `pytest test_word_writer.py -v`
Expected: all PASS

- [ ] **Step 6: Commit**

```bash
git add word_writer.py test_word_writer.py
git commit -m "feat: add Session.fill_tokens for generic token-based document filling"
```

---

### Task 2: `ProjectInfo.research_location`

**Files:**
- Modify: `excel_reader.py`
- Test: `test_excel_reader.py`

**Interfaces:**
- Consumes: existing `_read_text(ws, index, code) -> str`, `_build_code_index(ws) -> dict`.
- Produces: `ProjectInfo.research_location: Optional[str]` — consumed by `tokens.build_common_tokens()` in Task 4.

- [ ] **Step 1: Write the failing tests**

```python
# test_excel_reader.py — add near test_optional_partner_org_is_none_when_blank

def test_research_location_is_none_when_field_absent(tmp_path):
    path = _build_minimal_workbook(tmp_path)
    data = excel_reader.load_project_data(path, "Test")
    assert data.research_location is None


def test_research_location_is_read_when_present(tmp_path):
    path = _build_minimal_workbook(tmp_path, overrides={"A07": ("tỉnh Thái Bình", None, None)})
    data = excel_reader.load_project_data(path, "Test")
    assert data.research_location == "tỉnh Thái Bình"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_excel_reader.py -k research_location -v`
Expected: FAIL with `TypeError: load_project_data() got an unexpected keyword argument` or `AttributeError` on `.research_location` (the field doesn't exist yet)

- [ ] **Step 3: Implement the field**

```python
# excel_reader.py — in the ProjectInfo dataclass, add after `partner_org: Optional[str]`
    research_location: Optional[str]
```

```python
# excel_reader.py — in load_project_data(), add near the partner_org read
    research_location = _read_text(ws, index, "A07") if "A07" in index else ""
```

```python
# excel_reader.py — in the ProjectInfo(...) construction, add near partner_org=
        research_location=research_location or None,
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_excel_reader.py -k research_location -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Run the full excel_reader test suite**

Run: `pytest test_excel_reader.py -v`
Expected: all PASS

- [ ] **Step 6: Commit**

```bash
git add excel_reader.py test_excel_reader.py
git commit -m "feat: read optional research_location field (A07) from checklist"
```

---

### Task 3: Add checklist field A07 to the live workbook

**Files:**
- Create: `migrate_add_research_location.py`
- Test: `test_migrate_add_research_location.py`
- Modify (by running the script): `Form checklist hồ sơ dự án.xlsx`

**Interfaces:**
- Consumes: `excel_reader.ProjectInfo.research_location` (Task 2).
- Produces: a live checklist workbook where both sheets have code `A07`, readable by `excel_reader.load_project_data`.

- [ ] **Step 1: Write the failing test**

```python
# test_migrate_add_research_location.py
from pathlib import Path

import openpyxl

import migrate_add_research_location as migrate
import excel_reader

CHECKLIST_PATH = Path(__file__).resolve().parent / excel_reader.CHECKLIST_FILENAME
SHEET_NAMES = ["Đề tài - Bánh ăn dặm VIAM 2027", "Đề tài - Mẫu trắng dự án mới"]


def test_add_research_location_field_adds_a07_to_both_sheets():
    migrate.add_research_location_field()

    wb = openpyxl.load_workbook(CHECKLIST_PATH)
    for sheet_name in SHEET_NAMES:
        ws = wb[sheet_name]
        codes = [row[0].value for row in ws.iter_rows(min_row=5, max_col=1)]
        assert "A07" in codes


def test_add_research_location_field_is_idempotent():
    migrate.add_research_location_field()
    migrate.add_research_location_field()

    wb = openpyxl.load_workbook(CHECKLIST_PATH)
    for sheet_name in SHEET_NAMES:
        ws = wb[sheet_name]
        codes = [row[0].value for row in ws.iter_rows(min_row=5, max_col=1)]
        assert codes.count("A07") == 1
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest test_migrate_add_research_location.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'migrate_add_research_location'`

- [ ] **Step 3: Implement the migration script**

Modeled directly on the existing `migrate_add_partner_org.py` (same repo, same pattern) — insert after A06 (row 11 today, i.e. `INSERT_AT_ROW = 12`, right before the `SEC_B` header), copy style from A06's row, set code/label/status-formula, then repair the `[CDE]<row>` formula references that shift below the insertion point.

```python
# migrate_add_research_location.py
import copy
import re
from pathlib import Path

import openpyxl

CHECKLIST_PATH = Path(__file__).resolve().parent / "Form checklist hồ sơ dự án.xlsx"
SHEET_NAMES = ["Đề tài - Bánh ăn dặm VIAM 2027", "Đề tài - Mẫu trắng dự án mới"]
INSERT_AT_ROW = 12
STYLE_SOURCE_ROW = 11


def _fix_status_formula_row_refs(ws) -> None:
    for row_cells in ws.iter_rows(min_row=5):
        code_cell = row_cells[0]
        if not isinstance(code_cell.value, str):
            continue
        formula_cell = row_cells[5]
        if not (isinstance(formula_cell.value, str) and formula_cell.value.startswith("=")):
            continue
        actual_row = code_cell.row
        formula_cell.value = re.sub(
            r"\b([CDE])(\d{1,3})\b",
            lambda m: f"{m.group(1)}{actual_row}",
            formula_cell.value,
        )


def _has_research_location_field(ws) -> bool:
    for row in ws.iter_rows(min_row=5, max_col=1):
        if row[0].value == "A07":
            return True
    return False


def add_research_location_field() -> None:
    wb = openpyxl.load_workbook(CHECKLIST_PATH)
    changed = False
    for sheet_name in SHEET_NAMES:
        ws = wb[sheet_name]
        if _has_research_location_field(ws):
            continue
        changed = True
        ws.insert_rows(INSERT_AT_ROW)

        for col in range(1, 7):
            src_cell = ws.cell(row=STYLE_SOURCE_ROW, column=col)
            dst_cell = ws.cell(row=INSERT_AT_ROW, column=col)
            dst_cell.font = copy.copy(src_cell.font)
            dst_cell.fill = copy.copy(src_cell.fill)
            dst_cell.border = copy.copy(src_cell.border)
            dst_cell.alignment = copy.copy(src_cell.alignment)

        ws.cell(row=INSERT_AT_ROW, column=1, value="A07")
        ws.cell(row=INSERT_AT_ROW, column=2, value="Địa điểm triển khai nghiên cứu (Tùy chọn)")
        ws.cell(
            row=INSERT_AT_ROW,
            column=6,
            value=f'=IF(ISBLANK(C{INSERT_AT_ROW}), "⚪ Tùy chọn (Trống)", "✅ Xong")',
        )

        _fix_status_formula_row_refs(ws)

    if changed:
        wb.save(CHECKLIST_PATH)


if __name__ == "__main__":
    add_research_location_field()
    print("Da them truong A07 'Dia diem trien khai nghien cuu' vao ca 2 sheet.")
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest test_migrate_add_research_location.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Run the script against the real checklist and confirm with the full excel_reader suite**

Run: `python migrate_add_research_location.py`
Expected output: `Da them truong A07 'Dia diem trien khai nghien cuu' vao ca 2 sheet.`

Then run: `pytest test_excel_reader.py -v`
Expected: all PASS (including the two new Task 2 tests, now exercised against the real, updated workbook too if you add a quick manual check — not required as an automated test since `_build_minimal_workbook` already covers the synthetic case).

- [ ] **Step 6: Commit**

```bash
git add migrate_add_research_location.py test_migrate_add_research_location.py "Form checklist hồ sơ dự án.xlsx"
git commit -m "feat: add optional checklist field A07 for research location"
```

---

### Task 4: `tokens.build_common_tokens()`

**Files:**
- Create: `tokens.py`
- Test: `test_tokens.py`

**Interfaces:**
- Consumes: `excel_reader.ProjectInfo` (all fields), `excel_reader.parse_timeline(text) -> tuple[str, str]`.
- Produces: `build_common_tokens(info: ProjectInfo) -> dict[str, str]` — consumed by `tao_ho_so_moi.py` and every `section_*.py` function in Task 9.

- [ ] **Step 1: Write the failing tests**

```python
# test_tokens.py
import dataclasses

import excel_reader
import tokens


def _make_info(**overrides):
    base = excel_reader.ProjectInfo(
        title="Đề tài test",
        research_type="TVCT_ĐGHQ",
        year=2027,
        host_org="Viện ABC",
        partner_org=None,
        research_location=None,
        timeline="Tháng 01/2027 đến tháng 12/2027",
        head=excel_reader.Person(name="Nguyễn Văn A", degree="TS.", org="Viện ABC"),
        co_head=None,
        project_secretary=None,
        researchers=[],
        ethics_committee=excel_reader.CommitteeData(chair=excel_reader.Person(name="Chủ tịch")),
        proposal_committee=excel_reader.CommitteeData(chair=excel_reader.Person(name="Chủ tịch")),
        acceptance_committee=excel_reader.CommitteeData(chair=excel_reader.Person(name="Chủ tịch")),
        head_cv_filename="cv.docx",
        expert_cvs=[],
    )
    return dataclasses.replace(base, **overrides)


def test_build_common_tokens_maps_scalar_fields():
    info = _make_info()
    t = tokens.build_common_tokens(info)
    assert t["{{TEN_DE_TAI}}"] == "Đề tài test"
    assert t["{{NAM}}"] == "2027"
    assert t["{{DON_VI_CHU_TRI}}"] == "Viện ABC"
    assert t["{{CHU_NHIEM_HO_TEN}}"] == "TS. Nguyễn Văn A"
    assert t["{{CHU_NHIEM_TEN}}"] == "Nguyễn Văn A"
    assert t["{{THOI_GIAN_BAT_DAU}}"] == "01/2027"
    assert t["{{THOI_GIAN_KET_THUC}}"] == "12/2027"


def test_build_common_tokens_blank_research_location_falls_back_to_dots():
    info = _make_info(research_location=None)
    t = tokens.build_common_tokens(info)
    assert t["{{DIA_DIEM_TRIEN_KHAI}}"] == "……………………………."


def test_build_common_tokens_uses_real_research_location_when_present():
    info = _make_info(research_location="tỉnh Thái Bình")
    t = tokens.build_common_tokens(info)
    assert t["{{DIA_DIEM_TRIEN_KHAI}}"] == "tỉnh Thái Bình"


def test_build_common_tokens_blank_secretary_and_co_head_are_empty_strings():
    info = _make_info(co_head=None, project_secretary=None)
    t = tokens.build_common_tokens(info)
    assert t["{{THU_KY_DE_TAI}}"] == ""
    assert t["{{DONG_CHU_NHIEM_TEN}}"] == ""


def test_build_common_tokens_fills_secretary_and_co_head_when_present():
    info = _make_info(
        co_head=excel_reader.Person(name="Đồng chủ nhiệm B"),
        project_secretary=excel_reader.Person(name="Thư ký C", degree="ThS."),
    )
    t = tokens.build_common_tokens(info)
    assert t["{{DONG_CHU_NHIEM_TEN}}"] == "Đồng chủ nhiệm B"
    assert t["{{THU_KY_DE_TAI}}"] == "ThS. Thư ký C"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_tokens.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'tokens'` (or a `TypeError` on `ProjectInfo(...)` about missing `expert_cvs`/`research_location` if Task 2/5 aren't done first — Task 4 must run after Tasks 2 and 5)

- [ ] **Step 3: Implement `build_common_tokens`**

```python
# tokens.py
from excel_reader import ProjectInfo, parse_timeline

DIA_DIEM_PLACEHOLDER = "……………………………."


def build_common_tokens(info: ProjectInfo) -> dict:
    start, end = parse_timeline(info.timeline)
    secretary = info.project_secretary
    return {
        "{{TEN_DE_TAI}}": info.title,
        "{{NAM}}": str(info.year),
        "{{DON_VI_CHU_TRI}}": info.host_org,
        "{{DON_VI_DOI_TAC}}": info.partner_org or "",
        "{{CHU_NHIEM_HO_TEN}}": f"{info.head.degree} {info.head.name}".strip(),
        "{{CHU_NHIEM_TEN}}": info.head.name,
        "{{DONG_CHU_NHIEM_TEN}}": info.co_head.name if info.co_head else "",
        "{{THU_KY_DE_TAI}}": f"{secretary.degree} {secretary.name}".strip() if secretary else "",
        "{{THOI_GIAN_BAT_DAU}}": start,
        "{{THOI_GIAN_KET_THUC}}": end,
        "{{DIA_DIEM_TRIEN_KHAI}}": info.research_location or DIA_DIEM_PLACEHOLDER,
    }
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_tokens.py -v`
Expected: PASS (6 passed)

- [ ] **Step 5: Commit**

```bash
git add tokens.py test_tokens.py
git commit -m "feat: add build_common_tokens as the single source of truth for template tokens"
```

---

### Task 5: `ExpertCvEntry` + `read_expert_cvs()` (PHẦN F02-F10)

**Files:**
- Modify: `excel_reader.py`
- Test: `test_excel_reader.py`

**Interfaces:**
- Consumes: `_cell_text(ws, row, col) -> str`, `_build_code_index(ws) -> dict`.
- Produces: `ExpertCvEntry(code: str, name: str, role: str, filename: str)` dataclass; `read_expert_cvs(ws, index) -> List[ExpertCvEntry]`; `ProjectInfo.expert_cvs: List[ExpertCvEntry]` — consumed by `tao_ho_so_moi.copy_expert_cvs()` in Task 7.

- [ ] **Step 1: Write the failing tests**

```python
# test_excel_reader.py — add near the F01 tests

def test_read_expert_cvs_skips_blank_filename_rows(tmp_path):
    path = _build_minimal_workbook(tmp_path)
    data = excel_reader.load_project_data(path, "Test")
    assert data.expert_cvs == []


def test_read_expert_cvs_reads_declared_rows(tmp_path):
    path = _build_minimal_workbook(
        tmp_path,
        overrides={
            "F03": ("Thư ký C", "Thư ký Đề tài", "cv_thuky.docx"),
            "F04": ("Chuyên gia D", "Ủy viên", "cv_d.docx"),
        },
    )
    data = excel_reader.load_project_data(path, "Test")
    codes = {e.code: e for e in data.expert_cvs}
    assert set(codes) == {"F03", "F04"}
    assert codes["F03"].name == "Thư ký C"
    assert codes["F03"].role == "Thư ký Đề tài"
    assert codes["F03"].filename == "cv_thuky.docx"


def test_real_checklist_expert_cvs_has_at_least_one_entry():
    data = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    assert len(data.expert_cvs) >= 1
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_excel_reader.py -k expert_cvs -v`
Expected: FAIL with `AttributeError: 'ProjectInfo' object has no attribute 'expert_cvs'`

- [ ] **Step 3: Implement `ExpertCvEntry` and `read_expert_cvs`**

```python
# excel_reader.py — add near the Person/CommitteeData dataclasses

@dataclass
class ExpertCvEntry:
    code: str
    name: str
    role: str
    filename: str
```

```python
# excel_reader.py — add as a module-level function, near parse_committee

def read_expert_cvs(ws, index: dict) -> List[ExpertCvEntry]:
    """Doc PHAN F, ma F02-F10 (F01 la CV chu nhiem, da co duong doc rieng bat buoc)."""
    entries = []
    for i in range(2, 11):
        code = f"F{i:02d}"
        row = index.get(code)
        if row is None:
            continue
        filename = _cell_text(ws, row, 5)
        if not filename:
            continue
        entries.append(
            ExpertCvEntry(
                code=code,
                name=_cell_text(ws, row, 3),
                role=_cell_text(ws, row, 4),
                filename=filename,
            )
        )
    return entries
```

```python
# excel_reader.py — in ProjectInfo dataclass, add after head_cv_filename: str
    expert_cvs: List[ExpertCvEntry]
```

```python
# excel_reader.py — in load_project_data(), in the ProjectInfo(...) construction, add after head_cv_filename=head_cv_filename,
        expert_cvs=read_expert_cvs(ws, index),
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_excel_reader.py -k expert_cvs -v`
Expected: PASS (3 passed) — the real-checklist test only passes once Task 6 has populated F04-F08 in the live workbook; if run before Task 6, temporarily skip it with `@pytest.mark.skip(reason="populated in Task 6")` and remove the skip once Task 6 lands. Prefer doing Task 6 first if strict green-at-every-step matters more than task ordering.

- [ ] **Step 5: Run the full excel_reader suite**

Run: `pytest test_excel_reader.py -v`
Expected: all PASS

- [ ] **Step 6: Commit**

```bash
git add excel_reader.py test_excel_reader.py
git commit -m "feat: read expert CV entries F02-F10 from checklist PHAN F"
```

---

### Task 6: Folder cleanup + populate real checklist F04-F08

**Files:**
- Move: 5 files from `Tài liệu tham khảo (không dùng tạo hồ sơ)/03. Công văn mời chuyên gia/` to `CV chuyên gia/`
- Delete: 7 redundant `.doc` files (list below)
- Modify: `Form checklist hồ sơ dự án.xlsx` (fill F04-F08 in the `Đề tài - Bánh ăn dặm VIAM 2027` sheet)
- Test: `test_tao_ho_so_moi.py` (one new integration-style assertion)

**Interfaces:**
- Consumes: nothing new.
- Produces: real checklist data usable by `test_real_checklist_expert_cvs_has_at_least_one_entry` (Task 5) and by `copy_expert_cvs` tests (Task 7).

- [ ] **Step 1: Move the 5 expert CV files**

```bash
git mv "Tài liệu tham khảo (không dùng tạo hồ sơ)/03. Công văn mời chuyên gia/TM-Gs. Nguyen Cong Khan.pdf" "CV chuyên gia/TM-Gs. Nguyen Cong Khan.pdf"
git mv "Tài liệu tham khảo (không dùng tạo hồ sơ)/03. Công văn mời chuyên gia/TM-PGs. Hoang Thi Thanh.pdf" "CV chuyên gia/TM-PGs. Hoang Thi Thanh.pdf"
git mv "Tài liệu tham khảo (không dùng tạo hồ sơ)/03. Công văn mời chuyên gia/TM-PGs. Nguyen Quang Dung.pdf" "CV chuyên gia/TM-PGs. Nguyen Quang Dung.pdf"
git mv "Tài liệu tham khảo (không dùng tạo hồ sơ)/03. Công văn mời chuyên gia/TM-PGs. Tran Quang Trung.pdf" "CV chuyên gia/TM-PGs. Tran Quang Trung.pdf"
git mv "Tài liệu tham khảo (không dùng tạo hồ sơ)/03. Công văn mời chuyên gia/TM-Ts. Nguyen Hung Long.pdf" "CV chuyên gia/TM-Ts. Nguyen Hung Long.pdf"
```

- [ ] **Step 2: Delete the 7 redundant `.doc` masters**

```bash
git rm "Tài liệu tham khảo (không dùng tạo hồ sơ)/01. Hồ sơ đạo đức đề cương/Bảng kiểm đánh giá đạo đức.doc"
git rm "Tài liệu tham khảo (không dùng tạo hồ sơ)/03. Công văn mời chuyên gia/Công văn mời chuyên gia.doc"
git rm "Tài liệu tham khảo (không dùng tạo hồ sơ)/04. Hồ sơ nghiệm thu/10. Biên bản họp HĐ nghiệm thu.doc"
git rm "Tài liệu tham khảo (không dùng tạo hồ sơ)/04. Hồ sơ nghiệm thu/11. Biên bản kiểm phiếu nghiệm thu.doc"
git rm "Tài liệu tham khảo (không dùng tạo hồ sơ)/04. Hồ sơ nghiệm thu/12. Quyết định công nhận kết quả đề tài.doc"
git rm "Tài liệu tham khảo (không dùng tạo hồ sơ)/04. Hồ sơ nghiệm thu/9. Quyết định thành lập HĐ nghiệm thu.doc"
git rm "Tài liệu tham khảo (không dùng tạo hồ sơ)/04. Hồ sơ nghiệm thu/Phiếu nhận xét nghiệm thu.doc"
```

- [ ] **Step 3: Populate F04-F08 in the real `Đề tài - Bánh ăn dặm VIAM 2027` sheet**

Run this one-off snippet directly (not saved as a script — a manual, reviewed data entry, same spirit as filling in a checklist by hand):

```python
import openpyxl

wb = openpyxl.load_workbook("Form checklist hồ sơ dự án.xlsx")
ws = wb["Đề tài - Bánh ăn dặm VIAM 2027"]
index = {row[0].value: row[0].row for row in ws.iter_rows(min_row=5, max_col=1) if isinstance(row[0].value, str)}

rows_to_fill = {
    "F04": ("Nguyễn Công Khẩn", "Chuyên gia hội đồng", "TM-Gs. Nguyen Cong Khan.pdf"),
    "F05": ("Hoàng Thị Thanh", "Chuyên gia hội đồng", "TM-PGs. Hoang Thi Thanh.pdf"),
    "F06": ("Nguyễn Quang Dũng", "Chuyên gia hội đồng", "TM-PGs. Nguyen Quang Dung.pdf"),
    "F07": ("Trần Quang Trung", "Chuyên gia hội đồng", "TM-PGs. Tran Quang Trung.pdf"),
    "F08": ("Nguyễn Hùng Long", "Chuyên gia hội đồng", "TM-Ts. Nguyen Hung Long.pdf"),
}
for code, (name, role, filename) in rows_to_fill.items():
    row = index[code]
    ws.cell(row=row, column=3, value=name)
    ws.cell(row=row, column=4, value=role)
    ws.cell(row=row, column=5, value=filename)

wb.save("Form checklist hồ sơ dự án.xlsx")
```

- [ ] **Step 4: Verify with the existing and Task 5 tests**

Run: `pytest test_excel_reader.py -v`
Expected: all PASS, including `test_real_checklist_expert_cvs_has_at_least_one_entry` now passing for real (remove any temporary `@pytest.mark.skip` added in Task 5 Step 4).

- [ ] **Step 5: Write and run a regression test confirming the moved files don't get swept into template copying**

```python
# test_tao_ho_so_moi.py — add near test_copy_templates_does_not_copy_archived_reference_files

def test_copy_templates_does_not_copy_relocated_expert_cvs(tmp_path):
    root = paths.project_root()

    tao_ho_so_moi.copy_templates(root, tmp_path)

    assert not (tmp_path / "CV chuyên gia").exists()
```

Run: `pytest test_tao_ho_so_moi.py -k relocated_expert_cvs -v`
Expected: PASS (the `CV chuyên gia/` folder was never inside a `- MẪU` folder, so `template_config.discover_copies` never touched it — this test documents that invariant now that the folder holds more files).

- [ ] **Step 6: Commit**

```bash
git add "CV chuyên gia" "Tài liệu tham khảo (không dùng tạo hồ sơ)" "Form checklist hồ sơ dự án.xlsx" test_tao_ho_so_moi.py
git commit -m "chore: relocate expert CV PDFs into CV chuyên gia/, drop redundant .doc masters, populate sample checklist F04-F08"
```

---

### Task 7: `tao_ho_so_moi.copy_expert_cvs()`

**Files:**
- Modify: `tao_ho_so_moi.py`
- Test: `test_tao_ho_so_moi.py`

**Interfaces:**
- Consumes: `excel_reader.ProjectInfo.expert_cvs: List[ExpertCvEntry]` (Task 5), real checklist data (Task 6).
- Produces: `copy_expert_cvs(root: Path, dest_root: Path, info: ProjectInfo) -> None` — called from `generate_all()` in Task 9.

- [ ] **Step 1: Write the failing tests**

```python
# test_tao_ho_so_moi.py — add near test_copy_head_cv_*

def test_copy_expert_cvs_copies_every_declared_file(tmp_path):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)

    tao_ho_so_moi.copy_expert_cvs(root, tmp_path, info)

    dest_dir = tmp_path / "03. Công văn mời chuyên gia"
    for entry in info.expert_cvs:
        assert (dest_dir / entry.filename).exists()


def test_copy_expert_cvs_raises_clear_error_when_file_missing(tmp_path):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    bad_entry = dataclasses.replace(info.expert_cvs[0], filename="không tồn tại.pdf")
    bad_info = dataclasses.replace(info, expert_cvs=[bad_entry])

    with pytest.raises(FileNotFoundError):
        tao_ho_so_moi.copy_expert_cvs(root, tmp_path, bad_info)


def test_copy_expert_cvs_does_nothing_when_list_is_empty(tmp_path):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    empty_info = dataclasses.replace(info, expert_cvs=[])

    tao_ho_so_moi.copy_expert_cvs(root, tmp_path, empty_info)

    assert not (tmp_path / "03. Công văn mời chuyên gia").exists()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_tao_ho_so_moi.py -k copy_expert_cvs -v`
Expected: FAIL with `AttributeError: module 'tao_ho_so_moi' has no attribute 'copy_expert_cvs'`

- [ ] **Step 3: Implement `copy_expert_cvs`**

```python
# tao_ho_so_moi.py — add after copy_head_cv

def copy_expert_cvs(root: Path, dest_root: Path, info) -> None:
    for entry in info.expert_cvs:
        src = root / "CV chuyên gia" / entry.filename
        if not src.exists():
            raise FileNotFoundError(
                f"Không tìm thấy file CV '{entry.filename}' (khai báo ở mã mục {entry.code} - "
                f"{entry.name}) trong thư mục 'CV chuyên gia/'. Vui lòng đặt đúng file vào đó "
                "hoặc sửa lại tên file trong checklist cho khớp."
            )
        dst = dest_root / "03. Công văn mời chuyên gia" / entry.filename
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_bytes(src.read_bytes())
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_tao_ho_so_moi.py -k copy_expert_cvs -v`
Expected: PASS (3 passed)

- [ ] **Step 5: Run the full tao_ho_so_moi suite**

Run: `pytest test_tao_ho_so_moi.py -v`
Expected: all PASS

- [ ] **Step 6: Commit**

```bash
git add tao_ho_so_moi.py test_tao_ho_so_moi.py
git commit -m "feat: add copy_expert_cvs to attach every declared PHAN F expert CV"
```

---

### Task 8: Migrate the 20 template masters to tokens

**Files:**
- Create: `migrate_templates_to_tokens.py`
- Test: `test_migrate_templates_to_tokens.py`
- Modify (by running the script, then committed as data): all 20 `.docx` files across the four `- MẪU` folders.

**Interfaces:**
- Consumes: `word_writer.Session(force_backend="docx")`, `.open()`, `.replace_text()`, `.save_close()` (existing, unchanged).
- Produces: 20 template masters whose text now contains `{{TOKEN}}` placeholders instead of literal old sample data — consumed at runtime by `fill_tokens` once Task 9 wires it in.

This script is self-contained (doesn't import from `tao_ho_so_moi.py`) so it keeps working standing alone even after Task 9 removes `TITLE_OLD` from that module — matching every other one-off `migrate_*.py` script in this repo.

- [ ] **Step 1: Write the failing test**

Test against an isolated fixture copy, not the real templates, so the suite stays fast and repeatable — verifies the *mechanism* (apply one mapping entry, confirm search text is gone and token is present), not the full 20-file mapping table (that part is verified by the one-time manual run in Step 5).

```python
# test_migrate_templates_to_tokens.py
import shutil
from pathlib import Path

import docx

import migrate_templates_to_tokens as migrate


def test_apply_mapping_replaces_old_text_with_token(tmp_path):
    src = tmp_path / "sample.docx"
    d = docx.Document()
    d.add_paragraph("Tên đề tài: OLD_TITLE_MARKER.")
    d.save(str(src))

    migrate.apply_mapping(src, [("OLD_TITLE_MARKER", "{{TEN_DE_TAI}}")])

    check = docx.Document(str(src))
    text = "\n".join(p.text for p in check.paragraphs)
    assert text == "Tên đề tài: {{TEN_DE_TAI}}."


def test_apply_mapping_raises_when_search_text_not_found(tmp_path):
    src = tmp_path / "sample.docx"
    d = docx.Document()
    d.add_paragraph("Không liên quan.")
    d.save(str(src))

    try:
        migrate.apply_mapping(src, [("KHONG_TON_TAI", "{{X}}")])
        assert False, "Ky vong RuntimeError khi khong tim thay chuoi can thay"
    except RuntimeError as e:
        assert "KHONG_TON_TAI" in str(e)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest test_migrate_templates_to_tokens.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'migrate_templates_to_tokens'`

- [ ] **Step 3: Implement `apply_mapping` and the full MIGRATIONS table**

The old-title constant is copied verbatim from `tao_ho_so_moi.TITLE_OLD` as it exists today (self-contained per the note above — do not import it).

```python
# migrate_templates_to_tokens.py
from pathlib import Path

import word_writer

TITLE_OLD = (
    "Đánh giá hiệu quả sản phẩm sữa dinh dưỡng pha sẵn KUN DOCTOR COLOSTRUM lên "
    "tình trạng dinh dưỡng, miễn dịch, tiêu hóa và giấc ngủ của trẻ từ 24 đến 72 tháng tuổi"
)

TOKEN_TEN_DE_TAI = "{{TEN_DE_TAI}}"
TOKEN_NAM = "{{NAM}}"
TOKEN_DON_VI_CHU_TRI = "{{DON_VI_CHU_TRI}}"
TOKEN_CHU_NHIEM_HO_TEN = "{{CHU_NHIEM_HO_TEN}}"
TOKEN_CHU_NHIEM_TEN = "{{CHU_NHIEM_TEN}}"
TOKEN_DONG_CHU_NHIEM_TEN = "{{DONG_CHU_NHIEM_TEN}}"
TOKEN_THU_KY_DE_TAI = "{{THU_KY_DE_TAI}}"

DAO_DUC = "01. Hồ sơ đạo đức đề cương - MẪU"
KHOA_HOC = "02. Hồ sơ khoa học đề cương - MẪU"
MOI_CHUYEN_GIA = "03. Công văn mời chuyên gia - MẪU"
NGHIEM_THU = "04. Hồ sơ nghiệm thu - MẪU"

# Moi entry: (duong dan tuong doi tinh tu goc du an, [(text cu, text moi da co token), ...])
MIGRATIONS = {
    f"{DAO_DUC}/00. QĐ Giao đề tài.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
    ],
    f"{DAO_DUC}/01. QĐTLHĐ đạo đức đề cương.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
    ],
    f"{DAO_DUC}/02. BB họp HĐ đạo đức.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Chủ nhiệm đề tài: Ts. Bs. Trương Hồng Sơn", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN}"),
        ("Cơ quan thực hiện đề tài: Viện Y học ứng dụng Việt Nam", f"Cơ quan thực hiện đề tài: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{DAO_DUC}/03. BB kiểm phiếu HĐ đạo đức.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Chủ nhiệm: Ts. Bs. Trương Hồng Sơn", f"Chủ nhiệm: {TOKEN_CHU_NHIEM_HO_TEN}"),
    ],
    f"{DAO_DUC}/04. QĐ chấp nhận đạo đức.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Chủ nhiệm đề tài: Ts. Bs. Trương Hồng Sơn", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN}"),
        ("Cơ quan thực hiện đề tài:  Viện Y học ứng dụng Việt Nam.", f"Cơ quan thực hiện đề tài:  {TOKEN_DON_VI_CHU_TRI}."),
    ],
    f"{DAO_DUC}/Bảng kiểm đánh giá đạo đức.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("2024", TOKEN_NAM),
        ("Chủ nhiệm đề tài: Ts. Bs. Trương Hồng Sơn", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN}"),
        ("Đơn vị chủ trì: Viện Y học ứng dụng Việt Nam", f"Đơn vị chủ trì: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{KHOA_HOC}/05. QĐ TLHĐ khoa học xét đề cương.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
    ],
    f"{KHOA_HOC}/06. BB họp thông qua đề cương.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
    ],
    f"{KHOA_HOC}/07. BB kiểm phiếu thông qua đề cương.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Chủ nhiệm đề tài: Ts. Bs. Trương Hồng Sơn", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN}"),
        ("Đơn vị thực hiện: Viện Y học ứng dụng Việt nam", f"Đơn vị thực hiện: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{KHOA_HOC}/08. QĐ phê duyệt đề tài.docx": [
        ("2025", TOKEN_NAM),
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("- Chủ nhiệm đề tài: Ts. Bs. Trương Hồng Sơn", f"- Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN}"),
        ("- Đơn vị thực hiện đề tài: Viện Y học ứng dụng Việt Nam", f"- Đơn vị thực hiện đề tài: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{KHOA_HOC}/Phiếu chấm điểm HĐ đề cương.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("2. Chủ nhiệm Đề tài: Ts. Bs. Trương Hồng Sơn", f"2. Chủ nhiệm Đề tài: {TOKEN_CHU_NHIEM_HO_TEN}"),
        ("3. Đơn vị chủ trì đề tài:  Viện Y học ứng dụng Việt Nam", f"3. Đơn vị chủ trì đề tài:  {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{KHOA_HOC}/Phiếu nhận xét đánh giá hồ sơ.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("- Chủ nhiệm đề tài: Ts. Bs Trương Hồng Sơn ", f"- Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN} "),
        ("- Đơn vị chủ trì đề tài: Viện Y học ứng dụng Việt Nam", f"- Đơn vị chủ trì đề tài: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{MOI_CHUYEN_GIA}/Công văn mời chuyên gia.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("2024", TOKEN_NAM),
        ("Chủ nhiệm đề tài: Ts. Bs. Trương Hồng Sơn.", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN}."),
        ("Thư ký đề tài: Ths. Lưu Liên Hương.", f"Thư ký đề tài: {TOKEN_THU_KY_DE_TAI}."),
        ("Đơn vị thực hiện đề tài: Viện Y học ứng dụng Việt Nam.", f"Đơn vị thực hiện đề tài: {TOKEN_DON_VI_CHU_TRI}."),
    ],
    f"{NGHIEM_THU}/9. Quyết định thành lập HĐ nghiệm thu.docx": [
        ("20xx", TOKEN_NAM),
        ("“Tên đề tài”", f"“{TOKEN_TEN_DE_TAI}”"),
    ],
    f"{NGHIEM_THU}/10. Biên bản họp HĐ nghiệm thu.docx": [
        ("20xx", TOKEN_NAM),
        ("1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_TEN}"),
        ("Đồng chủ nhiệm đề tài: Tên 2", f"Đồng chủ nhiệm đề tài: {TOKEN_DONG_CHU_NHIEM_TEN}"),
        ("Đơn vị chủ trì: Viện Y học ứng dụng Việt Nam.", f"Đơn vị chủ trì: {TOKEN_DON_VI_CHU_TRI}."),
    ],
    f"{NGHIEM_THU}/11. Biên bản kiểm phiếu nghiệm thu.docx": [
        ("20xx", TOKEN_NAM),
        ("1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_TEN}"),
        ("Đồng chủ nhiệm đề tài: Tên 2", f"Đồng chủ nhiệm đề tài: {TOKEN_DONG_CHU_NHIEM_TEN}"),
        ("Đơn vị chủ trì: Viện Y học ứng dụng Việt Nam", f"Đơn vị chủ trì: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{NGHIEM_THU}/12. Quyết định công nhận kết quả đề tài.docx": [
        ("20XX", TOKEN_NAM),
        ("20xx", TOKEN_NAM),
        ("“Tên đề tài”", f"“{TOKEN_TEN_DE_TAI}”"),
    ],
    f"{NGHIEM_THU}/Phiếu chấm điểm nghiệm thu (TNLS).docx": [
        ("1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_TEN}"),
    ],
    f"{NGHIEM_THU}/Phiếu chấm điểm nghiệm thu (TVCT_ĐGHQ).docx": [
        ("1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_TEN}"),
    ],
    f"{NGHIEM_THU}/Phiếu ký nhận tiền.docx": [
        (
            "“Đánh giá hiệu quả sản phẩm thực phẩm chức năng Viên nang Đông trùng hạ thảo CordySen”",
            f"“{TOKEN_TEN_DE_TAI}”",
        ),
    ],
    f"{NGHIEM_THU}/Phiếu nhận xét nghiệm thu.docx": [
        ("20xx", TOKEN_NAM),
        ("Tên đề tài: Tên đề tài", f"Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm: Tên 1", f"Chủ nhiệm: {TOKEN_CHU_NHIEM_TEN}"),
        ("Đồng chủ nhiệm: Tên 2", f"Đồng chủ nhiệm: {TOKEN_DONG_CHU_NHIEM_TEN}"),
        ("Đơn vị chủ trì: Viện Y học ứng dụng Việt Nam", f"Đơn vị chủ trì: {TOKEN_DON_VI_CHU_TRI}"),
    ],
}


def apply_mapping(path: Path, mapping: list) -> None:
    session = word_writer.Session(force_backend="docx")
    try:
        doc = session.open(path)
        for old_text, new_text in mapping:
            found = session.replace_text(doc, old_text, new_text, warn_if_missing=False)
            if not found:
                raise RuntimeError(f"Khong tim thay chuoi can thay trong {path.name}: {old_text!r}")
        session.save_close(doc)
    finally:
        session.quit()


def migrate_all(root: Path) -> None:
    for rel_path, mapping in MIGRATIONS.items():
        apply_mapping(root / rel_path, mapping)
        print(f"Da migrate: {rel_path}")


if __name__ == "__main__":
    migrate_all(Path(__file__).resolve().parent)
    print("XONG. Da chuyen toan bo file mau sang dung token.")
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest test_migrate_templates_to_tokens.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Run the script against the real 20 template masters**

Run: `python migrate_templates_to_tokens.py`
Expected: 20 lines of `Da migrate: <path>` followed by `XONG. Da chuyen toan bo file mau sang dung token.` — if any line instead raises `RuntimeError: Khong tim thay chuoi can thay...`, the search text in `MIGRATIONS` doesn't exactly match that file's real content (whitespace, punctuation, or a prior manual edit) — open that one file, find the actual text, and fix the corresponding tuple before re-running (the script is not idempotent — a partially-migrated file will fail on its own already-replaced entries on a second run, so restore from git first: `git checkout -- "<path>"`, fix the mapping, re-run).

- [ ] **Step 6: Manually spot-check all 20 migrated files**

Open each of the 20 `.docx` files listed in `MIGRATIONS` (Word or any viewer) and visually confirm every `{{...}}` token renders as clean, unfragmented text — no stray `{{TEN` / `DE_TAI}}` splits, no leftover old sample text. This check exists because of `python-docx`'s run-splitting behavior (spec §4/§10) — it must be done before committing.

- [ ] **Step 7: Commit**

```bash
git add migrate_templates_to_tokens.py test_migrate_templates_to_tokens.py "01. Hồ sơ đạo đức đề cương - MẪU" "02. Hồ sơ khoa học đề cương - MẪU" "03. Công văn mời chuyên gia - MẪU" "04. Hồ sơ nghiệm thu - MẪU"
git commit -m "feat: migrate all 20 template masters from literal old text to {{TOKEN}} placeholders"
```

---

### Task 9: Rewire orchestrator + all four section modules to use tokens

**Files:**
- Modify: `tao_ho_so_moi.py`, `section_dao_duc.py`, `section_khoa_hoc.py`, `section_moi_chuyen_gia.py`, `section_nghiem_thu.py`
- Test: `test_tao_ho_so_moi.py`, `test_section_dao_duc.py`, `test_section_khoa_hoc.py`, `test_section_moi_chuyen_gia.py`, `test_section_nghiem_thu.py`

**Interfaces:**
- Consumes: `tokens.build_common_tokens(info) -> dict[str, str]` (Task 4), `Session.fill_tokens(doc, tokens) -> set[str]` (Task 1), `copy_expert_cvs` (Task 7), the 20 tokenized masters (Task 8).
- Produces: `section_dao_duc.generate(session, dest_dir, info, common_tokens)`, `section_khoa_hoc.generate(session, dest_dir, info, common_tokens)`, `section_moi_chuyen_gia.generate(session, dest_dir, info, common_tokens)` — same signature shape, `title_old: str` replaced by `common_tokens: dict`. `section_nghiem_thu.generate(session, dest_dir, info)` keeps its existing signature (it never took `title_old`) but its internal functions now also receive `common_tokens` via a module-level pattern matching the others — see Step 3.

This is one coupled task because `tao_ho_so_moi.generate_all()` calls all four `generate()` functions with a matching signature — changing one without the others breaks the tool. Work through the steps below in order; only the final step commits.

- [ ] **Step 1: Update `tao_ho_so_moi.py`'s tests first (drives the interface change)**

```python
# test_tao_ho_so_moi.py — remove the two-line TITLE_OLD reference if any test used it directly (none currently do; skip if absent)
# Update test_generate_all_* tests: no signature change needed there (they call generate_all(root, dest_root, info, session), unchanged) — but add one new assertion confirming expert CVs land in the output:

def test_generate_all_copies_expert_cvs_into_output(tmp_path):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    dest_root = tmp_path / "Hồ sơ output"

    session = word_writer.Session(force_backend="docx")
    try:
        tao_ho_so_moi.generate_all(root, dest_root, info, session)
    finally:
        session.quit()

    for entry in info.expert_cvs:
        assert (dest_root / "03. Công văn mời chuyên gia" / entry.filename).exists()
```

Run: `pytest test_tao_ho_so_moi.py -k copies_expert_cvs_into_output -v`
Expected: FAIL (generate_all doesn't call copy_expert_cvs yet)

- [ ] **Step 2: Update `tao_ho_so_moi.py`**

Remove `TITLE_OLD` and thread `common_tokens` through `generate_all`, add the `copy_expert_cvs` call:

```python
# tao_ho_so_moi.py — remove these two module-level lines entirely:
# TITLE_OLD = (
#     "Đánh giá hiệu quả sản phẩm sữa dinh dưỡng pha sẵn KUN DOCTOR COLOSTRUM lên "
#     "tình trạng dinh dưỡng, miễn dịch, tiêu hóa và giấc ngủ của trẻ từ 24 đến 72 tháng tuổi"
# )
```

```python
# tao_ho_so_moi.py — add to the imports
import tokens
```

```python
# tao_ho_so_moi.py — rewrite generate_all's body (docstring unchanged, keep it)
def generate_all(root: Path, dest_root: Path, info, session: word_writer.Session) -> None:
    staging_root = Path(tempfile.mkdtemp(prefix="tao_ho_so_"))
    try:
        common_tokens = tokens.build_common_tokens(info)

        print("Dang sao chep file mau...")
        copy_templates(root, staging_root)

        print("Dang sao chep CV chu nhiem de tai...")
        copy_head_cv(root, staging_root, info)

        print("Dang sao chep CV chuyen gia da khai...")
        copy_expert_cvs(root, staging_root, info)

        print("Dang sinh ho so dao duc...")
        section_dao_duc.generate(session, staging_root / "01. Hồ sơ đạo đức đề cương", info, common_tokens)

        print("Dang sinh ho so khoa hoc de cuong...")
        section_khoa_hoc.generate(session, staging_root / "02. Hồ sơ khoa học đề cương", info, common_tokens)

        print("Dang sinh cong van moi chuyen gia...")
        section_moi_chuyen_gia.generate(session, staging_root / "03. Công văn mời chuyên gia", info, common_tokens)

        print("Dang sinh ho so nghiem thu...")
        section_nghiem_thu.generate(session, staging_root / "04. Hồ sơ nghiệm thu", info, common_tokens)

        shutil.copytree(staging_root, dest_root, dirs_exist_ok=True)
    except Exception:
        print(f"  [LUU Y] Cac file da xu ly tam thoi con luu tai: {staging_root} (de kiem tra loi)")
        raise
    else:
        shutil.rmtree(staging_root, ignore_errors=True)
```

- [ ] **Step 3: Rewrite `section_dao_duc.py`**

```python
# section_dao_duc.py — full file
from pathlib import Path

import committee_writer
import word_writer
from excel_reader import ProjectInfo, parse_timeline

ROLES = ["Chủ tịch Hội đồng", "Thành viên", "Thành viên", "Thành viên", "Thành viên"]


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, common_tokens: dict) -> None:
    _quyet_dinh_giao_de_tai(session, dest_dir, info, common_tokens)
    _qdtlhd_dao_duc(session, dest_dir, info, common_tokens)
    _bb_hop_hd_dao_duc(session, dest_dir, info, common_tokens)
    _bb_kiem_phieu_hd_dao_duc(session, dest_dir, info, common_tokens)
    _qd_chap_nhan_dao_duc(session, dest_dir, info, common_tokens)
    _bang_kiem_danh_gia_dao_duc(session, dest_dir, info, common_tokens)


def _quyet_dinh_giao_de_tai(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "00. QĐ Giao đề tài.docx")
    session.fill_tokens(doc, common_tokens)

    head_text = f"{info.head.degree} {info.head.name}".strip()
    head_org = f" - {info.head.org}" if info.head.org else ""
    session.set_cell(doc, 3, 2, 3, f"Chủ nhiệm đề tài: \r{head_text}{head_org}.")

    members_text = "\r".join(f"{p.degree} {p.name}".strip() for p in info.researchers)
    session.set_cell(doc, 3, 3, 3, f"Thành viên thực hiện:\r{members_text}")

    session.save_close(doc)


def _qdtlhd_dao_duc(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "01. QĐTLHĐ đạo đức đề cương.docx")
    session.fill_tokens(doc, common_tokens)
    committee_writer.write_committee_roster(
        session, doc, 2, info.ethics_committee, roles=ROLES, name_col=1, org_col=2, role_col=3
    )
    committee_writer.write_committee_secretaries(
        session, doc, 3, info.ethics_committee, name_col=1, org_col=2
    )
    session.save_close(doc)


def _bb_hop_hd_dao_duc(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "02. BB họp HĐ đạo đức.docx")
    session.fill_tokens(doc, common_tokens)
    chair = info.ethics_committee.chair
    session.replace_text(doc, "PGs. Ts. Hoàng Thị Thanh", f"{chair.degree} {chair.name}".strip())
    session.replace_text(
        doc,
        "Quyết định số: 04/QĐ-YHUD/2024 ngày 19 tháng 04 năm 2024",
        f"Quyết định số: ……/QĐ-YHUD/{info.year} ngày …… tháng …… năm {info.year}",
    )
    session.replace_text(doc, "Thời gian: ngày 25 tháng 04 năm 2024", f"Thời gian: ngày …… tháng …… năm {info.year}")
    session.save_close(doc)


def _bb_kiem_phieu_hd_dao_duc(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "03. BB kiểm phiếu HĐ đạo đức.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _qd_chap_nhan_dao_duc(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "04. QĐ chấp nhận đạo đức.docx")
    session.fill_tokens(doc, common_tokens)
    chair = info.ethics_committee.chair
    session.set_cell(doc, 2, 1, 2, f"CHỦ TỊCH HỘI ĐỒNG\r{chair.degree} {chair.name}".strip())
    session.save_close(doc)


def _bang_kiem_danh_gia_dao_duc(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Bảng kiểm đánh giá đạo đức.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)
```

Note what disappeared versus the original: the `parse_timeline` import/call inside `_qd_chap_nhan_dao_duc` (now covered by `{{THOI_GIAN_BAT_DAU}}`/`{{THOI_GIAN_KET_THUC}}` in `common_tokens`), the `"Địa điểm triển khai nghiên cứu"` dots-only replacement (now `{{DIA_DIEM_TRIEN_KHAI}}`, real data when A07 is filled — the bug from the original audit), and the standalone `"2024"`/`title_old` calls in every function (now covered by `fill_tokens`). `parse_timeline` is no longer imported in this file — remove it from the `from excel_reader import ...` line, keep only `ProjectInfo`.

- [ ] **Step 4: Rewrite `section_khoa_hoc.py`**

```python
# section_khoa_hoc.py — full file
from pathlib import Path

import committee_writer
import word_writer
from excel_reader import ProjectInfo

ROLES = ["Chủ tịch Hội đồng", "Phản biện 1", "Phản biện 2", "Ủy viên", "Ủy viên"]


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, common_tokens: dict) -> None:
    _qdtlhd_khoa_hoc(session, dest_dir, info, common_tokens)
    _bb_hop_thong_qua_de_cuong(session, dest_dir, info, common_tokens)
    _bb_kiem_phieu_thong_qua_de_cuong(session, dest_dir, info, common_tokens)
    _qd_phe_duyet_de_tai(session, dest_dir, info, common_tokens)
    _phieu_cham_diem_hd_de_cuong(session, dest_dir, info, common_tokens)
    _phieu_nhan_xet_danh_gia_ho_so(session, dest_dir, info, common_tokens)


def _qdtlhd_khoa_hoc(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "05. QĐ TLHĐ khoa học xét đề cương.docx")
    session.fill_tokens(doc, common_tokens)
    committee_writer.write_committee_roster(
        session, doc, 2, info.proposal_committee, roles=ROLES, name_col=1, org_col=2, role_col=3
    )
    committee_writer.write_committee_secretaries(
        session, doc, 3, info.proposal_committee, name_col=2, org_col=3
    )
    session.save_close(doc)


def _bb_hop_thong_qua_de_cuong(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "06. BB họp thông qua đề cương.docx")
    session.fill_tokens(doc, common_tokens)
    chair = info.proposal_committee.chair
    session.replace_text(
        doc,
        "PGs. Ts. Hoàng Thị Thanh - Chủ tịch Hội đồng điều khiển phiên họp",
        f"{chair.degree} {chair.name} - Chủ tịch Hội đồng điều khiển phiên họp".strip(),
    )
    session.save_close(doc)


def _bb_kiem_phieu_thong_qua_de_cuong(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "07. BB kiểm phiếu thông qua đề cương.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _qd_phe_duyet_de_tai(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "08. QĐ phê duyệt đề tài.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _phieu_cham_diem_hd_de_cuong(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Phiếu chấm điểm HĐ đề cương.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _phieu_nhan_xet_danh_gia_ho_so(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Phiếu nhận xét đánh giá hồ sơ.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)
```

Note: `_qd_phe_duyet_de_tai` no longer needs `parse_timeline` (covered by `{{THOI_GIAN_BAT_DAU}}`/`{{THOI_GIAN_KET_THUC}}`) — remove that import too.

- [ ] **Step 5: Rewrite `section_moi_chuyen_gia.py`**

```python
# section_moi_chuyen_gia.py — full file
from pathlib import Path

import word_writer
from excel_reader import ProjectInfo


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, common_tokens: dict) -> None:
    doc = session.open(dest_dir / "Công văn mời chuyên gia.docx")

    session.fill_tokens(doc, common_tokens)

    session.replace_text(
        doc,
        "Suy dinh dưỡng ở trẻ em dưới 5 tuổi – đặc biệt là suy dinh dưỡng thấp còi vẫn là một vấn đề sức khỏe cộng đồng.",
        "[Bổ sung bối cảnh/lý do triển khai dự án tại đây]",
    )
    session.replace_text(
        doc,
        "Một trong những giải pháp làm giảm tình trạng suy dinh dưỡng ở trẻ dưới 5 tuổi là sử dụng các sản phẩm bổ sung dinh dưỡng trong hệ thống trường mầm non.",
        "",
    )
    session.replace_text(
        doc,
        "Nhằm đánh giá tình trạng suy dinh dưỡng ở trẻ dưới 5 tuổi và hiệu quả của sản phẩm bổ sung dinh dưỡng LOF KUN COLOSTRUM, Viện Y học ứng dụng Việt Nam tiến hành triển khai nghiên cứu",
        "Viện Y học ứng dụng Việt Nam tiến hành triển khai đề tài",
    )
    session.replace_text(
        doc,
        "Thời gian: 9 giờ 00 – sáng thứ 7 ngày 07 tháng 12 năm 2024.",
        f"Thời gian: …… giờ ……, ngày …… tháng …… năm {info.year}.",
    )

    session.save_close(doc)
```

Note: the third bespoke `replace_text` call above ("Nhằm đánh giá...") still hardcodes "Viện Y học ứng dụng Việt Nam" in its *replacement* text — this one is intentionally left as-is (not tokenized) because it's rewriting a whole introductory sentence, not a name field, and the sentence was never in the audited bug list; revisit only if a future project needs a different host org's wording here specifically. The `"2024"` standalone occurrence used previously for the title-block year is now covered by `fill_tokens`.

- [ ] **Step 6: Rewrite `section_nghiem_thu.py`**

```python
# section_nghiem_thu.py — full file
from pathlib import Path

import committee_writer
import word_writer
from excel_reader import ProjectInfo

ROLES = ["Chủ tịch\rHội đồng", "Phản biện 1", "Phản biện 2", "Ủy viên", "Uỷ viên"]

SCORING_FORM_FILENAMES = {
    "TVCT_ĐGHQ": "Phiếu chấm điểm nghiệm thu (TVCT_ĐGHQ).docx",
    "TNLS": "Phiếu chấm điểm nghiệm thu (TNLS).docx",
}


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, common_tokens: dict) -> None:
    _quyet_dinh_thanh_lap(session, dest_dir, info, common_tokens)
    _bb_hop_hd_nghiem_thu(session, dest_dir, info, common_tokens)
    _bb_kiem_phieu_nghiem_thu(session, dest_dir, info, common_tokens)
    _qd_cong_nhan_ket_qua(session, dest_dir, info, common_tokens)
    _phieu_cham_diem_nghiem_thu(session, dest_dir, info, common_tokens)
    _phieu_ky_nhan_tien(session, dest_dir, info, common_tokens)
    _phieu_nhan_xet_nghiem_thu(session, dest_dir, info, common_tokens)


def _quyet_dinh_thanh_lap(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "9. Quyết định thành lập HĐ nghiệm thu.docx")
    session.fill_tokens(doc, common_tokens)
    committee_writer.write_committee_roster(
        session, doc, 2, info.acceptance_committee, roles=ROLES, name_col=2, org_col=3, role_col=4
    )
    committee_writer.write_committee_secretaries(
        session, doc, 3, info.acceptance_committee, name_col=1, org_col=2, number_prefix=True
    )
    session.save_close(doc)


def _bb_hop_hd_nghiem_thu(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "10. Biên bản họp HĐ nghiệm thu.docx")
    session.fill_tokens(doc, common_tokens)
    member_count = committee_writer.roster_size(info.acceptance_committee)
    session.replace_text(
        doc,
        "5. Số thành viên Hội đồng theo quyết định là …… người",
        f"5. Số thành viên Hội đồng theo quyết định là {member_count:02d} người",
    )
    session.save_close(doc)


def _bb_kiem_phieu_nghiem_thu(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "11. Biên bản kiểm phiếu nghiệm thu.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _qd_cong_nhan_ket_qua(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "12. Quyết định công nhận kết quả đề tài.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _phieu_cham_diem_nghiem_thu(session, dest_dir, info, common_tokens):
    if info.research_type not in SCORING_FORM_FILENAMES:
        raise ValueError(
            f"Loại hình nghiên cứu '{info.research_type}' (mã A02) không hợp lệ - "
            f"chỉ chấp nhận {sorted(SCORING_FORM_FILENAMES)}"
        )

    selected_filename = SCORING_FORM_FILENAMES[info.research_type]
    for research_type, filename in SCORING_FORM_FILENAMES.items():
        if filename != selected_filename:
            unused_path = dest_dir / filename
            if unused_path.exists():
                unused_path.unlink()

    doc = session.open(dest_dir / selected_filename)
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _phieu_ky_nhan_tien(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Phiếu ký nhận tiền.docx")
    session.fill_tokens(doc, common_tokens)
    secretary = info.acceptance_committee.secretaries[0]
    session.set_cell(doc, 2, 7, 2, f"{secretary.degree} {secretary.name}".strip())
    session.save_close(doc)


def _phieu_nhan_xet_nghiem_thu(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Phiếu nhận xét nghiệm thu.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)
```

Note the fix inside `_phieu_ky_nhan_tien`: `set_cell(doc, 2, 7, 2, "Hoàng Hà Linh")` (a hardcoded name unrelated to any Excel field, per the original code's own comment) is replaced with the acceptance committee's first secretary — resolved with the user during planning (spec revision, §Ambiguity resolution in this plan's Task 9 context). `_qd_cong_nhan_ket_qua` no longer needs `replace_text_any` (the "20XX"/"20xx" split is now a single `{{NAM}}` token after Task 8's migration) — `word_writer.Session.replace_text_any` itself stays in `word_writer.py`, just unused by this file now; do not delete it, it's still directly tested by `test_word_writer.py`.

- [ ] **Step 7: Update the four section test files**

Each currently asserts against literal old text being replaced; update to build a `common_tokens` dict via `tokens.build_common_tokens(info)` and assert the token surface renders correctly. Read each existing `test_section_*.py` file first (not reproduced in full here — they follow the same fixture-based pattern as `test_word_writer.py`/`test_excel_reader.py` already shown in this plan) and mechanically:
1. Replace any `title_old` argument passed into a `generate(...)` call with `tokens.build_common_tokens(info)`.
2. Replace assertions like `"OLD_TITLE" not in text` / `"NEW_TITLE" in text` with the equivalent against the real `info.title` value, since the token is now filled with real data end-to-end rather than a separate old/new string pair.
3. Add `import tokens` to each test file's imports.

Run: `pytest test_section_dao_duc.py test_section_khoa_hoc.py test_section_moi_chuyen_gia.py test_section_nghiem_thu.py -v`
Expected: after the mechanical updates, all PASS.

- [ ] **Step 8: Run the entire test suite**

Run: `pytest -v`
Expected: all PASS, zero failures, zero errors. This is the single point where the whole coupled change is verified end-to-end.

- [ ] **Step 9: Commit**

```bash
git add tao_ho_so_moi.py section_dao_duc.py section_khoa_hoc.py section_moi_chuyen_gia.py section_nghiem_thu.py test_tao_ho_so_moi.py test_section_dao_duc.py test_section_khoa_hoc.py test_section_moi_chuyen_gia.py test_section_nghiem_thu.py
git commit -m "refactor: rewire orchestrator and all sections onto common_tokens/fill_tokens, drop TITLE_OLD, fix acceptance-committee-secretary hardcode"
```

---

### Task 10: Excel dropdowns — `capnhat_danh_sach_cv.py`

**Files:**
- Create: `capnhat_danh_sach_cv.py`
- Create: `capnhat_danh_sach_cv.bat`
- Test: `test_capnhat_danh_sach_cv.py`

**Interfaces:**
- Consumes: filesystem listing of `CV chuyên gia/`.
- Produces: a `_Lists` hidden sheet in the checklist workbook + Data Validation on A02 and every F01-F10 filename cell (column 5) in the `Đề tài - Mẫu trắng dự án mới` sheet.

- [ ] **Step 1: Write the failing tests**

```python
# test_capnhat_danh_sach_cv.py
from pathlib import Path

import openpyxl

import capnhat_danh_sach_cv as refresh


def _make_checklist(tmp_path, cv_filenames):
    cv_dir = tmp_path / "CV chuyên gia"
    cv_dir.mkdir()
    for name in cv_filenames:
        (cv_dir / name).write_text("x")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Đề tài - Mẫu trắng dự án mới"
    ws.cell(row=6, column=1, value="A01")
    ws.cell(row=7, column=1, value="A02")
    ws.cell(row=67, column=1, value="F01")
    ws.cell(row=68, column=1, value="F02")
    path = tmp_path / "checklist.xlsx"
    wb.save(path)
    return path, cv_dir


def test_refresh_writes_cv_filenames_into_hidden_lists_sheet(tmp_path):
    checklist_path, cv_dir = _make_checklist(tmp_path, ["a.docx", "b.pdf"])

    refresh.refresh(checklist_path, cv_dir)

    wb = openpyxl.load_workbook(checklist_path)
    assert "_Lists" in wb.sheetnames
    values = {c.value for c in wb["_Lists"]["A"] if c.value}
    assert values == {"a.docx", "b.pdf"}


def test_refresh_adds_data_validation_to_f01_and_f02_filename_cells(tmp_path):
    checklist_path, cv_dir = _make_checklist(tmp_path, ["a.docx"])

    refresh.refresh(checklist_path, cv_dir)

    wb = openpyxl.load_workbook(checklist_path)
    ws = wb["Đề tài - Mẫu trắng dự án mới"]
    ranges = [str(dv.sqref) for dv in ws.data_validations.dataValidation]
    assert any("E67" in r for r in ranges)
    assert any("E68" in r for r in ranges)


def test_refresh_adds_static_data_validation_to_a02(tmp_path):
    checklist_path, cv_dir = _make_checklist(tmp_path, [])

    refresh.refresh(checklist_path, cv_dir)

    wb = openpyxl.load_workbook(checklist_path)
    ws = wb["Đề tài - Mẫu trắng dự án mới"]
    a02_validations = [dv for dv in ws.data_validations.dataValidation if "C7" in str(dv.sqref)]
    assert len(a02_validations) == 1
    assert "TVCT_ĐGHQ" in a02_validations[0].formula1
    assert "TNLS" in a02_validations[0].formula1


def test_refresh_is_idempotent_on_lists_sheet(tmp_path):
    checklist_path, cv_dir = _make_checklist(tmp_path, ["a.docx"])

    refresh.refresh(checklist_path, cv_dir)
    refresh.refresh(checklist_path, cv_dir)

    wb = openpyxl.load_workbook(checklist_path)
    assert wb.sheetnames.count("_Lists") == 1
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_capnhat_danh_sach_cv.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'capnhat_danh_sach_cv'`

- [ ] **Step 3: Implement `capnhat_danh_sach_cv.py`**

Note on cell addressing: column 3 = name, column 5 = filename (confirmed against the live workbook — see plan header). A02's editable value cell is column 3 of row 7 (`C7`); the F-row filename cells are column 5 (`E<row>`) of whichever row holds each `F0X` code.

```python
# capnhat_danh_sach_cv.py
from pathlib import Path

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

CHECKLIST_PATH = Path(__file__).resolve().parent / "Form checklist hồ sơ dự án.xlsx"
CV_DIR = Path(__file__).resolve().parent / "CV chuyên gia"
TEMPLATE_SHEET = "Đề tài - Mẫu trắng dự án mới"
LISTS_SHEET = "_Lists"
RESEARCH_TYPES = ["TVCT_ĐGHQ", "TNLS"]


def _refresh_lists_sheet(wb, cv_filenames: list) -> int:
    if LISTS_SHEET in wb.sheetnames:
        del wb[LISTS_SHEET]
    ws = wb.create_sheet(LISTS_SHEET)
    ws.sheet_state = "hidden"
    for i, name in enumerate(sorted(cv_filenames), start=1):
        ws.cell(row=i, column=1, value=name)
    return len(cv_filenames)


def _clear_existing_validations(ws, cell_refs: set) -> None:
    keep = []
    for dv in ws.data_validations.dataValidation:
        if not any(ref in str(dv.sqref) for ref in cell_refs):
            keep.append(dv)
    ws.data_validations.dataValidation = keep


def _build_code_index(ws) -> dict:
    return {
        row[0].value: row[0].row
        for row in ws.iter_rows(min_row=5, max_col=1)
        if isinstance(row[0].value, str)
    }


def refresh(checklist_path: Path = CHECKLIST_PATH, cv_dir: Path = CV_DIR) -> None:
    cv_filenames = [p.name for p in cv_dir.iterdir() if p.is_file()]

    wb = openpyxl.load_workbook(checklist_path)
    n = _refresh_lists_sheet(wb, cv_filenames)

    ws = wb[TEMPLATE_SHEET]
    index = _build_code_index(ws)

    f_cell_refs = {f"E{index[f'F{i:02d}']}" for i in range(1, 11) if f"F{i:02d}" in index}
    a02_cell_ref = {f"C{index['A02']}"} if "A02" in index else set()
    _clear_existing_validations(ws, f_cell_refs | a02_cell_ref)

    if n > 0:
        cv_dv = DataValidation(type="list", formula1=f"='{LISTS_SHEET}'!$A$1:$A${n}", allow_blank=True)
        ws.add_data_validation(cv_dv)
        for ref in f_cell_refs:
            cv_dv.add(ws[ref])

    if a02_cell_ref:
        research_type_dv = DataValidation(
            type="list", formula1=f'"{",".join(RESEARCH_TYPES)}"', allow_blank=True
        )
        ws.add_data_validation(research_type_dv)
        for ref in a02_cell_ref:
            research_type_dv.add(ws[ref])

    wb.save(checklist_path)


if __name__ == "__main__":
    refresh()
    print("Da cap nhat danh sach dropdown CV va kieu nghien cuu trong checklist.")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_capnhat_danh_sach_cv.py -v`
Expected: PASS (4 passed)

- [ ] **Step 5: Create the `.bat` wrapper**

```bat
@echo off
echo Dang cap nhat danh sach dropdown CV chuyen gia va kieu nghien cuu...
python capnhat_danh_sach_cv.py
pause
```

Save as `capnhat_danh_sach_cv.bat` (same directory as `setup.bat`).

- [ ] **Step 6: Run the script against the real checklist**

Run: `python capnhat_danh_sach_cv.py`
Expected: `Da cap nhat danh sach dropdown CV va kieu nghien cuu trong checklist.`

Then run: `pytest test_excel_reader.py test_tao_ho_so_moi.py -v` to confirm the checklist still loads correctly after gaining a `_Lists` sheet and validation objects.
Expected: all PASS.

- [ ] **Step 7: Commit**

```bash
git add capnhat_danh_sach_cv.py capnhat_danh_sach_cv.bat test_capnhat_danh_sach_cv.py "Form checklist hồ sơ dự án.xlsx"
git commit -m "feat: add CV filename and research-type Excel dropdowns via a manually-run refresh script"
```

---

### Task 11: New template-authoring guide + HUONG_DAN.md pointer

**Files:**
- Create: `HUONG_DAN_LAM_MAU_MOI.md`
- Modify: `HUONG_DAN.md` (§6)

**Interfaces:**
- Consumes: the token vocabulary from `tokens.py` (Task 4).
- Produces: documentation only, no code interface.

- [ ] **Step 1: Write `HUONG_DAN_LAM_MAU_MOI.md`**

```markdown
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

## Cách 1 — Mẫu chỉ cần token dùng chung (không cần viết code)

1. Đặt file `.docx` vào đúng thư mục `- MẪU` tương ứng, đặt tên file **giống
   hệt** tên sẽ xuất hiện trong hồ sơ đầu ra (không có hậu tố `" - MẪU"`).
2. Gõ trực tiếp các token cần dùng (ví dụ `{{TEN_DE_TAI}}`, `{{NAM}}`) vào
   đúng chỗ trong nội dung file Word.
3. Xong — công cụ tự quét thư mục `- MẪU` và tự điền token, không cần sửa
   file `.py` nào.

## Cách 2 — Mẫu cần dữ liệu riêng (bảng hội đồng, chọn file theo điều kiện...)

1. Làm bước 1-2 ở Cách 1 cho các token dùng chung.
2. Mở file `section_*.py` tương ứng với phần hồ sơ đó, viết một hàm
   `_ten_ham(session, dest_dir, info, common_tokens)`:
   ```python
   def _ten_ham(session, dest_dir, info, common_tokens):
       doc = session.open(dest_dir / "Tên file.docx")
       session.fill_tokens(doc, common_tokens)
       # phần logic riêng của mẫu này, ví dụ bảng hội đồng:
       committee_writer.write_committee_roster(session, doc, 2, info.ethics_committee, ...)
       session.save_close(doc)
   ```
3. Gọi hàm này trong `generate()` của file đó.
4. Nếu mẫu cần một trường dữ liệu chưa có trong checklist Excel: thêm mã
   mục mới vào `Form checklist hồ sơ dự án.xlsx` (theo mẫu
   `migrate_add_research_location.py`), đọc trường đó trong
   `excel_reader.py`, rồi quyết định: nếu trường đó dùng chung cho nhiều
   mẫu → thêm vào `build_common_tokens()` trong `tokens.py`; nếu chỉ dùng
   riêng cho 1 mẫu → truyền trực tiếp trong hàm `_ten_ham` như Cách 2.
5. Viết/cập nhật test tương ứng trong `test_section_*.py`.
6. Chạy thử toàn bộ `python tao_ho_so_moi.py` với một sheet Excel thử
   nghiệm trước khi dùng cho dự án thật.

## Trường không thể biết trước lúc sinh hồ sơ (ngày họp, số quyết định...)

Không cần token — gõ thẳng dấu `……` (chấm chấm) vào đúng chỗ trong file
mẫu, để người dùng tự điền tay sau khi hồ sơ được tạo ra.

## Thêm token dùng chung mới

Đặt tên `{{VIET_HOA_CO_GACH_DUOI}}`, có ý nghĩa rõ ràng bằng tiếng Việt.
Thêm vào bảng ở đầu tài liệu này và vào `build_common_tokens()` trong
`tokens.py` cùng lúc, để tài liệu này luôn khớp với code thật.
```

- [ ] **Step 2: Update `HUONG_DAN.md` §6**

Read the current `HUONG_DAN.md` §6 (§6.2 in particular) and replace its content with a short pointer, keeping §6.1 (folder/file naming convention) unchanged since it's still accurate:

```markdown
### 6.2 Khi cần thêm một mẫu tài liệu hoàn toàn mới

Xem hướng dẫn chi tiết tại [`HUONG_DAN_LAM_MAU_MOI.md`](HUONG_DAN_LAM_MAU_MOI.md)
— quy trình đã đổi sang dùng token `{{TEN_BIEN}}` thay vì tìm-thay theo câu
chữ mẫu cũ.
```

- [ ] **Step 3: Verify the guide's token table matches `tokens.py` exactly**

Run: `python -c "import tokens, inspect; print(inspect.getsource(tokens.build_common_tokens))"` and manually diff the token names against the table in `HUONG_DAN_LAM_MAU_MOI.md` — every key in the function must appear in the table and vice versa. No automated test needed for a docs file; this is a one-time manual check.

- [ ] **Step 4: Commit**

```bash
git add HUONG_DAN_LAM_MAU_MOI.md HUONG_DAN.md
git commit -m "docs: add token-based template-authoring guide, point HUONG_DAN.md at it"
```

---

## Self-Review Notes

*(kept here as the completed self-review record, not a to-do)*

- **Spec coverage:** §3 → Tasks 1, 4, 9. §4 → Task 8. §5 → Tasks 2, 3, 5. §6 → Task 10. §7 → Tasks 5, 6, 7. §8 → Task 6. §9 → Task 11. §10 → tests embedded in every task + Task 8 Step 6. §11 (backward compatibility) → every checklist-reading change uses the existing `if "<code>" in index` optional-blank-safe pattern (Tasks 2, 5); the coupled single-commit risk is called out explicitly at the top of Task 9.
- **Placeholder scan:** no TBD/TODO; the one open question during drafting (which secretary field feeds `Phiếu ký nhận tiền.docx`) was resolved with the user before writing Task 9 and the resolution is baked directly into Task 9 Step 6's code and note.
- **Type consistency:** `common_tokens: dict` is threaded with the same name and shape from `tokens.build_common_tokens()` (Task 4) through `tao_ho_so_moi.generate_all()` (Task 9 Step 2) into every `section_*.generate()` call (Task 9 Steps 3-6) — verified the parameter name and call sites match across all five files. `ExpertCvEntry` fields (`code, name, role, filename`) are consistent between `excel_reader.py` (Task 5) and its use in `tao_ho_so_moi.copy_expert_cvs` (Task 7). `Session.fill_tokens` signature (`doc, tokens -> set[str]`) is identical between its Task 1 definition and every call site in Task 9.
