# Excel Token Architecture (Phase 0 + Phase 1 + Phase 1.7) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reproduce and document item C's "tên dài quá" bug (Phase 0), then move the `{{TOKEN}}` ↔ checklist-code mapping out of Python (`tokens.py`) into a new Excel sheet `_Tokens` (Phase 1) so Excel is the single source of truth for tokens, and add a shared, dropdown-driven person registry sheet `_NhanSu` (Phase 1.7) so any reusable person (chủ nhiệm, thư ký, committee members) is entered once and picked by dropdown everywhere instead of retyped.

**Architecture:** `token_rules.py` reads a new hidden `_Tokens` sheet (token name → checklist code → transform "kind") and resolves it into a plain `{token: value}` dict at `load_project_data()` time, stored on `ProjectInfo.common_tokens`. `tokens.build_common_tokens(info)` shrinks to a 1-line passthrough, keeping its public signature stable so no section module or existing section-level test needs to change. Separately, `_NhanSu` is a hidden sheet of people (name → degree/org/contact), wired into every project sheet's person rows via Excel `DataValidation` dropdown (name) + `VLOOKUP` formulas (degree/org) — pure Excel-side change, no `excel_reader.py` reading logic changes needed.

**Tech Stack:** Python 3, `openpyxl` (Excel read/write), `python-docx` + `pywin32`/Word COM (`word_writer.py`), `pytest`.

**Spec:** `C:\Users\Kien\.claude\plans\xem-c-c-b-o-c-o-golden-lynx.md` (full 8-phase plan approved by the user; this document implements Phase 0 + Phase 1 + Phase 1.7 only — the foundation. Later phases get their own plan documents once this lands, since their exact details depend on what this phase actually produces.)

## Global Constraints

- Keep `tokens.build_common_tokens(info: ProjectInfo) -> dict` name/signature/return-shape stable — `tao_ho_so_moi.py` and all 4 `section_*.py` modules depend on it unchanged.
- Every new hidden config sheet (`_Tokens`, `_NhanSu`) follows the existing `_Lists` sheet convention: `sheet_state = "hidden"`, wholesale-recreatable by its migration script, `Path(__file__).resolve().parent / excel_reader.CHECKLIST_FILENAME` as the default checklist path with an overridable parameter for testing.
- All new migration scripts must be idempotent (safe to re-run) and must never silently destroy user-entered data — verify with an explicit idempotency test per script (matches `test_migrate_add_research_location.py`'s existing pattern).
- Console messages printed by scripts follow the existing convention in this repo: plain ASCII (no diacritics) for top-level `print()` status lines in `if __name__ == "__main__":` blocks (matches `migrate_add_research_location.py`, `capnhat_danh_sach_cv.py`), but Vietnamese with full diacritics inside error messages raised as exceptions (matches `excel_reader.py`).
- No VBA/macros — all Excel-side automation (dropdowns, lookups) must be plain `openpyxl` `DataValidation` objects and standard Excel formulas (`VLOOKUP`, `IFERROR`), consistent with the existing `_Lists`/`capnhat_danh_sach_cv.py` dropdown mechanism.
- Run `pytest` after every task; all existing tests must keep passing (no regressions), especially `test_excel_reader.py` and the 4 `test_section_*.py` files, none of which should need modification in this plan.

---

### Task 1: Phase 0 — reproduce and document item C's "tên dài quá" error

**Files:**
- Create (temporary, deleted at end of task): `spike_item_c_repro.py`
- Create (kept): `.superpowers/sdd/2026-08-27-item-c-spike-findings.md`

**Interfaces:**
- Consumes: `excel_reader.load_project_data`, `section_khoa_hoc._bb_kiem_phieu_thong_qua_de_cuong`, `tokens.build_common_tokens`, `word_writer.Session(force_backend="com")` — all existing, unchanged.
- Produces: a findings document (`.superpowers/sdd/2026-08-27-item-c-spike-findings.md`) that Phase 7's future plan will read to decide the real fix. No production code changes.

- [ ] **Step 1: Write the repro script**

```python
# spike_item_c_repro.py
"""Spike thu: tai hien loi 'ten dai qua' o file 07 (item C trong bao cao loi).
Chay thu cong (can Word COM cai san), khong phai mot phan bo test tu dong."""
import dataclasses
import shutil
import tempfile
from pathlib import Path

import excel_reader
import paths
import section_khoa_hoc
import tokens
import word_writer

LONG_TITLE = "Nghiên cứu đánh giá hiệu quả " + "và tính an toàn " * 15 + "của sản phẩm"

root = paths.project_root()
checklist_path = root / excel_reader.CHECKLIST_FILENAME
info = excel_reader.load_project_data(checklist_path, "Đề tài - Bánh ăn dặm VIAM 2027")
long_info = dataclasses.replace(info, title=LONG_TITLE)

cases = [
    ("short_local_path", Path(tempfile.mkdtemp(prefix="spike_short_"))),
    ("long_nested_path", root / ("Ho so tam rat dai de kiem tra duong dan " * 4) / "02. Ho so khoa hoc de cuong"),
]

for label, base_dir in cases:
    base_dir.mkdir(parents=True, exist_ok=True)
    src = root / "02. Hồ sơ khoa học đề cương - MẪU" / "07. BB kiểm phiếu thông qua đề cương.docx"
    shutil.copy2(src, base_dir / src.name)

    session = word_writer.Session(force_backend="com")
    try:
        print(f"--- {label}: {base_dir} ---")
        section_khoa_hoc._bb_kiem_phieu_thong_qua_de_cuong(
            session, base_dir, long_info, tokens.build_common_tokens(long_info)
        )
        print(f"{label}: OK, khong loi")
    except Exception as e:
        print(f"{label}: LOI - {type(e).__name__}: {e}")
    finally:
        session.quit()
```

- [ ] **Step 2: Run it**

Run: `python spike_item_c_repro.py`
(Requires a local Word install with COM available. If Word COM is not available in the execution environment, note that explicitly in the findings file instead of fabricating output, and mark Phase 7 as blocked-on-environment rather than blocked-on-unknown-cause.)

- [ ] **Step 3: Write the findings file**

Create `.superpowers/sdd/2026-08-27-item-c-spike-findings.md` containing: the exact exception type + message text printed for each of the two cases, which hypothesis it supports (Windows/OneDrive path-length limit vs. Word COM `Find.Execute` ~255-char search-string limit — see the approved plan's Phase 0/7 sections for the two hypotheses), and a one-line recommendation for what Phase 7's fix should target.

- [ ] **Step 4: Clean up the throwaway script and its output folders**

```bash
rm spike_item_c_repro.py
rm -rf "Ho so tam rat dai de kiem tra duong dan Ho so tam rat dai de kiem tra duong dan Ho so tam rat dai de kiem tra duong dan Ho so tam rat dai de kiem tra duong dan"
```
(Delete whatever the actual created long-path folder name was — it's a diacritic-free literal built from the script's `cases` list above.)

- [ ] **Step 5: Commit only the findings file**

```bash
git add ".superpowers/sdd/2026-08-27-item-c-spike-findings.md"
git commit -m "docs: record item C 'tên dài quá' spike findings"
```

---

### Task 2: `_Tokens` sheet migration script

**Files:**
- Create: `migrate_add_tokens_sheet.py`
- Test: `test_migrate_add_tokens_sheet.py`

**Interfaces:**
- Produces: `TOKENS_SHEET_NAME = "_Tokens"`, `TOKEN_SPECS: list[tuple[str, str, str, str, str]]` (name, code, kind, param, note), `add_tokens_sheet(checklist_path: Path = CHECKLIST_PATH) -> None`. Task 3 (`token_rules.py`) reads this sheet by name and column order at runtime — it does not import this module.

- [ ] **Step 1: Write the failing test**

```python
# test_migrate_add_tokens_sheet.py
import openpyxl
import pytest

import migrate_add_tokens_sheet as migrate


def test_add_tokens_sheet_creates_sheet_with_all_default_tokens(tmp_path):
    checklist_path = tmp_path / "checklist.xlsx"
    openpyxl.Workbook().save(checklist_path)

    migrate.add_tokens_sheet(checklist_path)

    wb = openpyxl.load_workbook(checklist_path)
    assert migrate.TOKENS_SHEET_NAME in wb.sheetnames
    ws = wb[migrate.TOKENS_SHEET_NAME]
    assert ws.sheet_state == "hidden"
    token_names = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
    assert "DONG_CHU_NHIEM_HO_TEN" in token_names
    assert len(token_names) == len(migrate.TOKEN_SPECS)


def test_add_tokens_sheet_is_idempotent(tmp_path):
    checklist_path = tmp_path / "checklist.xlsx"
    openpyxl.Workbook().save(checklist_path)

    migrate.add_tokens_sheet(checklist_path)
    migrate.add_tokens_sheet(checklist_path)

    wb = openpyxl.load_workbook(checklist_path)
    ws = wb[migrate.TOKENS_SHEET_NAME]
    token_names = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
    assert len(token_names) == len(migrate.TOKEN_SPECS)


def test_add_tokens_sheet_row_maps_kind_and_code_correctly(tmp_path):
    checklist_path = tmp_path / "checklist.xlsx"
    openpyxl.Workbook().save(checklist_path)

    migrate.add_tokens_sheet(checklist_path)

    wb = openpyxl.load_workbook(checklist_path)
    ws = wb[migrate.TOKENS_SHEET_NAME]
    rows = {ws.cell(row=r, column=1).value: r for r in range(2, ws.max_row + 1)}
    r = rows["DIA_DIEM_TRIEN_KHAI"]
    assert ws.cell(row=r, column=2).value == "A07"
    assert ws.cell(row=r, column=3).value == "raw_or_placeholder"
    assert ws.cell(row=r, column=4).value == "……………………………"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest test_migrate_add_tokens_sheet.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'migrate_add_tokens_sheet'`

- [ ] **Step 3: Write the implementation**

```python
# migrate_add_tokens_sheet.py
from pathlib import Path

import openpyxl

import excel_reader

CHECKLIST_PATH = Path(__file__).resolve().parent / excel_reader.CHECKLIST_FILENAME
TOKENS_SHEET_NAME = "_Tokens"

HEADERS = ["token_name", "code", "kind", "param", "note"]

TOKEN_SPECS = [
    ("TEN_DE_TAI", "A01", "raw", "", "Tên đề tài"),
    ("NAM", "A03", "raw", "", "Năm thực hiện hồ sơ"),
    ("DON_VI_CHU_TRI", "A04", "raw", "", "Cơ quan chủ trì"),
    ("DON_VI_DOI_TAC", "A06", "raw", "", "Cơ quan phối hợp"),
    ("CHU_NHIEM_HO_TEN", "B01", "person_ho_ten", "", "Chủ nhiệm đề tài - có học hàm/học vị"),
    ("CHU_NHIEM_TEN", "B01", "person_ten", "", "Chủ nhiệm đề tài - chỉ tên"),
    ("DONG_CHU_NHIEM_TEN", "B02", "person_ten", "", "Đồng chủ nhiệm đề tài - chỉ tên"),
    ("DONG_CHU_NHIEM_HO_TEN", "B02", "person_ho_ten", "", "Đồng chủ nhiệm đề tài - có học hàm/học vị"),
    ("THU_KY_DE_TAI", "B03", "person_ho_ten", "", "Thư ký đề tài"),
    ("THOI_GIAN_BAT_DAU", "A05", "timeline_start", "", "Mốc bắt đầu (MM/YYYY)"),
    ("THOI_GIAN_KET_THUC", "A05", "timeline_end", "", "Mốc kết thúc (MM/YYYY)"),
    ("DIA_DIEM_TRIEN_KHAI", "A07", "raw_or_placeholder", "……………………………", "Địa điểm triển khai nghiên cứu"),
]


def _write_tokens_sheet(wb) -> None:
    if TOKENS_SHEET_NAME in wb.sheetnames:
        del wb[TOKENS_SHEET_NAME]
    ws = wb.create_sheet(TOKENS_SHEET_NAME)
    ws.sheet_state = "hidden"
    for col, header in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=col, value=header)
    for row_i, spec in enumerate(TOKEN_SPECS, start=2):
        for col_i, value in enumerate(spec, start=1):
            ws.cell(row=row_i, column=col_i, value=value)


def add_tokens_sheet(checklist_path: Path = CHECKLIST_PATH) -> None:
    wb = openpyxl.load_workbook(checklist_path)
    _write_tokens_sheet(wb)
    wb.save(checklist_path)


if __name__ == "__main__":
    add_tokens_sheet()
    print("Da tao/cap nhat sheet _Tokens voi 12 token mac dinh.")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_migrate_add_tokens_sheet.py -v`
Expected: 3 passed

- [ ] **Step 5: Run the migration against the real, live checklist**

```bash
python migrate_add_tokens_sheet.py
```
Expected output: `Da tao/cap nhat sheet _Tokens voi 12 token mac dinh.` — this creates the `_Tokens` sheet in `Form checklist hồ sơ dự án.xlsx` at repo root, which Task 5's real-sheet regression test depends on.

- [ ] **Step 6: Commit**

```bash
git add migrate_add_tokens_sheet.py test_migrate_add_tokens_sheet.py "Form checklist hồ sơ dự án.xlsx"
git commit -m "feat: add _Tokens sheet migration with 12 default token definitions"
```

---

### Task 3: Token resolution engine (`token_rules.py`) + `excel_reader.py` wiring

**Files:**
- Create: `token_rules.py`
- Test: `test_token_rules.py`
- Modify: `excel_reader.py:84-98` (rename `_read_person`→`read_person`, `_read_text`→`read_text`, update internal call sites), `excel_reader.py:49-66` (`ProjectInfo` dataclass — append `common_tokens` field), `excel_reader.py:150-200` (`load_project_data` — call `token_rules.resolve_tokens` and populate `common_tokens`)

**Interfaces:**
- Consumes: `excel_reader.read_text(ws, index, code) -> str`, `excel_reader.read_person(ws, index, code) -> Optional[Person]`, `excel_reader.parse_timeline(text) -> tuple[str, str]` (all renamed/existing), `migrate_add_tokens_sheet.TOKENS_SHEET_NAME` value (`"_Tokens"`, hardcoded independently in this module — no import dependency on Task 2's script, only on the sheet existing in the workbook by that name at runtime).
- Produces: `token_rules.resolve_tokens(wb, ws, index: dict) -> dict[str, str]`, `token_rules.TRANSFORMS: dict[str, Callable]` (6 keys: `raw`, `raw_or_placeholder`, `person_ho_ten`, `person_ten`, `timeline_start`, `timeline_end`). `ProjectInfo.common_tokens: dict` (new field, default `{}`).

- [ ] **Step 1: Write the failing tests**

```python
# test_token_rules.py
import openpyxl
import pytest

import token_rules


def _code_index(ws):
    return {
        row[0].value: row[0].row
        for row in ws.iter_rows(min_row=5, max_col=1)
        if isinstance(row[0].value, str) and not row[0].value.startswith("SEC_")
    }


def _build_workbook(tmp_path, token_rows, project_rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Đề tài - Test"
    row_num = 5
    for code, name, degree, org in project_rows:
        ws.cell(row=row_num, column=1, value=code)
        ws.cell(row=row_num, column=3, value=name)
        ws.cell(row=row_num, column=4, value=degree)
        ws.cell(row=row_num, column=5, value=org)
        row_num += 1

    tokens_ws = wb.create_sheet(token_rules.TOKENS_SHEET_NAME)
    tokens_ws.cell(row=1, column=1, value="token_name")
    for i, (name, code, kind, param) in enumerate(token_rows, start=2):
        tokens_ws.cell(row=i, column=1, value=name)
        tokens_ws.cell(row=i, column=2, value=code)
        tokens_ws.cell(row=i, column=3, value=kind)
        tokens_ws.cell(row=i, column=4, value=param)

    path = tmp_path / "checklist.xlsx"
    wb.save(path)
    return path


def test_resolve_tokens_raw_kind(tmp_path):
    path = _build_workbook(
        tmp_path,
        token_rows=[("TEN_DE_TAI", "A01", "raw", "")],
        project_rows=[("A01", "Đề tài mẫu", None, None)],
    )
    wb = openpyxl.load_workbook(path)
    ws = wb["Đề tài - Test"]
    result = token_rules.resolve_tokens(wb, ws, _code_index(ws))
    assert result["{{TEN_DE_TAI}}"] == "Đề tài mẫu"


def test_resolve_tokens_raw_or_placeholder_kind_uses_param_when_blank(tmp_path):
    path = _build_workbook(
        tmp_path,
        token_rows=[("DIA_DIEM", "A07", "raw_or_placeholder", "……")],
        project_rows=[("A07", None, None, None)],
    )
    wb = openpyxl.load_workbook(path)
    ws = wb["Đề tài - Test"]
    result = token_rules.resolve_tokens(wb, ws, _code_index(ws))
    assert result["{{DIA_DIEM}}"] == "……"


def test_resolve_tokens_person_ho_ten_kind_combines_degree_and_name(tmp_path):
    path = _build_workbook(
        tmp_path,
        token_rows=[("CHU_NHIEM_HO_TEN", "B01", "person_ho_ten", "")],
        project_rows=[("B01", "Nguyễn Văn A", "TS.", "Viện ABC")],
    )
    wb = openpyxl.load_workbook(path)
    ws = wb["Đề tài - Test"]
    result = token_rules.resolve_tokens(wb, ws, _code_index(ws))
    assert result["{{CHU_NHIEM_HO_TEN}}"] == "TS. Nguyễn Văn A"


def test_resolve_tokens_person_ho_ten_kind_blank_person_is_empty_string(tmp_path):
    path = _build_workbook(
        tmp_path,
        token_rows=[("DONG_CHU_NHIEM_HO_TEN", "B02", "person_ho_ten", "")],
        project_rows=[("B02", None, None, None)],
    )
    wb = openpyxl.load_workbook(path)
    ws = wb["Đề tài - Test"]
    result = token_rules.resolve_tokens(wb, ws, _code_index(ws))
    assert result["{{DONG_CHU_NHIEM_HO_TEN}}"] == ""


def test_resolve_tokens_person_ten_kind_is_bare_name(tmp_path):
    path = _build_workbook(
        tmp_path,
        token_rows=[("CHU_NHIEM_TEN", "B01", "person_ten", "")],
        project_rows=[("B01", "Nguyễn Văn A", "TS.", "Viện ABC")],
    )
    wb = openpyxl.load_workbook(path)
    ws = wb["Đề tài - Test"]
    result = token_rules.resolve_tokens(wb, ws, _code_index(ws))
    assert result["{{CHU_NHIEM_TEN}}"] == "Nguyễn Văn A"


def test_resolve_tokens_timeline_start_and_end_kinds(tmp_path):
    path = _build_workbook(
        tmp_path,
        token_rows=[
            ("BAT_DAU", "A05", "timeline_start", ""),
            ("KET_THUC", "A05", "timeline_end", ""),
        ],
        project_rows=[("A05", "Tháng 01/2027 đến tháng 12/2027", None, None)],
    )
    wb = openpyxl.load_workbook(path)
    ws = wb["Đề tài - Test"]
    result = token_rules.resolve_tokens(wb, ws, _code_index(ws))
    assert result["{{BAT_DAU}}"] == "01/2027"
    assert result["{{KET_THUC}}"] == "12/2027"


def test_resolve_tokens_unknown_kind_raises_value_error(tmp_path):
    path = _build_workbook(
        tmp_path,
        token_rows=[("FOO", "A01", "not_a_real_kind", "")],
        project_rows=[("A01", "x", None, None)],
    )
    wb = openpyxl.load_workbook(path)
    ws = wb["Đề tài - Test"]
    with pytest.raises(ValueError):
        token_rules.resolve_tokens(wb, ws, _code_index(ws))


def test_resolve_tokens_missing_tokens_sheet_returns_empty_dict(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Đề tài - Test"
    ws.cell(row=5, column=1, value="A01")
    ws.cell(row=5, column=3, value="x")
    path = tmp_path / "no_tokens_sheet.xlsx"
    wb.save(path)

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2["Đề tài - Test"]
    result = token_rules.resolve_tokens(wb2, ws2, _code_index(ws2))
    assert result == {}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_token_rules.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'token_rules'`

- [ ] **Step 3: Rename `_read_text`/`_read_person` to public names in `excel_reader.py`**

Modify `excel_reader.py:84-98` — before:
```python
def _read_person(ws, index: dict, code: str) -> Optional[Person]:
    row = index.get(code)
    if row is None:
        raise KeyError(f"Không tìm thấy mã mục '{code}' trong checklist")
    name = _cell_text(ws, row, 3)
    if not name:
        return None
    return Person(name=name, degree=_cell_text(ws, row, 4), org=_cell_text(ws, row, 5))


def _read_text(ws, index: dict, code: str) -> str:
    row = index.get(code)
    if row is None:
        raise KeyError(f"Không tìm thấy mã mục '{code}' trong checklist")
    return _cell_text(ws, row, 3)
```
after:
```python
def read_person(ws, index: dict, code: str) -> Optional[Person]:
    row = index.get(code)
    if row is None:
        raise KeyError(f"Không tìm thấy mã mục '{code}' trong checklist")
    name = _cell_text(ws, row, 3)
    if not name:
        return None
    return Person(name=name, degree=_cell_text(ws, row, 4), org=_cell_text(ws, row, 5))


def read_text(ws, index: dict, code: str) -> str:
    row = index.get(code)
    if row is None:
        raise KeyError(f"Không tìm thấy mã mục '{code}' trong checklist")
    return _cell_text(ws, row, 3)
```

Then update every internal call site in the same file: `parse_committee` (calls `_read_person` 3×), `read_expert_cvs` uses `_cell_text` directly (no change), `load_project_data` (calls `_read_person` for `head`/`co_head`/`project_secretary`, `_read_text` for `title`/`partner_org`/`research_location`/`timeline`/`research_type`/`host_org`) — rename every `_read_person(` → `read_person(` and `_read_text(` → `read_text(` in this file (mechanical rename, `_cell_text`/`_build_code_index` stay private/unchanged).

- [ ] **Step 4: Write `token_rules.py`**

```python
# token_rules.py
from dataclasses import dataclass
from typing import Callable, Dict, List

import excel_reader

TOKENS_SHEET_NAME = "_Tokens"


@dataclass
class TokenSpec:
    name: str
    code: str
    kind: str
    param: str


def _load_token_specs(wb) -> List[TokenSpec]:
    ws = wb[TOKENS_SHEET_NAME]
    specs = []
    for row in ws.iter_rows(min_row=2):
        name = row[0].value
        if not name:
            continue
        specs.append(
            TokenSpec(
                name=name,
                code=row[1].value,
                kind=row[2].value,
                param=row[3].value or "",
            )
        )
    return specs


def _resolve_raw(ws, index, spec: TokenSpec) -> str:
    return excel_reader.read_text(ws, index, spec.code)


def _resolve_raw_or_placeholder(ws, index, spec: TokenSpec) -> str:
    value = excel_reader.read_text(ws, index, spec.code)
    return value or spec.param


def _resolve_person_ho_ten(ws, index, spec: TokenSpec) -> str:
    person = excel_reader.read_person(ws, index, spec.code)
    if person is None:
        return ""
    return f"{person.degree} {person.name}".strip()


def _resolve_person_ten(ws, index, spec: TokenSpec) -> str:
    person = excel_reader.read_person(ws, index, spec.code)
    return person.name if person else ""


def _resolve_timeline_start(ws, index, spec: TokenSpec) -> str:
    text = excel_reader.read_text(ws, index, spec.code)
    start, _end = excel_reader.parse_timeline(text)
    return start


def _resolve_timeline_end(ws, index, spec: TokenSpec) -> str:
    text = excel_reader.read_text(ws, index, spec.code)
    _start, end = excel_reader.parse_timeline(text)
    return end


TRANSFORMS: Dict[str, Callable] = {
    "raw": _resolve_raw,
    "raw_or_placeholder": _resolve_raw_or_placeholder,
    "person_ho_ten": _resolve_person_ho_ten,
    "person_ten": _resolve_person_ten,
    "timeline_start": _resolve_timeline_start,
    "timeline_end": _resolve_timeline_end,
}


def resolve_tokens(wb, ws, index: dict) -> Dict[str, str]:
    """Tra ve {} neu workbook chua co sheet _Tokens - giu tuong thich nguoc
    cho cac workbook/fixture kiem thu chua duoc migrate."""
    if TOKENS_SHEET_NAME not in wb.sheetnames:
        return {}

    result = {}
    for spec in _load_token_specs(wb):
        transform = TRANSFORMS.get(spec.kind)
        if transform is None:
            raise ValueError(
                f"Token '{spec.name}' trong sheet _Tokens dùng kind '{spec.kind}' không hợp lệ - "
                f"chỉ chấp nhận {sorted(TRANSFORMS)}"
            )
        result[f"{{{{{spec.name}}}}}"] = transform(ws, index, spec)
    return result
```

- [ ] **Step 5: Wire into `excel_reader.py`**

Modify `excel_reader.py:49-66` (`ProjectInfo` dataclass) — append at the very end (after `expert_cvs`, since dataclass fields with defaults must come after fields without):
```python
@dataclass
class ProjectInfo:
    title: str
    research_type: str
    year: int
    host_org: str
    partner_org: Optional[str]
    research_location: Optional[str]
    timeline: str
    head: Person
    co_head: Optional[Person]
    project_secretary: Optional[Person]
    researchers: List[Person]
    ethics_committee: CommitteeData
    proposal_committee: CommitteeData
    acceptance_committee: CommitteeData
    head_cv_filename: str
    expert_cvs: List[ExpertCvEntry]
    common_tokens: dict = field(default_factory=dict)
```
(`field` is already imported at the top of `excel_reader.py` via `from dataclasses import dataclass, field`.)

Modify `excel_reader.py:150-200` (`load_project_data`) — after the `index = _build_code_index(ws)` line and before the final `return ProjectInfo(...)`, add:
```python
    # Import cuc bo de tranh vong lap import (token_rules import excel_reader
    # de dung read_text/read_person/parse_timeline).
    import token_rules
    common_tokens = token_rules.resolve_tokens(wb, ws, index)
```
Then add `common_tokens=common_tokens,` as the last keyword argument in the `return ProjectInfo(...)` call.

- [ ] **Step 6: Run tests to verify they pass**

Run: `pytest test_token_rules.py test_excel_reader.py -v`
Expected: all pass — `test_excel_reader.py` must still pass unmodified (its synthetic `_build_minimal_workbook` fixture has no `_Tokens` sheet, exercising the `return {}` early-out path in `resolve_tokens`).

- [ ] **Step 7: Commit**

```bash
git add excel_reader.py token_rules.py test_token_rules.py
git commit -m "feat: add token_rules resolver, wire into excel_reader.load_project_data"
```

---

### Task 4: Shrink `tokens.py` to a stable shim

**Files:**
- Modify: `tokens.py` (full rewrite, 22 lines → 4 lines)
- Modify: `test_tokens.py` (full rewrite)

**Interfaces:**
- Consumes: `ProjectInfo.common_tokens` (from Task 3).
- Produces: `tokens.build_common_tokens(info: ProjectInfo) -> dict` — same name/signature as before; every caller (`tao_ho_so_moi.py`, all `test_section_*.py`) needs zero changes.

- [ ] **Step 1: Write the failing tests**

```python
# test_tokens.py
import excel_reader
import tokens


def _make_info(common_tokens):
    return excel_reader.ProjectInfo(
        title="t",
        research_type="TNLS",
        year=2027,
        host_org="h",
        partner_org=None,
        research_location=None,
        timeline="Tháng 01/2027 đến tháng 12/2027",
        head=excel_reader.Person(name="A"),
        co_head=None,
        project_secretary=None,
        researchers=[],
        ethics_committee=excel_reader.CommitteeData(chair=excel_reader.Person(name="C")),
        proposal_committee=excel_reader.CommitteeData(chair=excel_reader.Person(name="C")),
        acceptance_committee=excel_reader.CommitteeData(chair=excel_reader.Person(name="C")),
        head_cv_filename="cv.docx",
        expert_cvs=[],
        common_tokens=common_tokens,
    )


def test_build_common_tokens_forwards_info_common_tokens():
    info = _make_info({"{{FOO}}": "bar"})
    assert tokens.build_common_tokens(info) == {"{{FOO}}": "bar"}


def test_build_common_tokens_returns_a_copy_not_the_same_object():
    info = _make_info({"{{FOO}}": "bar"})
    result = tokens.build_common_tokens(info)
    result["{{FOO}}"] = "changed"
    assert info.common_tokens["{{FOO}}"] == "bar"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_tokens.py -v`
Expected: FAIL — old `tokens.py` doesn't accept `common_tokens` as a `ProjectInfo` kwarg yet from this test's perspective... actually it will fail with `AssertionError` since old `build_common_tokens` computes its own dict from other fields, ignoring `common_tokens` entirely, so results won't match `{"{{FOO}}": "bar"}`.

- [ ] **Step 3: Rewrite `tokens.py`**

```python
# tokens.py
from excel_reader import ProjectInfo


def build_common_tokens(info: ProjectInfo) -> dict:
    return dict(info.common_tokens)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_tokens.py -v`
Expected: 2 passed

- [ ] **Step 5: Run the full existing suite to confirm no regressions**

Run: `pytest -v`
Expected: all pass, including all 4 `test_section_*.py` files (they go through the real `load_project_data → build_common_tokens` path end-to-end and don't hand-construct `ProjectInfo`, so they're unaffected by the new required-with-default field).

- [ ] **Step 6: Commit**

```bash
git add tokens.py test_tokens.py
git commit -m "refactor: shrink tokens.build_common_tokens to a thin ProjectInfo.common_tokens passthrough"
```

---

### Task 5: Real-checklist regression test for the new tokens

**Files:**
- Modify: `test_token_rules.py` (append one test)

**Interfaces:**
- Consumes: the live `Form checklist hồ sơ dự án.xlsx` (already migrated by Task 2, Step 5), `excel_reader.load_project_data`, `paths.project_root()`.

- [ ] **Step 1: Write the failing test**

Append to `test_token_rules.py`:
```python
import paths

CHECKLIST_PATH = paths.project_root() / excel_reader.CHECKLIST_FILENAME
SHEET_VIAM = "Đề tài - Bánh ăn dặm VIAM 2027"


def test_real_checklist_resolves_new_and_existing_tokens_correctly():
    import excel_reader as er

    data = er.load_project_data(CHECKLIST_PATH, SHEET_VIAM)

    assert data.common_tokens["{{CHU_NHIEM_HO_TEN}}"] == f"{data.head.degree} {data.head.name}".strip()
    assert data.common_tokens["{{DONG_CHU_NHIEM_HO_TEN}}"] == (
        f"{data.co_head.degree} {data.co_head.name}".strip() if data.co_head else ""
    )
    assert data.common_tokens["{{THU_KY_DE_TAI}}"] == (
        f"{data.project_secretary.degree} {data.project_secretary.name}".strip()
        if data.project_secretary
        else ""
    )
```
(Add `import excel_reader` at the top of the file if not already present from Task 3's tests — it is.)

- [ ] **Step 2: Run test to verify it fails or passes**

Run: `pytest test_token_rules.py::test_real_checklist_resolves_new_and_existing_tokens_correctly -v`
Expected: PASS immediately (no new implementation needed — this test only exists to lock in end-to-end correctness against the real, already-migrated checklist file as a permanent regression guard). If it fails, it means Task 2 Step 5 (running the migration against the live file) was skipped or the live checklist's B01/B02/B03 data doesn't match expectations — fix by re-running `python migrate_add_tokens_sheet.py` and re-checking.

- [ ] **Step 3: Run the full suite one more time**

Run: `pytest -v`
Expected: all pass.

- [ ] **Step 4: Commit**

```bash
git add test_token_rules.py
git commit -m "test: lock in real-checklist token resolution as a regression guard"
```

---

### Task 6: `_NhanSu` shared person-registry sheet + seed migration

**Files:**
- Create: `migrate_add_nhan_su_sheet.py`
- Test: `test_migrate_add_nhan_su_sheet.py`

**Interfaces:**
- Produces: `NHAN_SU_SHEET_NAME = "_NhanSu"`, `HEADERS` (10 columns: `ten, hoc_ham_hoc_vi, don_vi, dia_chi, sdt, email, cccd, mst, so_tk, ngan_hang`), `add_nhan_su_sheet(checklist_path: Path = CHECKLIST_PATH) -> None`. Task 7 imports `NHAN_SU_SHEET_NAME` and calls `add_nhan_su_sheet` directly.

- [ ] **Step 1: Write the failing tests**

```python
# test_migrate_add_nhan_su_sheet.py
import openpyxl

import migrate_add_nhan_su_sheet as migrate


def _build_project_sheet(wb, sheet_name, rows):
    ws = wb.create_sheet(sheet_name)
    row_num = 5
    for code, name, degree, org in rows:
        ws.cell(row=row_num, column=1, value=code)
        ws.cell(row=row_num, column=3, value=name)
        ws.cell(row=row_num, column=4, value=degree)
        ws.cell(row=row_num, column=5, value=org)
        row_num += 1
    return ws


def test_add_nhan_su_sheet_creates_hidden_sheet_with_headers(tmp_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _build_project_sheet(wb, "Đề tài - Bánh ăn dặm VIAM 2027", [("B01", "Trương Hồng Sơn", "TS.BS.", "VIAM")])
    _build_project_sheet(wb, "Đề tài - Mẫu trắng dự án mới", [])
    path = tmp_path / "checklist.xlsx"
    wb.save(path)

    migrate.add_nhan_su_sheet(path)

    result = openpyxl.load_workbook(path)
    assert migrate.NHAN_SU_SHEET_NAME in result.sheetnames
    ws = result[migrate.NHAN_SU_SHEET_NAME]
    assert ws.sheet_state == "hidden"
    assert [ws.cell(row=1, column=c).value for c in range(1, 4)] == ["ten", "hoc_ham_hoc_vi", "don_vi"]


def test_add_nhan_su_sheet_seeds_distinct_names_from_both_project_sheets(tmp_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _build_project_sheet(
        wb,
        "Đề tài - Bánh ăn dặm VIAM 2027",
        [("B01", "Trương Hồng Sơn", "TS.BS.", "VIAM"), ("C01", "Nguyễn Công Khẩn", "GS.TS.", "Hội đồng Đạo đức")],
    )
    _build_project_sheet(wb, "Đề tài - Mẫu trắng dự án mới", [("B01", "Trương Hồng Sơn", "TS.BS.", "VIAM")])
    path = tmp_path / "checklist.xlsx"
    wb.save(path)

    migrate.add_nhan_su_sheet(path)

    result = openpyxl.load_workbook(path)
    ws = result[migrate.NHAN_SU_SHEET_NAME]
    names = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
    assert names.count("Trương Hồng Sơn") == 1
    assert "Nguyễn Công Khẩn" in names


def test_add_nhan_su_sheet_is_idempotent(tmp_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _build_project_sheet(wb, "Đề tài - Bánh ăn dặm VIAM 2027", [("B01", "Trương Hồng Sơn", "TS.BS.", "VIAM")])
    _build_project_sheet(wb, "Đề tài - Mẫu trắng dự án mới", [])
    path = tmp_path / "checklist.xlsx"
    wb.save(path)

    migrate.add_nhan_su_sheet(path)
    migrate.add_nhan_su_sheet(path)

    result = openpyxl.load_workbook(path)
    ws = result[migrate.NHAN_SU_SHEET_NAME]
    names = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
    assert names.count("Trương Hồng Sơn") == 1


def test_add_nhan_su_sheet_preserves_manually_added_contact_info(tmp_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _build_project_sheet(wb, "Đề tài - Bánh ăn dặm VIAM 2027", [("B01", "Trương Hồng Sơn", "TS.BS.", "VIAM")])
    _build_project_sheet(wb, "Đề tài - Mẫu trắng dự án mới", [])
    path = tmp_path / "checklist.xlsx"
    wb.save(path)
    migrate.add_nhan_su_sheet(path)

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2[migrate.NHAN_SU_SHEET_NAME]
    ws2.cell(row=2, column=4, value="Số 47 Đặng Văn Ngữ, Hà Nội")
    wb2.save(path)

    migrate.add_nhan_su_sheet(path)

    result = openpyxl.load_workbook(path)
    ws = result[migrate.NHAN_SU_SHEET_NAME]
    assert ws.cell(row=2, column=4).value == "Số 47 Đặng Văn Ngữ, Hà Nội"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_migrate_add_nhan_su_sheet.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'migrate_add_nhan_su_sheet'`

- [ ] **Step 3: Write the implementation**

```python
# migrate_add_nhan_su_sheet.py
from pathlib import Path

import openpyxl

import excel_reader

CHECKLIST_PATH = Path(__file__).resolve().parent / excel_reader.CHECKLIST_FILENAME
NHAN_SU_SHEET_NAME = "_NhanSu"
SHEET_NAMES = ["Đề tài - Bánh ăn dặm VIAM 2027", "Đề tài - Mẫu trắng dự án mới"]

HEADERS = ["ten", "hoc_ham_hoc_vi", "don_vi", "dia_chi", "sdt", "email", "cccd", "mst", "so_tk", "ngan_hang"]

PERSON_CODE_PREFIXES = ("B", "C", "D", "E")


def _build_code_index(ws) -> dict:
    return {
        row[0].value: row[0].row
        for row in ws.iter_rows(min_row=5, max_col=1)
        if isinstance(row[0].value, str) and not row[0].value.startswith("SEC_")
    }


def _collect_existing_people(wb) -> dict:
    """Quet 2 sheet du an hien co, gom (ten -> hoc ham, don vi) de seed
    _NhanSu - tranh viec chuyen cot hoc ham/don vi sang cong thuc tra cuu
    lam mat du lieu da nhap cua cac du an cu."""
    people = {}
    for sheet_name in SHEET_NAMES:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        index = _build_code_index(ws)
        for code, row in index.items():
            if code[0] not in PERSON_CODE_PREFIXES:
                continue
            name = ws.cell(row=row, column=3).value
            if not name or not str(name).strip():
                continue
            name = str(name).strip()
            if name in people:
                continue
            degree = ws.cell(row=row, column=4).value or ""
            org = ws.cell(row=row, column=5).value or ""
            people[name] = (degree, org)
    return people


def add_nhan_su_sheet(checklist_path: Path = CHECKLIST_PATH) -> None:
    wb = openpyxl.load_workbook(checklist_path)

    if NHAN_SU_SHEET_NAME in wb.sheetnames:
        ws = wb[NHAN_SU_SHEET_NAME]
        existing_names = {
            ws.cell(row=r, column=1).value
            for r in range(2, ws.max_row + 1)
            if ws.cell(row=r, column=1).value
        }
    else:
        ws = wb.create_sheet(NHAN_SU_SHEET_NAME)
        ws.sheet_state = "hidden"
        for col, header in enumerate(HEADERS, start=1):
            ws.cell(row=1, column=col, value=header)
        existing_names = set()

    people = _collect_existing_people(wb)
    next_row = ws.max_row + 1 if ws.max_row >= 1 else 2
    for name, (degree, org) in people.items():
        if name in existing_names:
            continue
        ws.cell(row=next_row, column=1, value=name)
        ws.cell(row=next_row, column=2, value=degree)
        ws.cell(row=next_row, column=3, value=org)
        next_row += 1
        existing_names.add(name)

    wb.save(checklist_path)


if __name__ == "__main__":
    add_nhan_su_sheet()
    print("Da tao/cap nhat sheet _NhanSu va seed du lieu tu 2 sheet du an hien co.")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_migrate_add_nhan_su_sheet.py -v`
Expected: 4 passed

- [ ] **Step 5: Run the migration against the real, live checklist**

```bash
python migrate_add_nhan_su_sheet.py
```
This seeds `_NhanSu` from every distinct name already typed across `Đề tài - Bánh ăn dặm VIAM 2027` and `Đề tài - Mẫu trắng dự án mới` — this seeding must happen **before** Task 7 converts any project-sheet degree/org cell into a `VLOOKUP` formula, otherwise those formulas would resolve against an empty registry.

- [ ] **Step 6: Commit**

```bash
git add migrate_add_nhan_su_sheet.py test_migrate_add_nhan_su_sheet.py "Form checklist hồ sơ dự án.xlsx"
git commit -m "feat: add _NhanSu shared person registry, seeded from existing project sheets"
```

---

### Task 7: Wire name-dropdown + degree/org lookup formulas onto every person row

**Files:**
- Create: `capnhat_nhan_su.py`
- Test: `test_capnhat_nhan_su.py`

**Interfaces:**
- Consumes: `migrate_add_nhan_su_sheet.NHAN_SU_SHEET_NAME`, `migrate_add_nhan_su_sheet.add_nhan_su_sheet` (Task 6).
- Produces: `wire_person_dropdowns(checklist_path: Path = CHECKLIST_PATH) -> None` — a standalone, manually-run maintenance script (same operational pattern as `capnhat_danh_sach_cv.py`, run via a paired `.bat` file if the user wants — not wired into `tao_ho_so_moi.py`'s automatic run path).

**Important caveat to note in `HUONG_DAN.md` during this task:** cells written as Excel formulas via `openpyxl` have no cached computed value until the workbook is opened and saved once in real Excel (openpyxl does not evaluate formulas). This matches the pre-existing behavior of this checklist's column-6 "TRẠNG THÁI KIỂM TRA (TỰ ĐỘNG)" status formulas — not a new risk, but worth calling out since it now also affects the degree/org columns.

- [ ] **Step 1: Write the failing tests**

```python
# test_capnhat_nhan_su.py
import openpyxl

import capnhat_nhan_su as wiring


def _build_project_sheet(wb, sheet_name, rows):
    ws = wb.create_sheet(sheet_name)
    row_num = 5
    for code, name, degree, org in rows:
        ws.cell(row=row_num, column=1, value=code)
        ws.cell(row=row_num, column=3, value=name)
        ws.cell(row=row_num, column=4, value=degree)
        ws.cell(row=row_num, column=5, value=org)
        row_num += 1
    return ws


def _build_checklist(tmp_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _build_project_sheet(
        wb,
        "Đề tài - Bánh ăn dặm VIAM 2027",
        [("B01", "Trương Hồng Sơn", "TS.BS.", "VIAM"), ("C01", "Nguyễn Công Khẩn", "GS.TS.", "Hội đồng Đạo đức")],
    )
    _build_project_sheet(wb, "Đề tài - Mẫu trắng dự án mới", [])
    path = tmp_path / "checklist.xlsx"
    wb.save(path)
    return path


def test_wire_person_dropdowns_adds_name_dropdown_on_person_rows(tmp_path):
    path = _build_checklist(tmp_path)

    wiring.wire_person_dropdowns(path)

    wb = openpyxl.load_workbook(path)
    ws = wb["Đề tài - Bánh ăn dặm VIAM 2027"]
    refs = {ref for dv in ws.data_validations.dataValidation for ref in str(dv.sqref).split()}
    assert "C5" in refs  # B01 row


def test_wire_person_dropdowns_sets_degree_and_org_lookup_formulas(tmp_path):
    path = _build_checklist(tmp_path)

    wiring.wire_person_dropdowns(path)

    wb = openpyxl.load_workbook(path)
    ws = wb["Đề tài - Bánh ăn dặm VIAM 2027"]
    degree_formula = ws.cell(row=5, column=4).value
    org_formula = ws.cell(row=5, column=5).value
    assert degree_formula.startswith("=IFERROR(VLOOKUP(C5,")
    assert org_formula.startswith("=IFERROR(VLOOKUP(C5,")


def test_wire_person_dropdowns_is_idempotent_no_duplicate_validations(tmp_path):
    path = _build_checklist(tmp_path)

    wiring.wire_person_dropdowns(path)
    wiring.wire_person_dropdowns(path)

    wb = openpyxl.load_workbook(path)
    ws = wb["Đề tài - Bánh ăn dặm VIAM 2027"]
    refs_c5 = [dv for dv in ws.data_validations.dataValidation if "C5" in str(dv.sqref).split()]
    assert len(refs_c5) == 1


def test_wire_person_dropdowns_seeds_nhan_su_before_wiring(tmp_path):
    path = _build_checklist(tmp_path)

    wiring.wire_person_dropdowns(path)

    wb = openpyxl.load_workbook(path)
    ws = wb[wiring.NHAN_SU_SHEET_NAME]
    names = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
    assert "Trương Hồng Sơn" in names
    assert "Nguyễn Công Khẩn" in names
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_capnhat_nhan_su.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'capnhat_nhan_su'`

- [ ] **Step 3: Write the implementation**

```python
# capnhat_nhan_su.py
from pathlib import Path

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

import excel_reader
import migrate_add_nhan_su_sheet as nhan_su

CHECKLIST_PATH = Path(__file__).resolve().parent / excel_reader.CHECKLIST_FILENAME
SHEET_NAMES = ["Đề tài - Bánh ăn dặm VIAM 2027", "Đề tài - Mẫu trắng dự án mới"]
NHAN_SU_SHEET_NAME = nhan_su.NHAN_SU_SHEET_NAME
PERSON_CODE_PREFIXES = ("B", "C", "D", "E")


def _build_code_index(ws) -> dict:
    return {
        row[0].value: row[0].row
        for row in ws.iter_rows(min_row=5, max_col=1)
        if isinstance(row[0].value, str) and not row[0].value.startswith("SEC_")
    }


def _nhan_su_row_count(wb) -> int:
    ws = wb[NHAN_SU_SHEET_NAME]
    return max(ws.max_row - 1, 0)


def _clear_existing_person_validations(ws, target_refs: set) -> None:
    keep = []
    for dv in ws.data_validations.dataValidation:
        dv_cells = set(str(dv.sqref).split())
        if not (dv_cells & target_refs):
            keep.append(dv)
    ws.data_validations.dataValidation = keep


def wire_person_dropdowns(checklist_path: Path = CHECKLIST_PATH) -> None:
    nhan_su.add_nhan_su_sheet(checklist_path)

    wb = openpyxl.load_workbook(checklist_path)
    n = _nhan_su_row_count(wb)
    if n == 0:
        wb.save(checklist_path)
        return

    for sheet_name in SHEET_NAMES:
        ws = wb[sheet_name]
        index = _build_code_index(ws)
        person_rows = [row for code, row in index.items() if code[0] in PERSON_CODE_PREFIXES]

        name_refs = {f"C{row}" for row in person_rows}
        _clear_existing_person_validations(ws, name_refs)

        name_dv = DataValidation(
            type="list", formula1=f"='{NHAN_SU_SHEET_NAME}'!$A$2:$A${n + 1}", allow_blank=True
        )
        ws.add_data_validation(name_dv)
        for row in person_rows:
            name_dv.add(ws.cell(row=row, column=3))
            ws.cell(
                row=row,
                column=4,
                value=f'=IFERROR(VLOOKUP(C{row},{NHAN_SU_SHEET_NAME}!$A:$C,2,FALSE),"")',
            )
            ws.cell(
                row=row,
                column=5,
                value=f'=IFERROR(VLOOKUP(C{row},{NHAN_SU_SHEET_NAME}!$A:$C,3,FALSE),"")',
            )

    wb.save(checklist_path)


if __name__ == "__main__":
    wire_person_dropdowns()
    print("Da gan dropdown chon ten + cong thuc tra cuu hoc ham/don vi cho toan bo dong nhan su.")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_capnhat_nhan_su.py -v`
Expected: 4 passed

- [ ] **Step 5: Run the full suite**

Run: `pytest -v`
Expected: all pass. Note: this step does **not** run `wire_person_dropdowns()` against the real live checklist yet — doing so converts real, currently-literal degree/org cells in the 2 real project sheets into formulas, which is a more consequential change to the live file the user edits by hand. Leave that for a deliberate, separate manual step the user confirms (see Step 6), since — per the caveat above — those formula cells will show blank in Excel until the user opens and saves the workbook once.

- [ ] **Step 6: Run against the real checklist only after confirming with the user**

```bash
python capnhat_nhan_su.py
```
Before running this against the live `Form checklist hồ sơ dự án.xlsx`, tell the user explicitly: this converts every existing person row's degree/org cells (columns D/E) in both project sheets from plain text into `VLOOKUP` formulas, and that opening the file in Excel once afterward is required for those formulas to display computed values. Get their go-ahead first (this is exactly the kind of "affects shared state the user edits by hand" action that warrants a check-in before acting, not a blanket default).

- [ ] **Step 7: Update `HUONG_DAN.md`**

Add a short new section describing the `_NhanSu` workflow: to add a new person, add them to the (hidden) `_NhanSu` sheet first (unhide it via Excel's sheet right-click menu, or ask the tool maintainer), then pick their name from the dropdown in column C of their row in the project sheet — degree (column D) and org (column E) auto-fill via lookup formula once a name is picked, and stay editable overrides are no longer needed/expected. Also update `HUONG_DAN_LAM_MAU_MOI.md`'s "thêm token mới" section to describe adding a row to `_Tokens` instead of editing `tokens.py`.

- [ ] **Step 8: Commit**

```bash
git add capnhat_nhan_su.py test_capnhat_nhan_su.py HUONG_DAN.md HUONG_DAN_LAM_MAU_MOI.md
git commit -m "feat: wire _NhanSu dropdown + lookup formulas onto every person row in the checklist"
```

---

## Self-Review Notes

- **Spec coverage:** Task 1 covers the approved plan's Phase 0. Tasks 2-5 cover Phase 1 (`_Tokens` sheet, `token_rules.py`, `excel_reader.py` wiring, `tokens.py` shim, real-sheet regression). Tasks 6-7 cover Phase 1.7 (`_NhanSu` sheet, dropdown/lookup wiring). Phases 2-8 of the approved plan are deliberately **out of scope** for this document — they get their own plan documents once this foundation lands, since several of their exact details (e.g. which literal strings remain in templates) should be re-verified against the state of the repo after this phase, not assumed now.
- **Placeholder scan:** every step has real, complete code; no "TODO"/"similar to Task N"/vague instructions remain.
- **Type/name consistency check:** `token_rules.resolve_tokens(wb, ws, index) -> dict[str, str]` (Task 3) is the same signature used in Task 3's own tests and referenced (unused directly, only via `ProjectInfo.common_tokens`) elsewhere. `ProjectInfo.common_tokens` (Task 3) is what `tokens.build_common_tokens` (Task 4) and the Task 5 regression test both read — consistent. `migrate_add_nhan_su_sheet.NHAN_SU_SHEET_NAME`/`add_nhan_su_sheet` (Task 6) are imported and called by name in `capnhat_nhan_su.py` (Task 7) exactly as defined.
