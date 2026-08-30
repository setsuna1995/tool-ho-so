# Multi-Page Expert Invitation Letters & Name-Based CV Matching Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the single-page, hand-addressed expert invitation letter with automatically-personalized, multi-page letters (one per external committee member, per council), add the missing nghiệm thu invitation letter, fix two confirmed hardcode gaps, and replace the manual "TÊN FILE CV" checklist column with automatic name-based CV matching.

**Architecture:** Two new standalone modules (`expert_invitation.py` for python-docx page-cloning, `cv_matching.py` for diacritics-insensitive filename matching) plug into existing `section_moi_chuyen_gia.py`/`tao_ho_so_moi.py` call sites. Checklist/template changes ride the project's existing one-off `migrate_*.py` script convention.

**Tech Stack:** Python 3, `python-docx` 1.2.0, `openpyxl`, `pywin32` (optional COM backend), `pytest`.

**Spec:** `.superpowers/sdd/2026-08-29-invitation-letter-and-cv-matching-design.md`

## Global Constraints

- Token syntax: `{{UPPER_SNAKE_CASE}}` (existing convention, unchanged).
- `expert_invitation.py` bypasses `word_writer.Session`/COM entirely and works directly with `python-docx`, regardless of the run's chosen backend — this is a deliberate, scoped exception (spec §3).
- CV name matching (`cv_matching.find_cv_file`) is diacritics-insensitive, case-insensitive, and requires the person's name to appear as a **contiguous word sequence, in order** inside the normalized filename (spec §7) — not just "all words present anywhere."
- No VBA/macros anywhere in the toolchain (existing project rule).
- Error messages are Vietnamese, in the existing tone (`"Không tìm thấy ... Vui lòng ..."`).
- All new one-off Excel migration scripts follow the exact `migrate_add_research_location.py`/`migrate_fix_f01_cv_filename.py` pattern: idempotent, operate on the real `Form checklist hồ sơ dự án.xlsx` in place, kept in the repo after running (not deleted).
- Every code file is UTF-8 with Vietnamese identifiers/strings exactly as given below — copy them verbatim, do not transliterate.

---

### Task 1: `cv_matching.py` — name-based CV file lookup

**Files:**
- Create: `cv_matching.py`
- Test: `test_cv_matching.py`

**Interfaces:**
- Produces: `find_cv_file(cv_dir: Path, person_name: str, context: str = "") -> Path`, raises `FileNotFoundError` on zero or multiple matches at whichever tier is used.

- [ ] **Step 1: Write the failing tests**

```python
# test_cv_matching.py
import pytest

import cv_matching


def _make_cv_dir(tmp_path, filenames):
    cv_dir = tmp_path / "CV chuyên gia"
    cv_dir.mkdir()
    for name in filenames:
        (cv_dir / name).write_text("x")
    return cv_dir


def test_finds_exact_match_with_diacritics(tmp_path):
    cv_dir = _make_cv_dir(tmp_path, ["Lý lịch khoa học - Trương Hồng Sơn.docx"])
    result = cv_matching.find_cv_file(cv_dir, "Trương Hồng Sơn")
    assert result.name == "Lý lịch khoa học - Trương Hồng Sơn.docx"


def test_falls_back_to_diacritics_insensitive_match_when_no_exact_match(tmp_path):
    cv_dir = _make_cv_dir(tmp_path, ["TM-Gs. Nguyen Cong Khan.pdf"])
    result = cv_matching.find_cv_file(cv_dir, "Nguyễn Công Khẩn")
    assert result.name == "TM-Gs. Nguyen Cong Khan.pdf"


def test_prefers_exact_diacritics_match_over_ambiguous_stripped_match(tmp_path):
    """Bo dau se khien 'Đặng Thị Bình' va 'Đăng Thị Bình' trung nhau (ca hai
    deu rut gon ve 'dang thi binh') - vong khop dung dau phai phan biet
    duoc 2 nguoi nay, khong duoc roi xuong vong bo dau va bao loi mo ho."""
    cv_dir = _make_cv_dir(tmp_path, ["CV Đặng Thị Bình.docx", "CV Đăng Thị Bình.docx"])
    result = cv_matching.find_cv_file(cv_dir, "Đặng Thị Bình")
    assert result.name == "CV Đặng Thị Bình.docx"


def test_raises_when_no_file_matches_at_either_tier(tmp_path):
    cv_dir = _make_cv_dir(tmp_path, ["TM-Gs. Nguyen Cong Khan.pdf"])
    with pytest.raises(FileNotFoundError, match="Lưu Liên Hương"):
        cv_matching.find_cv_file(cv_dir, "Lưu Liên Hương")


def test_raises_when_multiple_files_match_with_no_diacritics_to_disambiguate(tmp_path):
    cv_dir = _make_cv_dir(tmp_path, ["CV Nguyen Van A - ban 1.docx", "CV Nguyen Van A - ban 2.docx"])
    with pytest.raises(FileNotFoundError, match="Nguyễn Văn A"):
        cv_matching.find_cv_file(cv_dir, "Nguyễn Văn A")


def test_context_appears_in_error_message(tmp_path):
    cv_dir = _make_cv_dir(tmp_path, [])
    with pytest.raises(FileNotFoundError, match="chủ nhiệm đề tài"):
        cv_matching.find_cv_file(cv_dir, "Ai Đó", context=" (chủ nhiệm đề tài)")


def test_requires_contiguous_word_order_not_scrambled(tmp_path):
    cv_dir = _make_cv_dir(tmp_path, ["Van Khan Cong Nguyen.pdf"])
    with pytest.raises(FileNotFoundError):
        cv_matching.find_cv_file(cv_dir, "Nguyễn Công Khẩn")
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_cv_matching.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'cv_matching'`

- [ ] **Step 3: Write the implementation**

```python
# cv_matching.py
import re
import unicodedata
from pathlib import Path


def _normalize(text: str, strip_diacritics: bool) -> str:
    if strip_diacritics:
        text = text.replace("đ", "d").replace("Đ", "D")
        text = "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))
    return re.sub(r"[^\w]+", " ", text, flags=re.UNICODE).strip().lower()


def _search(cv_dir: Path, person_name: str, strip_diacritics: bool):
    target = _normalize(person_name, strip_diacritics)
    candidates = sorted((f for f in cv_dir.iterdir() if f.is_file()), key=lambda f: f.name)
    return [f for f in candidates if target in _normalize(f.stem, strip_diacritics)]


def _not_found_error(cv_dir: Path, person_name: str, context: str) -> FileNotFoundError:
    return FileNotFoundError(
        f"Không tìm thấy file CV nào khớp tên '{person_name}'{context} trong thư mục "
        f"'{cv_dir.name}/'. Vui lòng đặt file CV có tên chứa '{person_name}' vào thư mục đó."
    )


def _ambiguous_error(cv_dir: Path, person_name: str, context: str, matches) -> FileNotFoundError:
    names = ", ".join(f"'{m.name}'" for m in matches)
    return FileNotFoundError(
        f"Tìm thấy nhiều hơn 1 file CV khớp tên '{person_name}'{context} trong thư mục "
        f"'{cv_dir.name}/': {names}. Vui lòng đổi tên file để chỉ còn đúng 1 file khớp."
    )


def find_cv_file(cv_dir: Path, person_name: str, context: str = "") -> Path:
    """Tim file trong cv_dir co ten (khong ke phan mo rong) chua cum
    `person_name` LIEN NHAU, dung thu tu. Uu tien khop DUNG DAU truoc (chi
    chuan hoa hoa/thuong + khoang trang, giu nguyen dau tieng Viet) - chi
    khi khong file nao khop dung dau moi thu lai sau khi bo dau (de van
    khop duoc file dat ten khong dau nhu "TM-Gs. Nguyen Cong Khan.pdf").
    Neu vong khop dung dau ra >1 ket qua, bao loi mo ho ngay, khong roi
    xuong vong bo dau (vi bo dau se khong lam het mo ho, chi lam mo ho
    hon). Nem FileNotFoundError neu 0 hoac >1 file khop o vong duoc dung."""
    exact_matches = _search(cv_dir, person_name, strip_diacritics=False)
    if len(exact_matches) == 1:
        return exact_matches[0]
    if len(exact_matches) > 1:
        raise _ambiguous_error(cv_dir, person_name, context, exact_matches)

    loose_matches = _search(cv_dir, person_name, strip_diacritics=True)
    if not loose_matches:
        raise _not_found_error(cv_dir, person_name, context)
    if len(loose_matches) > 1:
        raise _ambiguous_error(cv_dir, person_name, context, loose_matches)
    return loose_matches[0]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_cv_matching.py -v`
Expected: 7 passed

- [ ] **Step 5: Commit**

```bash
git add cv_matching.py test_cv_matching.py
git commit -m "$(cat <<'EOF'
feat: add two-tier (exact-diacritics then diacritics-insensitive) CV name matching

Exact-diacritics pass runs first so two different names that would
collide once stripped (e.g. Đặng Thị Bình / Đăng Thị Bình) still
resolve correctly; the diacritics-stripped pass only kicks in when the
exact pass finds zero matches, for filenames with no Vietnamese
diacritics at all (e.g. TM-Gs. Nguyen Cong Khan.pdf).
EOF
)"
```

---

### Task 2: `committee_writer.py` — make `org_col` optional

**Files:**
- Modify: `committee_writer.py:7-29`
- Test: `test_committee_writer.py`

**Interfaces:**
- Consumes: nothing new.
- Produces: `write_committee_roster(..., org_col: Optional[int] = 2, ...)` — org cell is written only when `org_col is not None`.

- [ ] **Step 1: Write the failing test**

Add to `test_committee_writer.py`:

```python
def test_write_committee_roster_skips_org_column_when_org_col_is_none(tmp_path):
    src = _table_fixture(tmp_path, rows=3, cols=3)
    session = word_writer.Session(force_backend="docx")
    try:
        doc = session.open(src)
        committee = _sample_committee()
        committee_writer.write_committee_roster(
            session, doc, 1, committee,
            roles=["Chủ tịch Hội đồng", "Thành viên", "Thành viên"],
            name_col=1, org_col=None, role_col=None,
        )
        session.save_close(doc)
    finally:
        session.quit()

    check = docx.Document(str(src))
    table = check.tables[0]
    assert table.cell(0, 0).text.strip() == "GS.TS. Nguyễn Công Khẩn"
    assert table.cell(0, 1).text.strip() == ""
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest test_committee_writer.py::test_write_committee_roster_skips_org_column_when_org_col_is_none -v`
Expected: FAIL — `org_col` currently always writes, cell(0,1) would end up non-empty (`"Hội đồng Đạo đức Y sinh Quốc gia"`), assertion `== ""` fails.

- [ ] **Step 3: Update the implementation**

In `committee_writer.py`, change the `write_committee_roster` signature and body:

```python
def write_committee_roster(
    session: word_writer.Session,
    doc: word_writer.OpenDoc,
    table_index: int,
    committee: CommitteeData,
    roles: List[str],
    name_col: int = 1,
    org_col: Optional[int] = 2,
    role_col: Optional[int] = None,
    start_row: int = 1,
) -> None:
    people = [committee.chair] + committee.reviewers + committee.members
    if len(people) != len(roles):
        raise ValueError(
            f"Số thành viên hội đồng ({len(people)}) không khớp số vai trò truyền vào ({len(roles)})"
        )
    for offset, (person, role) in enumerate(zip(people, roles)):
        row = start_row + offset
        display_name = f"{person.degree} {person.name}".strip()
        session.set_cell(doc, table_index, row, name_col, display_name)
        if org_col is not None:
            session.set_cell(doc, table_index, row, org_col, person.org)
        if role_col is not None:
            session.set_cell(doc, table_index, row, role_col, role)
```

- [ ] **Step 4: Run all committee_writer tests to verify they pass**

Run: `pytest test_committee_writer.py -v`
Expected: 6 passed (5 existing + 1 new)

- [ ] **Step 5: Commit**

```bash
git add committee_writer.py test_committee_writer.py
git commit -m "$(cat <<'EOF'
feat: make write_committee_roster's org_col optional

Needed for tables that have no "Đơn vị" column (Phiếu ký nhận tiền.docx),
matching the existing Optional pattern already used by role_col.
EOF
)"
```

---

### Task 3: `expert_invitation.py` — multi-page letter generation

**Files:**
- Create: `expert_invitation.py`
- Test: `test_expert_invitation.py`

**Interfaces:**
- Consumes: `excel_reader.Person`, `word_writer._docx_replace_in_paragraph` (existing private helper, reused as-is).
- Produces: `generate_multi_page_letter(path: Path, recipients: list[Person], common_tokens: dict) -> bool`.

- [ ] **Step 1: Write the failing tests**

```python
# test_expert_invitation.py
import docx
from docx.oxml.ns import qn

import expert_invitation
from excel_reader import Person


def _make_template(tmp_path):
    path = tmp_path / "letter.docx"
    doc = docx.Document()
    doc.add_paragraph("Đề tài: {{TEN_DE_TAI}}")
    doc.add_paragraph("Kính gửi: {{CHUYEN_GIA_HO_TEN}}, {{CHUYEN_GIA_DON_VI}}.")
    doc.save(str(path))
    return path


def test_returns_false_and_leaves_file_untouched_when_no_recipients(tmp_path):
    path = _make_template(tmp_path)
    original_text = [p.text for p in docx.Document(str(path)).paragraphs]

    result = expert_invitation.generate_multi_page_letter(path, [], {"{{TEN_DE_TAI}}": "Đề tài X"})

    assert result is False
    assert [p.text for p in docx.Document(str(path)).paragraphs] == original_text


def test_generates_one_page_per_recipient_with_correct_names(tmp_path):
    path = _make_template(tmp_path)
    recipients = [
        Person("Hoàng Thị Thanh", "PGs. Ts.", "Hội đồng Đạo đức trong nghiên cứu y sinh học Quốc gia"),
        Person("Nguyễn Công Khẩn", "Gs. Ts.", "Hiệp hội Sữa Việt Nam"),
    ]

    result = expert_invitation.generate_multi_page_letter(path, recipients, {"{{TEN_DE_TAI}}": "Đề tài X"})

    assert result is True
    doc = docx.Document(str(path))
    full_text = [p.text for p in doc.paragraphs]
    assert full_text.count("Đề tài: Đề tài X") == 2
    assert "Kính gửi: PGs. Ts. Hoàng Thị Thanh, Hội đồng Đạo đức trong nghiên cứu y sinh học Quốc gia." in full_text
    assert "Kính gửi: Gs. Ts. Nguyễn Công Khẩn, Hiệp hội Sữa Việt Nam." in full_text
    assert "{{CHUYEN_GIA_HO_TEN}}" not in "\n".join(full_text)


def test_inserts_exactly_one_page_break_between_two_recipients(tmp_path):
    path = _make_template(tmp_path)  # template has 2 paragraphs
    recipients = [Person("Người Một", "", "Đơn vị 1"), Person("Người Hai", "", "Đơn vị 2")]

    expert_invitation.generate_multi_page_letter(path, recipients, {"{{TEN_DE_TAI}}": "X"})

    doc = docx.Document(str(path))
    assert len(doc.paragraphs) == 2 * 2 + 1  # 2 template paragraphs x 2 pages + 1 page-break paragraph


def test_sect_pr_remains_last_body_element(tmp_path):
    path = _make_template(tmp_path)
    recipients = [Person("Người Một", "", "Đơn vị 1")]

    expert_invitation.generate_multi_page_letter(path, recipients, {})

    doc = docx.Document(str(path))
    assert doc.element.body[-1].tag == qn("w:sectPr")
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_expert_invitation.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'expert_invitation'`

- [ ] **Step 3: Write the implementation**

```python
# expert_invitation.py
import copy
from pathlib import Path
from typing import List

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from excel_reader import Person
from word_writer import _docx_replace_in_paragraph


def _wrap_element(element, doc):
    if element.tag == qn("w:p"):
        return Paragraph(element, doc)
    if element.tag == qn("w:tbl"):
        return Table(element, doc)
    return None


def _apply_tokens_to_elements(elements, doc, tokens: dict) -> None:
    for element in elements:
        wrapped = _wrap_element(element, doc)
        if wrapped is None:
            continue
        if isinstance(wrapped, Paragraph):
            for find, value in tokens.items():
                _docx_replace_in_paragraph(wrapped, find, value)
        else:
            for row in wrapped.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for find, value in tokens.items():
                            _docx_replace_in_paragraph(paragraph, find, value)


def _page_break_element():
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    return p


def generate_multi_page_letter(path: Path, recipients: List[Person], common_tokens: dict) -> bool:
    """Mo file .docx tai `path` (da duoc copy san tu thu muc MAU), nhan ban
    toan bo noi dung than file (tru sectPr, phai luon la phan tu cuoi cung
    theo chuan OOXML) thanh len(recipients) trang - moi trang danh cho 1
    nguoi, dien ca common_tokens lan token rieng trang
    ({{CHUYEN_GIA_HO_TEN}}, {{CHUYEN_GIA_DON_VI}}) - roi ghi de luu lai
    dung `path`.

    Tra ve False (khong dong/luu gi ca, giu nguyen file mau chua dien) neu
    `recipients` rong.
    """
    if not recipients:
        return False

    doc = docx.Document(str(path))
    body = doc.element.body

    sect_pr = body.find(qn("w:sectPr"))
    template_elements = [el for el in list(body) if el is not sect_pr]
    for element in template_elements:
        body.remove(element)

    def _insert(element):
        if sect_pr is not None:
            sect_pr.addprevious(element)
        else:
            body.append(element)

    for index, person in enumerate(recipients):
        if index > 0:
            _insert(_page_break_element())

        cloned_elements = []
        for element in template_elements:
            clone = copy.deepcopy(element)
            _insert(clone)
            cloned_elements.append(clone)

        page_tokens = {
            "{{CHUYEN_GIA_HO_TEN}}": f"{person.degree} {person.name}".strip(),
            "{{CHUYEN_GIA_DON_VI}}": person.org,
        }
        _apply_tokens_to_elements(cloned_elements, doc, {**common_tokens, **page_tokens})

    doc.save(str(path))
    return True
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_expert_invitation.py -v`
Expected: 4 passed

- [ ] **Step 5: Commit**

```bash
git add expert_invitation.py test_expert_invitation.py
git commit -m "$(cat <<'EOF'
feat: add multi-page expert invitation letter generation

Clones a template's body N times (one per recipient) with page breaks,
filling per-page recipient tokens plus the usual common tokens. Bypasses
Session/COM deliberately (see spec §3) since this is structural XML
cloning, not search-replace.
EOF
)"
```

---

### Task 4: `excel_reader.py` — CV declarations keyed by name, not filename

**Files:**
- Modify: `excel_reader.py:33-38,49-67,129-148,151-207`
- Modify: `test_excel_reader.py`
- Delete: `migrate_fix_f01_cv_filename.py`, `test_migrate_fix_f01_cv_filename.py` (their sole purpose — fixing F01's now-unused filename column — no longer applies once `head_cv_filename` is removed)

**Interfaces:**
- Consumes: nothing new.
- Produces: `ExpertCvEntry(code, name, role)` (drops `filename`); `ProjectInfo` drops `head_cv_filename`; `read_expert_cvs` includes a row when its **name** (column 3) is non-blank.

- [ ] **Step 1: Write the failing tests**

In `test_excel_reader.py`:
1. Delete `test_head_cv_filename_is_read_from_f01`, `test_missing_f01_cv_filename_raises_value_error`, `test_real_checklist_head_cv_filename`.
2. Replace `test_read_expert_cvs_skips_blank_filename_rows` and `test_read_expert_cvs_reads_declared_rows` with:

```python
def test_read_expert_cvs_skips_blank_name_rows(tmp_path):
    path = _build_minimal_workbook(tmp_path, overrides={"F03": (None, None, None)})
    data = excel_reader.load_project_data(path, "Test")
    assert data.expert_cvs == []


def test_read_expert_cvs_reads_declared_rows_by_name(tmp_path):
    path = _build_minimal_workbook(
        tmp_path,
        overrides={
            "F03": ("Thư ký C", "Thư ký Đề tài", None),
            "F04": ("Chuyên gia D", "Ủy viên", None),
        },
    )
    data = excel_reader.load_project_data(path, "Test")
    codes = {e.code: e for e in data.expert_cvs}
    assert set(codes) == {"F03", "F04"}
    assert codes["F03"].name == "Thư ký C"
    assert codes["F03"].role == "Thư ký Đề tài"
```

`test_real_checklist_expert_cvs_has_at_least_one_entry` stays unchanged (still valid — doesn't touch `filename`).

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_excel_reader.py -v`
Expected: FAIL — `AttributeError: 'ExpertCvEntry' object has no attribute 'filename'` does not yet apply (field still exists), but the two new tests fail because rows are still filtered on filename, not name (`F03`/`F04` overrides leave filename blank so they're currently skipped regardless of name).

- [ ] **Step 3: Update the implementation**

In `excel_reader.py`, change `ExpertCvEntry`:

```python
@dataclass
class ExpertCvEntry:
    code: str
    name: str
    role: str
```

Change `ProjectInfo` — remove the `head_cv_filename: str` field (keep field order for the rest unchanged).

Change `read_expert_cvs`:

```python
def read_expert_cvs(ws, index: dict) -> List[ExpertCvEntry]:
    """Doc PHAN F, ma F02-F10 (F01 la chu nhiem, CV cua chu nhiem khop truc
    tiep qua info.head.name, khong doc rieng F01 nua). Dong nao ten (cot 3)
    de trong thi bo qua - khong con phu thuoc vao cot filename nua."""
    entries = []
    for i in range(2, 11):
        code = f"F{i:02d}"
        row = index.get(code)
        if row is None:
            continue
        name = _cell_text(ws, row, 3)
        if not name:
            continue
        entries.append(ExpertCvEntry(code=code, name=name, role=_cell_text(ws, row, 4)))
    return entries
```

In `load_project_data`, delete the two lines:

```python
    head_cv_filename = _cell_text(ws, index["F01"], 5) if "F01" in index else ""
    if not head_cv_filename:
        raise ValueError("Tên file CV của chủ nhiệm đề tài (F01) là bắt buộc nhưng đang trống")
```

and delete `head_cv_filename=head_cv_filename,` from the returned `ProjectInfo(...)` call.

Delete `migrate_fix_f01_cv_filename.py` and `test_migrate_fix_f01_cv_filename.py` entirely (their target field no longer exists).

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_excel_reader.py -v`
Expected: all pass (2 new, minus 3 deleted)

Run: `pytest -v` (full suite) and confirm the only remaining failures are in `test_tao_ho_so_moi.py` (fixed in Task 6) and `test_section_moi_chuyen_gia.py` (fixed in Task 9) — both currently reference `head_cv_filename`/`entry.filename`.

- [ ] **Step 5: Commit**

```bash
git add excel_reader.py test_excel_reader.py
git rm migrate_fix_f01_cv_filename.py test_migrate_fix_f01_cv_filename.py
git commit -m "$(cat <<'EOF'
feat: read PHẦN F CV declarations by name instead of filename

ExpertCvEntry drops its filename field; ProjectInfo drops
head_cv_filename entirely. Actual file resolution moves to
cv_matching.find_cv_file, called from tao_ho_so_moi.py (next commit).
Removes migrate_fix_f01_cv_filename.py, whose sole purpose (fixing F01's
now-unused filename cell) no longer applies.
EOF
)"
```

---

### Task 5: Fix live checklist F03 CV gap

**Files:**
- Create: `migrate_clear_f03_missing_cv.py`
- Test: `test_migrate_clear_f03_missing_cv.py`

**Context:** the live `Form checklist hồ sơ dự án.xlsx`, sheet `Đề tài - Bánh ăn dặm VIAM 2027`, row F03 declares "Lưu Liên Hương" (Thư ký Đề tài) but `CV chuyên gia/` has no file matching her name. Under the old filename-column system this was silently skipped (blank filename cell); under the new name-keyed system (Task 4) her declared name alone is now enough to require a resolvable CV, so `cv_matching.find_cv_file` will raise for this row once Task 6 wires it in. This is a real, pre-existing data gap the audit surfaced (spec §7) — resolve it by clearing the declaration rather than fabricating a CV file. **This edits real project data, not a template** — the script's docstring records exactly why and how to reverse it.

**Interfaces:**
- Produces: `clear_f03_pending_cv() -> None`.

- [ ] **Step 1: Write the failing tests**

```python
# test_migrate_clear_f03_missing_cv.py
import migrate_clear_f03_missing_cv as migrate
from excel_reader import load_project_data

CHECKLIST_PATH = migrate.CHECKLIST_PATH
SHEET_VIAM = migrate.SHEET_VIAM


def test_f03_no_longer_declared_after_clearing():
    migrate.clear_f03_pending_cv()

    data = load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    assert "F03" not in {e.code for e in data.expert_cvs}


def test_running_migration_twice_is_safe():
    migrate.clear_f03_pending_cv()
    migrate.clear_f03_pending_cv()

    data = load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    assert "F03" not in {e.code for e in data.expert_cvs}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_migrate_clear_f03_missing_cv.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'migrate_clear_f03_missing_cv'`

- [ ] **Step 3: Write the implementation**

```python
# migrate_clear_f03_missing_cv.py
from pathlib import Path

import openpyxl

CHECKLIST_PATH = Path(__file__).resolve().parent / "Form checklist hồ sơ dự án.xlsx"
SHEET_VIAM = "Đề tài - Bánh ăn dặm VIAM 2027"


def _find_row(ws, code: str) -> int:
    for row in ws.iter_rows(min_row=5, max_col=1):
        if row[0].value == code:
            return row[0].row
    raise ValueError(f"Không tìm thấy mã mục '{code}' trong sheet '{ws.title}'")


def clear_f03_pending_cv() -> None:
    """F03 (Luu Lien Huong, Thu ky De tai) da khai ten nhung chua co file CV
    khop trong 'CV chuyen gia/' - xoa ten/vai tro de tranh loi
    FileNotFoundError khi sinh ho so (spec §7). De khai lai: dat file CV
    that cua co vao 'CV chuyen gia/', roi go lai ten + vai tro vao F03
    trong sheet nay."""
    wb = openpyxl.load_workbook(CHECKLIST_PATH)
    ws = wb[SHEET_VIAM]
    row = _find_row(ws, "F03")
    if ws.cell(row=row, column=3).value is None:
        return
    ws.cell(row=row, column=3, value=None)
    ws.cell(row=row, column=4, value=None)
    wb.save(CHECKLIST_PATH)


if __name__ == "__main__":
    clear_f03_pending_cv()
    print("Da xoa ten F03 (chua co file CV khop) khoi sheet du an VIAM 2027.")
```

- [ ] **Step 4: Run tests to verify they pass, then run the migration**

Run: `pytest test_migrate_clear_f03_missing_cv.py -v`
Expected: 2 passed

Run: `python migrate_clear_f03_missing_cv.py`
Expected: prints `Da xoa ten F03 (chua co file CV khop) khoi sheet du an VIAM 2027.`

- [ ] **Step 5: Commit**

```bash
git add migrate_clear_f03_missing_cv.py test_migrate_clear_f03_missing_cv.py "Form checklist hồ sơ dự án.xlsx"
git commit -m "$(cat <<'EOF'
fix: clear F03's undeclared CV name in the live VIAM 2027 checklist

Lưu Liên Hương's name was declared with no matching file in
CV chuyên gia/ - silently skipped under the old filename-column system,
but the new name-keyed CV matching (this branch) requires it to resolve.
Re-declare her once a real CV file is added, per the script's docstring.
EOF
)"
```

---

### Task 6: `tao_ho_so_moi.py` — resolve CVs via name matching

**Files:**
- Modify: `tao_ho_so_moi.py:1-18,65-92`
- Modify: `test_tao_ho_so_moi.py`

**Interfaces:**
- Consumes: `cv_matching.find_cv_file` (Task 1), `ExpertCvEntry` without `filename` (Task 4).
- Produces: `copy_head_cv`/`copy_expert_cvs` unchanged signatures, new resolution behavior.

- [ ] **Step 1: Write the failing tests**

In `test_tao_ho_so_moi.py`, add `import cv_matching` at the top, then replace the CV-related tests:

```python
def test_copy_head_cv_copies_file_matching_head_name(tmp_path):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    expected = cv_matching.find_cv_file(root / "CV chuyên gia", info.head.name)

    tao_ho_so_moi.copy_head_cv(root, tmp_path, info)

    assert (tmp_path / "01. Hồ sơ đạo đức đề cương" / expected.name).exists()


def test_copy_head_cv_raises_clear_error_when_no_cv_matches_head_name(tmp_path):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    bad_info = dataclasses.replace(info, head=dataclasses.replace(info.head, name="Người Không Tồn Tại"))

    with pytest.raises(FileNotFoundError):
        tao_ho_so_moi.copy_head_cv(root, tmp_path, bad_info)


def test_copy_expert_cvs_copies_every_declared_entry(tmp_path):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)

    tao_ho_so_moi.copy_expert_cvs(root, tmp_path, info)

    dest_dir = tmp_path / "03. Công văn mời chuyên gia"
    for entry in info.expert_cvs:
        expected = cv_matching.find_cv_file(root / "CV chuyên gia", entry.name)
        assert (dest_dir / expected.name).exists()


def test_copy_expert_cvs_raises_clear_error_when_no_cv_matches(tmp_path):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    good_entry = info.expert_cvs[0]
    bad_entry = dataclasses.replace(info.expert_cvs[1], name="Người Không Tồn Tại")
    bad_info = dataclasses.replace(info, expert_cvs=[good_entry, bad_entry])

    with pytest.raises(FileNotFoundError):
        tao_ho_so_moi.copy_expert_cvs(root, tmp_path, bad_info)

    assert not (tmp_path / "03. Công văn mời chuyên gia").exists()
```

Keep `test_copy_expert_cvs_does_nothing_when_list_is_empty` unchanged. Replace `test_generate_all_copies_expert_cvs_into_output` and `test_generate_all_keeps_staging_dir_and_reports_path_on_failure`:

```python
def test_generate_all_copies_expert_cvs_into_output(tmp_path):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    dest_root = tmp_path / "Hồ sơ output"

    session = word_writer.Session(force_backend="docx")
    try:
        tao_ho_so_moi.generate_all(root, dest_root, info, session)
    finally:
        session.quit()

    assert info.expert_cvs
    for entry in info.expert_cvs:
        expected = cv_matching.find_cv_file(root / "CV chuyên gia", entry.name)
        assert (dest_root / "03. Công văn mời chuyên gia" / expected.name).exists()


def test_generate_all_keeps_staging_dir_and_reports_path_on_failure(tmp_path, monkeypatch, capsys):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    bad_info = dataclasses.replace(info, head=dataclasses.replace(info.head, name="Người Không Tồn Tại"))
    dest_root = tmp_path / "Hồ sơ output"
    staging_dir = tmp_path / "staging"

    def fake_mkdtemp(prefix=None):
        staging_dir.mkdir(exist_ok=True)
        return str(staging_dir)

    monkeypatch.setattr(tao_ho_so_moi.tempfile, "mkdtemp", fake_mkdtemp)

    session = word_writer.Session(force_backend="docx")
    try:
        with pytest.raises(FileNotFoundError):
            tao_ho_so_moi.generate_all(root, dest_root, bad_info, session)
    finally:
        session.quit()

    assert staging_dir.exists()
    captured = capsys.readouterr()
    assert str(staging_dir) in captured.out
    assert not dest_root.exists()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_tao_ho_so_moi.py -v`
Expected: FAIL — `copy_head_cv`/`copy_expert_cvs` still read `info.head_cv_filename`/`entry.filename`, which no longer exist (`AttributeError`).

- [ ] **Step 3: Update the implementation**

In `tao_ho_so_moi.py`, add `import cv_matching` near the other imports, then replace both functions:

```python
def copy_head_cv(root: Path, dest_root: Path, info) -> None:
    cv_dir = root / "CV chuyên gia"
    src = cv_matching.find_cv_file(cv_dir, info.head.name, context=" (chủ nhiệm đề tài)")
    dst = dest_root / "01. Hồ sơ đạo đức đề cương" / src.name
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_bytes(src.read_bytes())


def copy_expert_cvs(root: Path, dest_root: Path, info) -> None:
    cv_dir = root / "CV chuyên gia"
    resolved = [
        cv_matching.find_cv_file(cv_dir, entry.name, context=f" (mã mục {entry.code} - {entry.role})")
        for entry in info.expert_cvs
    ]

    for src in resolved:
        dst = dest_root / "03. Công văn mời chuyên gia" / src.name
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_bytes(src.read_bytes())
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_tao_ho_so_moi.py -v`
Expected: all pass

- [ ] **Step 5: Commit**

```bash
git add tao_ho_so_moi.py test_tao_ho_so_moi.py
git commit -m "$(cat <<'EOF'
feat: resolve CV attachments via name matching instead of filename column
EOF
)"
```

---

### Task 7: `section_nghiem_thu.py` — wire missing `Phiếu ký nhận tiền.docx` rows

**Files:**
- Modify: `section_nghiem_thu.py:86-91`
- Modify: `test_section_nghiem_thu.py`

**Interfaces:**
- Consumes: `committee_writer.write_committee_roster` with `org_col=None` (Task 2).

- [ ] **Step 1: Write the failing test**

Add to `test_section_nghiem_thu.py`:

```python
def test_generate_writes_acceptance_committee_roster_in_payment_slip(dest_dir, info):
    session = word_writer.Session(force_backend="docx")
    try:
        section_nghiem_thu.generate(session, dest_dir, info, tokens.build_common_tokens(info))
    finally:
        session.quit()

    doc = docx.Document(str(dest_dir / "Phiếu ký nhận tiền.docx"))
    table = doc.tables[1]
    committee = info.acceptance_committee
    people = [committee.chair] + committee.reviewers + committee.members
    for offset, person in enumerate(people):
        expected = f"{person.degree} {person.name}".strip()
        assert table.cell(1 + offset, 1).text.strip() == expected
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest test_section_nghiem_thu.py::test_generate_writes_acceptance_committee_roster_in_payment_slip -v`
Expected: FAIL — cells 1-5 still hold the old sample project's committee names.

- [ ] **Step 3: Update the implementation**

In `section_nghiem_thu.py`, update `_phieu_ky_nhan_tien`:

```python
def _phieu_ky_nhan_tien(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Phiếu ký nhận tiền.docx")
    session.fill_tokens(doc, common_tokens)
    committee_writer.write_committee_roster(
        session, doc, 2, info.acceptance_committee, roles=ROLES,
        name_col=2, org_col=None, role_col=None, start_row=2,
    )
    secretary = info.acceptance_committee.secretaries[0]
    session.set_cell(doc, 2, 7, 2, secretary.name)
    session.save_close(doc)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_section_nghiem_thu.py -v`
Expected: all pass (existing `test_generate_writes_secretary_name_without_degree_in_payment_slip` must still pass unchanged)

- [ ] **Step 5: Commit**

```bash
git add section_nghiem_thu.py test_section_nghiem_thu.py
git commit -m "$(cat <<'EOF'
fix: wire acceptance committee roster into Phiếu ký nhận tiền.docx

Only the secretary's row was filled; the other 5 signer rows (chair +
reviewers + members) still showed the old sample project's names.
EOF
)"
```

---

### Task 8: Invitation letter template content — tokenize + create nghiệm thu variant

**Files:**
- Create: `migrate_tokenize_invitation_letters.py`
- Test: `test_migrate_tokenize_invitation_letters.py`
- Modify (via script, not by hand): `03. Công văn mời chuyên gia - MẪU/Công văn mời chuyên gia.docx`
- Create (via script): `03. Công văn mời chuyên gia - MẪU/Công văn mời chuyên gia nghiệm thu.docx`

**Interfaces:**
- Consumes: `word_writer._docx_replace_in_paragraph`.
- Produces: both `.docx` masters contain `{{CHUYEN_GIA_HO_TEN}}`, `{{CHUYEN_GIA_DON_VI}}`, `{{DAU_MOI_LIEN_HE}}` in place of the old dotted blanks / static contact line / old-sample-project intro text.

- [ ] **Step 1: Write the failing tests**

```python
# test_migrate_tokenize_invitation_letters.py
import docx

import migrate_tokenize_invitation_letters as migrate


def test_de_cuong_letter_has_no_leftover_sample_text():
    migrate.tokenize_de_cuong_letter()

    doc = docx.Document(str(migrate.DE_CUONG_PATH))
    full_text = "\n".join(p.text for p in doc.paragraphs)
    assert "LOF KUN COLOSTRUM" not in full_text
    assert "Lê Minh Khánh" not in full_text
    assert "{{CHUYEN_GIA_HO_TEN}}" in full_text
    assert "{{CHUYEN_GIA_DON_VI}}" in full_text
    assert "{{DAU_MOI_LIEN_HE}}" in full_text


def test_nghiem_thu_letter_created_with_nghiem_thu_wording():
    migrate.tokenize_de_cuong_letter()
    migrate.create_nghiem_thu_letter()

    doc = docx.Document(str(migrate.NGHIEM_THU_PATH))
    full_text = "\n".join(p.text for p in doc.paragraphs)
    assert "NGHIỆM THU" in full_text
    assert "HỘI ĐỒNG KHOA HỌC VÀ HỘI ĐỒNG ĐẠO ĐỨC" not in full_text
    assert "{{CHUYEN_GIA_HO_TEN}}" in full_text
    assert "{{DAU_MOI_LIEN_HE}}" in full_text


def test_de_cuong_tokenization_is_idempotent():
    migrate.tokenize_de_cuong_letter()
    migrate.tokenize_de_cuong_letter()

    doc = docx.Document(str(migrate.DE_CUONG_PATH))
    full_text = "\n".join(p.text for p in doc.paragraphs)
    assert full_text.count("{{CHUYEN_GIA_HO_TEN}}") == 3
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_migrate_tokenize_invitation_letters.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'migrate_tokenize_invitation_letters'`

- [ ] **Step 3: Write the implementation**

```python
# migrate_tokenize_invitation_letters.py
import docx

import paths
from word_writer import _docx_replace_in_paragraph

MAU_DIR = paths.project_root() / "03. Công văn mời chuyên gia - MẪU"
DE_CUONG_PATH = MAU_DIR / "Công văn mời chuyên gia.docx"
NGHIEM_THU_PATH = MAU_DIR / "Công văn mời chuyên gia nghiệm thu.docx"

DE_CUONG_REPLACEMENTS = [
    (
        "Suy dinh dưỡng ở trẻ em dưới 5 tuổi – đặc biệt là suy dinh dưỡng thấp còi vẫn là một vấn đề sức khỏe cộng đồng. "
        "Một trong những giải pháp làm giảm tình trạng suy dinh dưỡng ở trẻ dưới 5 tuổi là sử dụng các sản phẩm bổ sung dinh dưỡng trong hệ thống trường mầm non. "
        "Nhằm đánh giá tình trạng suy dinh dưỡng ở trẻ dưới 5 tuổi và hiệu quả của sản phẩm bổ sung dinh dưỡng LOF KUN COLOSTRUM, Viện Y học ứng dụng Việt Nam tiến hành triển khai nghiên cứu",
        "[Bổ sung bối cảnh/lý do triển khai dự án tại đây] {{DON_VI_CHU_TRI}} tiến hành triển khai đề tài",
    ),
    (
        "Viện Y học ứng dụng Việt Nam trân trọng kính mời .................................................. tham gia Hội đồng khoa học xét duyệt đề cương và Hội đồng đạo đức nghiên cứu.",
        "{{DON_VI_CHU_TRI}} trân trọng kính mời {{CHUYEN_GIA_HO_TEN}} tham gia Hội đồng khoa học xét duyệt đề cương và Hội đồng đạo đức nghiên cứu.",
    ),
    (
        "Xin gửi hồ sơ nghiên cứu để .............................................................. đọc và cho ý kiến đóng góp, phản biện.",
        "Xin gửi hồ sơ nghiên cứu để {{CHUYEN_GIA_HO_TEN}} đọc và cho ý kiến đóng góp, phản biện.",
    ),
    (
        "Mọi thông tin chi tiết xin vui lòng liên hệ: Ông Lê Minh Khánh, Trung tâm nghiên cứu - Viện Y học ứng dụng Việt Nam (Email: leminhkhanh@viam.vn - Điện thoại: 096.3355.652). ",
        "Mọi thông tin chi tiết xin vui lòng liên hệ: {{DAU_MOI_LIEN_HE}}.",
    ),
]

NGHIEM_THU_REPLACEMENTS = [
    ("THAM GIA HỘI ĐỒNG KHOA HỌC VÀ HỘI ĐỒNG ĐẠO ĐỨC", "THAM GIA HỘI ĐỒNG NGHIỆM THU ĐỀ TÀI"),
    (
        "V/v: Mời chuyên gia tham gia Hội đồng khoa học và Hội đồng đạo đức",
        "V/v: Mời chuyên gia tham gia Hội đồng nghiệm thu đề tài",
    ),
    (
        "{{DON_VI_CHU_TRI}} trân trọng kính mời {{CHUYEN_GIA_HO_TEN}} tham gia Hội đồng khoa học xét duyệt đề cương và Hội đồng đạo đức nghiên cứu.",
        "{{DON_VI_CHU_TRI}} trân trọng kính mời {{CHUYEN_GIA_HO_TEN}} tham gia Hội đồng nghiệm thu đề tài.",
    ),
    (
        "Xin gửi hồ sơ nghiên cứu để {{CHUYEN_GIA_HO_TEN}} đọc và cho ý kiến đóng góp, phản biện.",
        "Xin gửi hồ sơ nghiệm thu để {{CHUYEN_GIA_HO_TEN}} đọc và cho ý kiến đánh giá, nghiệm thu.",
    ),
]

DOTS_REPLACEMENT = "{{CHUYEN_GIA_HO_TEN}}, {{CHUYEN_GIA_DON_VI}}."


def _is_dots_placeholder(text: str) -> bool:
    stripped = text.strip()
    return len(stripped) > 20 and set(stripped) <= {"."}


def _apply_replacements(doc, replacements) -> None:
    for paragraph in doc.paragraphs:
        for find, replace in replacements:
            _docx_replace_in_paragraph(paragraph, find, replace)
        if _is_dots_placeholder(paragraph.text):
            _docx_replace_in_paragraph(paragraph, paragraph.text, DOTS_REPLACEMENT)


def tokenize_de_cuong_letter() -> None:
    doc = docx.Document(str(DE_CUONG_PATH))
    _apply_replacements(doc, DE_CUONG_REPLACEMENTS)
    doc.save(str(DE_CUONG_PATH))


def create_nghiem_thu_letter() -> None:
    doc = docx.Document(str(DE_CUONG_PATH))
    _apply_replacements(doc, NGHIEM_THU_REPLACEMENTS)
    doc.save(str(NGHIEM_THU_PATH))


if __name__ == "__main__":
    tokenize_de_cuong_letter()
    create_nghiem_thu_letter()
    print("Da token hoa thu moi de cuong va tao file thu moi nghiem thu.")
```

- [ ] **Step 4: Run tests, then run the script for real, then visually spot-check**

Run: `pytest test_migrate_tokenize_invitation_letters.py -v`
Expected: 3 passed

Run: `python migrate_tokenize_invitation_letters.py`
Expected: prints `Da token hoa thu moi de cuong va tao file thu moi nghiem thu.`; creates the new `Công văn mời chuyên gia nghiệm thu.docx` file.

Open both `.docx` files in Word if available. Without Word, dump their text instead:

```bash
python -c "
import docx
for path in ['03. Công văn mời chuyên gia - MẪU/Công văn mời chuyên gia.docx',
             '03. Công văn mời chuyên gia - MẪU/Công văn mời chuyên gia nghiệm thu.docx']:
    print('=' * 20, path)
    doc = docx.Document(path)
    for p in doc.paragraphs:
        if p.text.strip():
            print(p.text)
"
```

Confirm every `{{...}}` token prints intact (e.g. `{{CHUYEN_GIA_HO_TEN}}`, not `{{CHUYEN_GIA_` + `HO_TEN}}` split across two visually-adjacent fragments) — per the known python-docx run-splitting risk documented in `HUONG_DAN.md`.

- [ ] **Step 5: Commit**

```bash
git add migrate_tokenize_invitation_letters.py test_migrate_tokenize_invitation_letters.py \
  "03. Công văn mời chuyên gia - MẪU/Công văn mời chuyên gia.docx" \
  "03. Công văn mời chuyên gia - MẪU/Công văn mời chuyên gia nghiệm thu.docx"
git commit -m "$(cat <<'EOF'
feat: tokenize invitation letter, add nghiệm thu invitation letter master

Replaces the old sample project's leftover intro text and static
contact line with tokens, and replaces the three dotted recipient
blanks with {{CHUYEN_GIA_HO_TEN}}/{{CHUYEN_GIA_DON_VI}} for per-page
auto-addressing. Adds the missing nghiệm thu invitation letter master,
auto-discovered by template_config.discover_copies like any other
- MẪU file.
EOF
)"
```

---

### Task 9: `section_moi_chuyen_gia.py` — filtering, dedup, and both letters

**Files:**
- Modify: `section_moi_chuyen_gia.py` (full rewrite)
- Modify: `test_section_moi_chuyen_gia.py` (full rewrite)

**Interfaces:**
- Consumes: `expert_invitation.generate_multi_page_letter` (Task 3), tokenized templates (Task 8).
- Produces: `generate(session, dest_dir, info, common_tokens) -> None` (signature unchanged, for `tao_ho_so_moi.py` compatibility); `_external_members(committee, host_org) -> List[Person]`, `_dedupe_people(people) -> List[Person]` (module-private, used directly by tests).

- [ ] **Step 1: Write the failing tests**

Replace the full contents of `test_section_moi_chuyen_gia.py`:

```python
import shutil

import docx
import pytest

import excel_reader
import paths
import section_moi_chuyen_gia
import tokens
import word_writer
from excel_reader import CommitteeData, Person

CHECKLIST_PATH = paths.project_root() / excel_reader.CHECKLIST_FILENAME
SHEET_VIAM = "Đề tài - Bánh ăn dặm VIAM 2027"

MAU_DIR = paths.project_root() / "03. Công văn mời chuyên gia - MẪU"
FILES = ["Công văn mời chuyên gia.docx", "Công văn mời chuyên gia nghiệm thu.docx"]


@pytest.fixture()
def dest_dir(tmp_path):
    for filename in FILES:
        shutil.copy2(MAU_DIR / filename, tmp_path / filename)
    return tmp_path


@pytest.fixture()
def info():
    return excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)


def test_de_cuong_letter_has_one_page_per_external_member(dest_dir, info):
    session = word_writer.Session(force_backend="docx")
    try:
        section_moi_chuyen_gia.generate(session, dest_dir, info, tokens.build_common_tokens(info))
    finally:
        session.quit()

    expected = section_moi_chuyen_gia._dedupe_people(
        section_moi_chuyen_gia._external_members(info.ethics_committee, info.host_org)
        + section_moi_chuyen_gia._external_members(info.proposal_committee, info.host_org)
    )
    doc = docx.Document(str(dest_dir / "Công văn mời chuyên gia.docx"))
    full_text = "\n".join(p.text for p in doc.paragraphs)
    for person in expected:
        assert f"{person.degree} {person.name}".strip() in full_text
    assert "{{CHUYEN_GIA_HO_TEN}}" not in full_text


def test_nghiem_thu_letter_only_pulls_from_acceptance_committee(dest_dir, info):
    session = word_writer.Session(force_backend="docx")
    try:
        section_moi_chuyen_gia.generate(session, dest_dir, info, tokens.build_common_tokens(info))
    finally:
        session.quit()

    doc = docx.Document(str(dest_dir / "Công văn mời chuyên gia nghiệm thu.docx"))
    full_text = "\n".join(p.text for p in doc.paragraphs)
    expected = section_moi_chuyen_gia._external_members(info.acceptance_committee, info.host_org)
    for person in expected:
        assert f"{person.degree} {person.name}".strip() in full_text


def test_external_members_excludes_host_org_and_secretaries():
    committee = CommitteeData(
        chair=Person("Chủ tịch Ngoài", "PGS.", "Hội đồng ngoài"),
        reviewers=[Person("Phản biện VIAM", "TS.", "Viện Y học ứng dụng Việt Nam")],
        members=[Person("Ủy viên Ngoài", "ThS.", "Đơn vị khác")],
        secretaries=[Person("Thư ký VIAM", "CN.", "Ngoài")],
    )
    result = section_moi_chuyen_gia._external_members(committee, "Viện Y học ứng dụng Việt Nam")
    names = {p.name for p in result}
    assert names == {"Chủ tịch Ngoài", "Ủy viên Ngoài"}
    assert "Thư ký VIAM" not in names


def test_dedupe_people_collapses_same_person_across_committees():
    person_a = Person("Người A", "TS.", "Đơn vị A")
    person_a_dup = Person("người a", "TS.", "đơn vị a")
    person_b = Person("Người B", "TS.", "Đơn vị B")
    result = section_moi_chuyen_gia._dedupe_people([person_a, person_b, person_a_dup])
    assert result == [person_a, person_b]


def test_generate_prints_notice_when_no_external_members(tmp_path, info, capsys):
    import dataclasses

    for filename in FILES:
        shutil.copy2(MAU_DIR / filename, tmp_path / filename)

    all_host_committee = CommitteeData(
        chair=Person("Chủ tịch VIAM", "TS.", info.host_org),
        secretaries=[Person("Thư ký VIAM", "CN.", info.host_org)],
    )
    no_external_info = dataclasses.replace(
        info,
        ethics_committee=all_host_committee,
        proposal_committee=all_host_committee,
        acceptance_committee=all_host_committee,
    )

    session = word_writer.Session(force_backend="docx")
    try:
        section_moi_chuyen_gia.generate(
            session, tmp_path, no_external_info, tokens.build_common_tokens(no_external_info)
        )
    finally:
        session.quit()

    captured = capsys.readouterr()
    assert "bo qua dien thu moi de cuong" in captured.out
    assert "bo qua dien thu moi nghiem thu" in captured.out
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_section_moi_chuyen_gia.py -v`
Expected: FAIL — `section_moi_chuyen_gia` has no `_external_members`/`_dedupe_people`, and `generate` still writes the old single dotted-blank letter.

- [ ] **Step 3: Write the implementation**

Replace the full contents of `section_moi_chuyen_gia.py`:

```python
from pathlib import Path
from typing import List

import expert_invitation
import word_writer
from excel_reader import CommitteeData, Person, ProjectInfo


def _external_members(committee: CommitteeData, host_org: str) -> List[Person]:
    host_org_norm = host_org.strip().lower()
    candidates = [committee.chair] + committee.reviewers + committee.members
    return [p for p in candidates if p.org.strip().lower() != host_org_norm]


def _dedupe_people(people: List[Person]) -> List[Person]:
    seen = set()
    result = []
    for person in people:
        key = (person.name.strip().lower(), person.org.strip().lower())
        if key in seen:
            continue
        seen.add(key)
        result.append(person)
    return result


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, common_tokens: dict) -> None:
    _thu_moi_de_cuong(dest_dir, info, common_tokens)
    _thu_moi_nghiem_thu(dest_dir, info, common_tokens)


def _thu_moi_de_cuong(dest_dir, info, common_tokens):
    recipients = _dedupe_people(
        _external_members(info.ethics_committee, info.host_org)
        + _external_members(info.proposal_committee, info.host_org)
    )
    generated = expert_invitation.generate_multi_page_letter(
        dest_dir / "Công văn mời chuyên gia.docx", recipients, common_tokens
    )
    if not generated:
        print(
            "  [LUU Y] Khong co chuyen gia ngoai don vi chu tri trong Hoi dong khoa hoc/dao duc - "
            "bo qua dien thu moi de cuong."
        )


def _thu_moi_nghiem_thu(dest_dir, info, common_tokens):
    recipients = _dedupe_people(_external_members(info.acceptance_committee, info.host_org))
    generated = expert_invitation.generate_multi_page_letter(
        dest_dir / "Công văn mời chuyên gia nghiệm thu.docx", recipients, common_tokens
    )
    if not generated:
        print(
            "  [LUU Y] Khong co chuyen gia ngoai don vi chu tri trong Hoi dong nghiem thu - "
            "bo qua dien thu moi nghiem thu."
        )
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest test_section_moi_chuyen_gia.py -v`
Expected: 5 passed

- [ ] **Step 5: Commit**

```bash
git add section_moi_chuyen_gia.py test_section_moi_chuyen_gia.py
git commit -m "$(cat <<'EOF'
feat: auto-address invitation letters to every external committee member

Replaces the single hand-addressed letter with per-council filtering
(excludes host-org staff and secretaries), cross-committee dedup for
the combined đề cương letter, and now also fills the nghiệm thu letter.
EOF
)"
```

---

### Task 10: Checklist changes — A08 "Đầu mối liên hệ" field + token

**Files:**
- Create: `migrate_add_contact_person.py`
- Test: `test_migrate_add_contact_person.py`
- Modify: `migrate_add_tokens_sheet.py:13-26`
- Modify: `test_migrate_add_tokens_sheet.py`
- Modify: `test_token_rules.py:178-202`

**Interfaces:**
- Produces: checklist code `A08` in both project sheets; common token `{{DAU_MOI_LIEN_HE}}` resolved purely through the existing `_Tokens`-sheet mechanism (no new Python token-resolution code).

- [ ] **Step 1: Write the failing tests**

```python
# test_migrate_add_contact_person.py
from pathlib import Path

import openpyxl
import pytest

import migrate_add_contact_person as migrate
import excel_reader

CHECKLIST_PATH = Path(__file__).resolve().parent / excel_reader.CHECKLIST_FILENAME
SHEET_NAMES = ["Đề tài - Bánh ăn dặm VIAM 2027", "Đề tài - Mẫu trắng dự án mới"]


def test_add_contact_person_field_adds_a08_to_both_sheets():
    migrate.add_contact_person_field()

    wb = openpyxl.load_workbook(CHECKLIST_PATH)
    for sheet_name in SHEET_NAMES:
        ws = wb[sheet_name]
        codes = [row[0].value for row in ws.iter_rows(min_row=5, max_col=1)]
        assert "A08" in codes


def test_add_contact_person_field_is_idempotent():
    migrate.add_contact_person_field()
    migrate.add_contact_person_field()

    wb = openpyxl.load_workbook(CHECKLIST_PATH)
    for sheet_name in SHEET_NAMES:
        ws = wb[sheet_name]
        codes = [row[0].value for row in ws.iter_rows(min_row=5, max_col=1)]
        assert codes.count("A08") == 1


def _code_index(ws):
    index = {}
    for row in ws.iter_rows(min_row=5, max_col=1):
        cell = row[0]
        if isinstance(cell.value, str) and not cell.value.startswith("SEC_"):
            index[cell.value] = cell.row
    return index


@pytest.mark.parametrize("sheet_name", SHEET_NAMES)
def test_shifted_rows_reference_their_own_row(sheet_name):
    migrate.add_contact_person_field()

    wb = openpyxl.load_workbook(CHECKLIST_PATH, data_only=False)
    ws = wb[sheet_name]
    index = _code_index(ws)
    b01_row = index["B01"]
    formula = ws.cell(row=b01_row, column=6).value
    assert f"C{b01_row}" in formula
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_migrate_add_contact_person.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'migrate_add_contact_person'`

- [ ] **Step 3: Write the implementation**

```python
# migrate_add_contact_person.py
import copy
import re
from pathlib import Path

import openpyxl

CHECKLIST_PATH = Path(__file__).resolve().parent / "Form checklist hồ sơ dự án.xlsx"
SHEET_NAMES = ["Đề tài - Bánh ăn dặm VIAM 2027", "Đề tài - Mẫu trắng dự án mới"]
INSERT_AT_ROW = 13
STYLE_SOURCE_ROW = 12


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


def _has_contact_person_field(ws) -> bool:
    for row in ws.iter_rows(min_row=5, max_col=1):
        if row[0].value == "A08":
            return True
    return False


def add_contact_person_field() -> None:
    wb = openpyxl.load_workbook(CHECKLIST_PATH)
    changed = False
    for sheet_name in SHEET_NAMES:
        ws = wb[sheet_name]
        if _has_contact_person_field(ws):
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

        ws.cell(row=INSERT_AT_ROW, column=1, value="A08")
        ws.cell(
            row=INSERT_AT_ROW,
            column=2,
            value="Đầu mối liên hệ (họ tên, đơn vị, email, điện thoại - Tùy chọn)",
        )
        ws.cell(
            row=INSERT_AT_ROW,
            column=6,
            value=f'=IF(ISBLANK(C{INSERT_AT_ROW}), "⚪ Tùy chọn (Trống)", "✅ Xong")',
        )

        _fix_status_formula_row_refs(ws)

    if changed:
        wb.save(CHECKLIST_PATH)


if __name__ == "__main__":
    add_contact_person_field()
    print("Da them truong A08 'Dau moi lien he' vao ca 2 sheet.")
```

In `migrate_add_tokens_sheet.py`, add one entry to `TOKEN_SPECS` (after the `DIA_DIEM_TRIEN_KHAI` line):

```python
    ("DAU_MOI_LIEN_HE", "A08", "raw_or_placeholder", "……", "Đầu mối liên hệ (thư mời chuyên gia)"),
```

and update the print message at the bottom to `"Da tao/cap nhat sheet _Tokens voi 13 token mac dinh."`.

Add to `test_migrate_add_tokens_sheet.py`:

```python
def test_add_tokens_sheet_row_maps_contact_person_token(tmp_path):
    checklist_path = tmp_path / "checklist.xlsx"
    openpyxl.Workbook().save(checklist_path)

    migrate.add_tokens_sheet(checklist_path)

    wb = openpyxl.load_workbook(checklist_path)
    ws = wb[migrate.TOKENS_SHEET_NAME]
    rows = {ws.cell(row=r, column=1).value: r for r in range(2, ws.max_row + 1)}
    r = rows["DAU_MOI_LIEN_HE"]
    assert ws.cell(row=r, column=2).value == "A08"
    assert ws.cell(row=r, column=3).value == "raw_or_placeholder"
    assert ws.cell(row=r, column=4).value == "……"
```

Add to `test_token_rules.py`, inside `test_real_checklist_resolves_new_and_existing_tokens_correctly`:

```python
    # A08 chua duoc dien trong checklist that -> dung placeholder tu cot `param`.
    assert data.common_tokens["{{DAU_MOI_LIEN_HE}}"] == "……"
```

- [ ] **Step 4: Run tests, then run both migrations against the live checklist**

Run: `pytest test_migrate_add_contact_person.py test_migrate_add_tokens_sheet.py -v`
Expected: all pass

Run: `python migrate_add_contact_person.py` then `python migrate_add_tokens_sheet.py`
Expected: `Da them truong A08 'Dau moi lien he' vao ca 2 sheet.` then `Da tao/cap nhat sheet _Tokens voi 13 token mac dinh.`

Run: `pytest test_token_rules.py -v`
Expected: all pass, including the new `{{DAU_MOI_LIEN_HE}}` assertion.

- [ ] **Step 5: Commit**

```bash
git add migrate_add_contact_person.py test_migrate_add_contact_person.py \
  migrate_add_tokens_sheet.py test_migrate_add_tokens_sheet.py test_token_rules.py \
  "Form checklist hồ sơ dự án.xlsx"
git commit -m "$(cat <<'EOF'
feat: add A08 "Đầu mối liên hệ" checklist field and its common token

Wired purely through the existing _Tokens-sheet mechanism, no new
Python token-resolution code needed.
EOF
)"
```

---

### Task 11: Retire the CV-filename dropdown (checklist + script)

**Files:**
- Create: `migrate_remove_cv_filename_column.py`
- Test: `test_migrate_remove_cv_filename_column.py`
- Delete: `capnhat_danh_sach_cv.py`, `capnhat_danh_sach_cv.bat`, `test_capnhat_danh_sach_cv.py`

**Interfaces:**
- Produces: `remove_cv_filename_column() -> None`.

- [ ] **Step 1: Write the failing tests**

```python
# test_migrate_remove_cv_filename_column.py
import openpyxl

import migrate_remove_cv_filename_column as migrate


def _code_index(ws):
    return {
        row[0].value: row[0].row
        for row in ws.iter_rows(min_row=5, max_col=1)
        if isinstance(row[0].value, str)
    }


def test_removes_e_column_validation_on_f_rows():
    migrate.remove_cv_filename_column()

    wb = openpyxl.load_workbook(migrate.CHECKLIST_PATH)
    for sheet_name in migrate.SHEET_NAMES:
        ws = wb[sheet_name]
        index = _code_index(ws)
        f01_row = index["F01"]
        all_refs = set()
        for dv in ws.data_validations.dataValidation:
            all_refs |= set(str(dv.sqref).split())
        assert f"E{f01_row}" not in all_refs


def test_f01_status_formula_no_longer_checks_role_or_filename_columns():
    migrate.remove_cv_filename_column()

    wb = openpyxl.load_workbook(migrate.CHECKLIST_PATH)
    ws = wb["Đề tài - Mẫu trắng dự án mới"]
    index = _code_index(ws)
    formula = ws.cell(row=index["F01"], column=6).value
    assert "ISBLANK(E" not in formula
    assert "ISBLANK(D" not in formula


def test_f04_status_formula_only_checks_name_column():
    migrate.remove_cv_filename_column()

    wb = openpyxl.load_workbook(migrate.CHECKLIST_PATH)
    ws = wb["Đề tài - Mẫu trắng dự án mới"]
    index = _code_index(ws)
    row = index["F04"]
    formula = ws.cell(row=row, column=6).value
    assert formula == f'=IF(ISBLANK(C{row}), "⚪ Tùy chọn (Trống)", "✅ Đã khai tên - CV sẽ tự khớp theo tên")'


def test_lists_sheet_removed():
    migrate.remove_cv_filename_column()

    wb = openpyxl.load_workbook(migrate.CHECKLIST_PATH)
    assert migrate.LISTS_SHEET not in wb.sheetnames


def test_is_idempotent():
    migrate.remove_cv_filename_column()
    migrate.remove_cv_filename_column()

    wb = openpyxl.load_workbook(migrate.CHECKLIST_PATH)
    ws = wb["Đề tài - Mẫu trắng dự án mới"]
    index = _code_index(ws)
    assert "F01" in index
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest test_migrate_remove_cv_filename_column.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'migrate_remove_cv_filename_column'`

- [ ] **Step 3: Write the implementation**

```python
# migrate_remove_cv_filename_column.py
from pathlib import Path

import openpyxl

CHECKLIST_PATH = Path(__file__).resolve().parent / "Form checklist hồ sơ dự án.xlsx"
SHEET_NAMES = ["Đề tài - Bánh ăn dặm VIAM 2027", "Đề tài - Mẫu trắng dự án mới"]
LISTS_SHEET = "_Lists"


def _build_code_index(ws) -> dict:
    return {
        row[0].value: row[0].row
        for row in ws.iter_rows(min_row=5, max_col=1)
        if isinstance(row[0].value, str)
    }


def _clear_e_column_validations(ws, target_refs: set) -> None:
    keep = []
    for dv in ws.data_validations.dataValidation:
        dv_cells = set(str(dv.sqref).split())
        if not (dv_cells & target_refs):
            keep.append(dv)
    ws.data_validations.dataValidation = keep


def _rewrite_status_formulas(ws, index: dict) -> None:
    f01_row = index.get("F01")
    if f01_row is not None:
        ws.cell(row=f01_row, column=6).value = (
            f'=IF(ISBLANK(C{f01_row}), "❌ CHƯA ĐIỀN THÔNG TIN CNĐT (BÁO LỖI)", '
            f'"✅ Đã khai chủ nhiệm đề tài - CV sẽ tự khớp theo tên")'
        )
    for i in range(2, 11):
        row = index.get(f"F{i:02d}")
        if row is None:
            continue
        ws.cell(row=row, column=6).value = (
            f'=IF(ISBLANK(C{row}), "⚪ Tùy chọn (Trống)", "✅ Đã khai tên - CV sẽ tự khớp theo tên")'
        )


def remove_cv_filename_column() -> None:
    wb = openpyxl.load_workbook(CHECKLIST_PATH)

    for sheet_name in SHEET_NAMES:
        ws = wb[sheet_name]
        index = _build_code_index(ws)
        e_refs = {f"E{index[f'F{i:02d}']}" for i in range(1, 11) if f"F{i:02d}" in index}
        _clear_e_column_validations(ws, e_refs)
        _rewrite_status_formulas(ws, index)

    if LISTS_SHEET in wb.sheetnames:
        del wb[LISTS_SHEET]

    wb.save(CHECKLIST_PATH)


if __name__ == "__main__":
    remove_cv_filename_column()
    print("Da go cot 'TEN FILE CV' (dropdown + cong thuc trang thai) khoi ca 2 sheet.")
```

Delete `capnhat_danh_sach_cv.py`, `capnhat_danh_sach_cv.bat`, `test_capnhat_danh_sach_cv.py` — their sole purpose (refreshing the CV-filename dropdown) no longer exists.

- [ ] **Step 4: Run tests, then run the migration against the live checklist**

Run: `pytest test_migrate_remove_cv_filename_column.py -v`
Expected: 6 passed

Run: `python migrate_remove_cv_filename_column.py`
Expected: prints `Da go cot 'TEN FILE CV' (dropdown + cong thuc trang thai) khoi ca 2 sheet.`

Run: `pytest -v` (full suite, excluding `test_capnhat_danh_sach_cv.py` which is now deleted)
Expected: all pass

- [ ] **Step 5: Commit**

```bash
git add migrate_remove_cv_filename_column.py test_migrate_remove_cv_filename_column.py \
  "Form checklist hồ sơ dự án.xlsx"
git rm capnhat_danh_sach_cv.py capnhat_danh_sach_cv.bat test_capnhat_danh_sach_cv.py
git commit -m "$(cat <<'EOF'
feat: remove CV-filename dropdown/validation, retire its refresh script

PHẦN F rows now only need a declared name (CV resolved automatically
via cv_matching.find_cv_file) - the filename column, its dropdown, and
capnhat_danh_sach_cv.py/.bat (which existed solely to keep that
dropdown in sync) no longer serve a purpose.
EOF
)"
```

---

### Task 12: Documentation updates

**Files:**
- Modify: `HUONG_DAN.md`
- Modify: `HUONG_DAN_LAM_MAU_MOI.md`

- [ ] **Step 1: Update `HUONG_DAN.md` §3.4**

Replace the CV-filename-typing instructions (currently: "Mở Excel, vào ô F01... gõ đúng tên file...") with:

```markdown
### 3.4 Chuẩn bị file CV chủ nhiệm đề tài và chuyên gia

Công cụ tự động đính kèm CV vào hồ sơ, **tự khớp theo tên** — không cần gõ tên file nữa.

1. Đặt file CV (`.docx`/`.pdf`) vào thư mục **`CV chuyên gia/`** ở thư mục gốc công cụ. Tên file chỉ cần **chứa** tên người đó (không dấu hay có dấu đều được, hoa/thường không quan trọng) — ví dụ file `"TM-Gs. Nguyen Cong Khan.pdf"` tự khớp với người tên "Nguyễn Công Khẩn" khai ở checklist.
2. Chủ nhiệm đề tài: chỉ cần tên ở mã mục **B01** đã điền — công cụ tự tìm CV khớp tên đó trong `CV chuyên gia/`.
3. Chuyên gia khác cần đính kèm CV: khai tên ở **PHẦN F, mã mục F02–F10** (tên + vai trò) — không còn cột "TÊN FILE CV" nữa. Dòng nào để trống tên thì bỏ qua.

Nếu một cái tên không khớp đúng 1 file (khớp 0 file hoặc khớp nhiều hơn 1 file), script sẽ báo lỗi rõ ràng ngay khi chạy, nêu rõ tên nào đang gây lỗi và cách khắc phục (xem mục 7.2d).
```

- [ ] **Step 2: Update `HUONG_DAN.md` §6.1**

Delete the entire "6.1 Script cần chạy LẠI mỗi khi đổi danh sách CV chuyên gia" section (about `capnhat_danh_sach_cv.bat`) — it no longer applies since there is no dropdown to refresh.

- [ ] **Step 3: Update `HUONG_DAN.md` §7.2d**

Replace with:

```markdown
### 7.2d Lỗi: Không tìm thấy (hoặc khớp nhiều hơn 1) file CV chuyên gia
**Dấu hiệu:** Script báo lỗi kiểu: "Không tìm thấy file CV nào khớp tên '...' trong thư mục 'CV chuyên gia/'" hoặc "Tìm thấy nhiều hơn 1 file CV khớp tên '...'".

**Nguyên nhân:** Tên khai trong checklist (B01 hoặc F02–F10) không khớp đúng 1 file trong thư mục `CV chuyên gia/` sau khi so khớp không dấu/không phân biệt hoa-thường.

**Cách khắc phục:**
1. Mở thư mục `CV chuyên gia/`, kiểm tra tên file CV thật.
2. Nếu không khớp: đổi tên file (hoặc tên khai trong checklist) sao cho tên người xuất hiện liền nhau, đúng thứ tự trong tên file.
3. Nếu khớp nhiều hơn 1 file: đổi tên bớt các file trùng khớp để chỉ còn đúng 1 file khớp tên đó.
4. Lưu rồi chạy lại.
```

- [ ] **Step 4: Update the one-off script list in `HUONG_DAN.md` §6**

In the bullet list under "Trong thư mục công cụ, có các file script chạy một lần sau:", remove the `migrate_fix_f01_cv_filename.py` line (script deleted in Task 4) and the `capnhat_danh_sach_cv.bat`-related content already removed in Step 2 above, and add:

```markdown
- migrate_add_contact_person.py
- migrate_remove_cv_filename_column.py
- migrate_tokenize_invitation_letters.py
```

- [ ] **Step 5: Update `HUONG_DAN_LAM_MAU_MOI.md` token table**

Add two rows to the "Bảng token dùng chung" table (after `{{DIA_DIEM_TRIEN_KHAI}}`):

```markdown
| `{{DAU_MOI_LIEN_HE}}` | `info.common_tokens` | A08 |
```

Add a new subsection after the existing token table, before "## Cách 1":

```markdown
## Token riêng theo trang (chỉ dùng trong thư mời chuyên gia)

Khác với bảng token dùng chung ở trên (1 giá trị/dự án, khai qua sheet
`_Tokens`), thư mời chuyên gia (`expert_invitation.py`) còn có 2 token
**riêng theo từng trang**, tính động lúc sinh hồ sơ theo từng người nhận,
không khai báo qua Excel:

| Token | Ý nghĩa |
|---|---|
| `{{CHUYEN_GIA_HO_TEN}}` | `"<học vị> <tên>"` của người nhận trang đó |
| `{{CHUYEN_GIA_DON_VI}}` | Đơn vị công tác của người nhận trang đó |
```

- [ ] **Step 6: Commit**

```bash
git add HUONG_DAN.md HUONG_DAN_LAM_MAU_MOI.md
git commit -m "$(cat <<'EOF'
docs: document name-based CV matching, A08 field, per-page tokens

Removes the now-obsolete CV-filename-dropdown instructions.
EOF
)"
```

---

### Task 13: Full pipeline verification

**Files:** none (verification only)

- [ ] **Step 1: Run the complete test suite**

Run: `pytest -v`
Expected: all tests pass, zero failures, zero errors.

- [ ] **Step 2: Run the full generator against the live sample project**

Run: `python tao_ho_so_moi.py "Đề tài - Bánh ăn dặm VIAM 2027"`
Expected: completes with `XONG. Bo ho so da tao tai: ...`, no `[CANH BAO]`/`[LOI]` lines about missing CVs or unresolved tokens.

- [ ] **Step 3: Manually verify the generated output**

Open the generated `Hồ sơ - .../03. Công văn mời chuyên gia/Công văn mời chuyên gia.docx` and confirm:
- One page per external (non-VIAM) member of the ethics + proposal committees, each addressed by their own name, no leftover `{{...}}` tokens, no leftover dots.
- The contact line shows either real data (if A08 was filled) or the `……` placeholder — not "Lê Minh Khánh".

Open `Công văn mời chuyên gia nghiệm thu.docx` and confirm it exists, has one page per external acceptance-committee member, and its wording says "NGHIỆM THU" not "khoa học và đạo đức".

Open `04. Hồ sơ nghiệm thu/Phiếu ký nhận tiền.docx` and confirm all 6 signer rows show real acceptance-committee names (chair, 4 reviewers/members, secretary) — none of the old sample project's names ("Phạm Văn Hoan", "Nguyễn Thị Lâm", etc.) remain.

- [ ] **Step 4: Report findings**

If every check in Steps 1-3 passes, the feature is complete — no commit needed for this task (verification-only).
