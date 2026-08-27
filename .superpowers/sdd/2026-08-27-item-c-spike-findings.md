# Phase 0 spike findings: item C's "tên dài quá" error

**Date:** 2026-08-27
**Script:** `spike_item_c_repro.py` (throwaway, deleted after this run — see plan Task 1)
**Environment:** Word COM available and working (`win32com.client.Dispatch('Word.Application')` succeeded, Word version 16.0). The spike ran to completion; no COM-unavailability blocker.

## Repro script

Ran the script exactly as specified in the plan's Task 1 brief, unmodified. It builds an
over-long title string (`LONG_TITLE`, ~282 characters — `"Nghiên cứu đánh giá hiệu quả "` +
`"và tính an toàn " * 15` + `"của sản phẩm"`) and calls
`section_khoa_hoc._bb_kiem_phieu_thong_qua_de_cuong` against two destination directories:

1. `short_local_path` — a fresh `tempfile.mkdtemp()` directory (short, no nesting).
2. `long_nested_path` — a deeply nested directory built under the repo root from
   `"Ho so tam rat dai de kiem tra duong dan " * 4` + `"02. Ho so khoa hoc de cuong"`.

## Case 1: `short_local_path` — reproduced the bug

Destination: `C:\Users\Kien\AppData\Local\Temp\spike_short_q7pqt20c`

Exact output printed by the script:

```
--- short_local_path: C:\Users\Kien\AppData\Local\Temp\spike_short_q7pqt20c ---
short_local_path: LOI - RuntimeError: Loi khi thay the '{{TEN_DE_TAI}}' trong file '07. BB kiểm phiếu thông qua đề cương.docx': (-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'String parameter too long.', 'wdmain11.chm', 25334, -2146822434), None)
```

This is the exact exception type + message text, verbatim, as printed by the script (no
paraphrasing).

**This confirms the bug reproduces on a short, non-nested local path.** The destination
directory path here is nowhere near any Windows/OneDrive path-length limit. The failure is a
COM error raised by Word itself ("String parameter too long") while replacing the
`{{TEN_DE_TAI}}` token, i.e. while performing a find/replace whose search-or-replacement
string is the ~282-character `LONG_TITLE`.

## Case 2: `long_nested_path` — did not reach Word at all; different failure, before the code under test ever ran

Destination attempted: `F:\tool-ho-so\.claude\worktrees\excel-token-architecture\Ho so tam rat dai de kiem tra duong dan Ho so tam rat dai de kiem tra duong dan Ho so tam rat dai de kiem tra duong dan Ho so tam rat dai de kiem tra duong dan \02. Ho so khoa hoc de cuong`

The script's own `base_dir.mkdir(parents=True, exist_ok=True)` call (which runs *before* the
`try`/`except` around the code under test) raised an unhandled exception and crashed the
script for this case, so `section_khoa_hoc._bb_kiem_phieu_thong_qua_de_cuong` and Word COM
were never invoked for this case at all:

```
Traceback (most recent call last):
  File "C:\Users\Kien\AppData\Local\Python\pythoncore-3.14-64\Lib\pathlib\__init__.py", line 1011, in mkdir
    os.mkdir(self, mode)
    ~~~~~~~~^^^^^^^^^^^^
FileNotFoundError: [WinError 3] The system cannot find the path specified: 'F:\\tool-ho-so\\.claude\\worktrees\\excel-token-architecture\\Ho so tam rat dai de kiem tra duong dan Ho so tam rat dai de kiem tra duong dan Ho so tam rat dai de kiem tra duong dan Ho so tam rat dai de kiem tra duong dan \\02. Ho so khoa hoc de cuong'
```

`WinError 3` ("The system cannot find the path specified") is the classic symptom of exceeding
the Windows `MAX_PATH` (260-character) limit on a filesystem/process where long-path support
isn't enabled — the OS refuses to create a directory whose full path is too long, rather than
returning a "path too long" error. This happened purely from directory nesting depth (worsened
by this run happening inside the deeply-nested worktree path
`F:\tool-ho-so\.claude\worktrees\excel-token-architecture\...`); it is unrelated to Word COM
or to the `LONG_TITLE` content.

Because the exception in this case happened outside the `try`/`except`, `session.quit()` for
this iteration's `word_writer.Session` was also never reached — but no `word_writer.Session`
had been constructed yet for this case (construction happens inside the `try` block, after the
`mkdir`), so no Word process was left orphaned. Confirmed via `tasklist | grep -i WINWORD`
after the run: no `WINWORD.exe` processes remained.

## Which hypothesis each case supports

The plan's Phase 0/7 sections framed two hypotheses:
- **H1: Windows/OneDrive path-length limit** (a path or nested-directory name too long).
- **H2: Word COM `Find.Execute` ~255-char search-string limit** (the title/token text itself
  too long for a single Find/Replace call).

- **Case 1 supports H2, and only H2.** The failure happened on a short, non-nested path, so
  path length cannot be the cause. The COM error text — "String parameter too long" — occurs
  while replacing `{{TEN_DE_TAI}}` with the ~282-character `LONG_TITLE`, which is consistent
  with Word's Find/Replace string-length ceiling (commonly cited around 255 characters for
  `Find.Execute`/`Selection.Find`-style calls).
- **Case 2 is a real Windows path-length failure (supports the general existence of an
  H1-type problem), but it is a different failure from item C's reported bug**: it never
  reaches the `{{TEN_DE_TAI}}` replacement code at all, and it is triggered by directory
  nesting the production code doesn't construct (the repro script's own artificially deep
  folder name), not by anything in `section_khoa_hoc.py` or `word_writer.py`.

**Item C's originally reported "tên dài quá" error corresponds to Case 1 / H2.** The
production code path that fails is the `{{TEN_DE_TAI}}` token replacement in
`section_khoa_hoc._bb_kiem_phieu_thong_qua_de_cuong`, and it fails purely because of the
*title string's* length, independent of destination path length.

## Recommendation for Phase 7

Phase 7's fix should target **H2**: the `{{TEN_DE_TAI}}` (and likely any other long-token)
replacement call in `word_writer.py` / `section_khoa_hoc.py` needs to stop relying on a single
Word COM Find/Replace call for arbitrarily long strings — e.g. chunk the replacement, or use a
COM API path that isn't subject to the ~255-character search/replacement string ceiling —
rather than treating this as a path-length problem.
