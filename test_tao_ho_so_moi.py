import dataclasses

import pytest

import excel_reader
import paths
import tao_ho_so_moi
import word_writer

CHECKLIST_PATH = paths.project_root() / excel_reader.CHECKLIST_FILENAME
SHEET_VIAM = "Đề tài - Bánh ăn dặm VIAM 2027"


def test_list_project_sheets_includes_known_sheets_and_excludes_config():
    sheets = tao_ho_so_moi.list_project_sheets(CHECKLIST_PATH)
    assert "Đề tài - Bánh ăn dặm VIAM 2027" in sheets
    assert "Đề tài - Mẫu trắng dự án mới" in sheets
    assert "Cấu hình mẫu" not in sheets


def test_choose_sheet_name_uses_command_line_argument(monkeypatch):
    monkeypatch.setattr("sys.argv", ["tao_ho_so_moi.py", "Đề tài - Bánh ăn dặm VIAM 2027"])
    assert tao_ho_so_moi.choose_sheet_name(CHECKLIST_PATH) == "Đề tài - Bánh ăn dặm VIAM 2027"


def test_choose_sheet_name_prompts_and_returns_selected_sheet(monkeypatch):
    monkeypatch.setattr("sys.argv", ["tao_ho_so_moi.py"])
    sheets = tao_ho_so_moi.list_project_sheets(CHECKLIST_PATH)
    target_index = sheets.index("Đề tài - Bánh ăn dặm VIAM 2027") + 1
    monkeypatch.setattr("builtins.input", lambda _: str(target_index))
    assert tao_ho_so_moi.choose_sheet_name(CHECKLIST_PATH) == "Đề tài - Bánh ăn dặm VIAM 2027"


def test_choose_sheet_name_reprompts_on_invalid_input(monkeypatch):
    monkeypatch.setattr("sys.argv", ["tao_ho_so_moi.py"])
    sheets = tao_ho_so_moi.list_project_sheets(CHECKLIST_PATH)
    target_index = sheets.index("Đề tài - Bánh ăn dặm VIAM 2027") + 1
    responses = iter(["0", "abc", str(target_index)])
    monkeypatch.setattr("builtins.input", lambda _: next(responses))
    assert tao_ho_so_moi.choose_sheet_name(CHECKLIST_PATH) == "Đề tài - Bánh ăn dặm VIAM 2027"


def test_copy_templates_discovers_files_from_mau_folders_without_excel_config(tmp_path):
    root = paths.project_root()

    tao_ho_so_moi.copy_templates(root, tmp_path)

    assert (tmp_path / "01. Hồ sơ đạo đức đề cương" / "00. QĐ Giao đề tài.docx").exists()
    assert (tmp_path / "04. Hồ sơ nghiệm thu" / "Phiếu chấm điểm nghiệm thu (TNLS).docx").exists()
    assert (tmp_path / "04. Hồ sơ nghiệm thu" / "Phiếu chấm điểm nghiệm thu (TVCT_ĐGHQ).docx").exists()


def test_copy_templates_does_not_copy_archived_reference_files(tmp_path):
    root = paths.project_root()

    tao_ho_so_moi.copy_templates(root, tmp_path)

    assert not (tmp_path / "01. Hồ sơ đạo đức đề cương" / "Lý lịch khoa học - Trương Hồng Sơn.docx").exists()
    assert not (tmp_path / "01. Hồ sơ đạo đức đề cương" / "Hồ sơ đạo đức COLOSTRUM .docx").exists()


def test_copy_head_cv_copies_file_named_in_f01(tmp_path):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)

    tao_ho_so_moi.copy_head_cv(root, tmp_path, info)

    dest = tmp_path / "01. Hồ sơ đạo đức đề cương" / info.head_cv_filename
    assert dest.exists()


def test_copy_head_cv_raises_clear_error_when_file_missing(tmp_path):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    bad_info = dataclasses.replace(info, head_cv_filename="không tồn tại.docx")

    with pytest.raises(FileNotFoundError):
        tao_ho_so_moi.copy_head_cv(root, tmp_path, bad_info)


def test_generate_all_stages_locally_then_copies_into_dest_root(tmp_path):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    dest_root = tmp_path / "Hồ sơ output"

    session = word_writer.Session(force_backend="docx")
    try:
        tao_ho_so_moi.generate_all(root, dest_root, info, session)
    finally:
        session.quit()

    assert (dest_root / "01. Hồ sơ đạo đức đề cương" / "00. QĐ Giao đề tài.docx").exists()
    assert (dest_root / "04. Hồ sơ nghiệm thu" / "Phiếu chấm điểm nghiệm thu (TVCT_ĐGHQ).docx").exists()


def test_generate_all_cleans_up_local_staging_dir_on_success(tmp_path, monkeypatch):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    dest_root = tmp_path / "Hồ sơ output"
    staging_dir = tmp_path / "staging"

    def fake_mkdtemp(prefix=None):
        staging_dir.mkdir(exist_ok=True)
        return str(staging_dir)

    monkeypatch.setattr(tao_ho_so_moi.tempfile, "mkdtemp", fake_mkdtemp)

    session = word_writer.Session(force_backend="docx")
    try:
        tao_ho_so_moi.generate_all(root, dest_root, info, session)
    finally:
        session.quit()

    assert not staging_dir.exists()


def test_generate_all_keeps_staging_dir_and_reports_path_on_failure(tmp_path, monkeypatch, capsys):
    root = paths.project_root()
    info = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    bad_info = dataclasses.replace(info, head_cv_filename="không tồn tại.docx")
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
