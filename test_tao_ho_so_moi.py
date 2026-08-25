import dataclasses

import pytest

import excel_reader
import paths
import tao_ho_so_moi

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
