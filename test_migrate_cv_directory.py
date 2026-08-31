# test_migrate_cv_directory.py
from pathlib import Path
import openpyxl
import migrate_cv_directory


def test_migrate_folder_renames_old_dir_and_deletes_tm_pdfs(tmp_path, monkeypatch):
    old_dir = tmp_path / "CV chuyên gia"
    new_dir = tmp_path / "Lý lịch khoa học"
    old_dir.mkdir()
    (old_dir / "Lý lịch khoa học - Trương Hồng Sơn.docx").write_text("cv")
    (old_dir / "TM-Gs. Nguyen Cong Khan.pdf").write_text("pdf")

    monkeypatch.setattr(migrate_cv_directory, "OLD_DIR", old_dir)
    monkeypatch.setattr(migrate_cv_directory, "NEW_DIR", new_dir)

    migrate_cv_directory.migrate_folder()

    assert not old_dir.exists()
    assert new_dir.exists()
    assert (new_dir / "Lý lịch khoa học - Trương Hồng Sơn.docx").exists()
    assert not (new_dir / "TM-Gs. Nguyen Cong Khan.pdf").exists()
