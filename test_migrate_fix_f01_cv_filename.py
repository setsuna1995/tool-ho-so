import migrate_fix_f01_cv_filename as migrate
from excel_reader import load_project_data

CHECKLIST_PATH = migrate.CHECKLIST_PATH
SHEET_VIAM = "Đề tài - Bánh ăn dặm VIAM 2027"


def test_f01_matches_real_cv_file_after_fix():
    migrate.fix_f01_cv_filename()

    data = load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    assert data.head_cv_filename == migrate.CORRECT_FILENAME

    cv_path = CHECKLIST_PATH.parent / "CV chuyên gia" / data.head_cv_filename
    assert cv_path.exists()


def test_running_migration_twice_is_safe():
    migrate.fix_f01_cv_filename()
    migrate.fix_f01_cv_filename()

    data = load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    assert data.head_cv_filename == migrate.CORRECT_FILENAME
