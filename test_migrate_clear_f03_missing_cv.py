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
