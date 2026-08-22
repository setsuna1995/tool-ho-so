import paths


def test_project_root_is_absolute():
    assert paths.project_root().is_absolute()


def test_project_root_contains_excel_checklist():
    root = paths.project_root()
    assert (root / "Form checklist hồ sơ dự án.xlsx").exists()
