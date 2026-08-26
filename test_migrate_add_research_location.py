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
