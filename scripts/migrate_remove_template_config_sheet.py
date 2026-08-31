from pathlib import Path

import openpyxl

CHECKLIST_PATH = Path(__file__).resolve().parent / "Form checklist hồ sơ dự án.xlsx"
CONFIG_SHEET_NAME = "Cấu hình mẫu"


def remove_template_config_sheet() -> None:
    wb = openpyxl.load_workbook(CHECKLIST_PATH)
    if CONFIG_SHEET_NAME not in wb.sheetnames:
        return
    del wb[CONFIG_SHEET_NAME]
    wb.save(CHECKLIST_PATH)


if __name__ == "__main__":
    remove_template_config_sheet()
    print(f"Da xoa sheet '{CONFIG_SHEET_NAME}' (neu co).")
