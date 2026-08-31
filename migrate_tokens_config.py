# migrate_tokens_config.py
import json
from pathlib import Path
import openpyxl

import excel_reader

CHECKLIST_PATH = Path(__file__).resolve().parent / excel_reader.CHECKLIST_FILENAME
CONFIG_PATH = Path(__file__).resolve().parent / "config_tokens.json"

TOKEN_SPECS = [
    {"token_name": "TEN_DE_TAI", "code": "A01", "kind": "raw", "param": "", "note": "Tên đề tài"},
    {"token_name": "NAM", "code": "A03", "kind": "raw", "param": "", "note": "Năm thực hiện hồ sơ"},
    {"token_name": "DON_VI_CHU_TRI", "code": "A04", "kind": "raw", "param": "", "note": "Cơ quan chủ trì"},
    {"token_name": "DON_VI_DOI_TAC", "code": "A06", "kind": "raw_or_placeholder", "param": "", "note": "Cơ quan phối hợp"},
    {"token_name": "CHU_NHIEM_HO_TEN", "code": "B01", "kind": "person_ho_ten", "param": "", "note": "Chủ nhiệm đề tài - có học hàm/học vị"},
    {"token_name": "CHU_NHIEM_TEN", "code": "B01", "kind": "person_ten", "param": "", "note": "Chủ nhiệm đề tài - chỉ tên"},
    {"token_name": "DONG_CHU_NHIEM_TEN", "code": "B02", "kind": "person_ten", "param": "", "note": "Đồng chủ nhiệm đề tài - chỉ tên"},
    {"token_name": "DONG_CHU_NHIEM_HO_TEN", "code": "B02", "kind": "person_ho_ten", "param": "", "note": "Đồng chủ nhiệm đề tài - có học hàm/học vị"},
    {"token_name": "THU_KY_DE_TAI", "code": "B03", "kind": "person_ho_ten", "param": "", "note": "Thư ký đề tài"},
    {"token_name": "THOI_GIAN_BAT_DAU", "code": "A05", "kind": "timeline_start", "param": "", "note": "Mốc bắt đầu (MM/YYYY)"},
    {"token_name": "THOI_GIAN_KET_THUC", "code": "A05", "kind": "timeline_end", "param": "", "note": "Mốc kết thúc (MM/YYYY)"},
    {"token_name": "DIA_DIEM_TRIEN_KHAI", "code": "A07", "kind": "raw_or_placeholder", "param": "……………………………", "note": "Địa điểm triển khai nghiên cứu"},
    {"token_name": "DAU_MOI_LIEN_HE", "code": "A08", "kind": "raw_or_placeholder", "param": "……", "note": "Đầu mối liên hệ (thư mời chuyên gia)"},
    {"token_name": "CHU_NHIEM_DON_VI", "code": "B01", "kind": "person_org", "param": "", "note": "Đơn vị công tác chủ nhiệm đề tài"},
    {"token_name": "DANH_SACH_NGHIEN_CUU_VIEN", "code": "B04", "kind": "numbered_researchers", "param": "", "note": "Danh sách thành viên nghiên cứu viên (có đánh số thứ tự)"},
    {"token_name": "CHU_TICH_HD_DAO_DUC", "code": "C01", "kind": "person_ho_ten", "param": "", "note": "Chủ tịch Hội đồng Đạo đức"},
    {"token_name": "CHU_TICH_HD_KHOA_HOC", "code": "D01", "kind": "person_ho_ten", "param": "", "note": "Chủ tịch Hội đồng Khoa học"},
    {"token_name": "CHU_TICH_HD_NGHIEM_THU", "code": "E01", "kind": "person_ho_ten", "param": "", "note": "Chủ tịch Hội đồng Nghiệm thu"},
    {"token_name": "CHU_TICH_HD_NGHIEM_THU_TEN", "code": "E01", "kind": "person_ten", "param": "", "note": "Chủ tịch Hội đồng Nghiệm thu - chỉ tên"},
]


def save_config_tokens(config_path: Path = CONFIG_PATH) -> None:
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(TOKEN_SPECS, f, ensure_ascii=False, indent=2)


def remove_tokens_sheet_from_checklist(checklist_path: Path = CHECKLIST_PATH) -> None:
    if not checklist_path.exists():
        return
    wb = openpyxl.load_workbook(checklist_path)
    if "_Tokens" in wb.sheetnames:
        del wb["_Tokens"]
        wb.save(checklist_path)


def migrate_all(checklist_path: Path = CHECKLIST_PATH, config_path: Path = CONFIG_PATH) -> None:
    save_config_tokens(config_path)
    remove_tokens_sheet_from_checklist(checklist_path)


if __name__ == "__main__":
    migrate_all()
    print("XONG: Da tao file config_tokens.json va tach hoan toan khoi file Excel checklist.")
