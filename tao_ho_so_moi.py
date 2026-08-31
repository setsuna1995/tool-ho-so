import io
import shutil
import sys
import tempfile
from pathlib import Path

import openpyxl

import cv_matching
import excel_reader
import paths
import section_dao_duc
import section_khoa_hoc
import section_moi_chuyen_gia
import section_nghiem_thu
import template_config
import tokens
import word_writer

PROJECT_SHEET_PREFIX = "Đề tài - "

ILLEGAL_FOLDER_CHARS = ':/\\*?"<>|'


def list_project_sheets(xlsx_path: Path) -> list:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    try:
        return [name for name in wb.sheetnames if name.startswith(PROJECT_SHEET_PREFIX)]
    finally:
        wb.close()


def choose_sheet_name(xlsx_path: Path) -> str:
    if len(sys.argv) > 1:
        return sys.argv[1]

    sheets = list_project_sheets(xlsx_path)
    if not sheets:
        raise RuntimeError(
            f"Khong tim thay sheet nao bat dau bang '{PROJECT_SHEET_PREFIX}' trong file {xlsx_path.name}."
        )
    if len(sheets) == 1:
        print(f"Dung sheet du an: '{sheets[0]}'")
        return sheets[0]

    print("Chon sheet du an muon tao ho so:")
    for i, name in enumerate(sheets, start=1):
        print(f"  {i}. {name}")

    while True:
        raw = input(f"Nhap so thu tu (1-{len(sheets)}): ").strip()
        if raw.isdigit() and 1 <= int(raw) <= len(sheets):
            return sheets[int(raw) - 1]
        print(f"  [LOI] Vui long nhap so tu 1 den {len(sheets)}.")


def copy_templates(root: Path, dest_root: Path) -> None:
    copies = template_config.discover_copies(root)
    for rel_src, rel_dst in copies:
        src = root / rel_src
        dst = dest_root / rel_dst
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_bytes(src.read_bytes())


def copy_head_cv(root: Path, dest_root: Path, info) -> None:
    cv_dir = root / "CV chuyên gia"
    src = cv_matching.find_cv_file(cv_dir, info.head.name, context=" (chủ nhiệm đề tài)")
    dst = dest_root / "01. Hồ sơ đạo đức đề cương" / src.name
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_bytes(src.read_bytes())


def copy_expert_cvs(root: Path, dest_root: Path, info) -> None:
    cv_dir = root / "CV chuyên gia"
    resolved = [
        cv_matching.find_cv_file(cv_dir, entry.name, context=f" (mã mục {entry.code} - {entry.role})")
        for entry in info.expert_cvs
    ]

    for src in resolved:
        dst = dest_root / "03. Công văn mời chuyên gia" / src.name
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_bytes(src.read_bytes())


def _sanitize_folder_name(name: str) -> str:
    """Thay cac ky tu khong hop le trong ten thu muc Windows bang khoang trang."""
    for ch in ILLEGAL_FOLDER_CHARS:
        name = name.replace(ch, " ")
    return name


def generate_all(root: Path, dest_root: Path, info, session: word_writer.Session) -> None:
    """Tao toan bo ho so o mot thu muc tam local, roi moi copy nguyen bo vao dest_root.

    Lam viec truc tiep tren thu muc dich khi no nam trong OneDrive/SharePoint de
    dong bo de bi AutoSave/Protected View/khoa file cua OneDrive lam gian doan
    thao tac COM cua Word giua chung - xu ly o thu muc tam local (khong dong bo)
    roi copy file thuan (khong qua Word) vao dich se tranh duoc van de nay.
    """
    staging_root = Path(tempfile.mkdtemp(prefix="tao_ho_so_"))
    try:
        common_tokens = tokens.build_common_tokens(info)

        print("Dang sao chep file mau...")
        copy_templates(root, staging_root)

        print("Dang sao chep CV chu nhiem de tai...")
        copy_head_cv(root, staging_root, info)

        print("Dang sao chep CV chuyen gia da khai...")
        copy_expert_cvs(root, staging_root, info)

        print("Dang sinh ho so dao duc...")
        section_dao_duc.generate(session, staging_root / "01. Hồ sơ đạo đức đề cương", info, common_tokens)

        print("Dang sinh ho so khoa hoc de cuong...")
        section_khoa_hoc.generate(session, staging_root / "02. Hồ sơ khoa học đề cương", info, common_tokens)

        print("Dang sinh cong van moi chuyen gia...")
        section_moi_chuyen_gia.generate(session, staging_root / "03. Công văn mời chuyên gia", info, common_tokens)

        print("Dang sinh ho so nghiem thu...")
        section_nghiem_thu.generate(session, staging_root / "04. Hồ sơ nghiệm thu", info, common_tokens)

        shutil.copytree(staging_root, dest_root, dirs_exist_ok=True)
    except Exception:
        print(f"  [LUU Y] Cac file da xu ly tam thoi con luu tai: {staging_root} (de kiem tra loi)")
        raise
    else:
        shutil.rmtree(staging_root, ignore_errors=True)


def main() -> None:
    root = paths.project_root()

    checklist_path = root / excel_reader.CHECKLIST_FILENAME
    sheet_name = None

    try:
        sheet_name = choose_sheet_name(checklist_path)

        print("Dang doc du lieu tu Excel checklist...")
        info = excel_reader.load_project_data(checklist_path, sheet_name)

        session = word_writer.Session()
        print(f"Che do ghi Word dang dung: {session.backend}")
        if session.backend == "docx":
            print(
                "  [LUU Y] Khong co Word COM - dung fallback python-docx thuan.\n"
                "  Mot vai cho xoa dong trong co the con sot dong trong, script se canh bao khi gap."
            )

        dest_dir_name = f"Hồ sơ - {_sanitize_folder_name(info.title)} ({info.year})"
        dest_root = root / dest_dir_name

        try:
            generate_all(root, dest_root, info, session)
        finally:
            session.quit()

        print(f"XONG. Bo ho so da tao tai: {dest_root}")

    except KeyError as e:
        # KeyError thuong xay ra khi SHEET_NAME sai ten, nhung cung co the la
        # ma muc (vd C09) bi thieu trong sheet dung - kiem tra ten sheet truoc
        # de dua ra thong bao dung trong tam.
        try:
            wb = openpyxl.load_workbook(root / excel_reader.CHECKLIST_FILENAME)
            sheet_names = wb.sheetnames
        except Exception:
            sheet_names = None

        if sheet_names is not None and sheet_name is not None and sheet_name not in sheet_names:
            print(f"[LOI] Khong tim thay sheet '{sheet_name}' trong file Excel checklist.")
            print("Cac sheet hien co trong file:")
            for name in sheet_names:
                print(f"  - {name}")
            print(
                "Neu ban truyen ten sheet qua dong lenh, hay kiem tra lai chinh ta "
                "(xem HUONG_DAN.md muc 7)."
            )
        else:
            print("Da xay ra loi khi tao ho so. Chi tiet loi:")
            print(f"  {e}")
            print("Vui long xem HUONG_DAN.md muc 7 de biet cach khac phuc.")
        sys.exit(1)

    except Exception as e:
        print("Da xay ra loi khi tao ho so. Chi tiet loi:")
        print(f"  {e}")
        print("Vui long xem HUONG_DAN.md muc 7 de biet cach khac phuc.")
        sys.exit(1)


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")
    main()
