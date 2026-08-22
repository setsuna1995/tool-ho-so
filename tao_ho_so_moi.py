import io
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import excel_reader
import paths
import section_dao_duc
import section_khoa_hoc
import section_moi_chuyen_gia
import section_nghiem_thu
import word_writer

SHEET_NAME = "Đề tài - Bánh ăn dặm VIAM 2027"

TITLE_OLD = (
    "Đánh giá hiệu quả sản phẩm sữa dinh dưỡng pha sẵn KUN DOCTOR COLOSTRUM lên "
    "tình trạng dinh dưỡng, miễn dịch, tiêu hóa và giấc ngủ của trẻ từ 24 đến 72 tháng tuổi"
)

COPIES = [
    ("01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM/00. QĐ Giao đề tài.docx", "01. Hồ sơ đạo đức đề cương/00. QĐ Giao đề tài.docx"),
    ("01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM/01. QĐTLHĐ đạo đức đề cương.docx", "01. Hồ sơ đạo đức đề cương/01. QĐTLHĐ đạo đức đề cương.docx"),
    ("01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM/02. BB họp HĐ đạo đức - KUN COLOSTRUM.docx", "01. Hồ sơ đạo đức đề cương/02. BB họp HĐ đạo đức.docx"),
    ("01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM/03. BB kiểm phiếu HĐ đạo đức.docx", "01. Hồ sơ đạo đức đề cương/03. BB kiểm phiếu HĐ đạo đức.docx"),
    ("01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM/04. Dr.Kun QĐ chấp nhận đạo đức.docx", "01. Hồ sơ đạo đức đề cương/04. QĐ chấp nhận đạo đức.docx"),
    ("01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM/Bảng kiểm đánh giá đạo đức.docx", "01. Hồ sơ đạo đức đề cương/Bảng kiểm đánh giá đạo đức.docx"),
    ("01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM/Truong_Hong_Son_ly-lich-khoa-hoc-2024.docx", "01. Hồ sơ đạo đức đề cương/Lý lịch khoa học - Trương Hồng Sơn.docx"),

    ("02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM/05. Dr.Kun QD TLHDKH đề cương.docx", "02. Hồ sơ khoa học đề cương/05. QĐ TLHĐ khoa học xét đề cương.docx"),
    ("02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM/06. Dr.Kun Bien ban hop thong qua de cuong de tai.docx", "02. Hồ sơ khoa học đề cương/06. BB họp thông qua đề cương.docx"),
    ("02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM/07. Dr.Kun Bien ban kiem phieu thong qua de cuong.docx", "02. Hồ sơ khoa học đề cương/07. BB kiểm phiếu thông qua đề cương.docx"),
    ("02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM/08. Dr.Kun QĐ phe-duyet-de-tai.docx", "02. Hồ sơ khoa học đề cương/08. QĐ phê duyệt đề tài.docx"),
    ("02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM/Dr.Kun Phieu cham diem HD de cuong.docx", "02. Hồ sơ khoa học đề cương/Phiếu chấm điểm HĐ đề cương.docx"),
    ("02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM/Dr.Kun Phieu nhan xet danh gia ho so.docx", "02. Hồ sơ khoa học đề cương/Phiếu nhận xét đánh giá hồ sơ.docx"),

    ("03. CV mời chuyên gia - mẫu COLOSTRUM/CV mời chuyên gia.docx", "03. Công văn mời chuyên gia/Công văn mời chuyên gia.docx"),

    ("04. Hồ sơ nghiệm thu/04. Hồ sơ nghiệm thu/9. Quyết định THÀNH LẬP HĐ nghiệm thu.docx", "04. Hồ sơ nghiệm thu/9. Quyết định thành lập HĐ nghiệm thu.docx"),
    ("04. Hồ sơ nghiệm thu/04. Hồ sơ nghiệm thu/10. Biên bản HỌP HĐ nghiệm thu.docx", "04. Hồ sơ nghiệm thu/10. Biên bản họp HĐ nghiệm thu.docx"),
    ("04. Hồ sơ nghiệm thu/04. Hồ sơ nghiệm thu/11. Biên bản KIỂM PHIẾU nghiệm thu.docx", "04. Hồ sơ nghiệm thu/11. Biên bản kiểm phiếu nghiệm thu.docx"),
    ("04. Hồ sơ nghiệm thu/04. Hồ sơ nghiệm thu/12. Quyết định công nhận kết quả đề tài.docx", "04. Hồ sơ nghiệm thu/12. Quyết định công nhận kết quả đề tài.docx"),
    ("04. Hồ sơ nghiệm thu/04. Hồ sơ nghiệm thu/Phiếu CHẤM ĐIỂM nghiệm thu-(TVCT_ĐGHQ).docx", "04. Hồ sơ nghiệm thu/Phiếu chấm điểm nghiệm thu (TVCT_ĐGHQ).docx"),
    ("04. Hồ sơ nghiệm thu/04. Hồ sơ nghiệm thu/Phiếu ký nhận tiền.docx", "04. Hồ sơ nghiệm thu/Phiếu ký nhận tiền.docx"),
    ("04. Hồ sơ nghiệm thu/04. Hồ sơ nghiệm thu/Phiếu NHẬN XÉT nghiệm thu.docx", "04. Hồ sơ nghiệm thu/Phiếu nhận xét nghiệm thu.docx"),
]


def copy_templates(root: Path, dest_root: Path) -> None:
    for rel_src, rel_dst in COPIES:
        src = root / rel_src
        dst = dest_root / rel_dst
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_bytes(src.read_bytes())


def main() -> None:
    root = paths.project_root()

    print("Dang doc du lieu tu Excel checklist...")
    info = excel_reader.load_project_data(root / excel_reader.CHECKLIST_FILENAME, SHEET_NAME)

    session = word_writer.Session()
    print(f"Che do ghi Word dang dung: {session.backend}")
    if session.backend == "docx":
        print(
            "  [LUU Y] Khong co Word COM - dung fallback python-docx thuan.\n"
            "  Mot vai cho xoa dong trong co the con sot dong trong, script se canh bao khi gap."
        )

    dest_dir_name = f"Hồ sơ - {info.title} ({info.year})"
    dest_root = root / dest_dir_name

    print(f"Dang sao chep file mau vao '{dest_dir_name}'...")
    copy_templates(root, dest_root)

    try:
        print("Dang sinh ho so dao duc...")
        section_dao_duc.generate(session, dest_root / "01. Hồ sơ đạo đức đề cương", info, TITLE_OLD)

        print("Dang sinh ho so khoa hoc de cuong...")
        section_khoa_hoc.generate(session, dest_root / "02. Hồ sơ khoa học đề cương", info, TITLE_OLD)

        print("Dang sinh cong van moi chuyen gia...")
        section_moi_chuyen_gia.generate(session, dest_root / "03. Công văn mời chuyên gia", info, TITLE_OLD)

        print("Dang sinh ho so nghiem thu...")
        section_nghiem_thu.generate(session, dest_root / "04. Hồ sơ nghiệm thu", info)
    finally:
        session.quit()

    print(f"XONG. Bo ho so da tao tai: {dest_root}")


if __name__ == "__main__":
    main()
