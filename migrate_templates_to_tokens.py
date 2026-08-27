from pathlib import Path

import word_writer

TITLE_OLD = (
    "Đánh giá hiệu quả sản phẩm sữa dinh dưỡng pha sẵn KUN DOCTOR COLOSTRUM lên "
    "tình trạng dinh dưỡng, miễn dịch, tiêu hóa và giấc ngủ của trẻ từ 24 đến 72 tháng tuổi"
)

TOKEN_TEN_DE_TAI = "{{TEN_DE_TAI}}"
TOKEN_NAM = "{{NAM}}"
TOKEN_DON_VI_CHU_TRI = "{{DON_VI_CHU_TRI}}"
TOKEN_CHU_NHIEM_HO_TEN = "{{CHU_NHIEM_HO_TEN}}"
TOKEN_CHU_NHIEM_TEN = "{{CHU_NHIEM_TEN}}"
TOKEN_DONG_CHU_NHIEM_TEN = "{{DONG_CHU_NHIEM_TEN}}"
TOKEN_THU_KY_DE_TAI = "{{THU_KY_DE_TAI}}"
TOKEN_THOI_GIAN_BAT_DAU = "{{THOI_GIAN_BAT_DAU}}"
TOKEN_THOI_GIAN_KET_THUC = "{{THOI_GIAN_KET_THUC}}"
TOKEN_DIA_DIEM_TRIEN_KHAI = "{{DIA_DIEM_TRIEN_KHAI}}"

DAO_DUC = "01. Hồ sơ đạo đức đề cương - MẪU"
KHOA_HOC = "02. Hồ sơ khoa học đề cương - MẪU"
MOI_CHUYEN_GIA = "03. Công văn mời chuyên gia - MẪU"
NGHIEM_THU = "04. Hồ sơ nghiệm thu - MẪU"

# Moi entry: (duong dan tuong doi tinh tu goc du an, [(text cu, text moi da co token), ...])
#
# QUAN TRONG ve thu tu trong moi danh sach: moi entry thay mot cau/cum tu CU
# THE phai dung TRUOC entry thay nam tran ("2024"/"2025"/"20xx"/"20XX") trong
# cung mot file. Ly do: entry nam tran khong neo ngu canh - no thay MOI cho
# xuat hien cua chuoi nam do trong toan bo file (ke ca nam nam lot trong mot
# cau khac), nen neu chay truoc, no se "an" mat chuoi cu the ma cac ham
# generate() trong section_*.py (Task 9) con tim bang session.replace_text()
# sau nay, khien cau do khong con duoc dien du lieu (loi am tham, khong crash).
#
# QUAN TRONG ve pham vi tim-thay cua tung entry ("neo nhan" vs "chi gia tri"):
# Mot so entry chi tim-thay CHI GIA TRI (vd "Ts. Bs. Trương Hồng Sơn") thay vi
# ca cum "Nhan: Gia tri" - dieu nay CHI an toan khi gia tri do xuat hien DUY
# NHAT MOT LAN trong toan bo file (kiem tra bang script rieng, xem bao cao
# Task 8 fix round 3). Khi gia tri xuat hien nhieu hon 1 lan trong file (vd
# "Viện Y học ứng dụng Việt Nam" lap lai o letterhead/chu ky), entry PHAI giu
# nguyen dang neo nhan ("Nhan: Gia tri") de tranh thay nham nhung cho khac -
# viec nay lam nhan bi mat dinh dang rieng (vd chu dam) trong bang nay, nhung
# do la dong bo da co tu truoc khi Task 8 chay (code section_*.py cu da tung
# thay ca cum "Nhan: Gia tri" nay qua COM khi sinh ho so, nen no CHUA BAO GIO
# giu duoc dinh dang rieng cho nhan - khong phai loi Task 8 gay ra).
MIGRATIONS = {
    f"{DAO_DUC}/00. QĐ Giao đề tài.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("2024", TOKEN_NAM),
    ],
    f"{DAO_DUC}/01. QĐTLHĐ đạo đức đề cương.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("2024", TOKEN_NAM),
    ],
    f"{DAO_DUC}/02. BB họp HĐ đạo đức.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Ts. Bs. Trương Hồng Sơn", TOKEN_CHU_NHIEM_HO_TEN),
        # "Viện Y học ứng dụng Việt Nam" xuat hien 3 lan trong file nay - giu
        # dang neo nhan.
        ("Cơ quan thực hiện đề tài: Viện Y học ứng dụng Việt Nam", f"Cơ quan thực hiện đề tài: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{DAO_DUC}/03. BB kiểm phiếu HĐ đạo đức.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Ts. Bs. Trương Hồng Sơn", TOKEN_CHU_NHIEM_HO_TEN),
        ("2024", TOKEN_NAM),
    ],
    f"{DAO_DUC}/04. QĐ chấp nhận đạo đức.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Ts. Bs. Trương Hồng Sơn", TOKEN_CHU_NHIEM_HO_TEN),
        # "Viện Y học ứng dụng Việt Nam" xuat hien 7 lan trong file nay (khong
        # chi trong dong "Co quan thuc hien de tai") - giu dang neo nhan.
        ("Cơ quan thực hiện đề tài:  Viện Y học ứng dụng Việt Nam.", f"Cơ quan thực hiện đề tài:  {TOKEN_DON_VI_CHU_TRI}."),
        ("Địa điểm triển khai nghiên cứu: tỉnh Thái Nguyên.", f"Địa điểm triển khai nghiên cứu: {TOKEN_DIA_DIEM_TRIEN_KHAI}."),
        (
            "Thời gian nghiên cứu: Từ 12/2024 đến 05/2024",
            f"Thời gian nghiên cứu: Từ {TOKEN_THOI_GIAN_BAT_DAU} đến {TOKEN_THOI_GIAN_KET_THUC}",
        ),
        ("2024", TOKEN_NAM),
    ],
    f"{DAO_DUC}/Bảng kiểm đánh giá đạo đức.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Ts. Bs. Trương Hồng Sơn", TOKEN_CHU_NHIEM_HO_TEN),
        ("Viện Y học ứng dụng Việt Nam", TOKEN_DON_VI_CHU_TRI),
        ("2024", TOKEN_NAM),
    ],
    f"{KHOA_HOC}/05. QĐ TLHĐ khoa học xét đề cương.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("2024", TOKEN_NAM),
    ],
    f"{KHOA_HOC}/06. BB họp thông qua đề cương.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("2024", TOKEN_NAM),
    ],
    f"{KHOA_HOC}/07. BB kiểm phiếu thông qua đề cương.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Ts. Bs. Trương Hồng Sơn", TOKEN_CHU_NHIEM_HO_TEN),
        ("Viện Y học ứng dụng Việt nam", TOKEN_DON_VI_CHU_TRI),
        ("2024", TOKEN_NAM),
    ],
    f"{KHOA_HOC}/08. QĐ phê duyệt đề tài.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Ts. Bs. Trương Hồng Sơn", TOKEN_CHU_NHIEM_HO_TEN),
        # "Viện Y học ứng dụng Việt Nam" xuat hien 6 lan trong file nay - giu
        # dang neo nhan.
        ("- Đơn vị thực hiện đề tài: Viện Y học ứng dụng Việt Nam", f"- Đơn vị thực hiện đề tài: {TOKEN_DON_VI_CHU_TRI}"),
        (
            "Thời gian thực hiện của đề tài: từ tháng 12/2024 đến tháng 05/2025",
            f"Thời gian thực hiện của đề tài: từ tháng {TOKEN_THOI_GIAN_BAT_DAU} đến tháng {TOKEN_THOI_GIAN_KET_THUC}",
        ),
        ("2025", TOKEN_NAM),
        ("2024", TOKEN_NAM),
    ],
    f"{KHOA_HOC}/Phiếu chấm điểm HĐ đề cương.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Ts. Bs. Trương Hồng Sơn", TOKEN_CHU_NHIEM_HO_TEN),
        ("Viện Y học ứng dụng Việt Nam", TOKEN_DON_VI_CHU_TRI),
        ("2024", TOKEN_NAM),
    ],
    f"{KHOA_HOC}/Phiếu nhận xét đánh giá hồ sơ.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Ts. Bs Trương Hồng Sơn", TOKEN_CHU_NHIEM_HO_TEN),
        ("Viện Y học ứng dụng Việt Nam", TOKEN_DON_VI_CHU_TRI),
        ("2024", TOKEN_NAM),
    ],
    f"{MOI_CHUYEN_GIA}/Công văn mời chuyên gia.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Ts. Bs. Trương Hồng Sơn", TOKEN_CHU_NHIEM_HO_TEN),
        ("Ths. Lưu Liên Hương", TOKEN_THU_KY_DE_TAI),
        # "Viện Y học ứng dụng Việt Nam" xuat hien 4 lan trong file nay - giu
        # dang neo nhan.
        ("Đơn vị thực hiện đề tài: Viện Y học ứng dụng Việt Nam.", f"Đơn vị thực hiện đề tài: {TOKEN_DON_VI_CHU_TRI}."),
        (
            "Nghiên cứu được triển khai trong 06 tháng, trong đó thời gian can thiệp là 04 tháng.",
            f"Thời gian thực hiện dự kiến: {TOKEN_THOI_GIAN_BAT_DAU} đến {TOKEN_THOI_GIAN_KET_THUC}.",
        ),
        (
            "Thời gian: 9 giờ 00 – sáng thứ 7 ngày 07 tháng 12 năm 2024.",
            f"Thời gian: …… giờ ……, ngày …… tháng …… năm {TOKEN_NAM}.",
        ),
        ("2024", TOKEN_NAM),
    ],
    f"{NGHIEM_THU}/9. Quyết định thành lập HĐ nghiệm thu.docx": [
        ("“Tên đề tài”", f"“{TOKEN_TEN_DE_TAI}”"),
        ("20xx", TOKEN_NAM),
    ],
    f"{NGHIEM_THU}/10. Biên bản họp HĐ nghiệm thu.docx": [
        ("1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_TEN}"),
        ("Đồng chủ nhiệm đề tài: Tên 2", f"Đồng chủ nhiệm đề tài: {TOKEN_DONG_CHU_NHIEM_TEN}"),
        # "Viện Y học ứng dụng Việt Nam" xuat hien 4 lan trong file nay - giu
        # dang neo nhan.
        ("Đơn vị chủ trì: Viện Y học ứng dụng Việt Nam.", f"Đơn vị chủ trì: {TOKEN_DON_VI_CHU_TRI}."),
        ("20xx", TOKEN_NAM),
    ],
    f"{NGHIEM_THU}/11. Biên bản kiểm phiếu nghiệm thu.docx": [
        ("1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_TEN}"),
        ("Đồng chủ nhiệm đề tài: Tên 2", f"Đồng chủ nhiệm đề tài: {TOKEN_DONG_CHU_NHIEM_TEN}"),
        # "Viện Y học ứng dụng Việt Nam" xuat hien 2 lan trong file nay - giu
        # dang neo nhan.
        ("Đơn vị chủ trì: Viện Y học ứng dụng Việt Nam", f"Đơn vị chủ trì: {TOKEN_DON_VI_CHU_TRI}"),
        ("20xx", TOKEN_NAM),
    ],
    f"{NGHIEM_THU}/12. Quyết định công nhận kết quả đề tài.docx": [
        ("“Tên đề tài”", f"“{TOKEN_TEN_DE_TAI}”"),
        # Gop "20XX" va "20xx" thanh MOT entry: COM Find.Execute chay khong
        # phan biet hoa/thuong (MatchCase=False), nen entry dau se da khop ca
        # hai bien the chu hoa/thuong - entry thu hai se khong con gi de thay
        # va lam apply_mapping bao RuntimeError.
        ("20xx", TOKEN_NAM),
    ],
    f"{NGHIEM_THU}/Phiếu chấm điểm nghiệm thu (TNLS).docx": [
        ("1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_TEN}"),
    ],
    f"{NGHIEM_THU}/Phiếu chấm điểm nghiệm thu (TVCT_ĐGHQ).docx": [
        ("1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_TEN}"),
    ],
    f"{NGHIEM_THU}/Phiếu ký nhận tiền.docx": [
        (
            "“Đánh giá hiệu quả sản phẩm thực phẩm chức năng Viên nang Đông trùng hạ thảo CordySen”",
            f"“{TOKEN_TEN_DE_TAI}”",
        ),
    ],
    f"{NGHIEM_THU}/Phiếu nhận xét nghiệm thu.docx": [
        ("Tên đề tài: Tên đề tài", f"Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm: Tên 1", f"Chủ nhiệm: {TOKEN_CHU_NHIEM_TEN}"),
        ("Đồng chủ nhiệm: Tên 2", f"Đồng chủ nhiệm: {TOKEN_DONG_CHU_NHIEM_TEN}"),
        # "Viện Y học ứng dụng Việt Nam" xuat hien 2 lan trong file nay - giu
        # dang neo nhan.
        ("Đơn vị chủ trì: Viện Y học ứng dụng Việt Nam", f"Đơn vị chủ trì: {TOKEN_DON_VI_CHU_TRI}"),
        ("20xx", TOKEN_NAM),
    ],
}


def apply_mapping(path: Path, mapping: list) -> None:
    # Bat buoc dung backend COM: fallback docx dung _docx_replace_in_paragraph,
    # ham nay ghi toan bo van ban moi vao run dau tien va xoa trang cac run con
    # lai - lam mat dinh dang ky tu (bold/italic/font size) cua nhung doan van
    # co nhieu run. COM's Find.Execute thay the tai cho va giu nguyen dinh dang
    # cac phan khong bi thay, dung nhu cach sinh ho so o production
    # (word_writer.Session() trong tao_ho_so_moi.py cung tu chon COM khi co
    # Word). Khong cho phep am tham roi ve docx - phai bao loi ro rang.
    session = word_writer.Session()
    if session.backend != "com":
        session.quit()
        raise RuntimeError(
            "Khong tim thay Word COM tren may nay - script nay BAT BUOC dung "
            "backend COM de giu nguyen dinh dang ky tu (bold/italic/font size) "
            "trong cac file mau khi migrate. Backend docx (fallback) se lam "
            "phang dinh dang cua moi doan van bi thay the. Cai/mo Microsoft "
            "Word roi chay lai."
        )
    try:
        doc = session.open(path)
        for old_text, new_text in mapping:
            found = session.replace_text(doc, old_text, new_text, warn_if_missing=False)
            if not found:
                raise RuntimeError(f"Khong tim thay chuoi can thay trong {path.name}: {old_text!r}")
        session.save_close(doc)
    finally:
        session.quit()


def migrate_all(root: Path) -> None:
    for rel_path, mapping in MIGRATIONS.items():
        apply_mapping(root / rel_path, mapping)
        print(f"Da migrate: {rel_path}")


if __name__ == "__main__":
    migrate_all(Path(__file__).resolve().parent)
    print("XONG. Da chuyen toan bo file mau sang dung token.")
