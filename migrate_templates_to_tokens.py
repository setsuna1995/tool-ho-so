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

DAO_DUC = "01. Hồ sơ đạo đức đề cương - MẪU"
KHOA_HOC = "02. Hồ sơ khoa học đề cương - MẪU"
MOI_CHUYEN_GIA = "03. Công văn mời chuyên gia - MẪU"
NGHIEM_THU = "04. Hồ sơ nghiệm thu - MẪU"

# Moi entry: (duong dan tuong doi tinh tu goc du an, [(text cu, text moi da co token), ...])
MIGRATIONS = {
    f"{DAO_DUC}/00. QĐ Giao đề tài.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
    ],
    f"{DAO_DUC}/01. QĐTLHĐ đạo đức đề cương.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
    ],
    f"{DAO_DUC}/02. BB họp HĐ đạo đức.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Chủ nhiệm đề tài: Ts. Bs. Trương Hồng Sơn", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN}"),
        ("Cơ quan thực hiện đề tài: Viện Y học ứng dụng Việt Nam", f"Cơ quan thực hiện đề tài: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{DAO_DUC}/03. BB kiểm phiếu HĐ đạo đức.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Chủ nhiệm: Ts. Bs. Trương Hồng Sơn", f"Chủ nhiệm: {TOKEN_CHU_NHIEM_HO_TEN}"),
    ],
    f"{DAO_DUC}/04. QĐ chấp nhận đạo đức.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Chủ nhiệm đề tài: Ts. Bs. Trương Hồng Sơn", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN}"),
        ("Cơ quan thực hiện đề tài:  Viện Y học ứng dụng Việt Nam.", f"Cơ quan thực hiện đề tài:  {TOKEN_DON_VI_CHU_TRI}."),
    ],
    f"{DAO_DUC}/Bảng kiểm đánh giá đạo đức.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("2024", TOKEN_NAM),
        ("Chủ nhiệm đề tài: Ts. Bs. Trương Hồng Sơn", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN}"),
        ("Đơn vị chủ trì: Viện Y học ứng dụng Việt Nam", f"Đơn vị chủ trì: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{KHOA_HOC}/05. QĐ TLHĐ khoa học xét đề cương.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
    ],
    f"{KHOA_HOC}/06. BB họp thông qua đề cương.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
    ],
    f"{KHOA_HOC}/07. BB kiểm phiếu thông qua đề cương.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("Chủ nhiệm đề tài: Ts. Bs. Trương Hồng Sơn", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN}"),
        ("Đơn vị thực hiện: Viện Y học ứng dụng Việt nam", f"Đơn vị thực hiện: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{KHOA_HOC}/08. QĐ phê duyệt đề tài.docx": [
        ("2025", TOKEN_NAM),
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("- Chủ nhiệm đề tài: Ts. Bs. Trương Hồng Sơn", f"- Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN}"),
        ("- Đơn vị thực hiện đề tài: Viện Y học ứng dụng Việt Nam", f"- Đơn vị thực hiện đề tài: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{KHOA_HOC}/Phiếu chấm điểm HĐ đề cương.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("2. Chủ nhiệm Đề tài: Ts. Bs. Trương Hồng Sơn", f"2. Chủ nhiệm Đề tài: {TOKEN_CHU_NHIEM_HO_TEN}"),
        ("3. Đơn vị chủ trì đề tài:  Viện Y học ứng dụng Việt Nam", f"3. Đơn vị chủ trì đề tài:  {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{KHOA_HOC}/Phiếu nhận xét đánh giá hồ sơ.docx": [
        ("2024", TOKEN_NAM),
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("- Chủ nhiệm đề tài: Ts. Bs Trương Hồng Sơn ", f"- Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN} "),
        ("- Đơn vị chủ trì đề tài: Viện Y học ứng dụng Việt Nam", f"- Đơn vị chủ trì đề tài: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{MOI_CHUYEN_GIA}/Công văn mời chuyên gia.docx": [
        (TITLE_OLD, TOKEN_TEN_DE_TAI),
        ("2024", TOKEN_NAM),
        ("Chủ nhiệm đề tài: Ts. Bs. Trương Hồng Sơn.", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_HO_TEN}."),
        ("Thư ký đề tài: Ths. Lưu Liên Hương.", f"Thư ký đề tài: {TOKEN_THU_KY_DE_TAI}."),
        ("Đơn vị thực hiện đề tài: Viện Y học ứng dụng Việt Nam.", f"Đơn vị thực hiện đề tài: {TOKEN_DON_VI_CHU_TRI}."),
    ],
    f"{NGHIEM_THU}/9. Quyết định thành lập HĐ nghiệm thu.docx": [
        ("20xx", TOKEN_NAM),
        ("“Tên đề tài”", f"“{TOKEN_TEN_DE_TAI}”"),
    ],
    f"{NGHIEM_THU}/10. Biên bản họp HĐ nghiệm thu.docx": [
        ("20xx", TOKEN_NAM),
        ("1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_TEN}"),
        ("Đồng chủ nhiệm đề tài: Tên 2", f"Đồng chủ nhiệm đề tài: {TOKEN_DONG_CHU_NHIEM_TEN}"),
        ("Đơn vị chủ trì: Viện Y học ứng dụng Việt Nam.", f"Đơn vị chủ trì: {TOKEN_DON_VI_CHU_TRI}."),
    ],
    f"{NGHIEM_THU}/11. Biên bản kiểm phiếu nghiệm thu.docx": [
        ("20xx", TOKEN_NAM),
        ("1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {TOKEN_CHU_NHIEM_TEN}"),
        ("Đồng chủ nhiệm đề tài: Tên 2", f"Đồng chủ nhiệm đề tài: {TOKEN_DONG_CHU_NHIEM_TEN}"),
        ("Đơn vị chủ trì: Viện Y học ứng dụng Việt Nam", f"Đơn vị chủ trì: {TOKEN_DON_VI_CHU_TRI}"),
    ],
    f"{NGHIEM_THU}/12. Quyết định công nhận kết quả đề tài.docx": [
        ("20XX", TOKEN_NAM),
        ("20xx", TOKEN_NAM),
        ("“Tên đề tài”", f"“{TOKEN_TEN_DE_TAI}”"),
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
        ("20xx", TOKEN_NAM),
        ("Tên đề tài: Tên đề tài", f"Tên đề tài: {TOKEN_TEN_DE_TAI}"),
        ("Chủ nhiệm: Tên 1", f"Chủ nhiệm: {TOKEN_CHU_NHIEM_TEN}"),
        ("Đồng chủ nhiệm: Tên 2", f"Đồng chủ nhiệm: {TOKEN_DONG_CHU_NHIEM_TEN}"),
        ("Đơn vị chủ trì: Viện Y học ứng dụng Việt Nam", f"Đơn vị chủ trì: {TOKEN_DON_VI_CHU_TRI}"),
    ],
}


def apply_mapping(path: Path, mapping: list) -> None:
    session = word_writer.Session(force_backend="docx")
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
