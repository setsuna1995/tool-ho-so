import shutil

import docx
import pytest

import excel_reader
import paths
import section_khoa_hoc
import word_writer

CHECKLIST_PATH = paths.project_root() / excel_reader.CHECKLIST_FILENAME
SHEET_VIAM = "Đề tài - Bánh ăn dặm VIAM 2027"
TITLE_OLD = (
    "Đánh giá hiệu quả sản phẩm sữa dinh dưỡng pha sẵn KUN DOCTOR COLOSTRUM lên "
    "tình trạng dinh dưỡng, miễn dịch, tiêu hóa và giấc ngủ của trẻ từ 24 đến 72 tháng tuổi"
)

SOURCE_DIR = paths.project_root() / "02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM"
RENAME_MAP = {
    "05. Dr.Kun QD TLHDKH đề cương.docx": "05. QĐ TLHĐ khoa học xét đề cương.docx",
    "06. Dr.Kun Bien ban hop thong qua de cuong de tai.docx": "06. BB họp thông qua đề cương.docx",
    "07. Dr.Kun Bien ban kiem phieu thong qua de cuong.docx": "07. BB kiểm phiếu thông qua đề cương.docx",
    "08. Dr.Kun QĐ phe-duyet-de-tai.docx": "08. QĐ phê duyệt đề tài.docx",
    "Dr.Kun Phieu cham diem HD de cuong.docx": "Phiếu chấm điểm HĐ đề cương.docx",
    "Dr.Kun Phieu nhan xet danh gia ho so.docx": "Phiếu nhận xét đánh giá hồ sơ.docx",
}


@pytest.fixture()
def dest_dir(tmp_path):
    for src_name, dst_name in RENAME_MAP.items():
        shutil.copy2(SOURCE_DIR / src_name, tmp_path / dst_name)
    return tmp_path


@pytest.fixture()
def info():
    return excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)


def test_generate_fixes_proposal_secretary_org(dest_dir, info):
    session = word_writer.Session(force_backend="docx")
    try:
        section_khoa_hoc.generate(session, dest_dir, info, TITLE_OLD)
    finally:
        session.quit()

    doc = docx.Document(str(dest_dir / "05. QĐ TLHĐ khoa học xét đề cương.docx"))
    secretary_table = doc.tables[2]
    assert secretary_table.cell(0, 1).text.strip() == "Hoàng Hà Linh"
    assert secretary_table.cell(0, 2).text.strip() == "Trung tâm NCKH - Viện VIAM"
