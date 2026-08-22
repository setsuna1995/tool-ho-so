import shutil
from pathlib import Path

import docx
import pytest

import excel_reader
import paths
import section_dao_duc
import word_writer

CHECKLIST_PATH = paths.project_root() / excel_reader.CHECKLIST_FILENAME
SHEET_VIAM = "Đề tài - Bánh ăn dặm VIAM 2027"
TITLE_OLD = (
    "Đánh giá hiệu quả sản phẩm sữa dinh dưỡng pha sẵn KUN DOCTOR COLOSTRUM lên "
    "tình trạng dinh dưỡng, miễn dịch, tiêu hóa và giấc ngủ của trẻ từ 24 đến 72 tháng tuổi"
)

SOURCE_DIR = paths.project_root() / "01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM"
FILES = [
    "00. QĐ Giao đề tài.docx",
    "01. QĐTLHĐ đạo đức đề cương.docx",
    "02. BB họp HĐ đạo đức - KUN COLOSTRUM.docx",
    "03. BB kiểm phiếu HĐ đạo đức.docx",
    "04. Dr.Kun QĐ chấp nhận đạo đức.docx",
    "Bảng kiểm đánh giá đạo đức.docx",
]
RENAME_MAP = {
    "02. BB họp HĐ đạo đức - KUN COLOSTRUM.docx": "02. BB họp HĐ đạo đức.docx",
    "04. Dr.Kun QĐ chấp nhận đạo đức.docx": "04. QĐ chấp nhận đạo đức.docx",
}


@pytest.fixture()
def dest_dir(tmp_path):
    for filename in FILES:
        dst_name = RENAME_MAP.get(filename, filename)
        shutil.copy2(SOURCE_DIR / filename, tmp_path / dst_name)
    return tmp_path


@pytest.fixture()
def info():
    return excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)


def test_generate_fixes_ethics_secretary_org(dest_dir, info):
    session = word_writer.Session(force_backend="docx")
    try:
        section_dao_duc.generate(session, dest_dir, info, TITLE_OLD)
    finally:
        session.quit()

    doc = docx.Document(str(dest_dir / "01. QĐTLHĐ đạo đức đề cương.docx"))
    secretary_table = doc.tables[2]
    assert secretary_table.cell(0, 0).text.strip() == "Hoàng Hà Linh"
    assert secretary_table.cell(0, 1).text.strip() == "Trung tâm NCKH - Viện VIAM"


def test_generate_replaces_title_everywhere(dest_dir, info):
    session = word_writer.Session(force_backend="docx")
    try:
        section_dao_duc.generate(session, dest_dir, info, TITLE_OLD)
    finally:
        session.quit()

    for filename in ["00. QĐ Giao đề tài.docx", "03. BB kiểm phiếu HĐ đạo đức.docx"]:
        doc = docx.Document(str(dest_dir / filename))
        full_text = "\n".join(p.text for p in doc.paragraphs)
        assert info.title in full_text
        assert TITLE_OLD not in full_text
