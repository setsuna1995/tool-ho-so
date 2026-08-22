import shutil

import docx
import pytest

import excel_reader
import paths
import section_moi_chuyen_gia
import word_writer

CHECKLIST_PATH = paths.project_root() / excel_reader.CHECKLIST_FILENAME
SHEET_VIAM = "Đề tài - Bánh ăn dặm VIAM 2027"
TITLE_OLD = (
    "Đánh giá hiệu quả sản phẩm sữa dinh dưỡng pha sẵn KUN DOCTOR COLOSTRUM lên "
    "tình trạng dinh dưỡng, miễn dịch, tiêu hóa và giấc ngủ của trẻ từ 24 đến 72 tháng tuổi"
)

SOURCE_FILE = paths.project_root() / "03. CV mời chuyên gia - mẫu COLOSTRUM" / "CV mời chuyên gia.docx"


@pytest.fixture()
def dest_dir(tmp_path):
    shutil.copy2(SOURCE_FILE, tmp_path / "Công văn mời chuyên gia.docx")
    return tmp_path


@pytest.fixture()
def info():
    return excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)


def test_generate_replaces_title_and_removes_colostrum_intro(dest_dir, info):
    session = word_writer.Session(force_backend="docx")
    try:
        section_moi_chuyen_gia.generate(session, dest_dir, info, TITLE_OLD)
    finally:
        session.quit()

    doc = docx.Document(str(dest_dir / "Công văn mời chuyên gia.docx"))
    full_text = "\n".join(p.text for p in doc.paragraphs)
    assert info.title in full_text
    assert "LOF KUN COLOSTRUM" not in full_text
    assert f"01/{info.year} đến 12/{info.year}" in full_text
