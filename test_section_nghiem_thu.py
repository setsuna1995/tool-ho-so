import shutil

import docx
import pytest

import excel_reader
import paths
import section_nghiem_thu
import word_writer

CHECKLIST_PATH = paths.project_root() / excel_reader.CHECKLIST_FILENAME
SHEET_VIAM = "Đề tài - Bánh ăn dặm VIAM 2027"

SOURCE_DIR = paths.project_root() / "04. Hồ sơ nghiệm thu" / "04. Hồ sơ nghiệm thu"
RENAME_MAP = {
    "9. Quyết định THÀNH LẬP HĐ nghiệm thu.docx": "9. Quyết định thành lập HĐ nghiệm thu.docx",
    "10. Biên bản HỌP HĐ nghiệm thu.docx": "10. Biên bản họp HĐ nghiệm thu.docx",
    "11. Biên bản KIỂM PHIẾU nghiệm thu.docx": "11. Biên bản kiểm phiếu nghiệm thu.docx",
    "12. Quyết định công nhận kết quả đề tài.docx": "12. Quyết định công nhận kết quả đề tài.docx",
    "Phiếu CHẤM ĐIỂM nghiệm thu-(TVCT_ĐGHQ).docx": "Phiếu chấm điểm nghiệm thu (TVCT_ĐGHQ).docx",
    "Phiếu ký nhận tiền.docx": "Phiếu ký nhận tiền.docx",
    "Phiếu NHẬN XÉT nghiệm thu.docx": "Phiếu nhận xét nghiệm thu.docx",
}


@pytest.fixture()
def dest_dir(tmp_path):
    for src_name, dst_name in RENAME_MAP.items():
        shutil.copy2(SOURCE_DIR / src_name, tmp_path / dst_name)
    return tmp_path


@pytest.fixture()
def info():
    return excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)


def test_generate_writes_correct_secretary_org(dest_dir, info):
    session = word_writer.Session(force_backend="docx")
    try:
        section_nghiem_thu.generate(session, dest_dir, info)
    finally:
        session.quit()

    doc = docx.Document(str(dest_dir / "9. Quyết định thành lập HĐ nghiệm thu.docx"))
    secretary_table = doc.tables[2]
    assert secretary_table.cell(0, 0).text.strip() == "1. Hoàng Hà Linh"
    assert secretary_table.cell(0, 1).text.strip() == "Viện Y học Ứng dụng Việt Nam"


def test_generate_uses_dynamic_member_count_not_hardcoded_05(dest_dir, info):
    session = word_writer.Session(force_backend="docx")
    try:
        section_nghiem_thu.generate(session, dest_dir, info)
    finally:
        session.quit()

    doc = docx.Document(str(dest_dir / "10. Biên bản họp HĐ nghiệm thu.docx"))
    full_text = "\n".join(p.text for p in doc.paragraphs)
    assert "là 05 người" in full_text
