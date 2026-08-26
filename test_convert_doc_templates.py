import pytest

import convert_doc_templates as cdt
import paths
import word_writer

pytestmark = pytest.mark.skipif(
    not word_writer.com_available(), reason="Can Word COM de chay test nay"
)


def _make_doc_fixture(session: word_writer.Session, docx_src, doc_dst) -> None:
    """Sinh fixture .doc tu ban .docx con song song, dung dinh dang Word cu.

    Repo khong con giu file .doc thuc te nao (Task 6 da xoa toan bo .doc du
    thua), nen test tu tao fixture .doc rieng bang SaveAs2 thay vi copy tu
    mot file .doc co san trong repo.
    """
    doc = session.open(docx_src)
    doc.handle.SaveAs2(str(doc_dst), FileFormat=word_writer.WD_FORMAT_DOC)
    doc.handle.Close()


def test_convert_one_preserves_table_count(tmp_path):
    root = paths.project_root()
    docx_src = (
        root / "01. Hồ sơ đạo đức đề cương - MẪU" / "Bảng kiểm đánh giá đạo đức.docx"
    )
    tmp_src = tmp_path / "Bảng kiểm đánh giá đạo đức.doc"
    tmp_dst = tmp_src.with_suffix(".docx")

    session = word_writer.Session(force_backend="com")
    try:
        _make_doc_fixture(session, docx_src, tmp_src)
        before, after = cdt.convert_one(session, tmp_src, tmp_dst)
    finally:
        session.quit()

    assert tmp_dst.exists()
    assert before == after
