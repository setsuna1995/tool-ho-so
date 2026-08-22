import shutil

import pytest

import convert_doc_templates as cdt
import paths
import word_writer

pytestmark = pytest.mark.skipif(
    not word_writer.com_available(), reason="Can Word COM de chay test nay"
)


def test_convert_one_preserves_table_count(tmp_path):
    root = paths.project_root()
    fixture_src = root / "01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM" / "Bảng kiểm đánh giá đạo đức.doc"
    tmp_src = tmp_path / "Bảng kiểm đánh giá đạo đức.doc"
    shutil.copy2(fixture_src, tmp_src)
    tmp_dst = tmp_src.with_suffix(".docx")

    session = word_writer.Session(force_backend="com")
    try:
        before, after = cdt.convert_one(session, tmp_src, tmp_dst)
    finally:
        session.quit()

    assert tmp_dst.exists()
    assert before == after
