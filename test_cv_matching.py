# test_cv_matching.py
import pytest

import cv_matching


def _make_cv_dir(tmp_path, filenames):
    cv_dir = tmp_path / "Lý lịch khoa học"
    cv_dir.mkdir()
    for name in filenames:
        (cv_dir / name).write_text("x")
    return cv_dir


def test_finds_exact_match_with_diacritics(tmp_path):
    cv_dir = _make_cv_dir(tmp_path, ["Lý lịch khoa học - Trương Hồng Sơn.docx"])
    result = cv_matching.find_cv_file(cv_dir, "Trương Hồng Sơn")
    assert result.name == "Lý lịch khoa học - Trương Hồng Sơn.docx"


def test_falls_back_to_diacritics_insensitive_match_when_no_exact_match(tmp_path):
    cv_dir = _make_cv_dir(tmp_path, ["TM-Gs. Nguyen Cong Khan.pdf"])
    result = cv_matching.find_cv_file(cv_dir, "Nguyễn Công Khẩn")
    assert result.name == "TM-Gs. Nguyen Cong Khan.pdf"


def test_prefers_exact_diacritics_match_over_ambiguous_stripped_match(tmp_path):
    """Bo dau se khien 'Đặng Thị Bình' va 'Đăng Thị Bình' trung nhau (ca hai
    deu rut gon ve 'dang thi binh') - vong khop dung dau phai phan biet
    duoc 2 nguoi nay, khong duoc roi xuong vong bo dau va bao loi mo ho."""
    cv_dir = _make_cv_dir(tmp_path, ["CV Đặng Thị Bình.docx", "CV Đăng Thị Bình.docx"])
    result = cv_matching.find_cv_file(cv_dir, "Đặng Thị Bình")
    assert result.name == "CV Đặng Thị Bình.docx"


def test_raises_when_no_file_matches_at_either_tier(tmp_path):
    cv_dir = _make_cv_dir(tmp_path, ["TM-Gs. Nguyen Cong Khan.pdf"])
    with pytest.raises(FileNotFoundError, match="Lưu Liên Hương"):
        cv_matching.find_cv_file(cv_dir, "Lưu Liên Hương")


def test_raises_when_multiple_files_match_with_no_diacritics_to_disambiguate(tmp_path):
    cv_dir = _make_cv_dir(tmp_path, ["CV Nguyen Van A - ban 1.docx", "CV Nguyen Van A - ban 2.docx"])
    with pytest.raises(FileNotFoundError, match="Nguyễn Văn A"):
        cv_matching.find_cv_file(cv_dir, "Nguyễn Văn A")


def test_context_appears_in_error_message(tmp_path):
    cv_dir = _make_cv_dir(tmp_path, [])
    with pytest.raises(FileNotFoundError, match="chủ nhiệm đề tài"):
        cv_matching.find_cv_file(cv_dir, "Ai Đó", context=" (chủ nhiệm đề tài)")


def test_requires_contiguous_word_order_not_scrambled(tmp_path):
    cv_dir = _make_cv_dir(tmp_path, ["Van Khan Cong Nguyen.pdf"])
    with pytest.raises(FileNotFoundError):
        cv_matching.find_cv_file(cv_dir, "Nguyễn Công Khẩn")
