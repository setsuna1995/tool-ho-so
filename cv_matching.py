# cv_matching.py
import re
import unicodedata
from pathlib import Path


def _normalize(text: str, strip_diacritics: bool) -> str:
    if strip_diacritics:
        text = text.replace("đ", "d").replace("Đ", "D")
        text = "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))
    return re.sub(r"[^\w]+", " ", text, flags=re.UNICODE).strip().lower()


def _search(cv_dir: Path, person_name: str, strip_diacritics: bool):
    target = _normalize(person_name, strip_diacritics)
    candidates = sorted((f for f in cv_dir.iterdir() if f.is_file()), key=lambda f: f.name)
    return [f for f in candidates if target in _normalize(f.stem, strip_diacritics)]


def _not_found_error(cv_dir: Path, person_name: str, context: str) -> FileNotFoundError:
    return FileNotFoundError(
        f"Không tìm thấy file CV nào khớp tên '{person_name}'{context} trong thư mục "
        f"'{cv_dir.name}/'. Vui lòng đặt file CV có tên chứa '{person_name}' vào thư mục đó."
    )


def _ambiguous_error(cv_dir: Path, person_name: str, context: str, matches) -> FileNotFoundError:
    names = ", ".join(f"'{m.name}'" for m in matches)
    return FileNotFoundError(
        f"Tìm thấy nhiều hơn 1 file CV khớp tên '{person_name}'{context} trong thư mục "
        f"'{cv_dir.name}/': {names}. Vui lòng đổi tên file để chỉ còn đúng 1 file khớp."
    )


def find_cv_file(cv_dir: Path, person_name: str, context: str = "") -> Path:
    """Tim file trong cv_dir co ten (khong ke phan mo rong) chua cum
    `person_name` LIEN NHAU, dung thu tu. Uu tien khop DUNG DAU truoc (chi
    chuan hoa hoa/thuong + khoang trang, giu nguyen dau tieng Viet) - chi
    khi khong file nao khop dung dau moi thu lai sau khi bo dau (de van
    khop duoc file dat ten khong dau nhu "TM-Gs. Nguyen Cong Khan.pdf").
    Neu vong khop dung dau ra >1 ket qua, bao loi mo ho ngay, khong roi
    xuong vong bo dau (vi bo dau se khong lam het mo ho, chi lam mo ho
    hon). Nem FileNotFoundError neu 0 hoac >1 file khop o vong duoc dung."""
    exact_matches = _search(cv_dir, person_name, strip_diacritics=False)
    if len(exact_matches) == 1:
        return exact_matches[0]
    if len(exact_matches) > 1:
        raise _ambiguous_error(cv_dir, person_name, context, exact_matches)

    loose_matches = _search(cv_dir, person_name, strip_diacritics=True)
    if not loose_matches:
        raise _not_found_error(cv_dir, person_name, context)
    if len(loose_matches) > 1:
        raise _ambiguous_error(cv_dir, person_name, context, loose_matches)
    return loose_matches[0]
