from pathlib import Path

import paths
import word_writer

DOC_FILES = [
    "01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM/Bảng kiểm đánh giá đạo đức.doc",
    "03. CV mời chuyên gia - mẫu COLOSTRUM/CV mời chuyên gia.doc",
    "04. Hồ sơ nghiệm thu/04. Hồ sơ nghiệm thu/9. Quyết định THÀNH LẬP HĐ nghiệm thu.doc",
    "04. Hồ sơ nghiệm thu/04. Hồ sơ nghiệm thu/10. Biên bản HỌP HĐ nghiệm thu.doc",
    "04. Hồ sơ nghiệm thu/04. Hồ sơ nghiệm thu/11. Biên bản KIỂM PHIẾU nghiệm thu.doc",
    "04. Hồ sơ nghiệm thu/04. Hồ sơ nghiệm thu/12. Quyết định công nhận kết quả đề tài.doc",
    "04. Hồ sơ nghiệm thu/04. Hồ sơ nghiệm thu/Phiếu NHẬN XÉT nghiệm thu.doc",
]


def convert_one(session: word_writer.Session, src: Path, dst: Path) -> tuple:
    doc = session.open(src)
    tables_before = doc.handle.Tables.Count
    doc.handle.SaveAs2(str(dst), FileFormat=word_writer.WD_FORMAT_DOCX)
    doc.handle.Close()

    check = session.open(dst)
    tables_after = check.handle.Tables.Count
    check.handle.Close()

    return tables_before, tables_after


def convert_all() -> list:
    if not word_writer.com_available():
        raise RuntimeError(
            "Can Microsoft Word tren may nay de chuyen doi file .doc sang .docx"
        )
    root = paths.project_root()
    session = word_writer.Session(force_backend="com")
    results = []
    try:
        for rel_path in DOC_FILES:
            src = root / rel_path
            dst = src.with_suffix(".docx")
            before, after = convert_one(session, src, dst)
            results.append((src.name, dst.name, before, after))
            if before != after:
                raise RuntimeError(
                    f"Convert '{src.name}' bi lech so bang: {before} -> {after}"
                )
    finally:
        session.quit()
    return results


if __name__ == "__main__":
    for src_name, dst_name, before, after in convert_all():
        print(f"{src_name} -> {dst_name} (bang: {before}/{after} khop)")
