from pathlib import Path

import paths
import template_config
import word_writer


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
    doc_files = template_config.discover_doc_files(root)
    session = word_writer.Session(force_backend="com")
    results = []
    try:
        for rel_path in doc_files:
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
