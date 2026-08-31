from pathlib import Path
from typing import List, Tuple

MAU_SUFFIX = " - MẪU"
LOCK_FILE_PREFIX = "~$"


def derive_dest_path(src_path: str) -> str:
    """Suy ra đường dẫn đích bằng cách bỏ hậu tố MAU_SUFFIX khỏi thư mục gốc (và bỏ prefix templates/ nếu có).

    Quy ước đặt tên: thư mục mẫu nguồn phải có dạng "<...>{MAU_SUFFIX}", còn
    tên file bên trong phải giống hệt tên file sẽ xuất hiện trong hồ sơ đầu
    ra.
    """
    clean_src = src_path.replace("templates/", "").replace("templates\\", "")
    top, _, rest = clean_src.partition("/")
    if not rest:
        top, _, rest = clean_src.partition("\\")
    if MAU_SUFFIX not in top:
        raise ValueError(
            f"Thư mục mẫu '{top}' không theo đúng quy ước đặt tên "
            f"(phải có hậu tố '{MAU_SUFFIX}') nên không thể tự suy ra đường dẫn đích."
        )
    dest_top = top.replace(MAU_SUFFIX, "")
    return f"{dest_top}/{rest}" if rest else dest_top


def _get_templates_root(root: Path) -> Path:
    tpl_dir = root / "templates"
    return tpl_dir if tpl_dir.exists() else root


def _discover_mau_dirs(root: Path) -> List[Path]:
    search_root = _get_templates_root(root)
    return sorted(
        (p for p in search_root.iterdir() if p.is_dir() and p.name.endswith(MAU_SUFFIX)),
        key=lambda p: p.name,
    )


def discover_copies(root: Path) -> List[Tuple[str, str]]:
    """Quét mọi thư mục "- MẪU" (trong templates/ hoặc gốc dự án), trả về các file .docx cần copy."""
    copies = []
    for mau_dir in _discover_mau_dirs(root):
        for file_path in sorted(mau_dir.glob("*.docx"), key=lambda p: p.name):
            if file_path.name.startswith(LOCK_FILE_PREFIX):
                continue
            rel_src = str(file_path.relative_to(root)).replace("\\", "/")
            copies.append((rel_src, derive_dest_path(rel_src)))
    return copies


def discover_doc_files(root: Path) -> List[str]:
    """Quét mọi thư mục "- MẪU", trả về file .doc chưa có bản .docx song song."""
    doc_files = []
    for mau_dir in _discover_mau_dirs(root):
        for file_path in sorted(mau_dir.glob("*.doc"), key=lambda p: p.name):
            if file_path.name.startswith(LOCK_FILE_PREFIX):
                continue
            if file_path.with_suffix(".docx").exists():
                continue
            rel_src = str(file_path.relative_to(root)).replace("\\", "/")
            doc_files.append(rel_src)
    return doc_files

