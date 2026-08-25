from pathlib import Path
from typing import List, Tuple

MAU_SUFFIX = " - MẪU"
LOCK_FILE_PREFIX = "~$"


def derive_dest_path(src_path: str) -> str:
    """Suy ra đường dẫn đích bằng cách bỏ hậu tố MAU_SUFFIX khỏi thư mục gốc.

    Quy ước đặt tên: thư mục mẫu nguồn phải có dạng "<...>{MAU_SUFFIX}", còn
    tên file bên trong phải giống hệt tên file sẽ xuất hiện trong hồ sơ đầu
    ra.
    """
    top, _, rest = src_path.partition("/")
    if MAU_SUFFIX not in top:
        raise ValueError(
            f"Thư mục mẫu '{top}' không theo đúng quy ước đặt tên "
            f"(phải có hậu tố '{MAU_SUFFIX}') nên không thể tự suy ra đường dẫn đích."
        )
    dest_top = top.replace(MAU_SUFFIX, "")
    return f"{dest_top}/{rest}" if rest else dest_top


def _discover_mau_dirs(root: Path) -> List[Path]:
    return sorted(
        (p for p in root.iterdir() if p.is_dir() and p.name.endswith(MAU_SUFFIX)),
        key=lambda p: p.name,
    )


def discover_copies(root: Path) -> List[Tuple[str, str]]:
    """Quét mọi thư mục "- MẪU" ở gốc dự án, trả về các file .docx cần copy.

    Đường dẫn đích tự suy ra bằng derive_dest_path - không cần khai báo thủ
    công trong Excel nữa: thêm file mẫu .docx đúng chỗ vào thư mục "- MẪU"
    tương ứng là công cụ tự nhận ra.
    """
    copies = []
    for mau_dir in _discover_mau_dirs(root):
        for file_path in sorted(mau_dir.glob("*.docx"), key=lambda p: p.name):
            if file_path.name.startswith(LOCK_FILE_PREFIX):
                continue
            rel_src = f"{mau_dir.name}/{file_path.name}"
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
            doc_files.append(f"{mau_dir.name}/{file_path.name}")
    return doc_files
