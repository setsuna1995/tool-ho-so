from pathlib import Path

import committee_writer
import word_writer
from excel_reader import ProjectInfo

ROLES = ["Chủ tịch Hội đồng", "Phản biện 1", "Phản biện 2", "Ủy viên", "Ủy viên"]


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, title_old: str) -> None:
    _qdtlhd_khoa_hoc(session, dest_dir, info, title_old)
    _bb_hop_thong_qua_de_cuong(session, dest_dir, info, title_old)
    _bb_kiem_phieu_thong_qua_de_cuong(session, dest_dir, info, title_old)
    _qd_phe_duyet_de_tai(session, dest_dir, info, title_old)
    _phieu_cham_diem_hd_de_cuong(session, dest_dir, info, title_old)
    _phieu_nhan_xet_danh_gia_ho_so(session, dest_dir, info, title_old)


def _qdtlhd_khoa_hoc(session, dest_dir, info, title_old):
    doc = session.open(dest_dir / "05. QĐ TLHĐ khoa học xét đề cương.docx")
    session.replace_text(doc, title_old, info.title)
    session.replace_text(doc, "2024", str(info.year))
    committee_writer.write_committee_roster(
        session, doc, 2, info.proposal_committee, roles=ROLES, name_col=1, org_col=2, role_col=3
    )
    committee_writer.write_committee_secretaries(
        session, doc, 3, info.proposal_committee, name_col=2, org_col=3
    )
    session.save_close(doc)


def _bb_hop_thong_qua_de_cuong(session, dest_dir, info, title_old):
    doc = session.open(dest_dir / "06. BB họp thông qua đề cương.docx")
    session.replace_text(
        doc,
        f"1. Tên đề tài nghiên cứu khoa học: {title_old}.",
        f"1. Tên đề tài nghiên cứu khoa học: {info.title}.",
    )
    session.replace_text(doc, "2024", str(info.year))
    chair = info.proposal_committee.chair
    session.replace_text(
        doc,
        "PGs. Ts. Hoàng Thị Thanh - Chủ tịch Hội đồng điều khiển phiên họp",
        f"{chair.degree} {chair.name} - Chủ tịch Hội đồng điều khiển phiên họp".strip(),
    )
    session.save_close(doc)


def _bb_kiem_phieu_thong_qua_de_cuong(session, dest_dir, info, title_old):
    doc = session.open(dest_dir / "07. BB kiểm phiếu thông qua đề cương.docx")
    session.replace_text(doc, f"Tên đề tài: {title_old}.", f"Tên đề tài: {info.title}.")
    session.replace_text(doc, "2024", str(info.year))
    session.save_close(doc)


def _qd_phe_duyet_de_tai(session, dest_dir, info, title_old):
    doc = session.open(dest_dir / "08. QĐ phê duyệt đề tài.docx")
    session.replace_text(doc, f"“{title_old}”.", f"“{info.title}”.")
    session.replace_text(
        doc,
        "Thời gian thực hiện của đề tài: từ tháng 12/2024 đến tháng 05/2025",
        f"Thời gian thực hiện của đề tài: từ tháng 01/{info.year} đến tháng 12/{info.year}",
    )
    session.replace_text(doc, "2025", str(info.year))
    session.replace_text(doc, "2024", str(info.year))
    session.save_close(doc)


def _phieu_cham_diem_hd_de_cuong(session, dest_dir, info, title_old):
    doc = session.open(dest_dir / "Phiếu chấm điểm HĐ đề cương.docx")
    session.replace_text(doc, f"1. Tên Đề tài: {title_old}.", f"1. Tên Đề tài: {info.title}.")
    session.replace_text(doc, "2024", str(info.year))
    session.save_close(doc)


def _phieu_nhan_xet_danh_gia_ho_so(session, dest_dir, info, title_old):
    doc = session.open(dest_dir / "Phiếu nhận xét đánh giá hồ sơ.docx")
    session.replace_text(doc, f"Tên đề tài: {title_old}.", f"Tên đề tài: {info.title}.")
    session.replace_text(doc, "2024", str(info.year))
    session.save_close(doc)
