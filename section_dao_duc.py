from pathlib import Path

import committee_writer
import word_writer
from excel_reader import ProjectInfo

ROLES = ["Chủ tịch Hội đồng", "Thành viên", "Thành viên", "Thành viên", "Thành viên"]


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, common_tokens: dict) -> None:
    _quyet_dinh_giao_de_tai(session, dest_dir, info, common_tokens)
    _qdtlhd_dao_duc(session, dest_dir, info, common_tokens)
    _bb_hop_hd_dao_duc(session, dest_dir, info, common_tokens)
    _bb_kiem_phieu_hd_dao_duc(session, dest_dir, info, common_tokens)
    _qd_chap_nhan_dao_duc(session, dest_dir, info, common_tokens)
    _bang_kiem_danh_gia_dao_duc(session, dest_dir, info, common_tokens)


def _quyet_dinh_giao_de_tai(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "00. QĐ Giao đề tài.docx")
    session.fill_tokens(doc, common_tokens)

    head_text = f"{info.head.degree} {info.head.name}".strip()
    head_org = f" - {info.head.org}" if info.head.org else ""
    session.set_cell(doc, 3, 2, 3, f"Chủ nhiệm đề tài: \r{head_text}{head_org}.")

    members_text = "\r".join(f"{p.degree} {p.name}".strip() for p in info.researchers)
    session.set_cell(doc, 3, 3, 3, f"Thành viên thực hiện:\r{members_text}")

    session.save_close(doc)


def _qdtlhd_dao_duc(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "01. QĐTLHĐ đạo đức đề cương.docx")
    session.fill_tokens(doc, common_tokens)
    committee_writer.write_committee_roster(
        session, doc, 2, info.ethics_committee, roles=ROLES, name_col=1, org_col=2, role_col=3
    )
    committee_writer.write_committee_secretaries(
        session, doc, 3, info.ethics_committee, name_col=1, org_col=2
    )
    session.save_close(doc)


def _bb_hop_hd_dao_duc(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "02. BB họp HĐ đạo đức.docx")
    session.fill_tokens(doc, common_tokens)
    chair = info.ethics_committee.chair
    session.replace_text(doc, "PGs. Ts. Hoàng Thị Thanh", f"{chair.degree} {chair.name}".strip())
    session.replace_text(
        doc,
        "Quyết định số: 04/QĐ-YHUD/2024 ngày 19 tháng 04 năm 2024",
        f"Quyết định số: ……/QĐ-YHUD/{info.year} ngày …… tháng …… năm {info.year}",
    )
    session.replace_text(doc, "Thời gian: ngày 25 tháng 04 năm 2024", f"Thời gian: ngày …… tháng …… năm {info.year}")
    session.save_close(doc)


def _bb_kiem_phieu_hd_dao_duc(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "03. BB kiểm phiếu HĐ đạo đức.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _qd_chap_nhan_dao_duc(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "04. QĐ chấp nhận đạo đức.docx")
    session.fill_tokens(doc, common_tokens)
    chair = info.ethics_committee.chair
    session.set_cell(doc, 2, 1, 2, f"CHỦ TỊCH HỘI ĐỒNG\r{chair.degree} {chair.name}".strip())
    session.save_close(doc)


def _bang_kiem_danh_gia_dao_duc(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Bảng kiểm đánh giá đạo đức.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)
