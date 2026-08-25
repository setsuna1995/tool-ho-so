from pathlib import Path

import committee_writer
import word_writer
from excel_reader import ProjectInfo, parse_timeline

ROLES = ["Chủ tịch Hội đồng", "Thành viên", "Thành viên", "Thành viên", "Thành viên"]


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, title_old: str) -> None:
    _quyet_dinh_giao_de_tai(session, dest_dir, info, title_old)
    _qdtlhd_dao_duc(session, dest_dir, info, title_old)
    _bb_hop_hd_dao_duc(session, dest_dir, info, title_old)
    _bb_kiem_phieu_hd_dao_duc(session, dest_dir, info, title_old)
    _qd_chap_nhan_dao_duc(session, dest_dir, info, title_old)
    _bang_kiem_danh_gia_dao_duc(session, dest_dir, info, title_old)


def _quyet_dinh_giao_de_tai(session, dest_dir, info, title_old):
    doc = session.open(dest_dir / "00. QĐ Giao đề tài.docx")
    session.replace_text(doc, "2024", str(info.year))
    session.replace_text(doc, title_old, info.title)

    head_text = f"{info.head.degree} {info.head.name}".strip()
    head_org = f" - {info.head.org}" if info.head.org else ""
    session.set_cell(doc, 3, 2, 3, f"Chủ nhiệm đề tài: \r{head_text}{head_org}.")

    members_text = "\r".join(f"{p.degree} {p.name}".strip() for p in info.researchers)
    session.set_cell(doc, 3, 3, 3, f"Thành viên thực hiện:\r{members_text}")

    session.save_close(doc)


def _qdtlhd_dao_duc(session, dest_dir, info, title_old):
    doc = session.open(dest_dir / "01. QĐTLHĐ đạo đức đề cương.docx")
    session.replace_text(doc, "2024", str(info.year))
    session.replace_text(doc, title_old, info.title)
    committee_writer.write_committee_roster(
        session, doc, 2, info.ethics_committee, roles=ROLES, name_col=1, org_col=2, role_col=3
    )
    committee_writer.write_committee_secretaries(
        session, doc, 3, info.ethics_committee, name_col=1, org_col=2
    )
    session.save_close(doc)


def _bb_hop_hd_dao_duc(session, dest_dir, info, title_old):
    doc = session.open(dest_dir / "02. BB họp HĐ đạo đức.docx")
    session.replace_text(doc, f"Tên đề tài: {title_old}.", f"Tên đề tài: {info.title}.")
    chair = info.ethics_committee.chair
    session.replace_text(doc, "PGs. Ts. Hoàng Thị Thanh", f"{chair.degree} {chair.name}".strip())
    session.replace_text(
        doc,
        "Quyết định số: 04/QĐ-YHUD/2024 ngày 19 tháng 04 năm 2024",
        f"Quyết định số: ……/QĐ-YHUD/{info.year} ngày …… tháng …… năm {info.year}",
    )
    session.replace_text(doc, "Thời gian: ngày 25 tháng 04 năm 2024", f"Thời gian: ngày …… tháng …… năm {info.year}")
    session.save_close(doc)


def _bb_kiem_phieu_hd_dao_duc(session, dest_dir, info, title_old):
    doc = session.open(dest_dir / "03. BB kiểm phiếu HĐ đạo đức.docx")
    session.replace_text(doc, "2024", str(info.year))
    session.replace_text(doc, f"Tên đề tài: {title_old}.", f"Tên đề tài: {info.title}.")
    session.save_close(doc)


def _qd_chap_nhan_dao_duc(session, dest_dir, info, title_old):
    doc = session.open(dest_dir / "04. QĐ chấp nhận đạo đức.docx")
    session.replace_text(doc, "Địa điểm triển khai nghiên cứu: tỉnh Thái Nguyên.", "Địa điểm triển khai nghiên cứu: ……………………………….")
    start, end = parse_timeline(info.timeline)
    session.replace_text(doc, "Thời gian nghiên cứu: Từ 12/2024 đến 05/2024", f"Thời gian nghiên cứu: Từ {start} đến {end}")
    session.replace_text(doc, "2024", str(info.year))
    session.replace_text(doc, f"“{title_old}”.", f"“{info.title}”.")
    chair = info.ethics_committee.chair
    session.set_cell(doc, 2, 1, 2, f"CHỦ TỊCH HỘI ĐỒNG\r{chair.degree} {chair.name}".strip())
    session.save_close(doc)


def _bang_kiem_danh_gia_dao_duc(session, dest_dir, info, title_old):
    doc = session.open(dest_dir / "Bảng kiểm đánh giá đạo đức.docx")
    session.replace_text(doc, f"Tên nghiên cứu: {title_old}.", f"Tên nghiên cứu: {info.title}.")
    session.replace_text(doc, "Ngày       tháng       năm 2024", f"Ngày       tháng       năm {info.year}")
    session.save_close(doc)
