from pathlib import Path

import committee_writer
import word_writer
from excel_reader import ProjectInfo

ROLES = ["Chủ tịch\rHội đồng", "Phản biện 1", "Phản biện 2", "Ủy viên", "Uỷ viên"]


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo) -> None:
    _quyet_dinh_thanh_lap(session, dest_dir, info)
    _bb_hop_hd_nghiem_thu(session, dest_dir, info)
    _bb_kiem_phieu_nghiem_thu(session, dest_dir, info)
    _qd_cong_nhan_ket_qua(session, dest_dir, info)
    _phieu_cham_diem_nghiem_thu(session, dest_dir, info)
    _phieu_ky_nhan_tien(session, dest_dir, info)
    _phieu_nhan_xet_nghiem_thu(session, dest_dir, info)


def _quyet_dinh_thanh_lap(session, dest_dir, info):
    doc = session.open(dest_dir / "9. Quyết định thành lập HĐ nghiệm thu.docx")
    session.replace_text(doc, "20xx", str(info.year))
    session.replace_text(doc, "“Tên đề tài”", f"“{info.title}”")
    committee_writer.write_committee_roster(
        session, doc, 2, info.acceptance_committee, roles=ROLES, name_col=2, org_col=3, role_col=4
    )
    committee_writer.write_committee_secretaries(
        session, doc, 3, info.acceptance_committee, name_col=1, org_col=2, number_prefix=True
    )
    session.save_close(doc)


def _bb_hop_hd_nghiem_thu(session, dest_dir, info):
    doc = session.open(dest_dir / "10. Biên bản họp HĐ nghiệm thu.docx")
    session.replace_text(doc, "20xx", str(info.year))
    session.replace_text(doc, "1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {info.title}")
    session.replace_text(doc, "Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {info.head.name}")
    co_head_name = info.co_head.name if info.co_head else ""
    session.replace_text(doc, "Đồng chủ nhiệm đề tài: Tên 2", f"Đồng chủ nhiệm đề tài: {co_head_name}")
    chair = info.acceptance_committee.chair
    session.replace_text(
        doc,
        "Tên 3 - Chủ tịch Hội đồng điều khiển phiên họp",
        f"{chair.name} - Chủ tịch Hội đồng điều khiển phiên họp",
    )
    member_count = committee_writer.roster_size(info.acceptance_committee)
    session.replace_text(
        doc,
        "5. Số thành viên Hội đồng theo quyết định là …… người",
        f"5. Số thành viên Hội đồng theo quyết định là {member_count:02d} người",
    )
    session.save_close(doc)


def _bb_kiem_phieu_nghiem_thu(session, dest_dir, info):
    doc = session.open(dest_dir / "11. Biên bản kiểm phiếu nghiệm thu.docx")
    session.replace_text(doc, "20xx", str(info.year))
    session.replace_text(doc, "1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {info.title}")
    session.replace_text(doc, "Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {info.head.name}")
    co_head_name = info.co_head.name if info.co_head else ""
    session.replace_text(doc, "Đồng chủ nhiệm đề tài: Tên 2", f"Đồng chủ nhiệm đề tài: {co_head_name}")
    session.save_close(doc)


def _qd_cong_nhan_ket_qua(session, dest_dir, info):
    doc = session.open(dest_dir / "12. Quyết định công nhận kết quả đề tài.docx")
    # Template nay dung "20XX" hoa o tieu de nhung "20xx" thuong o bang tieu de
    # (khac voi cac file 9/10/11 dung "20xx" thuong nhat quan o moi noi) - can
    # thay ca hai de dung tren backend docx (case-sensitive).
    session.replace_text(doc, "20XX", str(info.year))
    session.replace_text(doc, "20xx", str(info.year))
    session.replace_text(doc, "“Tên đề tài”", f"“{info.title}”")
    session.save_close(doc)


def _phieu_cham_diem_nghiem_thu(session, dest_dir, info):
    doc = session.open(dest_dir / "Phiếu chấm điểm nghiệm thu (TVCT_ĐGHQ).docx")
    session.replace_text(doc, "1. Tên đề tài: Tên đề tài", f"1. Tên đề tài: {info.title}")
    session.replace_text(doc, "Chủ nhiệm đề tài: Tên 1", f"Chủ nhiệm đề tài: {info.head.name}")
    session.save_close(doc)


def _phieu_ky_nhan_tien(session, dest_dir, info):
    doc = session.open(dest_dir / "Phiếu ký nhận tiền.docx")
    session.replace_text(
        doc,
        "“Đánh giá hiệu quả sản phẩm thực phẩm chức năng Viên nang Đông trùng hạ thảo CordySen”",
        f"“{info.title}”",
    )
    # Ten thu ky o day giu nguyen literal nhu ban goc - khong du can cu de anh xa
    # sang ma muc Excel nao (nam ngoai pham vi loi duoc bao cao).
    session.set_cell(doc, 2, 7, 2, "Hoàng Hà Linh")
    session.save_close(doc)


def _phieu_nhan_xet_nghiem_thu(session, dest_dir, info):
    doc = session.open(dest_dir / "Phiếu nhận xét nghiệm thu.docx")
    session.replace_text(doc, "20xx", str(info.year))
    session.replace_text(doc, "Tên đề tài: Tên đề tài", f"Tên đề tài: {info.title}")
    session.replace_text(doc, "Chủ nhiệm: Tên 1", f"Chủ nhiệm: {info.head.name}")
    co_head_name = info.co_head.name if info.co_head else ""
    session.replace_text(doc, "Đồng chủ nhiệm: Tên 2", f"Đồng chủ nhiệm: {co_head_name}")
    session.save_close(doc)
