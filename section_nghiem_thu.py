from pathlib import Path

import committee_writer
import word_writer
from excel_reader import ProjectInfo

ROLES = ["Chủ tịch\rHội đồng", "Phản biện 1", "Phản biện 2", "Ủy viên", "Uỷ viên"]

SCORING_FORM_FILENAMES = {
    "TVCT_ĐGHQ": "Phiếu chấm điểm nghiệm thu (TVCT_ĐGHQ).docx",
    "TNLS": "Phiếu chấm điểm nghiệm thu (TNLS).docx",
}


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, common_tokens: dict) -> None:
    _quyet_dinh_thanh_lap(session, dest_dir, info, common_tokens)
    _bb_hop_hd_nghiem_thu(session, dest_dir, info, common_tokens)
    _bb_kiem_phieu_nghiem_thu(session, dest_dir, info, common_tokens)
    _qd_cong_nhan_ket_qua(session, dest_dir, info, common_tokens)
    _phieu_cham_diem_nghiem_thu(session, dest_dir, info, common_tokens)
    _phieu_ky_nhan_tien(session, dest_dir, info, common_tokens)
    _phieu_nhan_xet_nghiem_thu(session, dest_dir, info, common_tokens)


def _quyet_dinh_thanh_lap(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "9. Quyết định thành lập HĐ nghiệm thu.docx")
    session.fill_tokens(doc, common_tokens)
    committee_writer.write_committee_roster(
        session, doc, 2, info.acceptance_committee, roles=ROLES, name_col=2, org_col=3, role_col=4
    )
    committee_writer.write_committee_secretaries(
        session, doc, 3, info.acceptance_committee, name_col=1, org_col=2, number_prefix=True
    )
    session.save_close(doc)


def _bb_hop_hd_nghiem_thu(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "10. Biên bản họp HĐ nghiệm thu.docx")
    session.fill_tokens(doc, common_tokens)
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


def _bb_kiem_phieu_nghiem_thu(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "11. Biên bản kiểm phiếu nghiệm thu.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _qd_cong_nhan_ket_qua(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "12. Quyết định công nhận kết quả đề tài.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _phieu_cham_diem_nghiem_thu(session, dest_dir, info, common_tokens):
    if info.research_type not in SCORING_FORM_FILENAMES:
        raise ValueError(
            f"Loại hình nghiên cứu '{info.research_type}' (mã A02) không hợp lệ - "
            f"chỉ chấp nhận {sorted(SCORING_FORM_FILENAMES)}"
        )

    selected_filename = SCORING_FORM_FILENAMES[info.research_type]
    for research_type, filename in SCORING_FORM_FILENAMES.items():
        if filename != selected_filename:
            unused_path = dest_dir / filename
            if unused_path.exists():
                unused_path.unlink()

    doc = session.open(dest_dir / selected_filename)
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _phieu_ky_nhan_tien(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Phiếu ký nhận tiền.docx")
    session.fill_tokens(doc, common_tokens)
    committee_writer.write_committee_roster(
        session, doc, 2, info.acceptance_committee, roles=ROLES,
        name_col=2, org_col=None, role_col=None, start_row=2,
    )
    secretary = info.acceptance_committee.secretaries[0]
    session.set_cell(doc, 2, 7, 2, secretary.name)
    session.save_close(doc)


def _phieu_nhan_xet_nghiem_thu(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Phiếu nhận xét nghiệm thu.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)
