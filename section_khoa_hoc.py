from pathlib import Path

import committee_writer
import word_writer
from excel_reader import ProjectInfo

ROLES = ["Chủ tịch Hội đồng", "Phản biện 1", "Phản biện 2", "Ủy viên", "Ủy viên"]


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, common_tokens: dict) -> None:
    _qdtlhd_khoa_hoc(session, dest_dir, info, common_tokens)
    _bb_hop_thong_qua_de_cuong(session, dest_dir, info, common_tokens)
    _bb_kiem_phieu_thong_qua_de_cuong(session, dest_dir, info, common_tokens)
    _qd_phe_duyet_de_tai(session, dest_dir, info, common_tokens)
    _phieu_cham_diem_hd_de_cuong(session, dest_dir, info, common_tokens)
    _phieu_nhan_xet_danh_gia_ho_so(session, dest_dir, info, common_tokens)


def _qdtlhd_khoa_hoc(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "05. QĐ TLHĐ khoa học xét đề cương.docx")
    session.fill_tokens(doc, common_tokens)
    committee_writer.write_committee_roster(
        session, doc, 2, info.proposal_committee, roles=ROLES, name_col=1, org_col=2, role_col=3
    )
    committee_writer.write_committee_secretaries(
        session, doc, 3, info.proposal_committee, name_col=2, org_col=3
    )
    session.save_close(doc)


def _bb_hop_thong_qua_de_cuong(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "06. BB họp thông qua đề cương.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _bb_kiem_phieu_thong_qua_de_cuong(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "07. BB kiểm phiếu thông qua đề cương.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _qd_phe_duyet_de_tai(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "08. QĐ phê duyệt đề tài.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _phieu_cham_diem_hd_de_cuong(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Phiếu chấm điểm HĐ đề cương.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


def _phieu_nhan_xet_danh_gia_ho_so(session, dest_dir, info, common_tokens):
    doc = session.open(dest_dir / "Phiếu nhận xét đánh giá hồ sơ.docx")
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)
