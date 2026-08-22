import pytest

import excel_reader
import paths

CHECKLIST_PATH = paths.project_root() / excel_reader.CHECKLIST_FILENAME
SHEET_VIAM = "Đề tài - Bánh ăn dặm VIAM 2027"
SHEET_BLANK = "Đề tài - Mẫu trắng dự án mới"


def test_load_project_data_title_and_year():
    data = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    assert data.title == "Tư vấn hiệu quả công thức sản phẩm Bánh ăn dặm VIAM"
    assert data.year == 2027


def test_secretary_org_differs_correctly_between_committees():
    data = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    assert data.ethics_committee.secretaries[0].org == "Trung tâm NCKH - Viện VIAM"
    assert data.proposal_committee.secretaries[0].org == "Trung tâm NCKH - Viện VIAM"
    assert data.acceptance_committee.secretaries[0].org == "Viện Y học Ứng dụng Việt Nam"


def test_optional_partner_org_is_none_when_blank():
    data = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    assert data.partner_org is None


def test_ethics_committee_has_required_counts():
    data = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    c = data.ethics_committee
    assert c.chair.name == "Nguyễn Công Khẩn"
    assert len(c.reviewers) == 2
    assert len(c.members) == 2
    assert len(c.secretaries) == 2


def test_head_and_researchers():
    data = excel_reader.load_project_data(CHECKLIST_PATH, SHEET_VIAM)
    assert data.head.name == "Trương Hồng Sơn"
    assert data.co_head is None
    names = [p.name for p in data.researchers]
    assert "Lê Việt Anh" in names
    assert "Lê Minh Khánh" in names


def test_blank_template_sheet_raises_on_missing_required_field():
    with pytest.raises(ValueError):
        excel_reader.load_project_data(CHECKLIST_PATH, SHEET_BLANK)
