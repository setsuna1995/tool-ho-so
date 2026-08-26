import dataclasses

import excel_reader
import tokens


def _make_info(**overrides):
    base = excel_reader.ProjectInfo(
        title="Đề tài test",
        research_type="TVCT_ĐGHQ",
        year=2027,
        host_org="Viện ABC",
        partner_org=None,
        research_location=None,
        timeline="Tháng 01/2027 đến tháng 12/2027",
        head=excel_reader.Person(name="Nguyễn Văn A", degree="TS.", org="Viện ABC"),
        co_head=None,
        project_secretary=None,
        researchers=[],
        ethics_committee=excel_reader.CommitteeData(chair=excel_reader.Person(name="Chủ tịch")),
        proposal_committee=excel_reader.CommitteeData(chair=excel_reader.Person(name="Chủ tịch")),
        acceptance_committee=excel_reader.CommitteeData(chair=excel_reader.Person(name="Chủ tịch")),
        head_cv_filename="cv.docx",
        expert_cvs=[],
    )
    return dataclasses.replace(base, **overrides)


def test_build_common_tokens_maps_scalar_fields():
    info = _make_info()
    t = tokens.build_common_tokens(info)
    assert t["{{TEN_DE_TAI}}"] == "Đề tài test"
    assert t["{{NAM}}"] == "2027"
    assert t["{{DON_VI_CHU_TRI}}"] == "Viện ABC"
    assert t["{{CHU_NHIEM_HO_TEN}}"] == "TS. Nguyễn Văn A"
    assert t["{{CHU_NHIEM_TEN}}"] == "Nguyễn Văn A"
    assert t["{{THOI_GIAN_BAT_DAU}}"] == "01/2027"
    assert t["{{THOI_GIAN_KET_THUC}}"] == "12/2027"


def test_build_common_tokens_blank_research_location_falls_back_to_dots():
    info = _make_info(research_location=None)
    t = tokens.build_common_tokens(info)
    assert t["{{DIA_DIEM_TRIEN_KHAI}}"] == "……………………………."


def test_build_common_tokens_uses_real_research_location_when_present():
    info = _make_info(research_location="tỉnh Thái Bình")
    t = tokens.build_common_tokens(info)
    assert t["{{DIA_DIEM_TRIEN_KHAI}}"] == "tỉnh Thái Bình"


def test_build_common_tokens_blank_secretary_and_co_head_are_empty_strings():
    info = _make_info(co_head=None, project_secretary=None)
    t = tokens.build_common_tokens(info)
    assert t["{{THU_KY_DE_TAI}}"] == ""
    assert t["{{DONG_CHU_NHIEM_TEN}}"] == ""


def test_build_common_tokens_fills_secretary_and_co_head_when_present():
    info = _make_info(
        co_head=excel_reader.Person(name="Đồng chủ nhiệm B"),
        project_secretary=excel_reader.Person(name="Thư ký C", degree="ThS."),
    )
    t = tokens.build_common_tokens(info)
    assert t["{{DONG_CHU_NHIEM_TEN}}"] == "Đồng chủ nhiệm B"
    assert t["{{THU_KY_DE_TAI}}"] == "ThS. Thư ký C"
