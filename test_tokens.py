import excel_reader
import tokens


def _make_info(common_tokens):
    return excel_reader.ProjectInfo(
        title="t",
        research_type="TNLS",
        year=2027,
        host_org="h",
        partner_org=None,
        research_location=None,
        timeline="Tháng 01/2027 đến tháng 12/2027",
        head=excel_reader.Person(name="A"),
        co_head=None,
        project_secretary=None,
        researchers=[],
        ethics_committee=excel_reader.CommitteeData(chair=excel_reader.Person(name="C")),
        proposal_committee=excel_reader.CommitteeData(chair=excel_reader.Person(name="C")),
        acceptance_committee=excel_reader.CommitteeData(chair=excel_reader.Person(name="C")),
        head_cv_filename="cv.docx",
        expert_cvs=[],
        common_tokens=common_tokens,
    )


def test_build_common_tokens_forwards_info_common_tokens():
    info = _make_info({"{{FOO}}": "bar"})
    assert tokens.build_common_tokens(info) == {"{{FOO}}": "bar"}


def test_build_common_tokens_returns_a_copy_not_the_same_object():
    info = _make_info({"{{FOO}}": "bar"})
    result = tokens.build_common_tokens(info)
    result["{{FOO}}"] = "changed"
    assert info.common_tokens["{{FOO}}"] == "bar"
