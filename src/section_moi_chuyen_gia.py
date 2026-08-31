from pathlib import Path
from typing import List

import expert_invitation
import word_writer
from excel_reader import CommitteeData, Person, ProjectInfo


def _external_members(committee: CommitteeData, host_org: str) -> List[Person]:
    host_org_norm = host_org.strip().lower()
    candidates = [committee.chair] + committee.reviewers + committee.members
    return [p for p in candidates if p.org.strip().lower() != host_org_norm]


def _dedupe_people(people: List[Person]) -> List[Person]:
    seen = set()
    result = []
    for person in people:
        key = (person.name.strip().lower(), person.org.strip().lower())
        if key in seen:
            continue
        seen.add(key)
        result.append(person)
    return result


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, common_tokens: dict) -> None:
    _thu_moi_de_cuong(dest_dir, info, common_tokens)
    _thu_moi_nghiem_thu(dest_dir, info, common_tokens)


def _thu_moi_de_cuong(dest_dir, info, common_tokens):
    recipients = _dedupe_people(
        _external_members(info.ethics_committee, info.host_org)
        + _external_members(info.proposal_committee, info.host_org)
    )
    generated = expert_invitation.generate_multi_page_letter(
        dest_dir / "Công văn mời chuyên gia.docx", recipients, common_tokens
    )
    if not generated:
        print(
            "  [LUU Y] Khong co chuyen gia ngoai don vi chu tri trong Hoi dong khoa hoc/dao duc - "
            "bo qua dien thu moi de cuong."
        )


def _thu_moi_nghiem_thu(dest_dir, info, common_tokens):
    recipients = _dedupe_people(_external_members(info.acceptance_committee, info.host_org))
    generated = expert_invitation.generate_multi_page_letter(
        dest_dir / "Công văn mời chuyên gia nghiệm thu.docx", recipients, common_tokens
    )
    if not generated:
        print(
            "  [LUU Y] Khong co chuyen gia ngoai don vi chu tri trong Hoi dong nghiem thu - "
            "bo qua dien thu moi nghiem thu."
        )
