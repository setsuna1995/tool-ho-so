from typing import List, Optional

import word_writer
from excel_reader import CommitteeData


def write_committee_roster(
    session: word_writer.Session,
    doc: word_writer.OpenDoc,
    table_index: int,
    committee: CommitteeData,
    roles: List[str],
    name_col: int = 1,
    org_col: int = 2,
    role_col: Optional[int] = None,
    start_row: int = 1,
) -> None:
    people = [committee.chair] + committee.reviewers + committee.members
    if len(people) != len(roles):
        raise ValueError(
            f"Số thành viên hội đồng ({len(people)}) không khớp số vai trò truyền vào ({len(roles)})"
        )
    for offset, (person, role) in enumerate(zip(people, roles)):
        row = start_row + offset
        display_name = f"{person.degree} {person.name}".strip()
        session.set_cell(doc, table_index, row, name_col, display_name)
        session.set_cell(doc, table_index, row, org_col, person.org)
        if role_col is not None:
            session.set_cell(doc, table_index, row, role_col, role)


def write_committee_secretaries(
    session: word_writer.Session,
    doc: word_writer.OpenDoc,
    table_index: int,
    committee: CommitteeData,
    name_col: int = 1,
    org_col: int = 2,
    number_prefix: bool = False,
    start_row: int = 1,
) -> None:
    for offset, person in enumerate(committee.secretaries):
        row = start_row + offset
        name = f"{offset + 1}. {person.name}" if number_prefix else person.name
        session.set_cell(doc, table_index, row, name_col, name)
        session.set_cell(doc, table_index, row, org_col, person.org)


def roster_size(committee: CommitteeData) -> int:
    return 1 + len(committee.reviewers) + len(committee.members)
