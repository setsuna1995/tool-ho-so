# token_rules.py
from dataclasses import dataclass
from typing import Callable, Dict, List

import excel_reader

TOKENS_SHEET_NAME = "_Tokens"


@dataclass
class TokenSpec:
    name: str
    code: str
    kind: str
    param: str


def _load_token_specs(wb) -> List[TokenSpec]:
    ws = wb[TOKENS_SHEET_NAME]
    specs = []
    for row in ws.iter_rows(min_row=2):
        name = row[0].value
        if not name:
            continue
        specs.append(
            TokenSpec(
                name=name,
                code=row[1].value,
                kind=row[2].value,
                param=row[3].value or "",
            )
        )
    return specs


def _resolve_raw(ws, index, spec: TokenSpec) -> str:
    return excel_reader.read_text(ws, index, spec.code)


def _resolve_raw_or_placeholder(ws, index, spec: TokenSpec) -> str:
    """Ma muc chua co trong checklist (ban copy cu) duoc coi nhu o trong,
    tra ve placeholder thay vi nem KeyError."""
    if spec.code not in index:
        return spec.param
    value = excel_reader.read_text(ws, index, spec.code)
    return value or spec.param


def _resolve_person_ho_ten(ws, index, spec: TokenSpec) -> str:
    person = excel_reader.read_person(ws, index, spec.code)
    if person is None:
        return ""
    return f"{person.degree} {person.name}".strip()


def _resolve_person_ten(ws, index, spec: TokenSpec) -> str:
    person = excel_reader.read_person(ws, index, spec.code)
    return person.name if person else ""


def _resolve_timeline_start(ws, index, spec: TokenSpec) -> str:
    text = excel_reader.read_text(ws, index, spec.code)
    start, _end = excel_reader.parse_timeline(text)
    return start


def _resolve_timeline_end(ws, index, spec: TokenSpec) -> str:
    text = excel_reader.read_text(ws, index, spec.code)
    _start, end = excel_reader.parse_timeline(text)
    return end


TRANSFORMS: Dict[str, Callable] = {
    "raw": _resolve_raw,
    "raw_or_placeholder": _resolve_raw_or_placeholder,
    "person_ho_ten": _resolve_person_ho_ten,
    "person_ten": _resolve_person_ten,
    "timeline_start": _resolve_timeline_start,
    "timeline_end": _resolve_timeline_end,
}


def resolve_tokens(wb, ws, index: dict) -> Dict[str, str]:
    """Tra ve {} neu workbook chua co sheet _Tokens - giu tuong thich nguoc
    cho cac workbook/fixture kiem thu chua duoc migrate."""
    if TOKENS_SHEET_NAME not in wb.sheetnames:
        return {}

    result = {}
    for spec in _load_token_specs(wb):
        transform = TRANSFORMS.get(spec.kind)
        if transform is None:
            raise ValueError(
                f"Token '{spec.name}' trong sheet _Tokens dùng kind '{spec.kind}' không hợp lệ - "
                f"chỉ chấp nhận {sorted(TRANSFORMS)}"
            )
        result[f"{{{{{spec.name}}}}}"] = transform(ws, index, spec)
    return result
