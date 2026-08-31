# token_rules.py
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Optional

import excel_reader
import paths

DEFAULT_CONFIG_PATH = paths.project_root() / "config_tokens.json"
TOKENS_SHEET_NAME = "_Tokens"


@dataclass
class TokenSpec:
    name: str
    code: str
    kind: str
    param: str
    note: str = ""


def load_token_specs(config_path: Optional[Path] = None) -> List[TokenSpec]:
    path = config_path or DEFAULT_CONFIG_PATH
    if not path.exists():
        return []
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return [
        TokenSpec(
            name=item.get("token_name", ""),
            code=item.get("code", ""),
            kind=item.get("kind", ""),
            param=item.get("param", "") or "",
            note=item.get("note", ""),
        )
        for item in data
        if item.get("token_name")
    ]


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


def _resolve_person_org(ws, index, spec: TokenSpec) -> str:
    person = excel_reader.read_person(ws, index, spec.code)
    return person.org if person else ""


def _resolve_numbered_researchers(ws, index, spec: TokenSpec) -> str:
    researchers = excel_reader.read_researchers(ws, index)
    return "\n".join(f"{i}. {p.degree} {p.name}".strip() for i, p in enumerate(researchers, start=1))


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
    "person_org": _resolve_person_org,
    "numbered_researchers": _resolve_numbered_researchers,
    "timeline_start": _resolve_timeline_start,
    "timeline_end": _resolve_timeline_end,
}


def resolve_tokens(arg1, arg2=None, arg3=None, config_path: Optional[Path] = None) -> Dict[str, str]:
    """Ho tro ca 2 chu ky goi de tuong thich:
    1. resolve_tokens(ws, index, config_path=...)
    2. resolve_tokens(wb, ws, index, config_path=...)
    """
    if hasattr(arg1, "sheetnames"):
        # goi kieu cu: (wb, ws, index)
        ws = arg2
        index = arg3 or {}
        # Neu workbook co sheet _Tokens va khong chi dinh config_path, load tu sheet do de backward compat
        if TOKENS_SHEET_NAME in arg1.sheetnames and config_path is None and not DEFAULT_CONFIG_PATH.exists():
            specs = _load_token_specs_from_sheet(arg1)
        else:
            specs = load_token_specs(config_path)
    else:
        # goi kieu moi: (ws, index)
        ws = arg1
        index = arg2 or {}
        specs = load_token_specs(config_path)

    result = {}
    for spec in specs:
        transform = TRANSFORMS.get(spec.kind)
        if transform is None:
            raise ValueError(
                f"Token '{spec.name}' trong config dùng kind '{spec.kind}' không hợp lệ - "
                f"chỉ chấp nhận {sorted(TRANSFORMS)}"
            )
        result[f"{{{{{spec.name}}}}}"] = transform(ws, index, spec)
    return result


def _load_token_specs_from_sheet(wb) -> List[TokenSpec]:
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
                note=row[4].value or "" if len(row) > 4 else "",
            )
        )
    return specs
