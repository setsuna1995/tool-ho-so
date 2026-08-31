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


def _resolve_single_researcher(ws, index, spec: TokenSpec) -> str:
    """Resolve a token for a specific numbered researcher.
    Supports spec.code or spec.name like:
    - B04_1, B04_2, ..., B04_20
    - B04, B05, ..., B20
    - NGHIEN_CUU_VIEN_1, ..., NGHIEN_CUU_VIEN_20
    """
    target_idx = None
    m = re.search(r"NGHIEN_CUU_VIEN_(\d+)", spec.name)
    if m:
        target_idx = int(m.group(1)) - 1
    elif spec.code:
        m2 = re.search(r"_(\d+)$", spec.code)
        if m2:
            target_idx = int(m2.group(1)) - 1
        else:
            m_b = re.match(r"^B(\d{2})$", spec.code)
            if m_b:
                b_num = int(m_b.group(1))
                if 4 <= b_num <= 20:
                    target_idx = b_num - 4

    if target_idx is None:
        return ""

    researchers = excel_reader.read_researchers(ws, index)
    if 0 <= target_idx < len(researchers):
        p = researchers[target_idx]
        if spec.kind in ("person_ten", "single_researcher_ten") or spec.name.endswith("_TEN"):
            return p.name
        if spec.kind in ("person_org", "single_researcher_org") or spec.name.endswith("_DON_VI"):
            return p.org
        return f"{p.degree} {p.name}".strip()
    return ""


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


import datetime
import re


def _format_vietnamese_date(val, fallback="ngày …… tháng …… năm") -> str:
    if val is None or val == "":
        return fallback
    if isinstance(val, (datetime.date, datetime.datetime)):
        return f"ngày {val.day:02d} tháng {val.month:02d} năm {val.year}"
    s = str(val).strip()
    if not s:
        return fallback
    m = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$", s)
    if m:
        d, mth, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"ngày {d:02d} tháng {mth:02d} năm {y}"
    m2 = re.match(r"^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})$", s)
    if m2:
        y, mth, d = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        return f"ngày {d:02d} tháng {mth:02d} năm {y}"
    return s


def _resolve_date_vietnamese(ws, index, spec: TokenSpec) -> str:
    if spec.code not in index:
        return spec.param
    row = excel_reader._resolve_row(index, spec.code)
    cell_val = ws.cell(row=row, column=3).value
    return _format_vietnamese_date(cell_val, fallback=spec.param)


TRANSFORMS: Dict[str, Callable] = {
    "raw": _resolve_raw,
    "raw_or_placeholder": _resolve_raw_or_placeholder,
    "date_vietnamese": _resolve_date_vietnamese,
    "person_ho_ten": _resolve_person_ho_ten,
    "person_ten": _resolve_person_ten,
    "person_org": _resolve_person_org,
    "numbered_researchers": _resolve_numbered_researchers,
    "timeline_start": _resolve_timeline_start,
    "timeline_end": _resolve_timeline_end,
    "single_researcher": _resolve_single_researcher,
    "single_researcher_ho_ten": _resolve_single_researcher,
    "single_researcher_ten": _resolve_single_researcher,
    "single_researcher_org": _resolve_single_researcher,
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

    # Tự động điền đầy đủ các token đích danh từng Nghiên cứu viên NGHIEN_CUU_VIEN_1..20
    if specs and (config_path is None or config_path == DEFAULT_CONFIG_PATH or any("NGHIEN_CUU_VIEN" in s.name for s in specs)):
        researchers = excel_reader.read_researchers(ws, index)
        for i in range(1, 21):
            if i <= len(researchers):
                p = researchers[i - 1]
                full_name = f"{p.degree} {p.name}".strip()
                name_only = p.name
                org_only = p.org
            else:
                full_name = ""
                name_only = ""
                org_only = ""

            result.setdefault(f"{{{{NGHIEN_CUU_VIEN_{i}}}}}", full_name)
            result.setdefault(f"{{{{NGHIEN_CUU_VIEN_{i}_HO_TEN}}}}", full_name)
            result.setdefault(f"{{{{NGHIEN_CUU_VIEN_{i}_TEN}}}}", name_only)
            result.setdefault(f"{{{{NGHIEN_CUU_VIEN_{i}_DON_VI}}}}", org_only)

    # Dynamic Token Discovery: Quét toàn bộ các dòng ở Cột A
    # Nếu có token mới dạng {{TOKEN_NAME}} hoặc TOKEN_NAME chưa có trong config
    known_keys = set(result.keys())
    for row in ws.iter_rows(min_row=3, max_col=5):
        cell_a = row[0].value
        if not cell_a or not isinstance(cell_a, str) or cell_a.startswith("SEC_"):
            continue
        raw_token = cell_a.strip()
        is_token_format = (raw_token.startswith("{{") and raw_token.endswith("}}")) or (
            raw_token.isupper()
            and not (len(raw_token) == 3 and raw_token[0] in "ABCDEF" and raw_token[1:].isdigit())
        )
        if is_token_format:
            token_clean = raw_token.replace("{{", "").replace("}}", "").strip()
            token_key = f"{{{{{token_clean}}}}}"
            if token_key not in known_keys:
                val_c = row[2].value if len(row) > 2 and row[2].value is not None else ""
                result[token_key] = "" if val_c is None else str(val_c).strip()
                known_keys.add(token_key)

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


def _person_to_dict(p) -> Optional[dict]:
    if p is None:
        return None
    degree = getattr(p, "degree", "") or ""
    name = getattr(p, "name", "") or ""
    org = getattr(p, "org", "") or ""
    return {
        "name": name,
        "degree": degree,
        "org": org,
        "full_name": f"{degree} {name}".strip(),
    }


def _committee_to_dict(c) -> Optional[dict]:
    if c is None:
        return None
    return {
        "chair": _person_to_dict(c.chair),
        "reviewers": [_person_to_dict(p) for p in (c.reviewers or [])],
        "members": [_person_to_dict(p) for p in (c.members or [])],
        "secretaries": [_person_to_dict(p) for p in (c.secretaries or [])],
    }


def build_template_context(info) -> dict:
    """Chuyển đổi đối tượng ProjectInfo thành dictionary context phục vụ Jinja2 / docxtpl và token mapping."""
    context = {}

    # 1. Các token phẳng truyền thống (bỏ dấu {{ }})
    if hasattr(info, "common_tokens") and info.common_tokens:
        for k, v in info.common_tokens.items():
            clean_key = k.strip("{}")
            context[clean_key] = v

    # 2. Cấu trúc đối tượng phân cấp project
    context["project"] = {
        "title": getattr(info, "title", ""),
        "year": getattr(info, "year", ""),
        "host_org": getattr(info, "host_org", ""),
        "partner_org": getattr(info, "partner_org", "") or "",
        "research_type": getattr(info, "research_type", ""),
        "research_location": getattr(info, "research_location", "") or "",
        "timeline": getattr(info, "timeline", ""),
        "head": _person_to_dict(getattr(info, "head", None)),
        "co_head": _person_to_dict(getattr(info, "co_head", None)),
        "project_secretary": _person_to_dict(getattr(info, "project_secretary", None)),
    }

    # 3. Danh sách nghiên cứu viên
    researchers = getattr(info, "researchers", []) or []
    context["researchers"] = [_person_to_dict(p) for p in researchers]

    # 4. Danh sách các hội đồng
    context["committees"] = {
        "ethics": _committee_to_dict(getattr(info, "ethics_committee", None)),
        "proposal": _committee_to_dict(getattr(info, "proposal_committee", None)),
        "acceptance": _committee_to_dict(getattr(info, "acceptance_committee", None)),
    }

    return context

