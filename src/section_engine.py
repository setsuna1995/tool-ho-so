# section_engine.py — Data-driven section generator
"""
Engine chung để sinh hồ sơ dựa trên cấu hình JSON thay vì hard-code Python.

Mỗi section (dao_duc, khoa_hoc, nghiem_thu, ...) được mô tả trong
section_config.json. Engine đọc config, dispatch từng template file
theo action type: fill_tokens, committee_roster, expert_invitation,
scoring_form, member_count, payment_slip.
"""
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

import committee_writer
import expert_invitation
import paths
import word_writer
from excel_reader import CommitteeData, ProjectInfo

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class TemplateSpec:
    filename: str
    action: str = "fill_tokens"
    # committee_roster
    roster_table_index: int = 0
    roster_kwargs: Dict[str, Any] = field(default_factory=dict)
    secretary_table_index: int = 0
    secretary_kwargs: Dict[str, Any] = field(default_factory=dict)
    # member_count
    member_count_find: str = ""
    member_count_replace: str = ""
    # scoring_form
    scoring_variants: Dict[str, str] = field(default_factory=dict)
    # payment_slip
    secretary_row: int = 0
    secretary_col: int = 0


@dataclass
class SectionConfig:
    id: str
    name: str
    template_dir: str
    output_dir: str
    committee_key: str
    default_roles: List[str]
    templates: List[TemplateSpec]


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------

_CONFIG_CACHE: Dict[str, SectionConfig] = {}


def _default_config_path() -> Path:
    return paths.project_root() / "section_config.json"


def _parse_template(raw: dict) -> TemplateSpec:
    return TemplateSpec(
        filename=raw["filename"],
        action=raw.get("action", "fill_tokens"),
        roster_table_index=raw.get("roster_table_index", 0),
        roster_kwargs=raw.get("roster_kwargs", {}),
        secretary_table_index=raw.get("secretary_table_index", 0),
        secretary_kwargs=raw.get("secretary_kwargs", {}),
        member_count_find=raw.get("member_count_find", ""),
        member_count_replace=raw.get("member_count_replace", ""),
        scoring_variants=raw.get("scoring_variants", {}),
        secretary_row=raw.get("secretary_row", 0),
        secretary_col=raw.get("secretary_col", 0),
    )


def _load_all(config_path: Path = None) -> Dict[str, SectionConfig]:
    path = config_path or _default_config_path()
    with open(path, encoding="utf-8") as f:
        raw_list = json.load(f)
    result = {}
    for raw in raw_list:
        templates = [_parse_template(t) for t in raw.get("templates", [])]
        cfg = SectionConfig(
            id=raw["id"],
            name=raw["name"],
            template_dir=raw["template_dir"],
            output_dir=raw["output_dir"],
            committee_key=raw["committee_key"],
            default_roles=raw["default_roles"],
            templates=templates,
        )
        result[cfg.id] = cfg
    return result


def load_section_config(section_id: str, config_path: Path = None) -> SectionConfig:
    """Trả về SectionConfig cho section_id, có cache."""
    cache_key = str(config_path or "default")
    if cache_key not in _CONFIG_CACHE or section_id not in _CONFIG_CACHE.get(cache_key, {}):
        all_configs = _load_all(config_path)
        _CONFIG_CACHE[cache_key] = all_configs
    configs = _CONFIG_CACHE.get(cache_key, {})
    if section_id not in configs:
        raise KeyError(
            f"Không tìm thấy section '{section_id}' trong config. "
            f"Các section hợp lệ: {list(configs.keys())}"
        )
    return configs[section_id]


def load_all_section_configs(config_path: Path = None) -> Dict[str, SectionConfig]:
    """Trả về dict tất cả section configs."""
    return _load_all(config_path)


def clear_cache():
    """Xóa cache config (dùng trong test)."""
    _CONFIG_CACHE.clear()


# ---------------------------------------------------------------------------
# Committee mapping
# ---------------------------------------------------------------------------

_COMMITTEE_MAP = {
    "ethics": "ethics_committee",
    "proposal": "proposal_committee",
    "acceptance": "acceptance_committee",
}


def _get_committee(info: ProjectInfo, committee_key: str) -> CommitteeData:
    attr = _COMMITTEE_MAP.get(committee_key)
    if attr is None:
        raise ValueError(
            f"committee_key '{committee_key}' không hợp lệ. "
            f"Chỉ chấp nhận: {list(_COMMITTEE_MAP.keys())}"
        )
    return getattr(info, attr)


def _external_members(committee: CommitteeData, host_org: str):
    host_org_norm = host_org.strip().lower()
    candidates = [committee.chair] + committee.reviewers + committee.members
    return [p for p in candidates if p.org.strip().lower() != host_org_norm]


# ---------------------------------------------------------------------------
# Action handlers
# ---------------------------------------------------------------------------

def _action_fill_tokens(session, doc, info, common_tokens, spec, committee, roles):
    """Chỉ fill token, không làm gì thêm."""
    session.fill_tokens(doc, common_tokens)


def _action_committee_roster(session, doc, info, common_tokens, spec, committee, roles):
    """Fill token + ghi danh sách hội đồng + thư ký vào bảng."""
    session.fill_tokens(doc, common_tokens)
    committee_writer.write_committee_roster(
        session, doc, spec.roster_table_index, committee,
        roles=roles, **spec.roster_kwargs,
    )
    if spec.secretary_table_index:
        committee_writer.write_committee_secretaries(
            session, doc, spec.secretary_table_index, committee,
            **spec.secretary_kwargs,
        )


def _action_member_count(session, doc, info, common_tokens, spec, committee, roles):
    """Fill token + thay số thành viên HĐ."""
    session.fill_tokens(doc, common_tokens)
    count = committee_writer.roster_size(committee)
    replace_text = spec.member_count_replace.format(count=count)
    session.replace_text(doc, spec.member_count_find, replace_text)


def _action_payment_slip(session, doc, info, common_tokens, spec, committee, roles):
    """Fill token + ghi danh sách HĐ + tên thư ký vào phiếu ký nhận tiền."""
    session.fill_tokens(doc, common_tokens)
    committee_writer.write_committee_roster(
        session, doc, spec.roster_table_index, committee,
        roles=roles, **spec.roster_kwargs,
    )
    if committee.secretaries:
        secretary = committee.secretaries[0]
        session.set_cell(doc, spec.roster_table_index, spec.secretary_row, spec.secretary_col, secretary.name)


# ---------------------------------------------------------------------------
# Main generate
# ---------------------------------------------------------------------------

def generate(
    session: word_writer.Session,
    dest_dir: Path,
    info: ProjectInfo,
    common_tokens: dict,
    config: SectionConfig,
) -> None:
    """Sinh tất cả file trong 1 section dựa trên SectionConfig."""
    committee = _get_committee(info, config.committee_key)
    roles = config.default_roles

    for spec in config.templates:
        if spec.action == "expert_invitation":
            _handle_expert_invitation(dest_dir, info, common_tokens, spec, committee)
            continue

        if spec.action == "scoring_form":
            _handle_scoring_form(session, dest_dir, info, common_tokens, spec, committee, roles)
            continue

        file_path = dest_dir / spec.filename
        if not file_path.exists():
            continue

        doc = session.open(file_path)
        handler = _ACTION_HANDLERS.get(spec.action, _action_fill_tokens)
        handler(session, doc, info, common_tokens, spec, committee, roles)
        session.save_close(doc)


def _handle_expert_invitation(dest_dir, info, common_tokens, spec, committee):
    """Xử lý thư mời chuyên gia — không mở qua word_writer Session."""
    letter_path = dest_dir / spec.filename
    if not letter_path.exists():
        return
    recipients = _external_members(committee, info.host_org)
    expert_invitation.generate_multi_page_letter(letter_path, recipients, common_tokens)


def _handle_scoring_form(session, dest_dir, info, common_tokens, spec, committee, roles):
    """Chọn phiếu chấm điểm theo research_type, xóa file thừa."""
    if info.research_type not in spec.scoring_variants:
        raise ValueError(
            f"Loại hình nghiên cứu '{info.research_type}' (mã A02) không hợp lệ - "
            f"chỉ chấp nhận {sorted(spec.scoring_variants)}"
        )
    selected_filename = spec.scoring_variants[info.research_type]
    for research_type, filename in spec.scoring_variants.items():
        if filename != selected_filename:
            unused_path = dest_dir / filename
            if unused_path.exists():
                unused_path.unlink()

    doc = session.open(dest_dir / selected_filename)
    session.fill_tokens(doc, common_tokens)
    session.save_close(doc)


_ACTION_HANDLERS = {
    "fill_tokens": _action_fill_tokens,
    "committee_roster": _action_committee_roster,
    "member_count": _action_member_count,
    "payment_slip": _action_payment_slip,
}
