from excel_reader import ProjectInfo, parse_timeline

DIA_DIEM_PLACEHOLDER = "……………………………"


def build_common_tokens(info: ProjectInfo) -> dict:
    start, end = parse_timeline(info.timeline)
    secretary = info.project_secretary
    return {
        "{{TEN_DE_TAI}}": info.title,
        "{{NAM}}": str(info.year),
        "{{DON_VI_CHU_TRI}}": info.host_org,
        "{{DON_VI_DOI_TAC}}": info.partner_org or "",
        "{{CHU_NHIEM_HO_TEN}}": f"{info.head.degree} {info.head.name}".strip(),
        "{{CHU_NHIEM_TEN}}": info.head.name,
        "{{DONG_CHU_NHIEM_TEN}}": info.co_head.name if info.co_head else "",
        "{{THU_KY_DE_TAI}}": f"{secretary.degree} {secretary.name}".strip() if secretary else "",
        "{{THOI_GIAN_BAT_DAU}}": start,
        "{{THOI_GIAN_KET_THUC}}": end,
        "{{DIA_DIEM_TRIEN_KHAI}}": info.research_location or DIA_DIEM_PLACEHOLDER,
    }
