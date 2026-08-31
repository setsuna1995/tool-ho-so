# dossier_packages.py
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

import paths

CONFIG_PACKAGES_PATH = paths.project_root() / "config_packages.json"


@dataclass
class DossierPackage:
    id: str
    name: str
    description: str
    sections: List[str]
    template_dirs: List[str] = field(default_factory=list)
    copy_head_cv: bool = True
    copy_expert_cvs: bool = True
    aliases: List[str] = field(default_factory=list)


# Danh sách các bộ hồ sơ chuẩn mặc định
_DEFAULT_PACKAGES = [
    DossierPackage(
        id="A",
        aliases=["A", "full", "BO_A", "Bộ A", "Bộ A - NCKH & Đánh giá công thức VIAM"],
        name="Bộ A: Hồ sơ NCKH & Đánh giá hiệu quả công thức (Trọn bộ đầy đủ)",
        description="Bao gồm Đạo đức, Khoa học và Nghiệm thu (kèm Thư mời chuyên gia & CV)",
        sections=["dao_duc", "khoa_hoc", "nghiem_thu"],
        template_dirs=[
            "01. Hồ sơ đạo đức đề cương - MẪU",
            "02. Hồ sơ khoa học đề cương - MẪU",
            "04. Hồ sơ nghiệm thu - MẪU",
        ],
        copy_head_cv=True,
        copy_expert_cvs=True,
    ),
    DossierPackage(
        id="A_dao_duc",
        aliases=["dao_duc", "A_dao_duc", "Bộ A - Đạo đức"],
        name="Bộ A (Phần 1): Hồ sơ Đạo đức đề cương",
        description="Gồm hồ sơ Hội đồng Đạo đức (kèm Thư mời chuyên gia đạo đức & CV Chủ nhiệm)",
        sections=["dao_duc"],
        template_dirs=[
            "01. Hồ sơ đạo đức đề cương - MẪU",
        ],
        copy_head_cv=True,
        copy_expert_cvs=True,
    ),
    DossierPackage(
        id="A_khoa_hoc",
        aliases=["khoa_hoc", "A_khoa_hoc", "Bộ A - Khoa học"],
        name="Bộ A (Phần 2): Hồ sơ Khoa học đề cương",
        description="Gồm hồ sơ Hội đồng Khoa học xét duyệt đề cương (kèm Thư mời chuyên gia khoa học)",
        sections=["khoa_hoc"],
        template_dirs=[
            "02. Hồ sơ khoa học đề cương - MẪU",
        ],
        copy_head_cv=False,
        copy_expert_cvs=True,
    ),
    DossierPackage(
        id="A_nghiem_thu",
        aliases=["nghiem_thu", "A_nghiem_thu", "Bộ A - Nghiệm thu"],
        name="Bộ A (Phần 3): Hồ sơ Nghiệm thu đề tài",
        description="Gồm hồ sơ Hội đồng Nghiệm thu kết quả (kèm Thư mời chuyên gia nghiệm thu)",
        sections=["nghiem_thu"],
        template_dirs=[
            "04. Hồ sơ nghiệm thu - MẪU",
        ],
        copy_head_cv=False,
        copy_expert_cvs=True,
    ),
    DossierPackage(
        id="B",
        aliases=["B", "BO_B", "Bộ B", "Bộ B - Thử nghiệm lâm sàng"],
        name="Bộ B: Hồ sơ Thử nghiệm lâm sàng (TNLS)",
        description="Bộ hồ sơ phục vụ đề tài Thử nghiệm lâm sàng / Can thiệp lâm sàng (Setup sẵn cho tương lai)",
        sections=["dao_duc", "khoa_hoc", "nghiem_thu"],
        template_dirs=[
            "01. Hồ sơ đạo đức đề cương - MẪU",
            "02. Hồ sơ khoa học đề cương - MẪU",
            "04. Hồ sơ nghiệm thu - MẪU",
        ],
        copy_head_cv=True,
        copy_expert_cvs=True,
    ),
    DossierPackage(
        id="C",
        aliases=["C", "BO_C", "Bộ C", "Bộ C - Nghiên cứu quan sát & Dịch tễ"],
        name="Bộ C: Hồ sơ Nghiên cứu quan sát & Khảo sát dịch tễ",
        description="Bộ hồ sơ phục vụ nghiên cứu quan sát, điều tra dịch tễ cộng đồng (Setup sẵn cho tương lai)",
        sections=["dao_duc", "khoa_hoc", "nghiem_thu"],
        template_dirs=[
            "01. Hồ sơ đạo đức đề cương - MẪU",
            "02. Hồ sơ khoa học đề cương - MẪU",
            "04. Hồ sơ nghiệm thu - MẪU",
        ],
        copy_head_cv=True,
        copy_expert_cvs=True,
    ),
    DossierPackage(
        id="D",
        aliases=["D", "BO_D", "Bộ D", "Bộ D - Chuyển giao công nghệ & Tư vấn"],
        name="Bộ D: Hồ sơ Tư vấn sản phẩm & Chuyển giao công nghệ",
        description="Bộ hồ sơ phục vụ dự án tư vấn công thức và chuyển giao công nghệ sản phẩm (Setup sẵn cho tương lai)",
        sections=["dao_duc", "khoa_hoc", "nghiem_thu"],
        template_dirs=[
            "01. Hồ sơ đạo đức đề cương - MẪU",
            "02. Hồ sơ khoa học đề cương - MẪU",
            "04. Hồ sơ nghiệm thu - MẪU",
        ],
        copy_head_cv=True,
        copy_expert_cvs=True,
    ),
]


BO_HO_SO_SHEET_NAME = "_BoHoSo"


def load_packages_from_sheet(wb) -> List[DossierPackage]:
    """Đọc cấu hình các bộ hồ sơ từ sheet _BoHoSo trong workbook Excel."""
    if wb is None or BO_HO_SO_SHEET_NAME not in wb.sheetnames:
        return []
    ws = wb[BO_HO_SO_SHEET_NAME]
    packages = []
    for r in range(2, ws.max_row + 1):
        pkg_id = ws.cell(row=r, column=1).value
        if not pkg_id:
            continue
        pkg_id = str(pkg_id).strip()
        name = str(ws.cell(row=r, column=2).value or "").strip()
        desc = str(ws.cell(row=r, column=3).value or "").strip()
        sections_str = str(ws.cell(row=r, column=4).value or "").strip()
        template_dirs_str = str(ws.cell(row=r, column=5).value or "").strip()
        sections = [s.strip() for s in sections_str.split(",") if s.strip()]
        template_dirs = [d.strip() for d in template_dirs_str.split(",") if d.strip()]
        packages.append(
            DossierPackage(
                id=pkg_id,
                name=name,
                description=desc,
                sections=sections,
                template_dirs=template_dirs,
                aliases=[pkg_id],
            )
        )
    return packages


def _load_packages_from_config() -> List[DossierPackage]:
    if not CONFIG_PACKAGES_PATH.exists():
        return _DEFAULT_PACKAGES
    try:
        with open(CONFIG_PACKAGES_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        packages = []
        for item in data:
            packages.append(
                DossierPackage(
                    id=item.get("id", ""),
                    name=item.get("name", ""),
                    description=item.get("description", ""),
                    sections=item.get("sections", []),
                    template_dirs=item.get("template_dirs", []),
                    copy_head_cv=item.get("copy_head_cv", True),
                    copy_expert_cvs=item.get("copy_expert_cvs", True),
                    aliases=item.get("aliases", []),
                )
            )
        return packages if packages else _DEFAULT_PACKAGES
    except Exception:
        return _DEFAULT_PACKAGES


_PACKAGE_REGISTRY: Dict[str, DossierPackage] = {}
_ALIAS_MAP: Dict[str, DossierPackage] = {}
_CUSTOM_PACKAGES: List[DossierPackage] = []


def _refresh_registry() -> None:
    global _PACKAGE_REGISTRY, _ALIAS_MAP
    _PACKAGE_REGISTRY.clear()
    _ALIAS_MAP.clear()
    packages = _load_packages_from_config() + _CUSTOM_PACKAGES
    for pkg in packages:
        _PACKAGE_REGISTRY[pkg.id] = pkg
        _ALIAS_MAP[pkg.id.lower()] = pkg
        for alias in pkg.aliases:
            _ALIAS_MAP[alias.lower()] = pkg


_refresh_registry()


def get_available_packages() -> List[DossierPackage]:
    """Trả về danh sách tất cả các bộ hồ sơ có sẵn trong hệ thống."""
    _refresh_registry()
    return list(_PACKAGE_REGISTRY.values())


def get_package(package_id: str) -> DossierPackage:
    """Lấy thông tin một bộ hồ sơ theo ID hoặc Alias (hỗ trợ A, full, B, C, D, v.v.)."""
    _refresh_registry()
    clean_id = package_id.strip().lower()
    pkg = _ALIAS_MAP.get(clean_id)
    if pkg is not None:
        return pkg

    # Thử tìm theo id trực tiếp
    pkg = _PACKAGE_REGISTRY.get(package_id)
    if pkg is not None:
        return pkg

    raise KeyError(
        f"Không tìm thấy bộ hồ sơ có mã '{package_id}'. "
        f"Các bộ hồ sơ hợp lệ: {list(_PACKAGE_REGISTRY.keys())}"
    )


def register_package(package: DossierPackage) -> None:
    """Đăng ký thêm một bộ hồ sơ mới vào hệ thống."""
    if package not in _CUSTOM_PACKAGES:
        _CUSTOM_PACKAGES.append(package)
    _PACKAGE_REGISTRY[package.id] = package
    _ALIAS_MAP[package.id.lower()] = package
    for alias in package.aliases:
        _ALIAS_MAP[alias.lower()] = package

