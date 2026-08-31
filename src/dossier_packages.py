# dossier_packages.py
from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class DossierPackage:
    id: str
    name: str
    description: str
    sections: List[str]
    template_dirs: List[str] = field(default_factory=list)
    copy_head_cv: bool = True
    copy_expert_cvs: bool = True


# Danh sách các bộ hồ sơ chuẩn mặc định
_DEFAULT_PACKAGES = [
    DossierPackage(
        id="full",
        name="Trọn bộ đầy đủ (Tất cả các hồ sơ)",
        description="Bao gồm Đạo đức, Khoa học và Nghiệm thu (đã tích hợp sẵn Thư mời chuyên gia)",
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
        id="dao_duc",
        name="Bộ hồ sơ Đạo đức đề cương",
        description="Gồm hồ sơ Hội đồng Đạo đức (kèm Thư mời chuyên gia đạo đức & CV Chủ nhiệm)",
        sections=["dao_duc"],
        template_dirs=[
            "01. Hồ sơ đạo đức đề cương - MẪU",
        ],
        copy_head_cv=True,
        copy_expert_cvs=True,
    ),
    DossierPackage(
        id="khoa_hoc",
        name="Bộ hồ sơ Khoa học đề cương",
        description="Gồm hồ sơ Hội đồng Khoa học xét duyệt đề cương (kèm Thư mời chuyên gia khoa học)",
        sections=["khoa_hoc"],
        template_dirs=[
            "02. Hồ sơ khoa học đề cương - MẪU",
        ],
        copy_head_cv=False,
        copy_expert_cvs=True,
    ),
    DossierPackage(
        id="nghiem_thu",
        name="Bộ hồ sơ Nghiệm thu đề tài",
        description="Gồm hồ sơ Hội đồng Nghiệm thu kết quả (kèm Thư mời chuyên gia nghiệm thu)",
        sections=["nghiem_thu"],
        template_dirs=[
            "04. Hồ sơ nghiệm thu - MẪU",
        ],
        copy_head_cv=False,
        copy_expert_cvs=True,
    ),
]


_PACKAGE_REGISTRY: Dict[str, DossierPackage] = {pkg.id: pkg for pkg in _DEFAULT_PACKAGES}


def get_available_packages() -> List[DossierPackage]:
    """Trả về danh sách tất cả các bộ hồ sơ có sẵn trong hệ thống."""
    return list(_PACKAGE_REGISTRY.values())


def get_package(package_id: str) -> DossierPackage:
    """Lấy thông tin một bộ hồ sơ theo ID."""
    pkg = _PACKAGE_REGISTRY.get(package_id)
    if pkg is None:
        raise KeyError(
            f"Không tìm thấy bộ hồ sơ có mã '{package_id}'. "
            f"Các bộ hồ sơ hợp lệ: {list(_PACKAGE_REGISTRY.keys())}"
        )
    return pkg


def register_package(package: DossierPackage) -> None:
    """Đăng ký thêm một bộ hồ sơ mới vào hệ thống."""
    _PACKAGE_REGISTRY[package.id] = package
