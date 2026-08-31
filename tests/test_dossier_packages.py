import pytest
import dossier_packages as dp


def test_get_available_packages_contains_standard_packages():
    packages = dp.get_available_packages()
    ids = [p.id for p in packages]
    assert "full" in ids
    assert "dao_duc" in ids
    assert "khoa_hoc" in ids
    assert "nghiem_thu" in ids


def test_get_package_by_id():
    full_pkg = dp.get_package("full")
    assert full_pkg.id == "full"
    assert "dao_duc" in full_pkg.sections
    assert "khoa_hoc" in full_pkg.sections
    assert "nghiem_thu" in full_pkg.sections
    assert full_pkg.copy_head_cv is True

    dao_duc_pkg = dp.get_package("dao_duc")
    assert dao_duc_pkg.id == "dao_duc"
    assert "dao_duc" in dao_duc_pkg.sections
    assert "nghiem_thu" not in dao_duc_pkg.sections
    assert dao_duc_pkg.copy_head_cv is True

    khoa_hoc_pkg = dp.get_package("khoa_hoc")
    assert "khoa_hoc" in khoa_hoc_pkg.sections
    assert "dao_duc" not in khoa_hoc_pkg.sections


def test_get_package_raises_on_invalid_id():
    with pytest.raises(KeyError):
        dp.get_package("non_existent_package")


def test_register_package_allows_future_custom_packages():
    custom = dp.DossierPackage(
        id="lam_sang",
        name="Bộ hồ sơ Thử nghiệm lâm sàng đặc biệt",
        description="Chuyên biệt cho nghiên cứu lâm sàng pha 2/3",
        sections=["dao_duc", "khoa_hoc"],
        template_dirs=["01. Hồ sơ đạo đức đề cương - MẪU"],
        copy_head_cv=True,
        copy_expert_cvs=False,
    )
    dp.register_package(custom)

    retrieved = dp.get_package("lam_sang")
    assert retrieved.name == "Bộ hồ sơ Thử nghiệm lâm sàng đặc biệt"
