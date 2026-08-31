import pytest
import dossier_packages as dp


def test_get_available_packages_contains_standard_packages():
    packages = dp.get_available_packages()
    ids = [p.id for p in packages]
    assert "A" in ids
    assert "B" in ids
    assert "C" in ids
    assert "D" in ids
    assert "A_dao_duc" in ids


def test_get_package_by_id_and_aliases():
    # Package A (trọn bộ hiện tại)
    pkg_a = dp.get_package("A")
    assert pkg_a.id == "A"
    assert "dao_duc" in pkg_a.sections
    assert "khoa_hoc" in pkg_a.sections
    assert "nghiem_thu" in pkg_a.sections

    # Backward compatibility with 'full' alias
    pkg_full = dp.get_package("full")
    assert pkg_full.id == "A"

    # Alias 'Bộ A'
    pkg_bo_a = dp.get_package("Bộ A")
    assert pkg_bo_a.id == "A"

    # Package B (Thử nghiệm lâm sàng)
    pkg_b = dp.get_package("B")
    assert pkg_b.id == "B"
    assert "Thử nghiệm lâm sàng" in pkg_b.name

    # Package C & D
    assert dp.get_package("C").id == "C"
    assert dp.get_package("D").id == "D"

    # Sub-packages
    dao_duc_pkg = dp.get_package("dao_duc")
    assert dao_duc_pkg.id == "A_dao_duc"
    assert "dao_duc" in dao_duc_pkg.sections
    assert "nghiem_thu" not in dao_duc_pkg.sections


def test_get_package_raises_on_invalid_id():
    with pytest.raises(KeyError):
        dp.get_package("non_existent_package_xyz")


def test_register_package_allows_future_custom_packages():
    custom = dp.DossierPackage(
        id="custom_special",
        name="Bộ hồ sơ Tùy chỉnh",
        description="Dành cho nghiên cứu đặc biệt",
        sections=["dao_duc"],
        template_dirs=["01. Hồ sơ đạo đức đề cương - MẪU"],
        aliases=["dac_biet", "Bộ Đặc Biệt"],
    )
    dp.register_package(custom)

    retrieved = dp.get_package("dac_biet")
    assert retrieved.name == "Bộ hồ sơ Tùy chỉnh"

