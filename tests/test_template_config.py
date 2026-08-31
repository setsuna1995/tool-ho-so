import pytest

import paths
import template_config


def _make_mau_folder(base, name, files):
    d = base / name
    d.mkdir()
    for filename in files:
        (d / filename).write_text("dummy", encoding="utf-8")
    return d


def test_derive_dest_path_strips_mau_suffix():
    dest = template_config.derive_dest_path("01. Hồ sơ đạo đức đề cương - MẪU/00. QĐ Giao đề tài.docx")
    assert dest == "01. Hồ sơ đạo đức đề cương/00. QĐ Giao đề tài.docx"


def test_derive_dest_path_supports_templates_prefix():
    dest = template_config.derive_dest_path("templates/01. Hồ sơ đạo đức đề cương - MẪU/00. QĐ Giao đề tài.docx")
    assert dest == "01. Hồ sơ đạo đức đề cương/00. QĐ Giao đề tài.docx"
    dest_win = template_config.derive_dest_path("templates\\01. Hồ sơ đạo đức đề cương - MẪU\\00. QĐ Giao đề tài.docx")
    assert dest_win == "01. Hồ sơ đạo đức đề cương/00. QĐ Giao đề tài.docx"


def test_derive_dest_path_raises_without_mau_suffix():
    with pytest.raises(ValueError):
        template_config.derive_dest_path("01. Hồ sơ đạo đức đề cương/00. QĐ Giao đề tài.docx")


def test_discover_copies_lists_docx_files_with_derived_dest(tmp_path):
    _make_mau_folder(tmp_path, "01. Hồ sơ - MẪU", ["a.docx", "b.docx"])

    copies = template_config.discover_copies(tmp_path)

    assert ("01. Hồ sơ - MẪU/a.docx", "01. Hồ sơ/a.docx") in copies
    assert ("01. Hồ sơ - MẪU/b.docx", "01. Hồ sơ/b.docx") in copies


def test_discover_copies_ignores_non_docx_and_lock_files(tmp_path):
    _make_mau_folder(tmp_path, "01. Hồ sơ - MẪU", ["a.docx", "a.doc", "notes.pptx", "~$a.docx"])

    copies = template_config.discover_copies(tmp_path)

    assert [src for src, _dst in copies] == ["01. Hồ sơ - MẪU/a.docx"]


def test_discover_copies_ignores_folders_without_mau_suffix(tmp_path):
    _make_mau_folder(tmp_path, "02. Không phải mẫu", ["a.docx"])

    copies = template_config.discover_copies(tmp_path)

    assert copies == []


def test_discover_doc_files_only_lists_doc_without_docx_twin(tmp_path):
    _make_mau_folder(tmp_path, "01. Hồ sơ - MẪU", ["a.doc", "a.docx", "b.doc"])

    doc_files = template_config.discover_doc_files(tmp_path)

    assert doc_files == ["01. Hồ sơ - MẪU/b.doc"]


def test_discover_doc_files_on_real_project_is_empty_after_cleanup():
    doc_files = template_config.discover_doc_files(paths.project_root())

    assert doc_files == []
