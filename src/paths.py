from pathlib import Path


def project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def templates_root() -> Path:
    tpl = project_root() / "templates"
    return tpl if tpl.exists() else project_root()


def template_dir(name: str) -> Path:
    d = templates_root() / name
    if d.exists():
        return d
    ref_d = project_root() / "references" / name
    if ref_d.exists():
        return ref_d
    return project_root() / name


def cv_dir() -> Path:
    tpl_cv = templates_root() / "Lý lịch khoa học"
    if tpl_cv.exists():
        return tpl_cv
    root_cv = project_root() / "Lý lịch khoa học"
    if root_cv.exists():
        return root_cv
    return tpl_cv



