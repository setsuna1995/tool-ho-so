from excel_reader import ProjectInfo


def build_common_tokens(info: ProjectInfo) -> dict:
    return dict(info.common_tokens)
