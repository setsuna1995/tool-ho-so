from excel_reader import ProjectInfo
import token_rules


def build_common_tokens(info: ProjectInfo) -> dict:
    return dict(info.common_tokens)


def build_template_context(info: ProjectInfo) -> dict:
    return token_rules.build_template_context(info)

