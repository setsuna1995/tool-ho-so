import section_engine
import word_writer
from pathlib import Path
from excel_reader import ProjectInfo


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, common_tokens: dict) -> None:
    config = section_engine.load_section_config("dao_duc")
    section_engine.generate(session, dest_dir, info, common_tokens, config)
