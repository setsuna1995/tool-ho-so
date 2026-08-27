from pathlib import Path

import word_writer
from excel_reader import ProjectInfo


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, common_tokens: dict) -> None:
    doc = session.open(dest_dir / "Công văn mời chuyên gia.docx")

    session.fill_tokens(doc, common_tokens)

    session.replace_text(
        doc,
        "Suy dinh dưỡng ở trẻ em dưới 5 tuổi – đặc biệt là suy dinh dưỡng thấp còi vẫn là một vấn đề sức khỏe cộng đồng.",
        "[Bổ sung bối cảnh/lý do triển khai dự án tại đây]",
    )
    session.replace_text(
        doc,
        "Một trong những giải pháp làm giảm tình trạng suy dinh dưỡng ở trẻ dưới 5 tuổi là sử dụng các sản phẩm bổ sung dinh dưỡng trong hệ thống trường mầm non.",
        "",
    )
    session.replace_text(
        doc,
        "Nhằm đánh giá tình trạng suy dinh dưỡng ở trẻ dưới 5 tuổi và hiệu quả của sản phẩm bổ sung dinh dưỡng LOF KUN COLOSTRUM, Viện Y học ứng dụng Việt Nam tiến hành triển khai nghiên cứu",
        "Viện Y học ứng dụng Việt Nam tiến hành triển khai đề tài",
    )

    session.save_close(doc)
