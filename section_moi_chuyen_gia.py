from pathlib import Path

import word_writer
from excel_reader import ProjectInfo, parse_timeline


def generate(session: word_writer.Session, dest_dir: Path, info: ProjectInfo, title_old: str) -> None:
    doc = session.open(dest_dir / "Công văn mời chuyên gia.docx")

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
    start, end = parse_timeline(info.timeline)
    session.replace_text(
        doc,
        "Nghiên cứu được triển khai trong 06 tháng, trong đó thời gian can thiệp là 04 tháng.",
        f"Thời gian thực hiện dự kiến: {start} đến {end}.",
    )
    session.replace_text(
        doc,
        "Thời gian: 9 giờ 00 – sáng thứ 7 ngày 07 tháng 12 năm 2024.",
        f"Thời gian: …… giờ ……, ngày …… tháng …… năm {info.year}.",
    )
    session.replace_text(doc, "2024", str(info.year))
    session.replace_text(doc, f"“{title_old}”.", f"“{info.title}”.")

    session.save_close(doc)
