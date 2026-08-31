# expert_invitation.py
import copy
from pathlib import Path
from typing import List

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from excel_reader import Person
from word_writer import _docx_replace_in_paragraph


def _wrap_element(element, doc):
    if element.tag == qn("w:p"):
        return Paragraph(element, doc)
    if element.tag == qn("w:tbl"):
        return Table(element, doc)
    return None


def _apply_tokens_to_elements(elements, doc, tokens: dict) -> None:
    for element in elements:
        wrapped = _wrap_element(element, doc)
        if wrapped is None:
            continue
        if isinstance(wrapped, Paragraph):
            for find, value in tokens.items():
                _docx_replace_in_paragraph(wrapped, find, value)
        else:
            for row in wrapped.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for find, value in tokens.items():
                            _docx_replace_in_paragraph(paragraph, find, value)


def _page_break_element():
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    return p


def generate_multi_page_letter(path: Path, recipients: List[Person], common_tokens: dict) -> bool:
    """Mo file .docx tai `path` (da duoc copy san tu thu muc MAU), nhan ban
    toan bo noi dung than file (tru sectPr, phai luon la phan tu cuoi cung
    theo chuan OOXML) thanh len(recipients) trang - moi trang danh cho 1
    nguoi, dien ca common_tokens lan token rieng trang
    ({{CHUYEN_GIA_HO_TEN}}, {{CHUYEN_GIA_DON_VI}}) - roi ghi de luu lai
    dung `path`.

    Tra ve False (khong dong/luu gi ca, giu nguyen file mau chua dien) neu
    `recipients` rong.
    """
    if not recipients:
        return False

    doc = docx.Document(str(path))
    body = doc.element.body

    sect_pr = body.find(qn("w:sectPr"))
    template_elements = [el for el in list(body) if el is not sect_pr]
    for element in template_elements:
        body.remove(element)

    def _insert(element):
        if sect_pr is not None:
            sect_pr.addprevious(element)
        else:
            body.append(element)

    for index, person in enumerate(recipients):
        if index > 0:
            _insert(_page_break_element())

        cloned_elements = []
        for element in template_elements:
            clone = copy.deepcopy(element)
            _insert(clone)
            cloned_elements.append(clone)

        page_tokens = {
            "{{CHUYEN_GIA_HO_TEN}}": f"{person.degree} {person.name}".strip(),
            "{{CHUYEN_GIA_DON_VI}}": person.org,
        }
        _apply_tokens_to_elements(cloned_elements, doc, {**common_tokens, **page_tokens})

    doc.save(str(path))
    return True
