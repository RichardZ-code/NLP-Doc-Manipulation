from docx import Document

from docx.text.paragraph import Paragraph
from docx.oxml.xmlchemy import OxmlElement


def insert_paragraph_after(paragraph, text = None, style = None):
    """
    Insert a new paragraph after the given paragraph
    """
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    if style is not None:
        new_para.style = style
    return new_para